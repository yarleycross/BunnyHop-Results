"""Fictional race coverage, including the unchanged results processor."""
from concurrent.futures import ThreadPoolExecutor
from io import BytesIO
from pathlib import Path
import sqlite3
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

import bunny_hop_processor as results
from live_timing.server import HEADERS, create_app


START = 1_800_000_000_000
ROOT = Path(__file__).resolve().parents[1]
LAN_IP = "192.168.8.12"


class LiveTimingTests(unittest.TestCase):
    def setUp(self):
        folder = self.enterContext(TemporaryDirectory())
        self.database = Path(folder) / "race.sqlite3"
        self.app = create_app(self.database, LAN_IP)
        self.client = self.app.test_client()
        self.clock = self.enterContext(patch("live_timing.server.now_ms", return_value=START))

    def start(self):
        response = self.client.post("/api/start", json={"start_ms": START})
        self.assertEqual(response.status_code, 201)

    def capture(self, stream, number, bib=None, client=None):
        return (client or self.client).post(
            f"/api/capture/{stream}", json={"request_id": f"{number:032x}", "bib": bib},
        )

    def capture_at(self, number, elapsed_ms):
        self.clock.return_value = START + elapsed_ms
        return self.capture("finish", number)

    def correct(self, sequence, bib=None, elapsed_seconds=None, void=False):
        response = self.client.post(f"/api/correction/{sequence}", json={
            "bib": bib, "elapsed_seconds": elapsed_seconds,
            "void": void, "reason": "Checked fictional manual record",
        })
        self.assertEqual(response.status_code, 201, response.get_json())

    def export(self, bibs=("101", "102", "103", "104", "105"), *, populated=False):
        book = Workbook()
        sheet = book.active
        sheet.title = "Results Entry"
        sheet.append(HEADERS)
        for number, bib in enumerate(bibs, 1):
            sheet.append([bib, "1:01" if populated else None, f"Fictional {number}",
                          "Runner", "F", "SF", None, f"example-{number}", number])
        upload = BytesIO()
        book.save(upload)
        upload.seek(0)
        return self.client.post("/export", data={"template": (upload, "Results Entry.xlsx")})

    def test_independent_sequences_durable_ack_and_idempotent_retry(self):
        self.assertEqual(self.capture("finish", 1).status_code, 409)
        self.start()
        self.assertEqual(self.client.post("/api/start", json={}).status_code, 409)
        bib = self.capture("bib", 2, "00102")
        self.assertEqual(bib.get_json()["sequence"], 1)
        self.assertEqual(self.capture_at(1, 4100).status_code, 201)
        self.capture_at(3, 5230)
        # A separate connection to the same file sees the original events at acknowledgement.
        with sqlite3.connect(self.database) as db:
            self.assertEqual(db.execute("SELECT COUNT(*) FROM captures").fetchone()[0], 3)
        reopened = create_app(self.database, LAN_IP).test_client()
        duplicate = self.capture("finish", 1, client=reopened)
        self.assertEqual(duplicate.status_code, 200)
        self.assertEqual(duplicate.get_json(), {"sequence": 1, "recorded_ms": START + 4100, "bib": None})
        self.assertEqual(self.capture("bib", 2, "999", client=reopened).status_code, 409)
        rows = reopened.get("/api/race").get_json()["rows"]
        self.assertEqual([(r["sequence"], r["elapsed_seconds"], r["bib"]) for r in rows],
                         [(1, 4, "102"), (2, 5, None)])

    def test_concurrent_requests_get_unique_consecutive_positions(self):
        self.start()

        def post(number):
            return self.capture("finish", number, client=self.app.test_client()).get_json()["sequence"]

        with ThreadPoolExecutor(max_workers=8) as pool:
            sequences = list(pool.map(post, range(20)))
        self.assertEqual(sorted(sequences), list(range(1, 21)))

    def test_failed_database_write_does_not_acknowledge_or_consume_position(self):
        self.start()
        with sqlite3.connect(self.database) as db:
            db.execute("""CREATE TRIGGER deny_capture BEFORE INSERT ON captures
                          BEGIN SELECT RAISE(ABORT, 'fictional storage failure'); END""")
        response = self.capture("finish", 1)
        self.assertEqual(response.status_code, 503)
        with sqlite3.connect(self.database) as db:
            db.execute("DROP TRIGGER deny_capture")
        self.assertEqual(self.capture("finish", 1).get_json()["sequence"], 1)

    def test_reconciliation_keeps_originals_and_export_preserves_ties(self):
        self.start()
        self.capture_at(1, 2_705_100)
        self.capture_at(2, 2_705_900)
        self.capture_at(3, 2_706_100)
        self.capture("bib", 4, "999")
        self.capture("bib", 5, "101")
        self.assertIn("Sequence 3", self.export().get_json()["error"])
        self.correct(1, bib="103")
        self.correct(1, bib="102")
        self.correct(3, void=True)
        self.correct(4, bib="104", elapsed_seconds=2707)
        rows = self.client.get("/api/race").get_json()["rows"]
        self.assertEqual((rows[0]["original_bib"], rows[0]["bib"]), ("999", "102"))
        self.assertEqual(rows[2]["finish_ms"], START + 2_706_100)
        with sqlite3.connect(self.database) as db:
            self.assertEqual(db.execute("SELECT bib FROM corrections WHERE sequence=1 ORDER BY revision")
                             .fetchall(), [("103",), ("102",)])

        output = self.export()
        self.assertEqual(output.status_code, 200, output.get_json())
        path = self.database.with_name("reconciled.xlsx")
        path.write_bytes(output.data)
        workbook = load_workbook(path, read_only=True)
        try:
            exported = list(workbook["Results Entry"].values)
        finally:
            workbook.close()
        self.assertEqual(exported[0], HEADERS)
        self.assertEqual([row[:2] for row in exported[1:]],
                         [("102", "45:05"), ("101", "45:05"), ("104", "45:07"), ("103", None), ("105", None)])
        classified = results.transform(path, ROOT / "config")
        results.verify_calculations(classified)
        self.assertEqual([(runner.bib, runner.seconds) for runner in classified[1]],
                         [("102", 2705), ("101", 2705), ("104", 2707)])

    def test_export_rejects_unresolved_conflicts(self):
        self.start()
        self.correct(1, bib="102", elapsed_seconds=5)
        self.correct(2, bib="102", elapsed_seconds=6)
        self.assertIn("more than one finish", self.export().get_json()["error"])
        self.correct(2, bib="101", elapsed_seconds=4)
        self.assertIn("earlier", self.export().get_json()["error"])
        self.correct(2, bib="100", elapsed_seconds=6)
        self.assertIn("missing", self.export().get_json()["error"])
        self.assertIn("blank", self.export(populated=True).get_json()["error"])

    def test_private_phone_routes_and_loopback_only_hq(self):
        self.start()
        phone = self.app.test_client()
        options = {
            "base_url": f"http://{LAN_IP}:8765",
            "environ_overrides": {"REMOTE_ADDR": "192.168.8.25"},
        }
        self.assertIn(b"FINISH", phone.get("/phone/finish", **options).data)
        self.assertIn(b'inputmode="numeric"', phone.get("/phone/bib", **options).data)
        self.assertEqual(phone.post("/api/capture/bib", json={
            "request_id": "a" * 32, "bib": "7",
        }, **options).status_code, 201)
        for endpoint in ("/", "/api/race"):
            self.assertEqual(phone.get(endpoint, **options).status_code, 403)
        for endpoint in ("/api/start", "/api/correction/1", "/export"):
            self.assertEqual(phone.post(endpoint, **options).status_code, 403)
        self.assertEqual(phone.get("/phone/finish", base_url="http://untrusted.example").status_code, 400)
        self.assertEqual(phone.get("/phone/finish", base_url=f"http://{LAN_IP}:8765",
                                   environ_overrides={"REMOTE_ADDR": "203.0.113.5"}).status_code, 403)
