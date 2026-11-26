# Bunny Hop Live Timing Prototype v0.1

A companion for finish-line timing, bib capture and reconciliation. The existing Bunny Hop 1.4.3 processor still handles entrant checks, classification and awards. Keep the established independent manual finish-line backup running as usual.

## Prepare the Windows HQ laptop

Install 64-bit Python 3.12. From the repository root in PowerShell:

```powershell
py -3.12 -m venv live_timing\.venv
.\live_timing\.venv\Scripts\python.exe -m pip install -r live_timing\requirements.txt
```

Prepare the dependencies before race day. For an offline install, run `py -3.12 -m pip download --only-binary=:all: -r live_timing\requirements.txt -d wheelhouse` on an internet-connected Windows computer with the same Python version and architecture. Copy the wheelhouse to the HQ laptop and add `--no-index --find-links .\wheelhouse` to the install command. Keep that directory outside the checkout. During use, the laptop and phones only need the local Wi-Fi.

Use the existing processor's **Check Entrants** workflow to generate the **Results Entry** workbook. Resolve entrant and bib issues there and set aside a copy with all Time cells blank. Keep real entrant workbooks and race databases on the laptop, outside Git. The workbook is selected only from HQ for export and is processed in memory.

## Connect and record

1. Connect the HQ laptop and both phones to the same trusted private Wi-Fi. On the laptop, run `ipconfig` and find the **IPv4 Address** under **Wireless LAN adapter Wi-Fi**. Open PowerShell in the repository root. Replace the example IPv4 address and choose a new database filename for each race or rehearsal.

   ```powershell
   .\live_timing\.venv\Scripts\python.exe .\live_timing\server.py --lan-ip 192.168.1.23 --db .\live_timing\data\race-day.sqlite3
   ```

2. Open **http://127.0.0.1:8765/** on the laptop. On phone 1 open the finish URL printed by the server, for example `http://192.168.1.23:8765/phone/finish`. On phone 2 open the bib URL, for example `http://192.168.1.23:8765/phone/bib`.
3. If Windows Firewall blocks the phones, allow inbound TCP 8765 on the **Private** profile only, limited to the local subnet and the laptop's Wi-Fi address. An administrator can add the rule below with the actual Wi-Fi address. Guest Wi-Fi or client isolation can also prevent the phones reaching the laptop. Use the trusted event Wi-Fi and avoid port forwarding.

   ```powershell
   New-NetFirewallRule -DisplayName "Bunny Hop Live Timing" -Direction Inbound -Action Allow -Protocol TCP -LocalPort 8765 -LocalAddress 192.168.1.23 -RemoteAddress LocalSubnet -Profile Private
   ```

4. Rehearse with fictional bibs on a separate database. Before race capture, wait for HQ to set **Set start now** at the official start signal, or use **Set earlier start** with the date and time recorded on the laptop clock. The start is saved once per database. Check the laptop's clock before the race, keep it awake, and leave the server running.
5. On phone 1, tap **FINISH** once per finisher. On phone 2, enter each bib in finish order and tap **Save bib**. Check the **Saved** sequence feedback each time. A pending save can be retried using the same submission ID. If it never reached the laptop, retry records the server time of the retry. Tell HQ about the delay so they can reconcile it against the manual backup. Resolve any unconfirmed save before moving on; use the manual record during a network outage or a burst the phones cannot keep up with.
6. HQ sees the two independent streams paired by sequence. This pairing is provisional. An omitted or extra tap can shift every later pair. Compare against the manual records even when the counts match.

The server listens on loopback and the selected Wi-Fi address. The phone pages accept submissions from the private LAN, so use a Wi-Fi network restricted to trusted devices. Only the laptop's loopback address can access HQ, the race table, corrections and export. Use the same database path if restarting during an event. Keep the file after the event for the raw records and correction history. For a new race, start with a new file and verify both phone tabs have no unconfirmed submission from the previous race.

## Reconcile and export

1. Have both phone operators stop recording and resolve any unconfirmed submissions. In HQ, compare the captured columns against the manual finish-line record.
2. Select **Edit** to override a bib or elapsed time (whole seconds), or void an extra record. Blank override fields use the captured values. For shifted streams, correct the affected sequence rows in order using the manual record. Use **Add missing finish** when a reconciled finish is needed after the last sequence. Every correction records a reason and a new SQLite revision. The captured rows and earlier revisions stay intact.
3. Choose the blank **Results Entry** workbook and select **Download reconciled workbook**. Resolve any reported missing, duplicate or unknown bibs, incomplete pairs or decreasing elapsed times and try again. Finish times are written as text such as `45:05`, using whole elapsed seconds truncated from the server timestamp. Finishers appear in reconciled sequence order, preserving order when the processor sees equal times. Other entrants retain blank times.
4. Open the downloaded workbook in the existing processor and run **Process Results**. Review the usual data checks and confirm the results before official publication.

## Test

From the repository root:

```powershell
.\live_timing\.venv\Scripts\python.exe -m unittest -v live_timing.test_live_timing
```

The prototype uses HTTP on trusted local Wi-Fi, one race per SQLite database, and the laptop clock. Capture requires a working connection to the laptop. Before a live event, rehearse using the actual Windows laptop, network, both phones and the manual backup.
