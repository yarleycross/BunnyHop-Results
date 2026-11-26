"""Non-destructive, fictional-data smoke test for the Bunny Hop program.

Runs inside a TemporaryDirectory, does not open or modify real race files.
Requires the same libraries and Python environment as the application.
"""
from datetime import date
import csv
from hashlib import sha256
from pathlib import Path
import subprocess
import sys
from tempfile import TemporaryDirectory

from openpyxl import Workbook, load_workbook
from bunny_hop_processor import check_entrants, easter_monday, parse_seconds


def test():
    assert easter_monday(2023) == date(2023, 4, 10)
    assert easter_monday(2024) == date(2024, 4, 1)
    assert easter_monday(2026) == date(2026, 4, 6)
    assert easter_monday(2027) == date(2027, 3, 29)
    assert parse_seconds(45.5) == 45 * 60 + 50
    assert parse_seconds(45.1) == 45 * 60 + 10
    assert parse_seconds('45.05') == 45 * 60 + 5
    assert parse_seconds('45.3') == 45 * 60 + 30

    with TemporaryDirectory(prefix='bunnyhop_install_test_') as td:
        folder = Path(td)
        fos = folder / 'fictional_FOS_participants_2027.xlsx'
        workbook = Workbook()
        ws = workbook.active
        ws.title = 'FOS Participants'
        ws.append(['Firstname', 'Lastname', 'Gender', 'DOB', 'AgeOnDay',
                   'Age on Race Day', 'ClubName', 'ep_id', 'RaceNumber',
                   'Email', 'Emergency Contact'])
        people = []
        for i in range(12):
            sex = 'M' if i < 6 else 'F'
            first = f'Fictional{i + 1}'
            last = 'MaleRunner' if sex == 'M' else 'FemaleRunner'
            dob = date(1980, 1, 2) if sex == 'M' else date(1972, 2, 2)
            age = 47 if sex == 'M' else 55
            club = 'Wells City Harriers' if i in (0, 1, 6, 7) else ''
            ws.append([first, last, sex, dob, age, age, club, 9001 + i,
                       None, 'private@example.test', 'PRIVATE CONTACT'])
            people.append((first, last, sex, club, 9001 + i))

        # FOS has NO race numbers; these are manually allocated in the Master list.
        master = workbook.create_sheet('Master list')
        master.append(['ClubName', 'RaceNumber', 'Lastname', 'Firstname',
                       'Gender', 'AgeCategory'])
        allocated = [521, 506, 540, 501, 530, 519, 511, 533, 503, 525, 507, 515]
        shuffle = [4, 10, 1, 8, 2, 11, 3, 6, 0, 9, 5, 7]
        for j in shuffle:
            first, last, sex, club, _id = people[j]
            master.append([club, allocated[j], last, first, sex,
                           'MV40' if sex == 'M' else 'FV50'])
        workbook.save(fos)
        original_hash = sha256(fos.read_bytes()).hexdigest()

        # Exercise raw FOS CSV separately: it has NO allocated race numbers,
        # no Excel worksheets, and may include private data not for results output.
        csv_file = folder / 'fictional_FOS_participants_2027.csv'
        with csv_file.open('w', encoding='utf-8-sig', newline='') as fp:
            writer = csv.writer(fp)
            writer.writerow(['Firstname','Lastname','Gender','DOB','AgeOnDay',
                             'Age on Race Day','ClubName','ep_id','RaceNumber',
                             'Email','Emergency Contact'])
            for first,last,sex,club,eid in people:
                writer.writerow([first,last,sex, '1980-01-02' if sex=='M' else '1972-02-02',
                                 47 if sex=='M' else 55, 47 if sex=='M' else 55,
                                 club,eid,'','private@example.test','PRIVATE CONTACT'])
        csv_hash = sha256(csv_file.read_bytes()).hexdigest()
        csv_report,csv_template,csv_summary=check_entrants(csv_file, '2027')
        assert csv_summary['records']==12, csv_summary
        assert csv_summary['race_numbers_allocated']==0, csv_summary
        assert csv_summary['errors']==0, csv_summary
        assert sha256(csv_file.read_bytes()).hexdigest()==csv_hash
        csv_book=load_workbook(csv_template,read_only=True,data_only=True)
        try:
            assert csv_book.active.max_row==13
            assert all(r[0] is None for r in list(csv_book.active.values)[1:])
        finally:
            csv_book.close()


        report, template, summary = check_entrants(fos, '2027')
        assert summary['records'] == 12, summary
        assert summary['race_numbers_allocated'] == 12, summary
        assert summary['errors'] == 0, summary
        assert report.exists() and template.exists()
        assert sha256(fos.read_bytes()).hexdigest() == original_hash, 'Original FOS workbook changed!'

        source = load_workbook(template)
        data = source.active
        assert data.max_row == 13
        columns = [c.value for c in data[1]]
        assert columns[:7] == ['RaceNumber', 'Time', 'Firstname', 'Lastname',
                               'Gender', 'AgeCategory', 'ClubName']
        assert {data.cell(n, 1).value for n in range(2, 14)} == set(map(str, allocated))
        # The times use the Bunny Hop MM.SS convention and deliberately differ
        # from the deliberately wrong RacePosition column added below.
        values = [44.59, 45.05, 45.1, 45.3, 45.5, 45.59,
                  46.0, 46.01, 46.02, 46.03, 46.04, 46.05]
        data.cell(1, 10).value = 'RacePosition'
        for n in range(2, 14):
            data.cell(n, 2).value = values[n - 2]
            data.cell(n, 10).value = 100 - n
        source.save(template)
        source.close()

        output = folder / 'classified_results.xlsx'
        cmd = [sys.executable, str(Path(__file__).parent / 'bunny_hop_processor.py'),
               '--input', str(template), '--year', '2027', '--output', str(output)]
        run = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
        if run.returncode:
            raise RuntimeError(f'Processor CLI failed:\n{run.stdout}\n{run.stderr}')
        assert output.is_file(), 'Results workbook not created'

        book = load_workbook(output, read_only=True, data_only=True)
        try:
            expected = {'Start Here', 'Overall Awards', 'Age Group Awards',
                        'All Runners', 'Age Groups Full', 'Overall Men',
                        'Overall Women', 'Club Results Full', 'Data Checks'}
            assert expected.issubset(set(book.sheetnames)), book.sheetnames
            results = list(book['All Runners'].values)[1:]
            assert len(results) == 12
            assert [r[0] for r in results] == list(range(1, 13))
            assert [r[7] for r in results[:5]] == ['44:59', '45:05',
                                                    '45:10', '45:30', '45:50']
            overall = list(book['Overall Awards'].values)[2:]
            ages = list(book['Age Group Awards'].values)[2:]
            assert len(overall) == 6 and len(ages) == 6
            overall_bibs = {str(r[3]) for r in overall}
            age_bibs = {str(r[4]) for r in ages}
            assert not overall_bibs.intersection(age_bibs)
            assert {r[0] for r in ages} == {'MV40', 'FV50'}
            assert ('Calculation self-checks', 'PASS') in list(book['Start Here'].values)
            assert not any(r[0] == 'ERROR' for r in list(book['Data Checks'].values)[2:])
            assert not any('private@example.test' in str(r) or 'PRIVATE CONTACT' in str(r)
                           for ws in book.worksheets for r in ws.values), 'Private entrant data leaked'
        finally:
            book.close()

        if sys.platform == 'win32':
            # In addition to importability, test that a usable Tk desktop window
            # can be created and dismissed on this Windows installation.
            import tkinter as tk
            root = tk.Tk()
            root.withdraw()
            root.update_idletasks()
            root.destroy()

    print('INSTALLATION TEST PASSED: raw FOS CSV and Excel Master, manual race numbers,')
    print('Easter ages, MM.SS time ordering, rankings, awards, privacy and GUI runtime.')


if __name__ == '__main__':
    test()
