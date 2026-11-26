"""Run with Python REGRESSION_TEST.py DATASET_FOLDER. Never edits fixtures."""
import csv, hashlib, json, sys, unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from datetime import date, time
from openpyxl import Workbook, load_workbook
import bunny_hop_processor as bh

DATA=Path(sys.argv.pop(1)) if len(sys.argv)>1 else Path('../datasets')
CFG=Path(__file__).parent/'config'
REPORT=[]

class Regressions(unittest.TestCase):
    def require_historical_files(self, filenames):
        missing = [name for name in filenames if not (DATA / name).is_file()]
        if missing:
            self.fail(f'Missing historical datasets in {DATA.resolve()}: {", ".join(missing)}')

    def test_preserved_historical_classifications(self):
        baseline=json.loads(Path(__file__).with_name('baseline_checksums.json').read_text())
        self.require_historical_files(baseline['expected'])
        for name,expected in baseline['expected'].items():
            path=DATA/name
            self.assertEqual(hashlib.sha256(path.read_bytes()).hexdigest(),expected['input_sha256'])
            result=bh.transform(path,CFG)
            payload=[[getattr(r,f) for f in baseline['fields']] for r in result[1]]
            self.assertEqual(hashlib.sha256(json.dumps(payload,ensure_ascii=True).encode()).hexdigest(),expected['classification_sha256'])

    def test_time_formats(self):
        for raw,expected in [(45.5,2750),(45.1,2710),('45.05',2705),('1:05:03',3903),(time(0,45,5),2705),(2705/86400,2705),('45.60',None),(-1,None),('DNF',None)]:
            self.assertEqual(bh.parse_seconds(raw),expected)

    def test_boundaries(self):
        cats=bh.DEFAULT_CATEGORIES
        for age,sex,cat in [(39,'M','SM'),(40,'M','MV40'),(49,'F','FV40'),(50,'F','FV50'),(60,'F','FV60'),(70,'F','FV70'),(80,'F','FV70'),(80,'M','MV80')]:
            self.assertEqual(bh.age_category_from_age(age,sex,cats),cat)
        self.assertEqual(bh.age_on_date(date(1986,4,7),date(2026,4,6)),39)

    def test_senior_mismatch_and_prizes(self):
        with TemporaryDirectory() as td:
            p=Path(td)/'input.xlsx';w=Workbook();s=w.active
            s.append(list(bh.MANDATORY)+['ClubName','RacePosition'])
            for n in range(1,9):
                s.append([n,40+n,'Runner',str(n),'F','SM' if n==4 else 'SF','Wells City Harriers' if n%2 else 'Westbury Harriers',100-n])
            w.save(p)
            result=bh.transform(p,CFG);bh.verify_calculations(result)
            self.assertTrue(any(i[1]=='Category/sex mismatch' for i in result[4]))
            self.assertEqual([r.bib for _,_,r in result[2]],['1','2','3'])
            self.assertEqual([r.bib for _,_,r in result[3]],['5','6','7'])
            self.assertEqual(len({r.club for r in result[1]}),2)

    def test_real_csvs(self):
        files = sorted(p for p in DATA.glob('*.csv') if p.is_file())
        if not files:
            self.fail(f'Missing historical datasets: no .csv files found in {DATA.resolve()}')
        for p in files:
            year=next(str(y) for y in range(2023,2027) if str(y) in p.name)
            before=hashlib.sha256(p.read_bytes()).hexdigest()
            entrants,issues,*_=bh.analyse_participants(p,bh.easter_monday(int(year)),CFG)
            with p.open(encoding='utf-8-sig',newline='') as fp:
                rows=list(csv.DictReader(fp))
            self.assertEqual(len(entrants),len(rows))
            self.assertEqual(hashlib.sha256(p.read_bytes()).hexdigest(),before)
            REPORT.append({'file':p.name,'records':len(entrants),'errors':sum(i[0]=='ERROR' for i in issues),'reviews':sum(i[0]=='REVIEW' for i in issues)})

    def test_real_results(self):
        cases=[('2024 BH results for testing.xlsx',134),('BH results for testing.xlsx',117),('Bunny_Hop_2025 Results Input.xlsx',117),('2026 race results for testing.xlsx',132),('shapwick-bunny-hop-2026_Results Input.xlsx',132),('BH_results_updated.xlsx',117)]
        self.require_historical_files(name for name, _ in cases)
        for name,count in cases:
            p=DATA/name
            before=hashlib.sha256(p.read_bytes()).hexdigest()
            r=bh.transform(p,CFG);bh.verify_calculations(r)
            self.assertEqual(len(r[1]),count,name)
            # Independent ordered bib/time reference from unprocessed cells.
            rows,_=bh.read_input(p)
            expected=sorted([(bh.parse_seconds(d['Time']),row,bh.normalise_bib(d['RaceNumber'])) for row,d in rows if bh.parse_seconds(d['Time']) is not None])
            self.assertEqual([(x.seconds,x.row,x.bib) for x in r[1]],expected)
            self.assertEqual(len(r[2]),6)
            if '2024' not in name:self.assertEqual(len(r[3]),30)
            self.assertFalse({x.bib for _,_,x in r[2]} & {x.bib for _,_,x in r[3]})
            with TemporaryDirectory() as td:
                out,stats=bh.process(p,Path(td)/'results.xlsx','2026',CFG)
                w=load_workbook(out,data_only=True)
                self.assertEqual(w['All Runners'].max_row-1,count)
                self.assertEqual([row[1] for row in list(w['All Runners'].values)[1:]],[x.bib for x in r[1]])
                for title in ('Shapwick','Street Striders'):
                    rows=list(w[title].iter_rows(min_row=3))
                    for sex in ('M','F'):
                        candidates=[row for row in rows if row[6].value==sex]
                        if candidates:self.assertEqual(candidates[0][0].fill.fgColor.rgb,'00FFFF00')
                w.close()
            self.assertEqual(hashlib.sha256(p.read_bytes()).hexdigest(),before)
            REPORT.append({'file':name,**stats})

    def test_real_master_and_standalone(self):
        name = 'shapwick-bunny-hop-2026_Participants_Master List.xlsx'
        self.require_historical_files([name])
        p=DATA/name
        es,issues,*_=bh.analyse_participants(p,date(2026,4,6),CFG)
        self.assertEqual(len(es),152);self.assertEqual(sum(bool(e['bib']) for e in es),152)
        w=load_workbook(p,read_only=True,data_only=True)
        master=list(w['Master list'].values);w.close()
        expected={bh.normalise_bib(row[0]):row[4] for row in master[1:]}
        self.assertEqual({e['bib']:e['category'] for e in es},expected)
        with TemporaryDirectory() as td:
            out=Path(td)/'Master.xlsx';w=Workbook()
            for row in master:w.active.append(row)
            # A deliberate manual correction must survive without DOB fields.
            w.active.cell(2,5,'MV80' if w.active.cell(2,4).value=='M' else 'FV70')
            w.save(out)
            report,entry,stats=bh.check_entrants(out,'2026',config_dir=CFG)
            self.assertEqual(stats['records'],152);self.assertEqual(stats['errors'],0)
            e=load_workbook(entry,data_only=True)
            self.assertEqual(e.active.cell(2,6).value,w.active.cell(2,5).value)
            e.close()
        REPORT.append({'file':p.name,'records':len(es),'errors':sum(i[0]=='ERROR' for i in issues),'reviews':sum(i[0]=='REVIEW' for i in issues)})

if __name__=='__main__':
    result=unittest.main(exit=False)
    Path(__file__).with_name('test-report.json').write_text(json.dumps(REPORT,indent=2),encoding='utf-8')
    sys.exit(not result.result.wasSuccessful())
