"""Run: python TEST_EXECUTABLE.py path/to/BunnyHop.exe path/to/datasets"""
from pathlib import Path
from tempfile import TemporaryDirectory
import subprocess,shutil,json,os,sys
from openpyxl import load_workbook

exe=Path(sys.argv[1]).resolve()
data=Path(sys.argv[2]).resolve()
report=[]
with TemporaryDirectory(prefix='bunnyhop_exe_test_') as tmp:
    folder=Path(tmp)
    env=os.environ.copy()
    for name in ('PYTHONHOME','PYTHONPATH','TCL_LIBRARY','TK_LIBRARY'):
        env.pop(name,None)
    env['PATH']=str(Path(os.environ.get('SystemRoot','C:/Windows'))/'System32')
    for name,year,count in [('bunny-hop-2023_Participants_12112026_1038.csv','2023',129),('bunny-hop-2024_Participants_12112026_1038.csv','2024',122),('shapwick-bunny-hop-2026_Participants_Master List.xlsx','2026',152)]:
        source=folder/name
        shutil.copy2(data/name,source)
        run=subprocess.run([str(exe),'--check-entrants',str(source),'--year',year],env=env,cwd=folder,timeout=120)
        assert run.returncode==0,(name,run.returncode)
        book=load_workbook(folder/f'BunnyHop_{year}_Results_Entry.xlsx',read_only=True,data_only=True)
        assert book.active.max_row-1==count,name
        book.close()
        report.append({'file':name,'records':count,'result':'PASS'})
    for name,year,count in [('2024 BH results for testing.xlsx','2024',134),('Bunny_Hop_2025 Results Input.xlsx','2025',117),('shapwick-bunny-hop-2026_Results Input.xlsx','2026',132)]:
        output=folder/f'{year}_classified.xlsx'
        run=subprocess.run([str(exe),'--input',str(data/name),'--output',str(output),'--year',year],env=env,cwd=folder,timeout=120)
        assert run.returncode==0,(name,run.returncode)
        book=load_workbook(output,read_only=True,data_only=True)
        assert book['All Runners'].max_row-1==count,name
        book.close()
        report.append({'file':name,'finishers':count,'result':'PASS'})
print(json.dumps(report,indent=2))
