param([string]$Python = 'python')
$ErrorActionPreference = 'Stop'
Push-Location $PSScriptRoot
try {
    & $Python -m venv .build-env
    if ($LASTEXITCODE) { throw 'Python with Tcl/Tk and venv is required for building.' }
    $buildPython = Join-Path $PSScriptRoot '.build-env\Scripts\python.exe'
    & $buildPython -m pip install -r requirements.txt pyinstaller==6.22.3
    if ($LASTEXITCODE) { throw 'Build dependency installation failed.' }
    & $buildPython -m PyInstaller --noconfirm --clean --onefile --windowed --name BunnyHop --add-data 'config;config' --hidden-import xlrd bunny_hop_processor.py
    if ($LASTEXITCODE) { throw 'Executable build failed.' }
    Write-Output 'Built dist\BunnyHop.exe. Run the acceptance tests before distribution.'
} finally { Pop-Location }
