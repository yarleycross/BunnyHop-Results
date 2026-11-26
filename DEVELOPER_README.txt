Bunny Hop 1.4.3 source

End users run the standalone EXE; this folder is for maintenance only.
Use Windows x64 Python 3.12 with Tcl/Tk. BUILD_WINDOWS.ps1 creates an isolated
build environment and builds a one-file windowed EXE. It downloads packages
from PyPI. No GitHub workflow is used.

Run SELF_TEST.py for the fictional end-to-end smoke test. Run
  python REGRESSION_TEST.py "path-to-copied-datasets"
for the actual-data regression tests. The test fixture files are retained in
the project work/datasets folder, not bundled into this source archive.
baseline_checksums.json pins the original fixture hashes and 1.4.2 results.
test-report.json describes the last run; it is not an independent certificate.

Build used Python 3.12.14, openpyxl 3.1.5, xlrd 2.0.2 and PyInstaller 6.22.3.
The Codex bundled Python was missing usable Tcl/Tk resources. For this build,
official Python 3.12.10 Tcl/Tk binaries/resources were extracted locally and
used with a copied build runtime; no end-user Python installation was added.
The project work folder retains build_app.py, BunnyHop.spec and build logs
for that environment-specific build. A complete ordinary Python installation
does not need that workaround. Source builds need their own Windows acceptance.

Review AUDIT_AND_TESTS.md and USER_GUIDE.txt supplied with the application.
The old Python-installing BAT launchers were intentionally not repackaged.
