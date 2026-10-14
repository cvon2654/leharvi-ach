@echo off
rem Builds Solver.exe (one file, runs offline, no Python needed on the PC that uses it).
rem Needs Python and internet once, on the PC that builds it.
cd /d "%~dp0"
python -m pip install vendor\openpyxl-3.1.5-py2.py3-none-any.whl vendor\et_xmlfile-2.0.0-py3-none-any.whl pyinstaller
python -m PyInstaller --noconfirm --onefile --windowed --name Solver run_solver.py
copy /y dist\Solver.exe Solver.exe
echo.
echo Done: Solver.exe is in this folder.
pause
