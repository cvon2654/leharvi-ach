@echo off
rem Solver - double-click to open, or drag an Excel file onto this icon.
set "HERE=%~dp0"
if exist "%HERE%Solver.exe" (
    start "" "%HERE%Solver.exe" %*
    exit /b
)
where pythonw >nul 2>nul
if %errorlevel%==0 (
    start "" pythonw "%HERE%run_solver.py" %*
    exit /b
)
where pyw >nul 2>nul
if %errorlevel%==0 (
    start "" pyw "%HERE%run_solver.py" %*
    exit /b
)
echo Solver needs Python (free, one-time install): https://www.python.org/downloads/
echo During install tick "Add python.exe to PATH". No internet is needed after that.
echo Or download Solver.exe (see README.md) and put it in this folder.
pause
