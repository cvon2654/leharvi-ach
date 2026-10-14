@echo off
rem Installs Solver into Excel: ribbon tab, macros, desktop shortcut. Close Excel first.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0excel\install.ps1"
