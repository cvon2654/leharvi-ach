@echo off
rem Removes Solver from Excel and this PC.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0excel\uninstall.ps1"
