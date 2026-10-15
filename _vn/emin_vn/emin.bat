@echo off
cd /d "%~dp0"
python emin.py %*
if errorlevel 1 pause
