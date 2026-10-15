@echo off
cd /d "%~dp0"
python milwaukee.py %*
if errorlevel 1 pause
