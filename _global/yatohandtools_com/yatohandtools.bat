@echo off
cd /d "%~dp0"
python yatohandtools.py %*
if errorlevel 1 pause
