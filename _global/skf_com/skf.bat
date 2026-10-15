@echo off
cd /d "%~dp0"
python skf.py %*
if errorlevel 1 pause
