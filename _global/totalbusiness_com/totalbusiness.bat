@echo off
cd /d "%~dp0"
python totalbusiness.py %*
if errorlevel 1 pause
