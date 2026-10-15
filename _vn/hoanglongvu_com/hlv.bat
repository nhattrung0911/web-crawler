@echo off
cd /d "%~dp0"
python hlv.py %*
if errorlevel 1 pause
