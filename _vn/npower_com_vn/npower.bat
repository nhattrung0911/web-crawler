@echo off
cd /d "%~dp0"
python npower.py %*
if errorlevel 1 pause
