@echo off
cd /d "%~dp0"
python htgoods.py %*
if errorlevel 1 pause
