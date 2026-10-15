@echo off
cd /d "%~dp0"
python bosi.py %*
if errorlevel 1 pause
