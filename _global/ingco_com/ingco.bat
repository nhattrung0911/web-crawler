@echo off
cd /d "%~dp0"
python ingco.py %*
if errorlevel 1 pause
