@echo off
cd /d "%~dp0"
python sieuthicongcu.py %*
if errorlevel 1 pause
