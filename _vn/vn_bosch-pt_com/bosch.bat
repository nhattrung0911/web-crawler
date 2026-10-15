@echo off
cd /d "%~dp0"
python bosch.py %*
if errorlevel 1 pause
