@echo off
cd /d "%~dp0"
python se.py %*
if errorlevel 1 pause
