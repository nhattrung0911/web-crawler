@echo off
cd /d "%~dp0"
python digikey.py %*
if errorlevel 1 pause
