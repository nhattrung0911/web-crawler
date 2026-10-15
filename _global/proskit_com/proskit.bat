@echo off
cd /d "%~dp0"
python proskit.py %*
if errorlevel 1 pause
