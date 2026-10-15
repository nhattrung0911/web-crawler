@echo off
cd /d "%~dp0"
python dungcuvang.py %*
if errorlevel 1 pause
