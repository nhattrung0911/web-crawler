@echo off
cd /d "%~dp0"
python toptul.py %*
if errorlevel 1 pause
