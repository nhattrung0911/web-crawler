@echo off
cd /d "%~dp0"
python tst.py %*
if errorlevel 1 pause
