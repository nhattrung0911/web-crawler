@echo off
cd /d "%~dp0"
python wadfow.py %*
if errorlevel 1 pause
