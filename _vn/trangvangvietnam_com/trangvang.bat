@echo off
cd /d "%~dp0"
python trangvang.py %*
if errorlevel 1 pause
