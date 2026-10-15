@echo off
cd /d "%~dp0"
python namtoanviet.py %*
if errorlevel 1 pause
