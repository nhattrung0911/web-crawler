@echo off
cd /d "%~dp0"
python totalpowertools.py %*
if errorlevel 1 pause
