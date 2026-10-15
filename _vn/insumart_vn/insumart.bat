@echo off
cd /d "%~dp0"
python insumart.py %*
if errorlevel 1 pause
