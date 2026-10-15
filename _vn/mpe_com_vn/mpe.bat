@echo off
cd /d "%~dp0"
python mpe.py %*
if errorlevel 1 pause
