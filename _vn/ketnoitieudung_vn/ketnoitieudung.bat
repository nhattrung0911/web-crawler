@echo off
cd /d "%~dp0"
python ketnoitieudung.py %*
if errorlevel 1 pause
