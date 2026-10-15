@echo off
cd /d "%~dp0"
python vongbicongnghiep.py %*
if errorlevel 1 pause
