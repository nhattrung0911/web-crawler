@echo off
cd /d "%~dp0"
streamlit run app.py %*
if errorlevel 1 pause
