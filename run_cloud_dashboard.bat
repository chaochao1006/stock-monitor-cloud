@echo off
cd /d "%~dp0"
set REPORT_BASE_DIR=%~dp0data
"D:\software\conda\python.exe" -m streamlit run app.py
pause
