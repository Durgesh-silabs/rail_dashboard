@echo off
set ROOT=%~dp0..
cd /d "%ROOT%"
if not exist reports mkdir reports
".venv\Scripts\python.exe" refresh_dashboard.py >> reports\dashboard_refresh.log 2>&1
