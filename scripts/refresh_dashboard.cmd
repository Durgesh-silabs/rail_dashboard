@echo off
cd /d C:\Projects\rail_dashboard
if not exist reports mkdir reports
".venv\Scripts\python.exe" refresh_dashboard.py >> reports\dashboard_refresh.log 2>&1
