@echo off
set ROOT=%~dp0..
cd /d "%ROOT%"
if not exist reports mkdir reports
".venv\Scripts\python.exe" triage_tracker\create_weekly_sheet.py >> reports\triage_effort_weekly.log 2>&1
