@echo off
REM Open the shared OneDrive workbook (same file for the whole team).
set ROOT=%~dp0..
cd /d "%ROOT%"
for /f "usebackq delims=" %%P in (`".venv\Scripts\python.exe" -c "import yaml; from pathlib import Path; print(yaml.safe_load((Path('triage_tracker/config.yaml')).read_text(encoding='utf-8'))['workbook'])"`) do set "WB=%%P"
if not exist "%WB%" (
  echo Workbook not found: %WB%
  echo Run: .venv\Scripts\python.exe triage_tracker\create_weekly_sheet.py
  pause
  exit /b 1
)
start "" "%WB%"
echo Opened shared workbook. In Excel turn ON AutoSave for live updates for everyone.
exit /b 0
