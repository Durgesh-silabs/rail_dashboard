@echo off
set ROOT=%~dp0..
cd /d "%ROOT%"
if not exist reports mkdir reports
echo.
echo === RAIL dashboard refresh ===
echo This usually takes 5-12 minutes (Jenkins + IoT + both Excel files).
echo Use VPN if Jenkins is internal. Close Excel first.
echo.
echo [1/2] Daily workbook...
".venv\Scripts\python.exe" refresh_dashboard.py
if errorlevel 1 (
    echo Daily refresh failed.
    pause
    exit /b 1
)
echo.
echo [2/2] Weekly workbook...
".venv\Scripts\python.exe" refresh_dashboard.py --weekly
if errorlevel 1 (
    echo Weekly refresh failed.
    pause
    exit /b 1
)
echo.
echo Done. Check OneDrive: RAIL-dashboard\RAIL-dashboard.xlsx
pause
