@echo off
REM New effort-tracker tab every Wednesday 00:00 (Wed–Wed report week).
set ROOT=%~dp0..
set SCRIPT=%~dp0triage_effort_weekly.cmd
set TASK=RAIL triage effort weekly tab

schtasks /Delete /TN "%TASK%" /F 2>nul

schtasks /Create /TN "%TASK%" /TR "\"%SCRIPT%\"" /SC WEEKLY /D WED /ST 00:00 /F

echo.
echo Created: %TASK%
echo   Every Wednesday at 12:00 AM (midnight, local)
echo   Script: %SCRIPT%
echo   Ensures Week tabs (Week 1..10+); active tab follows Wed-Wed calendar
echo.
echo PC must be ON and you must be SIGNED IN at midnight (or task runs when you next sign in).
echo Close RAIL-triage-effort.xlsx before the run if you want a clean save.
echo Log: %ROOT%\reports\triage_effort_weekly.log
echo.
echo Test now: schtasks /Run /TN "%TASK%"
pause
