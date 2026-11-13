@echo off
REM Daily 11:00 refresh. No password: runs while you are signed in (lock screen OK on many PCs).
set ROOT=%~dp0..
set SCRIPT=%~dp0refresh_all_scheduled.cmd
set TASK=RAIL dashboard refresh daily

schtasks /Delete /TN "%TASK%" /F 2>nul
schtasks /Delete /TN "RAIL dashboard refresh" /F 2>nul
schtasks /Delete /TN "RAIL dashboard refresh Thursday" /F 2>nul

schtasks /Create /TN "%TASK%" /TR "\"%SCRIPT%\"" /SC DAILY /ST 11:00 /F

echo.
echo Created: %TASK%
echo   Every day at 11:00 AM
echo   Script: %SCRIPT%
echo.
echo PC must be ON and you must be SIGNED IN (not logged out).
echo Close Excel before 11:00. Connect VPN for Jenkins.
echo Logs: %ROOT%\reports\
echo.
echo Test now: schtasks /Run /TN "%TASK%"
pause
