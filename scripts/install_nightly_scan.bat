@echo off
REM Creates (or updates) the Windows Task Scheduler task that runs the nightly scan.
REM The time comes from the first argument (default 01:00), e.g.:  install_nightly_scan.bat 01:00
REM The task runs every night; nightly_scan.bat itself skips the nights not listed in SETTINGS.json.
setlocal
set T=%~1
if "%T%"=="" set T=01:00
schtasks /Create /F /TN "Job Search Copilot nightly scan" /SC DAILY /ST %T% /TR "\"%~dp0nightly_scan.bat\""
if errorlevel 1 (
  echo Could not create the scheduled task. Try running this file again as administrator.
  exit /b 1
)
echo The nightly scan is scheduled every night at %T%.
echo To let it wake the computer: Task Scheduler, find "Job Search Copilot nightly scan", Properties,
echo Conditions tab, tick "Wake the computer to run this task".
endlocal
