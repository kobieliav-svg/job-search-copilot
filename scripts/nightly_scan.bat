@echo off
REM Nightly job scan, started by Windows Task Scheduler (install_nightly_scan.bat sets that up).
REM Runs scripts\run_scan.py (every source enabled in profile\SETTINGS.json, the title filter and the
REM description step) and appends everything to output\logs\nightly_scan.log. The morning agent reads
REM output\logs\scan_status.json, scores the new jobs and refreshes the dashboard.
REM Relative paths only: this file works wherever the project folder is.
setlocal
set PYTHONIOENCODING=utf-8
set PYTHONUTF8=1
cd /d "%~dp0.."
if not exist "output\logs" mkdir "output\logs"

REM Only on the nights listed in SETTINGS.json (schedule.scan_nights). Windows runs a missed scan
REM when the computer wakes, so a skipped night stays skipped. To scan any time:  nightly_scan.bat manual
if /i "%~1"=="manual" goto dayok
".venv\Scripts\python.exe" "scripts\settings.py" scan-tonight
if errorlevel 1 (
  echo ==== SKIPPED %DATE% %TIME% - not a scan night in SETTINGS.json ====>> "output\logs\nightly_scan.log"
  goto :eof
)
:dayok

REM Never run two scans at once.
powershell -NoProfile -Command "if (Get-CimInstance Win32_Process -Filter \"Name='python.exe'\" | Where-Object { $_.CommandLine -like '*run_scan.py*' }) { exit 1 } else { exit 0 }"
if errorlevel 1 (
  echo ==== SKIPPED %DATE% %TIME% - a scan is already running ====>> "output\logs\nightly_scan.log"
  goto :eof
)

REM After waking from sleep Windows may have no network yet. Wait up to 30 minutes for it.
set /a NETTRIES=0
:waitnet
powershell -NoProfile -Command "try { [System.Net.Dns]::GetHostAddresses('www.google.com') | Out-Null; exit 0 } catch { exit 1 }"
if not errorlevel 1 goto netok
set /a NETTRIES+=1
if %NETTRIES% GEQ 30 goto nonet
ping -n 61 127.0.0.1 >nul
goto waitnet
:nonet
for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format s"') do set NOWISO=%%i
echo ==== NO INTERNET %DATE% %TIME% - waited 30 minutes for a network connection; scan not started ====>> "output\logs\nightly_scan.log"
echo {"state": "no_internet", "started_at": "%NOWISO%", "finished_at": "%NOWISO%"}> "output\logs\scan_status.json"
goto :eof
:netok
if %NETTRIES% GTR 0 echo ==== NETWORK READY after %NETTRIES% minute(s) ====>> "output\logs\nightly_scan.log"
echo ==== START %DATE% %TIME% ====>> "output\logs\nightly_scan.log"
".venv\Scripts\python.exe" "scripts\run_scan.py" >> "output\logs\nightly_scan.log" 2>&1
echo ==== END %DATE% %TIME% exit=%ERRORLEVEL% ====>> "output\logs\nightly_scan.log"
endlocal
