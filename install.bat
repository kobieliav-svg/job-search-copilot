@echo off
REM Job Search Copilot - full setup on Windows: Python environment, libraries, and the browser
REM the scanners use. Run it once from the project folder (double-click, or ask Claude to run it).
setlocal
cd /d "%~dp0"
echo == Job Search Copilot setup ==

where python >nul 2>nul
if errorlevel 1 (
    echo Python was not found. Install Python 3.10 or newer from https://www.python.org/downloads/
    echo and tick "Add python.exe to PATH" during the install. Then run this file again.
    exit /b 1
)

echo -- Creating the Python environment (.venv)
python -m venv .venv
call .venv\Scripts\activate.bat

echo -- Installing the Python libraries
python -m pip install --upgrade pip >nul
pip install -r requirements.txt

echo -- Installing the browser used by the scanners (Chromium)
python -m playwright install chromium

echo -- Creating your profile files and pipeline\jobs.xlsx (existing files are never overwritten)
python scripts\init_profile.py

where soffice >nul 2>nul
if errorlevel 1 (
    echo -- Optional: LibreOffice was not found. Word files still work; LibreOffice is only used to
    echo    count resume pages. Install it from https://www.libreoffice.org if you want that check.
)

echo.
echo Setup complete. Go back to Claude and say: "The install finished. Continue the full setup."
endlocal
