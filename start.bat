@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>&1
if %errorlevel%==0 (
  set PY=py -3
) else (
  set PY=python
)

echo Installing packages...
%PY% -m pip install -r requirements.txt
if errorlevel 1 (
  echo Failed to install Python packages. Install Python 3 from https://www.python.org/downloads/ and tick "Add python.exe to PATH".
  pause
  exit /b 1
)

echo.
echo Starting Daily Sales Register at http://127.0.0.1:5050
%PY% app.py
pause
