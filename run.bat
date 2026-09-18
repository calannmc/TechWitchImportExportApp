@echo off
setlocal
cd /d "%~dp0"
title TechWitch Zendesk Helper

where py >nul 2>&1 && (set "PY=py -3") || (set "PY=python")
%PY% --version >nul 2>&1
if errorlevel 1 (
  echo Python 3 was not found. Install it from https://www.python.org/downloads/
  echo and tick "Add python.exe to PATH" during setup, then run this file again.
  pause
  exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
  echo Creating virtual environment...
  %PY% -m venv .venv || (echo Failed to create .venv & pause & exit /b 1)
)

echo Installing / checking dependencies...
".venv\Scripts\python.exe" -m pip install --quiet --upgrade pip
".venv\Scripts\python.exe" -m pip install --quiet -r requirements.txt || (echo pip install failed & pause & exit /b 1)

echo.
echo Starting TechWitch Zendesk Helper. Close this window to stop it.
".venv\Scripts\python.exe" app.py
pause
