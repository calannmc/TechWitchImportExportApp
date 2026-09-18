@echo off
setlocal
cd /d "%~dp0"
title Build TechWitch Zendesk Helper exe

where py >nul 2>&1 && (set "PY=py -3") || (set "PY=python")

if not exist ".venv\Scripts\python.exe" (
  echo Creating virtual environment...
  %PY% -m venv .venv || (echo Failed to create .venv & pause & exit /b 1)
)

echo Installing build dependencies...
".venv\Scripts\python.exe" -m pip install --quiet --upgrade pip
".venv\Scripts\python.exe" -m pip install --quiet -r requirements-build.txt || (echo pip install failed & pause & exit /b 1)

echo Building...
if exist build rmdir /s /q build
".venv\Scripts\python.exe" -m PyInstaller --noconfirm --clean TechWitchImportExportApp.spec || (echo Build failed & pause & exit /b 1)

echo.
echo Done. The exe is in dist\TechWitchImportExportApp.exe
echo Copy it anywhere; it creates exports\ and uploads\ folders next to itself.
pause
