@echo off
setlocal
title UDISE Enrolment Automation
cd /d "%~dp0"

rem ---- 1. Find Python ----
set "PY="
py -3 -c "import sys" >nul 2>nul && set "PY=py -3"
if not defined PY python -c "import sys" >nul 2>nul && set "PY=python"
if not defined PY goto :nopython

rem ---- 2. Create private environment - first run only ----
if exist ".venv\Scripts\python.exe" ".venv\Scripts\python.exe" -c "pass" >nul 2>nul || rmdir /s /q ".venv" >nul 2>nul
if not exist ".venv\Scripts\python.exe" (
  echo First run: setting up, this takes 1-2 minutes and needs internet...
  %PY% -m venv .venv
  if errorlevel 1 goto :fail
)

rem ---- 3. Install libraries when missing or when an update changed them ----
fc /b "app\requirements.txt" ".venv\installed-requirements.txt" >nul 2>nul
if errorlevel 1 (
  ".venv\Scripts\python.exe" -m pip install --disable-pip-version-check -q -r app\requirements.txt
  if errorlevel 1 goto :fail
  copy /y "app\requirements.txt" ".venv\installed-requirements.txt" >nul
)

rem ---- 4. Tell if a newer version is on GitHub - never blocks, silent when offline ----
powershell -NoProfile -ExecutionPolicy Bypass -File "app\update.ps1" -CheckOnly

rem ---- 5. Run ----
".venv\Scripts\python.exe" app\kys_enrolment.py
echo.
pause
exit /b 0

:nopython
echo Python is not installed. Trying to install it automatically...
winget install -e --id Python.Python.3.12 --scope user --silent --accept-package-agreements --accept-source-agreements
echo.
echo If the install finished, CLOSE this window and double-click start.bat again.
echo If it failed, install Python from https://www.python.org/downloads/
echo and tick "Add python.exe to PATH" during setup.
pause
exit /b 1

:fail
echo.
echo Setup failed. Check the internet connection and run start.bat again.
pause
exit /b 1
