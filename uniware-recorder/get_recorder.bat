@echo off
rem Another PC (e.g. the backup PC): copy this file and start_recorder.bat
rem to a folder of their own (e.g. C:\UniwareRecorder) and double-click it.
rem First time it fetches only the recorder from git (not the SnappySjaak
rem app) and installs it; after that it fetches the newest code and actions.
rem Needs Git for Windows and Python 3.12 (python.org, "Add to PATH").
cd /d "%~dp0"
setlocal
set "REPO=https://github.com/tribalSpy/SnapySjaak.git"
set "DIR=%~dp0SnapySjaak"

where git >nul 2>nul || (echo Git is not installed: get it from git-scm.com and run this again. & pause & exit /b 1)

if not exist "%DIR%\.git" (
  echo Fetching the recorder...
  git clone --filter=blob:none --no-checkout "%REPO%" "%DIR%" || (echo Fetching failed -- sign in to GitHub when asked. & pause & exit /b 1)
  git -C "%DIR%" sparse-checkout set uniware-recorder
  git -C "%DIR%" checkout main || (pause & exit /b 1)
) else (
  echo Fetching the newest version...
  git -C "%DIR%" pull --ff-only || (echo Update failed: an action changed on this PC? See README, "Another PC". & pause & exit /b 1)
)

set "APP=%DIR%\uniware-recorder"
if not exist "%APP%\.venv\Scripts\python.exe" (
  echo Installing...
  where py >nul 2>nul && (py -3.12 -m venv "%APP%\.venv" || py -3 -m venv "%APP%\.venv") || python -m venv "%APP%\.venv"
)
"%APP%\.venv\Scripts\python.exe" -m pip install -q --upgrade pip
"%APP%\.venv\Scripts\python.exe" -m pip install -q -r "%APP%\requirements.txt"
if errorlevel 1 (
  echo.
  echo Install failed. winsdk needs Python 3.9 - 3.12: install Python 3.12 from python.org and run this again.
) else (
  echo.
  echo Ready. Start it with start_recorder.bat
)
pause
