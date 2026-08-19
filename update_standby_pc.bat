@echo off
setlocal
title SnappySjaak - Office PC Standby Update
cd /d "%~dp0"

echo ============================================================
echo  SnappySjaak - Office PC Standby Update
echo ============================================================
echo.
echo Before running this: close the window running start_standby.bat
echo (or press Ctrl+C in it) so files aren't in use while updating.
echo Your data (standby-cache\, start_standby.bat, credentials\) is
echo untouched by any of this -- only the app code/dependencies update.
echo.
pause

if not exist "%~dp0.git" (
    echo.
    echo This folder isn't a git clone, so there's nothing to pull.
    echo To get updates in the future, set this PC up via "git clone"
    echo instead of a plain folder copy, then re-run this script.
    echo ^(You'll still need to copy credentials\service_account.json
    echo over separately after cloning -- git never tracks it.^)
    pause
    exit /b 1
)

echo.
echo --- Pulling latest code ---
echo.
git pull
if errorlevel 1 (
    echo.
    echo git pull failed -- resolve this manually before continuing
    echo ^(e.g. uncommitted local changes in this folder^), then re-run.
    pause
    exit /b 1
)

echo.
echo --- Updating Node dependencies ---
echo.
where node >nul 2>nul
if errorlevel 1 (
    echo Node.js not found on PATH -- skipping. Run setup_standby_pc.bat first.
    goto :after_npm
)
pushd "%~dp0shadow-app"
call npm ci
popd
:after_npm

echo.
echo --- Updating Python dependencies ---
echo.
if not exist "%~dp0.venv\Scripts\pip.exe" (
    echo No .venv found -- skipping. Run setup_standby_pc.bat first.
    goto :after_pip
)
call "%~dp0.venv\Scripts\pip.exe" install -r "%~dp0requirements.txt"
:after_pip

echo.
echo ============================================================
echo  Update finished. Double-click start_standby.bat to resume.
echo ============================================================
echo.
pause
exit /b 0
