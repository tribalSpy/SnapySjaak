@echo off
setlocal enabledelayedexpansion
title SnappySjaak - Office PC Standby Setup
cd /d "%~dp0"

echo ============================================================
echo  SnappySjaak - Office PC Standby Setup
echo ============================================================
echo.
echo Run this ONLY on the office PC that will act as the backup
echo standby. It checks/installs prerequisites, installs app
echo dependencies, and generates the backup encryption keypair.
echo See docs\disaster-recovery-runbook.md for the full picture.
echo.
pause

echo.
echo --- Step 1: Checking prerequisites ---
echo.

where node >nul 2>nul
if errorlevel 1 (
    echo [MISSING] Node.js
    call :install_with_winget "Node.js LTS" "OpenJS.NodeJS.LTS"
) else (
    echo [OK] Node.js:
    node --version
)

where python >nul 2>nul
if errorlevel 1 (
    echo [MISSING] Python
    call :install_with_winget "Python 3" "Python.Python.3.12"
) else (
    echo [OK] Python:
    python --version
)

where psql >nul 2>nul
if errorlevel 1 (
    echo [MISSING] PostgreSQL command-line tools
    echo   PostgreSQL needs a superuser password set during install, so this
    echo   step is NOT silent -- an installer window will open if you continue.
    set /p DO_PG_INSTALL="  Attempt to install PostgreSQL now via winget? (y/n): "
    if /i "!DO_PG_INSTALL!"=="y" (
        call :install_with_winget "PostgreSQL" "PostgreSQL.PostgreSQL.16"
        echo   IMPORTANT: remember the superuser password you just set --
        echo   you will need it for DATABASE_URL in step 3 of the runbook.
    ) else (
        echo   Skipped. Install manually from https://www.postgresql.org/download/windows/
        echo   then re-run this script.
    )
) else (
    echo [OK] PostgreSQL client tools found.
)

echo.
echo If anything just got installed above, CLOSE this window and
echo re-run the script once so the new tools are picked up on PATH.
echo.
pause

echo.
echo --- Step 2: Installing app dependencies ---
echo.

where node >nul 2>nul
if errorlevel 1 (
    echo Node.js is not on PATH yet -- skipping npm install. Re-run after installing it.
    goto :after_npm
)
echo Installing Node dependencies (shadow-app)...
pushd "%~dp0shadow-app"
call npm ci
popd
:after_npm

where python >nul 2>nul
if errorlevel 1 (
    echo Python is not on PATH yet -- skipping Python setup. Re-run after installing it.
    goto :after_python
)
echo Setting up Python virtual environment...
if not exist "%~dp0.venv" (
    python -m venv "%~dp0.venv"
)
echo Installing Python dependencies...
call "%~dp0.venv\Scripts\pip.exe" install --upgrade pip
call "%~dp0.venv\Scripts\pip.exe" install -r "%~dp0requirements.txt"
:after_python

echo.
echo --- Step 3: Generating the backup encryption keypair ---
echo.
where node >nul 2>nul
if errorlevel 1 (
    echo Node.js is required for this step. Install it, then re-run this script.
    pause
    exit /b 1
)

echo ============================================================
echo  COPY BOTH KEYS BELOW NOW -- this is the only time they are shown.
echo ============================================================
node "%~dp0shadow-app\server\backup\generate-keypair.js"
echo ============================================================
echo.
echo  BACKUP_PUBLIC_KEY   -^> paste into Render's dashboard environment variables
echo  BACKUP_PRIVATE_KEY  -^> keep ONLY on this PC. Never on Render, never in git.
echo                         Also save a copy somewhere safe (e.g. a password
echo                         manager) -- losing it means losing every backup
echo                         ever taken.
echo.
echo ============================================================
echo.
echo Setup finished. Next: follow docs\disaster-recovery-runbook.md
echo section 1 (steps 5-8) to set the remaining environment variables
echo and start the app.
echo.
pause
exit /b 0

:install_with_winget
set "FRIENDLY_NAME=%~1"
set "PACKAGE_ID=%~2"
where winget >nul 2>nul
if errorlevel 1 (
    echo   winget is not available on this PC.
    echo   Install %FRIENDLY_NAME% manually, then re-run this script.
    exit /b 0
)
echo   Installing %FRIENDLY_NAME% via winget (this may open its own window)...
winget install --id %PACKAGE_ID% -e --accept-source-agreements --accept-package-agreements
if errorlevel 1 (
    echo   Automatic install of %FRIENDLY_NAME% did not finish cleanly.
    echo   Install it manually if it's still missing, then re-run this script.
)
exit /b 0
