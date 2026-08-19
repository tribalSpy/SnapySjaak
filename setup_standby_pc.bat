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
echo dependencies, generates the backup secrets, and writes a
echo start_standby.bat with most of them already filled in.
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
        echo   you will need it for DATABASE_URL in start_standby.bat later.
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
echo --- Step 3: Generating secrets and start_standby.bat ---
echo.
where node >nul 2>nul
if errorlevel 1 (
    echo Node.js is required for this step. Install it, then re-run this script.
    pause
    exit /b 1
)

echo ============================================================
echo  COPY WHAT'S PRINTED BELOW NOW -- this is the only time the
echo  private key is shown. It has also been saved into
echo  start_standby.bat for you.
echo ============================================================
node "%~dp0shadow-app\server\backup\generate-keypair.js"
echo ============================================================
echo.
echo Next steps (also in docs\disaster-recovery-runbook.md section 1):
echo   1. Paste BACKUP_PUBLIC_KEY and BACKUP_AGENT_API_KEY (printed above)
echo      into Render's dashboard environment variables.
echo   2. Create an empty Postgres database and note its connection string.
echo   3. Copy credentials\service_account.json from your local dev setup
echo      into this folder.
echo   4. Open start_standby.bat and fill in RENDER_BACKUP_BASE_URL and
echo      DATABASE_URL.
echo   5. Double-click start_standby.bat to start the standby.
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
