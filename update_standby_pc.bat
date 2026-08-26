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
if not exist "%~dp0.venv\Scripts\python.exe" (
    echo No .venv found -- skipping. Run setup_standby_pc.bat first.
    goto :after_pip
)
REM A .venv can go stale if Python was moved, reinstalled, or upgraded since
REM it was created (its pip/python shebangs point at an absolute path that no
REM longer exists) -- verify it still actually runs before trusting it,
REM rather than silently failing partway through pip install.
"%~dp0.venv\Scripts\python.exe" --version >nul 2>nul
if errorlevel 1 (
    echo Existing .venv looks broken ^(Python it points at is gone or moved^) --
    echo recreating it from scratch.
    rmdir /s /q "%~dp0.venv"
    where python >nul 2>nul
    if errorlevel 1 (
        echo Python is not on PATH -- can't recreate .venv. Run setup_standby_pc.bat first.
        goto :after_pip
    )
    python -m venv "%~dp0.venv"
    call "%~dp0.venv\Scripts\pip.exe" install --upgrade pip
)
call "%~dp0.venv\Scripts\pip.exe" install -r "%~dp0requirements.txt"
:after_pip

echo.
echo --- Applying Postgres schema migrations ---
echo.
where node >nul 2>nul
if errorlevel 1 (
    echo Node.js not found on PATH -- skipping. Migrations will still run
    echo automatically next time start_standby.bat starts the app.
    goto :after_migrations
)
REM DATABASE_URL only lives inside start_standby.bat, not in this script's
REM own environment -- pull it out so this step actually runs instead of
REM always no-op'ing (migrations still apply automatically at next startup
REM either way, this just gets the schema current sooner).
if exist "%~dp0start_standby.bat" (
    for /f "tokens=1,* delims==" %%a in ('findstr /b /c:"set DATABASE_URL=" "%~dp0start_standby.bat"') do set "DATABASE_URL=%%b"
)
node "%~dp0shadow-app\server\apply-migrations.js"
:after_migrations

echo.
echo ============================================================
for /f %%v in ('git rev-parse --short HEAD') do echo  Updated to commit %%v.
echo  Update finished. Double-click start_standby.bat to resume.
echo ============================================================
echo.
pause
exit /b 0
