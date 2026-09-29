@echo off
setlocal enabledelayedexpansion

rem Standalone bootstrap script -- copy this ONE file to a PC that doesn't
rem sync the Synology Drive folder (USB stick, email, download it from
rem GitHub directly), then run it there. First run clones the repo; every
rem run after that just pulls the latest changes.
rem
rem Usage:
rem   pull_code.bat                  clones/pulls into .\SnapySjaak
rem   pull_code.bat C:\some\path     clones/pulls into that folder instead

set REPO_URL=https://github.com/tribalSpy/SnapySjaak.git
set REPO_DIR=%~dp0SnapySjaak
if not "%~1"=="" set REPO_DIR=%~1

where git >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Git was not found on PATH.
    echo Install Git for Windows from https://git-scm.com/download/win
    echo then re-run this script.
    pause
    exit /b 1
)

if exist "%REPO_DIR%\.git" (
    echo Repo already exists at %REPO_DIR% -- pulling latest changes...
    pushd "%REPO_DIR%"
    git fetch origin
    git checkout main
    git pull --ff-only origin main
    if errorlevel 1 (
        echo.
        echo [WARNING] Pull failed -- most likely there are local changes in
        echo this folder that conflict with what's on GitHub. This script
        echo never discards local work automatically. To see what's local:
        echo   git status
        echo Commit or "git stash" those changes, then re-run this script.
    )
    popd
) else (
    if exist "%REPO_DIR%" (
        echo [ERROR] %REPO_DIR% already exists but isn't a git repo.
        echo Move or delete it, or pass a different target folder as the
        echo first argument to this script.
        pause
        exit /b 1
    )
    echo Cloning into %REPO_DIR% ...
    git clone "%REPO_URL%" "%REPO_DIR%"
)

echo.
echo Done. Code is at %REPO_DIR%
echo Next: run %REPO_DIR%\setup_gpu_pc.bat
pause
