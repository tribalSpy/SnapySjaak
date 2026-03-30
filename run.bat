@echo off
setlocal

cd /d "%~dp0"

set "VENV_PYTHON=%CD%\.venv\Scripts\python.exe"
set "SYSTEM_PYTHON=C:\Python313\python.exe"
set "PYTHON_EXE="
set "SOURCE_LABEL="

if exist "%VENV_PYTHON%" (
    "%VENV_PYTHON%" -c "import dotenv, streamlit" >nul 2>nul
    if not errorlevel 1 (
        set "PYTHON_EXE=%VENV_PYTHON%"
        set "SOURCE_LABEL=project virtual environment"
    )
)

if not defined PYTHON_EXE if exist "%SYSTEM_PYTHON%" (
    "%SYSTEM_PYTHON%" -c "import dotenv, streamlit" >nul 2>nul
    if not errorlevel 1 (
        set "PYTHON_EXE=%SYSTEM_PYTHON%"
        set "SOURCE_LABEL=local Python"
    )
)

if defined PYTHON_EXE (
    echo Using %SOURCE_LABEL%...
) else (
    echo Could not find a usable Python environment.
    echo Checked:
    echo   %VENV_PYTHON%
    echo and
    echo   %SYSTEM_PYTHON%
    echo.
    echo Install dependencies with one of:
    echo   "%SYSTEM_PYTHON%" -m pip install --user -r requirements.txt
    echo or
    echo   "%SYSTEM_PYTHON%" -m venv .venv
    echo   .venv\Scripts\python.exe -m pip install -r requirements.txt
    pause
    exit /b 1
)

echo Starting Streamlit dashboard...
"%PYTHON_EXE%" -m streamlit run app.py

pause
