@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo ============================================
echo  Shelf Count GPU PC setup
echo ============================================
echo.

rem label-studio's Django dependency chain (django-environ in particular)
rem breaks on very new Python releases -- confirmed on 3.14 with a
rem "cannot import name 'find_loader' from pkgutil" ImportError (that name
rem was removed in 3.14). Prefer 3.12 via the py launcher when available so
rem a PC with only a bleeding-edge Python on PATH doesn't hit this blind.
set "PYTHON_CMD="
where py >nul 2>nul
if not errorlevel 1 (
    py -3.12 --version >nul 2>nul
    if not errorlevel 1 (
        set "PYTHON_CMD=py -3.12"
    ) else (
        py -3.11 --version >nul 2>nul
        if not errorlevel 1 (
            set "PYTHON_CMD=py -3.11"
        )
    )
)

if not defined PYTHON_CMD (
    where python >nul 2>nul
    if errorlevel 1 (
        echo [ERROR] Python was not found on PATH.
        echo Install Python 3.12 from https://www.python.org/downloads/windows/
        echo and make sure to check "Add python.exe to PATH" during install.
        echo Then re-run this script.
        pause
        exit /b 1
    )
    set "PYTHON_CMD=python"
    echo   Using system "python" -- no py launcher with 3.11/3.12 found.
    echo   label-studio is known to break on very new Python versions
    echo   -- confirmed on 3.14. If setup fails later with a pkgutil/environ
    echo   ImportError, install Python 3.12 via:
    echo     winget install --id Python.Python.3.12 -e
    echo   then delete shelf-training\.venv and re-run this script -- it
    echo   will pick 3.12 automatically via the py launcher next time.
) else (
    echo   Using "!PYTHON_CMD!" for the virtual environment.
)

echo [1/6] Creating shared virtual environment at shelf-training\.venv ...
if not exist "shelf-training\.venv" (
    !PYTHON_CMD! -m venv shelf-training\.venv
) else (
    echo   already exists, skipping.
)

call shelf-training\.venv\Scripts\activate.bat
python -m pip install --upgrade pip >nul

echo.
echo [2/6] Installing dataset-collection dependencies (Google Drive, dotenv) ...
pip install -r shelf-training\requirements.txt

echo.
echo [3/6] Installing labeling dependencies (Label Studio) -- this one is large, please wait ...
pip install -r shelf-training\labeling\requirements.txt

echo.
echo [4/6] PyTorch / Ultralytics
echo   ultralytics needs PyTorch. If this PC has an NVIDIA GPU, installing
echo   the default pip version now may silently give you a CPU-only build.
echo.
set /p HAS_CUDA_TORCH="Have you already installed a CUDA-enabled PyTorch build in this venv? [y/N]: "
if /i not "!HAS_CUDA_TORCH!"=="y" (
    echo.
    echo   Open https://pytorch.org/get-started/locally/ , pick your CUDA
    echo   version, and run the install command it gives you in THIS window
    echo   -- the venv here is already active -- before continuing.
    echo   No GPU on this PC? Just continue -- CPU training works, only slower.
    pause
)
echo Installing ultralytics ...
pip install -r shelf-training\training\requirements.txt

echo.
echo [5/6] Scaffolding config.json files from their examples (existing ones are never overwritten) ...
call :copy_config "llm-poller-app\config.example.json" "llm-poller-app\config.json"
call :copy_config "shelf-poller-app\config.example.json" "shelf-poller-app\config.json"
call :copy_config "shelf-training\data\config.example.json" "shelf-training\data\config.json"
call :copy_config "shelf-training\labeling\config.example.json" "shelf-training\labeling\config.json"
call :copy_config "shelf-training\training\config.example.json" "shelf-training\training\config.json"

echo.
echo [6/6] Checking for repo-root .env (Google Drive credentials) ...
if exist ".env" (
    echo   .env found.
) else (
    echo   [WARNING] No .env file at repo root. shelf-training\data\collect.py
    echo   needs GOOGLE_SERVICE_ACCOUNT_JSON and GOOGLE_DRIVE_ROOT_FOLDER_ID in it.
    echo   Copy the repo's .env from the main PC/server -- it holds real credentials,
    echo   never commit it to git -- send it directly, not through this repo.
)

echo.
echo ============================================
echo  Setup finished. Still needed by hand:
echo ============================================
echo  1. Fill in shelf-training\data\config.json      (server_url, api_key, drive_root_folder_id)
echo  2. Fill in llm-poller-app\config.json and shelf-poller-app\config.json
echo     (server_url, api_key, agent_name, pc_name, model_name)
echo  3. Confirm Ollama has the vision model installed: ollama list
echo  4. Run shelf-training\labeling\start_label_studio.bat once, grab a
echo     Personal Access Token from the Label Studio UI, and fill it into
echo     shelf-training\labeling\config.json
echo.
echo See GPU_PC_SETUP.md for the full plan and what needs to be sent separately.
pause
goto :eof

:copy_config
if not exist "%~2" (
    if exist "%~1" (
        copy "%~1" "%~2" >nul
        echo   created %~2
    )
) else (
    echo   %~2 already exists, left untouched
)
goto :eof
