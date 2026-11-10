@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

rem Installs what the trained shelf model needs (ultralytics + PyTorch) into
rem the SAME python the poller runs it with: config.json's
rem "trained_model_python". Fixes "trained model counted no photos: No module
rem named 'ultralytics'" on the Shelf count page. Run it on the GPU PC.

set "PY="
if exist "config.json" (
    for /f "usebackq delims=" %%P in (`powershell -NoProfile -Command "(Get-Content -Raw 'config.json' | ConvertFrom-Json).trained_model_python"`) do set "PY=%%P"
)
if not defined PY set "PY=%~dp0..\shelf-training\.venv\Scripts\python.exe"

echo Trained model python: !PY!
if not exist "!PY!" (
    echo.
    echo [ERROR] That python does not exist on this PC.
    echo Run setup_gpu_pc.bat in the repo root first ^(it creates shelf-training\.venv^),
    echo then set "trained_model_python" in shelf-poller-app\config.json to
    echo   ...\shelf-training\.venv\Scripts\python.exe
    pause
    exit /b 1
)

echo.
echo [1/3] PyTorch
"!PY!" -c "import torch; print('  torch', torch.__version__, '- GPU:', torch.cuda.is_available())" 2>nul
if errorlevel 1 (
    echo   PyTorch is not installed in this python.
    echo   NVIDIA GPU? Install the CUDA build first, in another window:
    echo     "!PY!" -m pip install torch torchvision --index-url https://download.pytorch.org/whl/cu124
    echo   ^(or the command https://pytorch.org/get-started/locally/ gives^).
    echo   No GPU? Just continue -- ultralytics then installs the CPU build ^(slower^).
    pause
)

echo.
echo [2/3] Installing ultralytics ...
"!PY!" -m pip install -r "%~dp0..\shelf-training\training\requirements.txt"
if errorlevel 1 (
    echo [ERROR] pip install failed -- see the lines above.
    pause
    exit /b 1
)

echo.
echo [3/3] Check
"!PY!" -c "import ultralytics, torch; print('  ultralytics', ultralytics.__version__, '/ torch', torch.__version__, '/ GPU:', torch.cuda.is_available())"
if errorlevel 1 (
    echo [ERROR] ultralytics still does not import -- see above.
    pause
    exit /b 1
)

echo.
echo Done. No restart needed: the poller starts the model fresh for each
echo job, so the next shelf count shows the trained model's numbers again.
pause
