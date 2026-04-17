@echo off
setlocal

REM Switch to script directory (supports spaces in path)
cd /d "%~dp0"

set "APP_FILE=app.py"
set "VENV_PYTHON=%~dp0.venv\Scripts\python.exe"
set "APP_URL=http://127.0.0.1:5000"

if not exist "%APP_FILE%" (
    echo [ERROR] app.py was not found. Put this script in the project root.
    pause
    exit /b 1
)

if exist "%VENV_PYTHON%" (
    set "PYTHON_EXE=%VENV_PYTHON%"
) else (
    set "PYTHON_EXE=python"
)

echo [INFO] Starting service...
start "" powershell -NoProfile -WindowStyle Hidden -Command "Start-Sleep -Seconds 4; Start-Process '%APP_URL%'"

echo [INFO] Browser will open automatically: %APP_URL%
echo [INFO] Keep this window open while the service is running.

"%PYTHON_EXE%" "%~dp0%APP_FILE%"

endlocal
