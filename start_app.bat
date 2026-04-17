@echo off
setlocal EnableExtensions

REM Switch to script directory (supports spaces in path)
cd /d "%~dp0"

set "APP_FILE=app.py"
set "VENV_PYTHON=%~dp0.venv\Scripts\python.exe"
set "APP_URL=http://127.0.0.1:5000"
set "BROWSER_WAIT_SECONDS=2"
set "BROWSER_READY_RETRIES=12"
set "PYTHON_EXE="
set "PYTHON_ARGS="
set "OPEN_BROWSER=1"

if /I "%~1"=="--no-browser" set "OPEN_BROWSER=0"

if not exist "%APP_FILE%" (
    echo [ERROR] app.py was not found. Put this script in the project root.
    pause
    exit /b 1
)

if exist "%VENV_PYTHON%" (
    set "PYTHON_EXE=%VENV_PYTHON%"
) else (
    where py >nul 2>nul
    if %ERRORLEVEL%==0 (
        set "PYTHON_EXE=py"
        set "PYTHON_ARGS=-3"
    ) else (
        where python >nul 2>nul
        if %ERRORLEVEL%==0 (
            set "PYTHON_EXE=python"
        )
    )
)

if "%PYTHON_EXE%"=="" (
    echo [ERROR] Python was not found. Please install Python 3 or create .venv first.
    echo [HINT] You can run: py -3 -m venv .venv
    pause
    exit /b 1
)

if /I "%~1"=="--check" (
    echo [OK] start_app.bat check passed.
    echo [INFO] Python: %PYTHON_EXE% %PYTHON_ARGS%
    exit /b 0
)

set "APP_DEBUG=0"
set "ENABLE_TEMPLATE_COM_THUMBNAIL=0"

echo [INFO] Starting service...
if "%OPEN_BROWSER%"=="1" (
    start "" powershell -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='SilentlyContinue'; $url='%APP_URL%'; Start-Sleep -Seconds %BROWSER_WAIT_SECONDS%; for($i=0; $i -lt %BROWSER_READY_RETRIES%; $i++){ try { $resp = Invoke-WebRequest -UseBasicParsing -Uri $url -TimeoutSec 2; if($resp.StatusCode -ge 200 -and $resp.StatusCode -lt 500){ break } } catch {}; Start-Sleep -Milliseconds 700 }; Start-Process $url"
)

if "%OPEN_BROWSER%"=="1" (
    echo [INFO] Browser will open automatically: %APP_URL%
) else (
    echo [INFO] Browser auto-open is disabled with --no-browser.
)
echo [INFO] Keep this window open while the service is running.

"%PYTHON_EXE%" %PYTHON_ARGS% "%~dp0%APP_FILE%"
set "EXIT_CODE=%ERRORLEVEL%"

if not "%EXIT_CODE%"=="0" (
    echo [ERROR] Service exited with code %EXIT_CODE%.
    pause
)

endlocal
