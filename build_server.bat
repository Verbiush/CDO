
@echo off
cd /d "%~dp0"
title Build CDO Server API
cls
echo ========================================================
echo        BUILD CDO SERVER API
echo ========================================================
echo.

:: 1. Verify Python
if exist "..\.venv\Scripts\python.exe" (
    set PYTHON_CMD=..\.venv\Scripts\python.exe
    echo [OK] Using virtual environment.
) else (
    python --version >nul 2>&1
    if %errorlevel% neq 0 (
        echo [ERROR] Python not found.
        pause
        exit /b
    )
    set PYTHON_CMD=python
    echo [OK] Using global Python.
)

:: 2. Install dependencies (PyInstaller if needed)
echo.
echo Checking dependencies...
%PYTHON_CMD% -m pip install pyinstaller requests pandas fastapi uvicorn
if %errorlevel% neq 0 (
    echo [ERROR] Failed to install dependencies.
    pause
    exit /b
)

:: 3. Run build script
echo.
echo Starting build process...
%PYTHON_CMD% build_server.py

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Build failed. Check errors above.
) else (
    echo.
    echo [SUCCESS] Executable created in 'dist'.
)

pause
