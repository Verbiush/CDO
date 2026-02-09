@echo off
cd /d "%~dp0"
title Iniciar Web App CDO

:: 1. Intentar usar el entorno virtual si existe
if exist "..\.venv\Scripts\python.exe" (
    echo Usando entorno virtual...
    "..\.venv\Scripts\python.exe" -m streamlit run src/app_web.py
    pause
    exit
)

:: 2. Si no, intentar usar python global
python --version >nul 2>&1
if %errorlevel% equ 0 (
    echo Usando Python global...
    python -m streamlit run src/app_web.py
    pause
    exit
)

:: 3. Si falla todo
echo [ERROR] No se encontro Python ni el entorno virtual.
echo Por favor ejecuta "Reparar_Entorno.bat" primero.
pause
