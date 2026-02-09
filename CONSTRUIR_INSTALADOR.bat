@echo off
cd /d "%~dp0"
title Construir Instalador CDO
cls
echo ========================================================
echo        CONSTRUCTOR DE INSTALADOR CDO
echo ========================================================
echo.

:: 1. Verificar Python
if exist "..\.venv\Scripts\python.exe" (
    set PYTHON_CMD=..\.venv\Scripts\python.exe
    echo [OK] Usando entorno virtual.
) else (
    python --version >nul 2>&1
    if %errorlevel% neq 0 (
        echo [ERROR] Python no encontrado. Ejecuta "Reparar_Entorno.bat".
        pause
        exit /b
    )
    set PYTHON_CMD=python
    echo [OK] Usando Python global.
)

:: 2. Instalar dependencias
echo.
echo Instalando dependencias necesarias...
%PYTHON_CMD% -m pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo [ERROR] Fallo al instalar dependencias.
    pause
    exit /b
)

:: 3. Ejecutar build.py
echo.
echo Iniciando proceso de construccion (build.py)...
%PYTHON_CMD% build.py

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] La construccion fallo. Revisa los errores arriba.
) else (
    echo.
    echo [EXITO] Instalador creado en la carpeta 'dist'.
)

pause
