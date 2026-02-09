@echo off
title Instalador CDO
cls
echo.
echo ========================================================
echo        PRE-INSTALADOR DE CLIENTE CDO
echo ========================================================
echo.
echo Comprobando requisitos del sistema...

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python no encontrado.
    echo Por favor, instale Python 3.10 o superior desde python.org
    echo Asegurese de marcar la casilla "Add Python to PATH".
    pause
    exit /b
)

echo [OK] Python detectado.
echo.
echo Iniciando Asistente Grafico...
python setup_wizard.py
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] No se pudo iniciar el asistente grafico.
    echo Intentando instalar dependencias basicas de GUI...
    pip install tk
    python setup_wizard.py
)

exit
