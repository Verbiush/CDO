@echo off
title Reparar Entorno Python
cls
echo ========================================================
echo        REPARADOR DE ENTORNO PYTHON
echo ========================================================
echo.
echo 1. Verificando instalacion de Python...

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python no encontrado en el sistema.
    echo.
    echo Se abrira la pagina de descarga de Python.
    echo POR FAVOR INSTALE PYTHON 3.10 o SUPERIOR.
    echo *** IMPORTANTE: Marque la casilla "Add Python to PATH" al instalar. ***
    echo.
    pause
    start https://www.python.org/downloads/
    exit /b
)

echo [OK] Python encontrado.
python --version
echo.

echo 2. Limpiando entorno virtual antiguo/roto...
if exist "..\.venv" (
    rmdir /s /q "..\.venv"
    echo [OK] Entorno .venv eliminado.
)

echo.
echo 3. Creando nuevo entorno virtual (.venv)...
python -m venv ..\.venv
if %errorlevel% neq 0 (
    echo [ERROR] No se pudo crear el entorno virtual.
    pause
    exit /b
)
echo [OK] Entorno virtual creado.

echo.
echo 4. Instalando dependencias (esto puede tardar)...
..\.venv\Scripts\python -m pip install --upgrade pip
..\.venv\Scripts\pip install -r requirements.txt

if %errorlevel% neq 0 (
    echo [ERROR] Fallo la instalacion de dependencias.
    pause
    exit /b
)

echo.
echo ========================================================
echo        REPARACION COMPLETADA EXITOSAMENTE
echo ========================================================
echo.
echo Ahora puedes ejecutar "Iniciar Web App.bat" o "INSTALAR.bat".
echo.
pause
