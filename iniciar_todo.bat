@echo off
echo ===================================================
echo      INICIANDO SISTEMA COMPLETO CDO ORGANIZER
echo ===================================================
echo.

cd /d "%~dp0"

:: 1. Verificar e iniciar API (Puerto 8000)
netstat -ano | findstr :8000 >nul
if %errorlevel% NEQ 0 (
    echo [1/3] Iniciando Servidor API...
    start "CDO API Bridge" /min python -m uvicorn src.server_api:app --host 0.0.0.0 --port 8000
    timeout /t 3 >nul
) else (
    echo [OK] Servidor API ya esta corriendo.
)

:: 2. Verificar e iniciar Agente Local (Conectado a API)
tasklist /FI "IMAGENAME eq python.exe" /FO CSV | findstr "main.py" >nul
if %errorlevel% NEQ 0 (
    echo [2/3] Iniciando Agente Local...
    cd src\local_agent
    start "CDO Agente Local" /min python main.py
    cd ..\..
    timeout /t 2 >nul
) else (
    echo [OK] Agente Local ya esta corriendo.
)

:: 3. Iniciar Interfaz Web (Puerto 8501)
netstat -ano | findstr :8501 >nul
if %errorlevel% NEQ 0 (
    echo [3/3] Iniciando Interfaz Web...
    start "CDO Web Interface" streamlit run src/app_web.py --server.port 8501 --server.address 0.0.0.0
) else (
    echo [OK] Interfaz Web ya esta corriendo.
)

echo.
echo SISTEMA INICIADO CORRECTAMENTE.
echo - Web Local: http://localhost:8501
echo - Web Remota: http://%COMPUTERNAME%:8501
echo.
pause
