@echo off
echo ===================================================
echo      INICIANDO CDO ORGANIZER (MODO SERVIDOR)
echo ===================================================

cd /d "%~dp0"

echo 1. Iniciando API Bridge (Puerto 8000)...
echo    Esta API permite la conexion de Agentes Remotos.
start "CDO API Bridge" python -m uvicorn src.server_api:app --host 0.0.0.0 --port 8000

timeout /t 3 /nobreak >nul

echo 2. Iniciando Interfaz Web (Puerto 8501)...
echo    Acceda a http://localhost:8501
start "CDO Web Interface" streamlit run src/app_web.py --server.port 8501 --server.address 0.0.0.0

echo.
echo Servidores en ejecucion.
echo - Web: http://localhost:8501
echo - API: http://localhost:8000/docs
echo.
echo NO CIERRE ESTA VENTANA ni las ventanas emergentes.
pause
