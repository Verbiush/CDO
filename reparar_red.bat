@echo off
echo ===================================================
echo      SOLUCIONADOR DE PROBLEMAS DE RED (CDO)
echo ===================================================
echo Este script abrira los puertos necesarios en el Firewall.
echo.
echo >> Solicitando permisos de Administrador...
echo.

netsh advfirewall firewall show rule name="CDO Web Interface" >nul
if %errorlevel% NEQ 0 (
    echo [+] Abriendo Puerto 8501 (Web)...
    netsh advfirewall firewall add rule name="CDO Web Interface" dir=in action=allow protocol=TCP localport=8501
) else (
    echo [OK] Puerto 8501 ya esta configurado.
)

netsh advfirewall firewall show rule name="CDO API Bridge" >nul
if %errorlevel% NEQ 0 (
    echo [+] Abriendo Puerto 8000 (API)...
    netsh advfirewall firewall add rule name="CDO API Bridge" dir=in action=allow protocol=TCP localport=8000
) else (
    echo [OK] Puerto 8000 ya esta configurado.
)

echo.
echo ===================================================
echo      REINICIANDO SERVICIOS LIMPIAMENTE
echo ===================================================
echo 1. Cerrando procesos antiguos...
taskkill /F /IM streamlit.exe >nul 2>&1
taskkill /F /IM python.exe /FI "WINDOWTITLE eq CDO API Bridge" >nul 2>&1
taskkill /F /IM python.exe /FI "WINDOWTITLE eq CDO Web Interface" >nul 2>&1

echo 2. Iniciando API (Puerto 8000)...
start "CDO API Bridge" python -m uvicorn src.server_api:app --host 0.0.0.0 --port 8000

echo 3. Iniciando Web App (Puerto 8501)...
timeout /t 2 >nul
start "CDO Web Interface" streamlit run src/app_web.py --server.port 8501 --server.address 0.0.0.0

echo.
echo LISTO!
echo - Prueba local: http://localhost:8501
echo - Prueba remota (busca tu IP): ipconfig
echo.
pause
