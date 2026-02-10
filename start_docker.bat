@echo off
echo ===================================================
echo      INICIANDO CDO ORGANIZER (MODO DOCKER)
echo ===================================================
echo Requisito: Docker Desktop debe estar instalado y corriendo.
echo.

cd /d "%~dp0"

echo 1. Construyendo y levantando contenedores...
docker-compose up --build

pause
