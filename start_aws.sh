#!/bin/bash
# Script de inicio para AWS
echo "Iniciando servicio de backend (API)..."
uvicorn src.server_api:app --host 0.0.0.0 --port 8000 &
PID_API=$!

echo "Iniciando servicio de frontend (Web)..."
streamlit run src/app_web.py --server.port=8501 --server.address=0.0.0.0 &
PID_WEB=$!

# Esperar a que terminen los procesos (o uno falle)
wait $PID_API $PID_WEB
