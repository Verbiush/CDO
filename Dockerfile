FROM python:3.10-slim

WORKDIR /app

# Instalar dependencias del sistema
# python3-tk: Para soporte básico de tkinter (aunque en server no muestra ventana)
# chromium: Para selenium si se usa en el servidor
RUN apt-get update && apt-get install -y \
    build-essential \
    curl \
    software-properties-common \
    python3-tk \
    chromium \
    chromium-driver \
    && rm -rf /var/lib/apt/lists/*

# Copiar requirements de servidor (sin pywin32)
COPY requirements_server.txt .

# Instalar dependencias de Python
RUN pip3 install --no-cache-dir -r requirements_server.txt

# Copiar todo el código
COPY . .

# Variables de entorno por defecto
ENV PORT=8501
ENV API_URL="http://backend:8000"

# El comando se define en docker-compose
CMD ["streamlit", "run", "src/app_web.py", "--server.port=8501", "--server.address=0.0.0.0"]
