FROM python:3.10-slim

WORKDIR /app

# Instalar dependencias del sistema
# python3-tk: Para soporte básico de tkinter (aunque en server no muestra ventana)
# chromium: Para selenium si se usa
RUN apt-get update && apt-get install -y \
    build-essential \
    curl \
    software-properties-common \
    python3-tk \
    chromium \
    chromium-driver \
    && rm -rf /var/lib/apt/lists/*

# Copiar requirements
COPY requirements_web.txt .

# Instalar dependencias de Python
RUN pip3 install --no-cache-dir -r requirements_web.txt

# Copiar todo el código
COPY . .

# Exponer el puerto de Streamlit
EXPOSE 8501

# Comando de inicio
ENTRYPOINT ["streamlit", "run", "src/app_web.py", "--server.port=8501", "--server.address=0.0.0.0"]
