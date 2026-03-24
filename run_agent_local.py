import os
import json
import subprocess
import sys

import socket
import time
import signal

# Configuration
SERVER_IP = "18.118.37.215"
SERVER_PORT = "8000"
SERVER_URL = f"http://{SERVER_IP}:{SERVER_PORT}"
USERNAME = "admin"
PASSWORD = "admin" # Default password, user should change this

def is_port_in_use(port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(('127.0.0.1', int(port))) == 0

def setup_agent():
    """
    Configura el agente: pregunta al usuario por la IP del servidor y la guarda.
    """
    DEFAULT_SERVER_IP = "18.118.37.215" # AWS
    
    # Intentar usar la carpeta local del proyecto si AppData falla
    try:
        config_dir = os.path.join(os.getenv('LOCALAPPDATA', os.path.expanduser("~")), 'CDO_Organizer')
        if not os.path.exists(config_dir):
            os.makedirs(config_dir, exist_ok=True)
        config_file = os.path.join(config_dir, 'agent_config.json')
        # Probar escritura
        with open(config_file, "a") as f:
            pass
    except Exception:
        # Si falla (por ejemplo en sandbox), usar directorio actual
        config_dir = os.getcwd()
        config_file = os.path.join(config_dir, 'agent_config.json')
        print(f"Advertencia: No se pudo usar AppData, usando configuración local en: {config_file}")

    config = {}
    if os.path.exists(config_file):
        try:
            with open(config_file, 'r') as f:
                config = json.load(f)
        except:
            pass

    # Force new IP since it changed in AWS
    server_ip = "18.118.37.215"
    print(f"Dirección del servidor actual: {server_ip}")
    
    # En modo desatendido o si ya existe config, podemos saltar la pregunta con un timeout o argumento
    # Por ahora, simplemente permitimos cambiarlo si el usuario presiona Enter rápido?
    # Para simplicidad, preguntamos siempre pero con valor por defecto
    
    # Si se pasa el argumento --start, saltar configuración
    if "--start" in sys.argv:
        print("Argumento --start detectado. Saltando configuración.")
        new_ip = ""
    else:
        try:
            new_ip = input(f"Ingrese IP del servidor [Enter para usar {server_ip}]: ").strip()
        except EOFError:
            new_ip = "" # Handle non-interactive environments

    if new_ip:
        server_ip = new_ip
    
    # Ensure port 8000
    if ":" in server_ip:
        if server_ip.endswith(":8501"):
             server_ip = server_ip.replace(":8501", ":8000")
             print(f"Corrigiendo puerto a 8000: {server_ip}")
        elif not server_ip.endswith(":8000"):
             # Assume user knows what they are doing if they specify another port, 
             # but warn if it looks like streamlit port
             pass
    else:
        server_ip = f"{server_ip}:8000"

    config['server_ip'] = server_ip
    
    try:
        with open(config_file, 'w') as f:
            json.dump(config, f)
        print("Configuración guardada.")
    except Exception as e:
         print(f"No se pudo guardar la configuración: {e}")

    return server_ip

def run_agent():
    server_ip = setup_agent()
    print(f"Iniciando agente conectado a {server_ip}...")
    
    # Path to main agent script
    agent_script = os.path.join("src", "local_agent", "main.py")
    if not os.path.exists(agent_script):
        print(f"Error: No se encuentra el script del agente en {agent_script}")
        return

    try:
        # Run the agent script
        subprocess.run([sys.executable, agent_script], check=True)
    except KeyboardInterrupt:
        print("\nAgente detenido por el usuario.")
    except Exception as e:
        print(f"Error ejecutando el agente: {e}")

if __name__ == "__main__":
    run_agent()
