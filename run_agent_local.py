import os
import json
import subprocess
import sys

import socket
import time
import signal

# Configuration
SERVER_IP = "3.142.164.128"
SERVER_PORT = "8000"
SERVER_URL = f"http://{SERVER_IP}:{SERVER_PORT}"
USERNAME = "admin"
PASSWORD = "admin" # Default password, user should change this

def is_port_in_use(port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(('127.0.0.1', int(port))) == 0

def setup_agent():
    print(f"Setting up Local Agent to connect to {SERVER_URL}...")
    
    # 0. Check and Start Local Server if needed
    # (Disabled for AWS connection)
    if SERVER_IP == "127.0.0.1":
        server_process = None
        if not is_port_in_use(SERVER_PORT):
            print(f"Local server not detected on port {SERVER_PORT}. Starting server...")
            try:
                # Start uvicorn in a separate process
                server_process = subprocess.Popen(
                    [sys.executable, "-m", "uvicorn", "src.server_api:app", "--host", "127.0.0.1", "--port", SERVER_PORT],
                    cwd=os.getcwd(),
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE
                )
                print("Server starting... waiting for port 8000...")
                
                # Wait for port to open
                for _ in range(10):
                    if is_port_in_use(SERVER_PORT):
                        print("Server is UP!")
                        break
                    time.sleep(1)
                else:
                    print("Warning: Server might not have started correctly. Continuing anyway...")
            except Exception as e:
                print(f"Failed to start local server: {e}")
        else:
            print(f"Local server detected on port {SERVER_PORT}. Connecting...")
    else:
        print(f"Connecting to Remote Server: {SERVER_URL}")
        server_process = None

    # 1. Create config directory
    log_dir = os.path.join(os.getenv('LOCALAPPDATA', os.path.expanduser("~")), 'CDO_Organizer')
    if not os.path.exists(log_dir):
        os.makedirs(log_dir, exist_ok=True)
        print(f"Created config directory: {log_dir}")
        
    config_file = os.path.join(log_dir, "agent_config.json")
    
    # 2. Write config file
    # --- Port Validation & Auto-Correction ---
    # The user might manually set SERVER_PORT to 8501 thinking it's the correct one.
    # We must prevent this because the Agent connects to the API (8000), not Streamlit (8501).
    
    if SERVER_PORT == "8501":
        print("--------------------------------------------------")
        print("⚠️  ADVERTENCIA: Puerto 8501 detectado (Web UI).")
        print("    El agente debe conectarse al puerto de la API (8000).")
        print("    Corrigiendo automáticamente a puerto 8000...")
        print("--------------------------------------------------")
        SERVER_PORT_FINAL = "8000"
    else:
        SERVER_PORT_FINAL = SERVER_PORT
        
    SERVER_URL_FINAL = f"http://{SERVER_IP}:{SERVER_PORT_FINAL}"
    
    config = {
        "server_url": SERVER_URL_FINAL,
        "username": USERNAME,
        "password": PASSWORD,
        "task_url": f"{SERVER_URL_FINAL}/tasks/poll",
        "result_url": f"{SERVER_URL_FINAL}/tasks"
    }
    
    with open(config_file, "w") as f:
        json.dump(config, f, indent=4)
    
    print(f"Configuration saved to: {config_file}")
    print("--------------------------------------------------")
    print(f"IMPORTANTE: El agente se conectará a: {SERVER_URL}")
    print("NO use el puerto 8501 (Web UI). Use el puerto 8000 (API).")
    print("--------------------------------------------------")
    
    # 3. Install dependencies if needed
    print("Checking dependencies...")
    try:
        import requests
        import pandas
        from PIL import Image
        from selenium import webdriver
    except ImportError:
        print("Installing required packages...")
        requirements_path = os.path.join("src", "local_agent", "requirements.txt")
        if os.path.exists(requirements_path):
            subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", requirements_path])
        else:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "requests", "pandas", "python-docx", "openpyxl", "Pillow", "selenium", "webdriver-manager"])
        
    # 4. Run the agent
    agent_script = os.path.join("src", "local_agent", "main.py")
    if not os.path.exists(agent_script):
        print(f"Error: Agent script not found at {agent_script}")
        if server_process:
            server_process.terminate()
        return
        
    print(f"Starting Agent from {agent_script}...")
    print("Press Ctrl+C to stop.")
    
    try:
        subprocess.run([sys.executable, agent_script])
    except KeyboardInterrupt:
        print("\nStopping...")
    finally:
        if server_process:
            print("Stopping local server...")
            server_process.terminate()
            server_process.wait()
            print("Server stopped.")

if __name__ == "__main__":
    setup_agent()
