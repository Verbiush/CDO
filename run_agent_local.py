import os
import json
import subprocess
import sys

# Configuration
SERVER_IP = "3.142.164.128"
SERVER_PORT = "8000"
SERVER_URL = f"http://{SERVER_IP}:{SERVER_PORT}"
USERNAME = "admin"
PASSWORD = "admin" # Default password, user should change this

def setup_agent():
    print(f"Setting up Local Agent to connect to {SERVER_URL}...")
    
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
        import fastapi
        import uvicorn
    except ImportError:
        print("Installing required packages...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "requests", "fastapi", "uvicorn", "pandas"])
        
    # 4. Run the agent
    agent_script = os.path.join("src", "local_agent", "main.py")
    if not os.path.exists(agent_script):
        print(f"Error: Agent script not found at {agent_script}")
        return
        
    print(f"Starting Agent from {agent_script}...")
    print("Press Ctrl+C to stop.")
    subprocess.run([sys.executable, agent_script])

if __name__ == "__main__":
    setup_agent()
