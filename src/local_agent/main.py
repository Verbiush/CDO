import sys
import io
import os
import time
import json
import logging
import threading
import requests
import base64
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
import uvicorn

# Fix for PyInstaller --noconsole mode
class NullWriter:
    def write(self, text): pass
    def flush(self): pass
    def isatty(self): return False
    def fileno(self): return -1

if sys.stdout is None: sys.stdout = NullWriter()
if sys.stderr is None: sys.stderr = NullWriter()

# Configure logging
try:
    log_dir = os.path.join(os.getenv('LOCALAPPDATA', os.path.expanduser("~")), 'CDO_Organizer')
    if not os.path.exists(log_dir):
        os.makedirs(log_dir, exist_ok=True)
    log_file = os.path.join(log_dir, "agent.log")
    config_file = os.path.join(log_dir, "agent_config.json")
except:
    log_file = "agent.log"
    config_file = "agent_config.json"

handlers = [logging.FileHandler(log_file)]
if not getattr(sys, 'frozen', False) or hasattr(sys.stdout, 'isatty'):
    handlers.append(logging.StreamHandler(sys.stdout))

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=handlers
)
logger = logging.getLogger("LocalAgent")

# Import Modules
try:
    sys.path.append(os.path.dirname(os.path.abspath(__file__)))
    # Also add parent dir if running from source
    sys.path.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
    from modules.ovida_validator import OvidaValidator
    # from modules.registraduria_validator import ValidatorRegistraduria
    # from modules.adres_validator import ValidatorAdres, ValidatorAdresWeb
except ImportError:
    logger.error("Could not import modules. Ensure 'modules' folder is present.")

# --- Polling Client for Remote Mode ---
class PollingClient(threading.Thread):
    def __init__(self, server_url, username, password):
        super().__init__()
        self.server_url = server_url
        self.username = username
        self.password = password
        self.running = True
        self.daemon = True
        self.drivers = {} # Store active validators/drivers: {'ovida': instance}

    def run(self):
        logger.info(f"Starting Polling Client to {self.server_url} as {self.username}")
        while self.running:
            try:
                self.poll()
            except Exception as e:
                logger.error(f"Polling error: {e}")
            time.sleep(2)

    def poll(self):
        try:
            resp = requests.get(
                f"{self.server_url}/tasks/poll",
                auth=(self.username, self.password),
                timeout=5
            )
            if resp.status_code == 200:
                data = resp.json()
                tasks = data.get("tasks", [])
                for task in tasks:
                    self.execute_task(task)
            else:
                logger.warning(f"Poll failed: {resp.status_code}")
        except requests.exceptions.ConnectionError:
            pass # Silent retry

    def execute_task(self, task):
        logger.info(f"Executing task: {task['command']}")
        task_id = task['id']
        command = task['command']
        params = task.get('params', {})
        
        result = {"success": False, "data": None, "error": None}
        status_code = "COMPLETED"

        try:
            if command == "PING":
                result["success"] = True
                result["data"] = "PONG"
            
            elif command == "LIST_FILES":
                path = params.get("path")
                if os.path.exists(path) and os.path.isdir(path):
                    files = []
                    for entry in os.scandir(path):
                        files.append({
                            "name": entry.name,
                            "is_dir": entry.is_dir(),
                            "path": entry.path,
                            "size": entry.stat().st_size if not entry.is_dir() else 0
                        })
                    result["success"] = True
                    result["data"] = files
                else:
                    result["error"] = "Path not found or not a directory"
                    status_code = "ERROR"

            elif command == "OVIDA_LAUNCH":
                try:
                    validator = OvidaValidator(headless=False)
                    validator.launch_browser()
                    validator.go_to_login()
                    self.drivers['ovida'] = validator
                    result["success"] = True
                    result["data"] = "Browser launched. Please login."
                except Exception as e:
                    result["error"] = str(e)
                    status_code = "ERROR"

            elif command == "OVIDA_PROCESS":
                validator = self.drivers.get('ovida')
                if not validator:
                    # Try to create new one if not exists (might fail login check)
                    validator = OvidaValidator(headless=False)
                    validator.launch_browser()
                    self.drivers['ovida'] = validator
                
                # Check login
                if not validator.check_login_status():
                   result["error"] = "Not logged in. Please login first."
                   status_code = "ERROR"
                else:
                    try:
                        df = pd.DataFrame(params.get('data', []))
                        col_map = params.get('col_map', {})
                        save_path = params.get('save_path')
                        
                        res = validator.process_massive(df, col_map, save_path)
                        result["success"] = True
                        result["data"] = res
                    except Exception as e:
                        result["error"] = str(e)
                        status_code = "ERROR"

            else:
                result["error"] = f"Unknown command: {command}"
                status_code = "ERROR"

        except Exception as e:
            result["error"] = str(e)
            status_code = "ERROR"
        
        # Send result back
        try:
            requests.post(
                f"{self.server_url}/tasks/{task_id}/result",
                json={"status": status_code, "result": result},
                auth=(self.username, self.password)
            )
        except Exception as e:
            logger.error(f"Failed to send result: {e}")

app = FastAPI(title="Local File System Agent", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def select_folder_dialog():
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    folder_path = filedialog.askdirectory()
    root.destroy()
    return folder_path

@app.get("/")
def read_root():
    return {"status": "running", "agent": "LocalFileSystemAgent"}

@app.post("/select-folder")
def select_folder():
    try:
        folder = select_folder_dialog()
        if not folder:
            return {"cancelled": True}
        return {"path": folder, "cancelled": False}
    except Exception as e:
        logger.error(f"Error selecting folder: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/list-files")
def list_files(path: str):
    if not os.path.exists(path): raise HTTPException(status_code=404, detail="Path not found")
    if not os.path.isdir(path): raise HTTPException(status_code=400, detail="Path is not a directory")
    try:
        files = []
        for entry in os.scandir(path):
            files.append({
                "name": entry.name,
                "is_dir": entry.is_dir(),
                "path": entry.path,
                "size": entry.stat().st_size if not entry.is_dir() else 0
            })
        return {"files": files}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    if os.path.exists(config_file):
        try:
            with open(config_file, "r") as f:
                config = json.load(f)
            server_url = config.get("server_url")
            username = config.get("username")
            password = config.get("password")
            if server_url and username and password:
                polling_client = PollingClient(server_url, username, password)
                polling_client.start()
        except Exception as e:
            logger.error(f"Error loading config: {e}")
            
    uvicorn.run(app, host="127.0.0.1", port=8989)
