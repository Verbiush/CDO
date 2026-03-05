import sys
import io
import os
import time
import json
import logging
import threading
import requests
import base64
import traceback
import multiprocessing
import tkinter as tk
from tkinter import scrolledtext, messagebox

# --- GUI Class ---
class AgentGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("CDO Agente Local")
        self.root.geometry("600x400")
        
        # Status Frame
        status_frame = tk.Frame(root)
        status_frame.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Label(status_frame, text="Estado:").pack(side=tk.LEFT)
        self.status_label = tk.Label(status_frame, text="Iniciando...", fg="blue", font=("Arial", 10, "bold"))
        self.status_label.pack(side=tk.LEFT, padx=5)
        
        # Log Area
        tk.Label(root, text="Registros (Logs):").pack(anchor=tk.W, padx=10)
        self.log_area = scrolledtext.ScrolledText(root, state='disabled', height=15)
        self.log_area.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Buttons
        btn_frame = tk.Frame(root)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)
        
        tk.Button(btn_frame, text="Ocultar Ventana", command=self.minimize_to_tray).pack(side=tk.LEFT)
        tk.Button(btn_frame, text="Abrir Logs", command=self.open_logs).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="Salir", command=self.on_closing, bg="#ffcccc").pack(side=tk.RIGHT)

        self.root.protocol("WM_DELETE_WINDOW", self.minimize_to_tray) # Default close minimizes
        
        # Redirect logging to GUI
        self.setup_logging()
        
    def setup_logging(self):
        class TextHandler(logging.Handler):
            def __init__(self, text_widget):
                logging.Handler.__init__(self)
                self.text_widget = text_widget

            def emit(self, record):
                msg = self.format(record)
                def append():
                    self.text_widget.configure(state='normal')
                    self.text_widget.insert(tk.END, msg + '\n')
                    self.text_widget.see(tk.END)
                    self.text_widget.configure(state='disabled')
                # Schedule update on main thread
                self.text_widget.after(0, append)

        text_handler = TextHandler(self.log_area)
        text_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        logging.getLogger().addHandler(text_handler)
        
    def set_status(self, text, color="black"):
        self.status_label.config(text=text, fg=color)
        
    def minimize_to_tray(self):
        self.root.iconify() # Just minimize for now, real tray needs pystray
        
    def open_logs(self):
        try:
            log_dir = os.path.join(os.getenv('LOCALAPPDATA', os.path.expanduser("~")), 'CDO_Organizer')
            os.startfile(log_dir)
        except:
            pass

    def on_closing(self):
        if messagebox.askokcancel("Salir", "¿Desea detener el Agente Local?"):
            self.root.destroy()
            os._exit(0) # Force exit

# --- CRITICAL: Early Error Logging Setup ---
# Setup a debug log in the same directory as the executable/script
try:
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    
    debug_log_path = os.path.join(base_dir, "agent_debug_crash.log")
    
    # Configure logging to write to this file immediately
    logging.basicConfig(
        filename=debug_log_path,
        level=logging.DEBUG,
        format='%(asctime)s - %(levelname)s - %(message)s',
        filemode='w' # Overwrite each run
    )
    logging.info("--- Agent Starting (Debug Mode) ---")
except Exception as e:
    # If we can't even log, we are in trouble. 
    pass

# Fix for PyInstaller --noconsole mode
class NullWriter:
    def write(self, text): 
        try:
            logging.info(f"STDOUT/ERR: {text.strip()}")
        except: pass
    def flush(self): pass
    def isatty(self): return False
    def fileno(self): return -1

# Redirect stdout/stderr to our logger/null
sys.stdout = NullWriter()
sys.stderr = NullWriter()

import pandas as pd
from tkinter import filedialog
try:
    from fastapi import FastAPI, HTTPException
    from fastapi.middleware.cors import CORSMiddleware
    import uvicorn
except ImportError as e:
    logging.critical(f"Failed to import FastAPI/Uvicorn: {e}")
    # Keep process alive to read log
    time.sleep(60)
    sys.exit(1)

# Configure standard logging (can coexist or override basicConfig)
# ... existing logging setup ...
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
    if not getattr(sys, 'frozen', False):
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
    def __init__(self, server_url, username, password, gui=None):
        super().__init__()
        self.server_url = server_url
        self.username = username
        self.password = password
        self.gui = gui
        self.running = True
        self.daemon = True
        self.drivers = {} # Store active validators/drivers: {'ovida': instance}

    def run(self):
        msg = f"Iniciando conexión a {self.server_url}..."
        logger.info(msg)
        if self.gui: self.gui.set_status(msg, "orange")
        
        while self.running:
            try:
                self.poll()
            except Exception as e:
                logger.error(f"Error de polling: {e}")
                if self.gui: self.gui.set_status("Error de Conexión", "red")
            time.sleep(2)

    def poll(self):
        try:
            resp = requests.get(
                f"{self.server_url}/tasks/poll",
                auth=(self.username, self.password),
                timeout=5
            )
            if resp.status_code == 200:
                if self.gui: self.gui.set_status("Conectado (Esperando tareas)", "green")
                data = resp.json()
                tasks = data.get("tasks", [])
                for task in tasks:
                    self.execute_task(task)
            else:
                logger.warning(f"Poll fallido: {resp.status_code}")
                if self.gui: self.gui.set_status(f"Error HTTP {resp.status_code}", "red")
        except requests.exceptions.ConnectionError:
            if self.gui: self.gui.set_status("Reconectando...", "orange")
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
            
            elif command == "SELECT_FOLDER":
                try:
                    root = tk.Tk()
                    root.withdraw()
                    root.attributes('-topmost', True)
                    folder_path = filedialog.askdirectory()
                    root.destroy()
                    
                    if folder_path:
                        result["success"] = True
                        result["data"] = folder_path
                    else:
                        result["success"] = False
                        result["data"] = None # Cancelled
                except Exception as e:
                    result["error"] = str(e)
                    status_code = "ERROR"

            elif command == "SELECT_FILE":
                try:
                    root = tk.Tk()
                    root.withdraw()
                    root.attributes('-topmost', True)
                    file_path = filedialog.askopenfilename()
                    root.destroy()
                    
                    if file_path:
                        result["success"] = True
                        result["data"] = file_path
                    else:
                        result["success"] = False
                        result["data"] = None
                except Exception as e:
                    result["error"] = str(e)
                    status_code = "ERROR"

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

import multiprocessing

if __name__ == "__main__":
    multiprocessing.freeze_support()
    
    # --- Start GUI ---
    root = tk.Tk()
    gui = AgentGUI(root)
    
    logger.info(f"Agent starting. Config file: {config_file}")
    
    try:
        # --- Create default config if missing ---
        if not os.path.exists(config_file):
            logger.info("Config file not found. Creating default configuration for AWS...")
            default_config = {
                "server_url": "http://3.142.164.128:8000",
                "username": "admin",  # Default user
                "password": "password" # Default password (user should change this via UI if needed)
            }
            try:
                with open(config_file, "w") as f:
                    json.dump(default_config, f, indent=4)
                logger.info(f"Created default config at {config_file}")
            except Exception as e:
                logger.error(f"Failed to create default config: {e}")

        if os.path.exists(config_file):
            try:
                with open(config_file, "r") as f:
                    config = json.load(f)
                
                # Use .get() to avoid KeyError if key is missing
                server_url = config.get("server_url", "")
                username = config.get("username", "")
                password = config.get("password", "")
                
                # --- Auto-fix port 8501 to 8000 ---
                if server_url and ":8501" in server_url:
                    logger.warning(f"Detected incorrect port 8501 in server_url: {server_url}. Switching to port 8000 for API connection.")
                    server_url = server_url.replace(":8501", ":8000")
                
                logger.info(f"Loaded config for user: {username} at {server_url}")
                
                if server_url and username and password:
                    try:
                        # Pass GUI reference to client
                        polling_client = PollingClient(server_url, username, password, gui=gui)
                        polling_client.start()
                    except Exception as pe:
                        logger.error(f"Failed to start PollingClient: {pe}")
                else:
                    logger.warning("Config missing server_url, username or password")
            except Exception as e:
                logger.error(f"Error loading config: {e}")
        else:
            logger.warning(f"Config file not found at {config_file}. Run setup or create file manually.")
                
        # Start Uvicorn in a separate thread to not block GUI
        def start_uvicorn():
            try:
                logger.info("Starting Uvicorn Server on 127.0.0.1:8989")
                config = uvicorn.Config(app, host="127.0.0.1", port=8989, log_level="info")
                server = uvicorn.Server(config)
                server.run()
            except Exception as ue:
                logger.critical(f"Uvicorn server crashed: {ue}")
                logger.critical(traceback.format_exc())

        uvicorn_thread = threading.Thread(target=start_uvicorn, daemon=True)
        uvicorn_thread.start()
            
    except Exception as global_e:
        logging.critical(f"CRITICAL GLOBAL ERROR: {global_e}")
        logging.critical(traceback.format_exc())
        
    # --- START MAIN LOOP ---
    root.mainloop()
