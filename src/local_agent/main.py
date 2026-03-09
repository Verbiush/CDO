import time
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import json
import os
import platform
import sys
import threading
import queue
import tkinter as tk
from tkinter import simpledialog, messagebox, scrolledtext, ttk
from getpass import getpass

# Add parent directory to path to allow imports from src (if running from source)
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)
if parent_dir not in sys.path:
    sys.path.insert(0, parent_dir)

CONFIG_FILE = "agent_config.json"

# --- FILE PROCESSING LOGIC ---

def recursive_update_cups(data, old_val, new_val):
    count = 0
    if isinstance(data, dict):
        for k, v in data.items():
            if k == "codServicio" and str(v).strip() == str(old_val).strip():
                data[k] = new_val
                count += 1
            elif isinstance(v, (dict, list)):
                count += recursive_update_cups(v, old_val, new_val)
    elif isinstance(data, list):
        for item in data:
            count += recursive_update_cups(item, old_val, new_val)
    return count

def process_update_cups(folder_path, old_val, new_val):
    count_files = 0
    total_changes = 0
    errors = []
    
    if not os.path.isdir(folder_path):
        return {"status": "error", "message": "Carpeta no válida"}

    files_to_process = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.lower().endswith('.json'):
                files_to_process.append(os.path.join(root, file))
    
    for file_path in files_to_process:
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
        
            changes = recursive_update_cups(data, old_val, new_val)
        
            if changes > 0:
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(data, f, indent=4, ensure_ascii=False)
                count_files += 1
                total_changes += changes
        except Exception as e:
            errors.append(f"{os.path.basename(file_path)}: {str(e)}")
            
    return {
        "count_files": count_files,
        "total_changes": total_changes,
        "errors": errors
    }

def get_config_path():
    # 1. Check current directory
    if os.path.exists(CONFIG_FILE):
        return CONFIG_FILE
    
    # 2. Check LocalAppData/CDO_Organizer (Standard install location)
    local_appdata = os.getenv('LOCALAPPDATA', os.path.expanduser("~"))
    app_dir = os.path.join(local_appdata, "CDO_Organizer")
    config_path = os.path.join(app_dir, CONFIG_FILE)
    
    if os.path.exists(config_path):
        return config_path
        
    # 3. Return default location for creation (AppData if possible, else CWD)
    if os.path.exists(app_dir):
        return config_path
    return CONFIG_FILE

def load_config():
    config_path = get_config_path()
    if os.path.exists(config_path):
        try:
            with open(config_path, 'r') as f:
                return json.load(f)
        except:
            return None
    return None

def save_config(url, username, password):
    config = {"server_url": url, "username": username, "password": password}
    config_path = get_config_path()
    
    # Ensure directory exists if saving to AppData
    os.makedirs(os.path.dirname(os.path.abspath(config_path)), exist_ok=True)
    
    with open(config_path, 'w') as f:
        json.dump(config, f)
    return config

# --- HELPER FUNCTIONS ---

def list_drives():
    drives = []
    if platform.system() == "Windows":
        import string
        from ctypes import windll
        bitmask = windll.kernel32.GetLogicalDrives()
        for letter in string.ascii_uppercase:
            if bitmask & 1:
                drives.append(f"{letter}:\\")
            bitmask >>= 1
    else:
        drives.append("/")
    return drives

def list_files(path):
    if not os.path.exists(path):
        return {"error": "Path not found"}
    
    items = []
    try:
        with os.scandir(path) as it:
            for entry in it:
                items.append({
                    "name": entry.name,
                    "is_dir": entry.is_dir(),
                    "path": entry.path
                })
        return items
    except Exception as e:
        return {"error": str(e)}

try:
    import bot_zeus
    from bot_zeus import abrir_navegador_inicial, obtener_driver
except ImportError:
    bot_zeus = None
    abrir_navegador_inicial = None
    obtener_driver = None

try:
    import pandas as pd
except ImportError:
    pd = None

def launch_browser(url=None, headless=False):
    if abrir_navegador_inicial is None:
        return {"error": "Bot Zeus module not found"}
    
    try:
        driver = obtener_driver()
        if driver:
            if url:
                driver.get(url)
            return {"status": "success", "message": "Browser launched"}
        else:
            success, msg = abrir_navegador_inicial()
            if success and url:
                driver = obtener_driver()
                if driver: driver.get(url)
            return {"status": "success" if success else "error", "message": msg}
    except Exception as e:
        return {"error": str(e)}

# --- WORKER THREAD ---

class AgentWorker(threading.Thread):
    def __init__(self, url, username, password, log_queue, status_queue):
        super().__init__()
        self.url = url
        self.username = username
        self.password = password
        self.auth = (username, password)
        self.log_queue = log_queue
        self.status_queue = status_queue
        self.running = True
        self.session = requests.Session()
        
        retry_strategy = Retry(
            total=5,
            backoff_factor=1,
            status_forcelist=[429, 500, 502, 503, 504],
            allowed_methods=["HEAD", "GET", "POST", "PUT", "DELETE", "OPTIONS", "TRACE"]
        )
        adapter = HTTPAdapter(max_retries=retry_strategy)
        self.session.mount("http://", adapter)
        self.session.mount("https://", adapter)
        self.session.headers.update({
            'User-Agent': 'CDO_Agent/1.0',
            'Connection': 'keep-alive'
        })

    def log(self, msg):
        self.log_queue.put(f"[{time.strftime('%H:%M:%S')}] {msg}")

    def update_status(self, status, color):
        self.status_queue.put((status, color))

    def process_task(self, task):
        command = task.get("command")
        params = task.get("params", {})
        
        self.log(f"Ejecutando comando: {command}")
        
        result = {"status": "COMPLETED", "result": None}
        
        try:
            if command == "list_drives":
                drives = list_drives()
                result["result"] = {"drives": drives}
                
            elif command == "list_files":
                path = params.get("path", "")
                files = list_files(path)
                result["result"] = {"items": files}
                
            elif command == "browse_folder":
                title = params.get("title", "Seleccionar Carpeta")
                
                # Use queue to request the main thread to open the dialog
                try:
                    request = {"type": "browse_folder", "title": title}
                    self.log_queue.put(("UI_REQUEST", request, task['id']))
                    return "UI_PENDING" # Signal that we are waiting for UI
                    
                except Exception as e:
                    result["status"] = "ERROR"
                    result["result"] = {"error": f"UI Error: {str(e)}"}

            elif command == "browse_file":
                title = params.get("title", "Seleccionar Archivo")
                file_types = params.get("file_types", [])
                request = {"type": "browse_file", "title": title, "file_types": file_types}
                self.log_queue.put(("UI_REQUEST", request, task['id']))
                return "UI_PENDING"

            elif command == "update_cups":
                path = params.get("path")
                old_val = params.get("old_val")
                new_val = params.get("new_val")
                
                if not path or not os.path.exists(path):
                    result["status"] = "ERROR"
                    result["result"] = {"error": "Ruta inválida o no encontrada"}
                else:
                    self.log(f"Procesando CUPS en: {path}")
                    res = process_update_cups(path, old_val, new_val)
                    result["result"] = res

            elif command == "launch_browser":
                url = params.get("url")
                res = launch_browser(url)
                result["result"] = res
                if "error" in res:
                    result["status"] = "ERROR"
            
            # ... (Add other commands like bot_zeus here if needed) ...
            else:
                result["status"] = "ERROR"
                result["result"] = {"error": f"Unknown command: {command}"}

        except Exception as e:
            result["status"] = "ERROR"
            result["result"] = {"error": str(e)}
            self.log(f"Error executing {command}: {e}")
            
        return result

    def run(self):
        self.log(f"Conectando a {self.url}...")
        error_count = 0
        
        while self.running:
            try:
                # Poll
                try:
                    res = self.session.get(f"{self.url}/tasks/poll", auth=self.auth, timeout=30)
                except requests.exceptions.Timeout:
                     # Just a timeout, perfectly normal for long polling
                     continue
                except requests.exceptions.ConnectionError:
                    self.update_status("Error de Conexión", "red")
                    self.log("No se puede conectar al servidor.")
                    time.sleep(5)
                    continue

                if res.status_code == 200:
                    self.update_status(f"Conectado: {self.username}", "green")
                    error_count = 0
                    data = res.json()
                    tasks = data.get("tasks", [])
                    
                    if tasks:
                        self.log(f"Recibidas {len(tasks)} tareas.")
                    
                    for task in tasks:
                        res_data = self.process_task(task)
                        
                        if res_data == "UI_PENDING":
                            continue # Main thread will handle submission
                            
                        # Submit result
                        try:
                            post_res = self.session.post(
                                f"{self.url}/tasks/{task['id']}/result",
                                json=res_data,
                                auth=self.auth,
                                timeout=30
                            )
                        except Exception as e:
                            self.log(f"Error enviando resultado: {e}")
                            
                elif res.status_code == 401:
                    self.update_status("Error de Autenticación", "red")
                    self.log("Credenciales inválidas.")
                    self.running = False # Stop retry
                    break
                else:
                    self.update_status(f"Error {res.status_code}", "orange")
                    self.log(f"Respuesta del servidor: {res.status_code}")
                    error_count += 1
                    
            except Exception as e:
                self.log(f"Error inesperado: {e}")
                self.update_status("Error Interno", "red")
                error_count += 1
            
            time.sleep(2)
        
        self.update_status("Desconectado", "gray")

# --- GUI APP ---

class AgentApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("CDO Agente Local")
        self.geometry("500x550")
        self.resizable(False, False)
        
        try:
            self.iconbitmap("assets/favicon.ico") # Optional if file exists
        except:
            pass

        self.worker = None
        self.log_queue = queue.Queue()
        self.status_queue = queue.Queue()
        
        self.ui_pending_tasks = {} # Store task_ids for UI requests

        self.create_widgets()
        self.load_settings()
        
        self.check_queues()

    def create_widgets(self):
        # Header
        header_frame = tk.Frame(self, bg="#0e1117", height=80)
        header_frame.pack(fill="x")
        
        title_label = tk.Label(header_frame, text="Agente CDO", font=("Arial", 20, "bold"), bg="#0e1117", fg="white")
        title_label.pack(pady=20)
        
        # Status
        self.status_var = tk.StringVar(value="Desconectado")
        self.status_label = tk.Label(self, textvariable=self.status_var, font=("Arial", 10, "bold"), fg="gray", bg="#f0f0f0")
        self.status_label.pack(fill="x", pady=5)
        
        # Settings Frame
        settings_frame = tk.LabelFrame(self, text="Configuración de Conexión", padx=10, pady=10)
        settings_frame.pack(padx=20, pady=10, fill="x")
        
        tk.Label(settings_frame, text="URL Servidor:").grid(row=0, column=0, sticky="w", pady=5)
        self.url_entry = tk.Entry(settings_frame, width=40)
        self.url_entry.grid(row=0, column=1, padx=5)
        self.url_entry.insert(0, "http://3.142.164.128:8000")
        
        tk.Label(settings_frame, text="Usuario:").grid(row=1, column=0, sticky="w", pady=5)
        self.user_entry = tk.Entry(settings_frame, width=40)
        self.user_entry.grid(row=1, column=1, padx=5)
        
        tk.Label(settings_frame, text="Contraseña:").grid(row=2, column=0, sticky="w", pady=5)
        self.pass_entry = tk.Entry(settings_frame, width=40, show="*")
        self.pass_entry.grid(row=2, column=1, padx=5)
        
        # Buttons
        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=10)
        
        self.connect_btn = tk.Button(btn_frame, text="Conectar", command=self.toggle_connection, bg="#007bff", fg="white", width=15)
        self.connect_btn.pack(side="left", padx=5)
        
        self.exit_btn = tk.Button(btn_frame, text="Salir", command=self.on_closing, width=10)
        self.exit_btn.pack(side="left", padx=5)
        
        # Log Area
        log_frame = tk.LabelFrame(self, text="Registro de Actividad", padx=5, pady=5)
        log_frame.pack(padx=20, pady=10, fill="both", expand=True)
        
        self.log_area = scrolledtext.ScrolledText(log_frame, height=10, state="disabled", font=("Consolas", 9))
        self.log_area.pack(fill="both", expand=True)
        
        # Protocol for closing
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def load_settings(self):
        config = load_config()
        if config:
            if "server_url" in config:
                self.url_entry.delete(0, tk.END)
                self.url_entry.insert(0, config["server_url"])
            if "username" in config:
                self.user_entry.delete(0, tk.END)
                self.user_entry.insert(0, config["username"])
            if "password" in config:
                self.pass_entry.delete(0, tk.END)
                self.pass_entry.insert(0, config["password"])

    def toggle_connection(self):
        if self.worker and self.worker.is_alive():
            # Disconnect
            self.worker.running = False
            self.worker.join(timeout=2)
            self.worker = None
            self.connect_btn.config(text="Conectar", bg="#007bff")
            self.url_entry.config(state="normal")
            self.user_entry.config(state="normal")
            self.pass_entry.config(state="normal")
            self.status_var.set("Desconectado")
            self.status_label.config(fg="gray")
            self.title("CDO Agente Local")
            self.log("Desconectado manualmente.")
        else:
            # Connect
            url = self.url_entry.get().strip()
            user = self.user_entry.get().strip()
            pwd = self.pass_entry.get().strip()
            
            if not url or not user or not pwd:
                messagebox.showerror("Error", "Todos los campos son obligatorios.")
                return
                
            # Normalize URL
            if not url.startswith("http"): url = "http://" + url
            url = url.rstrip("/")
            
            save_config(url, user, pwd)
            
            self.url_entry.config(state="disabled")
            self.user_entry.config(state="disabled")
            self.pass_entry.config(state="disabled")
            self.connect_btn.config(text="Desconectar", bg="#dc3545")
            
            self.title(f"CDO Agente Local - {user}")
            
            self.worker = AgentWorker(url, user, pwd, self.log_queue, self.status_queue)
            self.worker.start()
            self.log("Iniciando conexión...")

    def log(self, msg):
        self.log_area.config(state="normal")
        self.log_area.insert(tk.END, msg + "\n")
        self.log_area.see(tk.END)
        self.log_area.config(state="disabled")

    def check_queues(self):
        # Check logs
        try:
            while True:
                item = self.log_queue.get_nowait()
                if isinstance(item, tuple) and item[0] == "UI_REQUEST":
                    # Handle UI Request in main thread
                    self.handle_ui_request(item[1], item[2])
                else:
                    self.log(item)
        except queue.Empty:
            pass
            
        # Check status
        try:
            while True:
                status_text, color = self.status_queue.get_nowait()
                self.status_var.set(status_text)
                self.status_label.config(fg=color)
        except queue.Empty:
            pass
            
        self.after(100, self.check_queues)

    def handle_ui_request(self, request, task_id):
        from tkinter import filedialog
        
        self.log(f"Solicitud de UI recibida: {request['type']}")
        self.attributes('-topmost', True) # Bring to front
        
        result_data = {"status": "COMPLETED", "result": None}
        
        try:
            if request['type'] == "browse_folder":
                path = filedialog.askdirectory(title=request['title'])
                result_data["result"] = {"path": path if path else None}
                
            elif request['type'] == "browse_file":
                file_types = request.get('file_types', [])
                ft_tuples = []
                if file_types:
                    for ft in file_types:
                        if len(ft) >= 2:
                            ft_tuples.append((ft[0], ft[1]))
                if not ft_tuples:
                    ft_tuples = [("All Files", "*.*")]
                    
                path = filedialog.askopenfilename(title=request['title'], filetypes=ft_tuples)
                result_data["result"] = {"path": path if path else None}
                
        except Exception as e:
            result_data["status"] = "ERROR"
            result_data["result"] = {"error": str(e)}
            self.log(f"Error en UI: {e}")
        
        self.attributes('-topmost', False)
        
        # Send result back via worker session (in a new thread to avoid blocking UI)
        threading.Thread(target=self.send_result_async, args=(task_id, result_data)).start()

    def send_result_async(self, task_id, result_data):
        if self.worker and self.worker.running:
            try:
                self.log("Enviando resultado de UI...")
                post_res = self.worker.session.post(
                    f"{self.worker.url}/tasks/{task_id}/result",
                    json=result_data,
                    auth=self.worker.auth,
                    timeout=30
                )
                if post_res.status_code != 200:
                    self.log(f"Error enviando resultado UI: {post_res.text}")
            except Exception as e:
                self.log(f"Error enviando resultado UI: {e}")

    def on_closing(self):
        if self.worker and self.worker.is_alive():
            if messagebox.askokcancel("Salir", "El agente está conectado. ¿Desea desconectar y salir?"):
                self.worker.running = False
                self.destroy()
        else:
            self.destroy()

def main():
    app = AgentApp()
    app.mainloop()

if __name__ == "__main__":
    main()
