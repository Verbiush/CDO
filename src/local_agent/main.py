import time
import requests
import json
import os
import platform
import sys
import threading
from getpass import getpass

# Add parent directory to path to allow imports from src (if running from source)
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)
if parent_dir not in sys.path:
    sys.path.insert(0, parent_dir)

def log_error(msg):
    try:
        with open("agent_error.log", "a") as f:
            f.write(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {msg}\n")
    except:
        pass

try:
    import bot_zeus
    # Ensure we have access to necessary functions
    from bot_zeus import abrir_navegador_inicial, obtener_driver
except ImportError as e:
    # Fallback if running standalone without full src context
    bot_zeus = None
    abrir_navegador_inicial = None
    obtener_driver = None
    log_error(f"Failed to import bot_zeus: {e}")

try:
    import pandas as pd
except ImportError as e:
    pd = None
    log_error(f"Failed to import pandas: {e}")

CONFIG_FILE = "agent_config.json"

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

def setup():
    print("\n=== Configuración del Agente CDO ===")
    print("Este agente conectará su PC con la nube para permitir funciones nativas.")
    
    default_url = "http://3.142.164.128:8000"
    
    # Check if we are in a GUI environment (no console)
    is_gui = False
    try:
        if sys.stdin is None or sys.stdin.closed:
            is_gui = True
    except:
        is_gui = True
        
    if is_gui:
        import tkinter as tk
        from tkinter import simpledialog, messagebox
        
        root = tk.Tk()
        root.withdraw() # Hide main window
        
        messagebox.showinfo("Configuración Requerida", "El agente no está configurado.\nPor favor ingrese los datos de conexión.")
        
        url = simpledialog.askstring("Configuración", f"URL del Servidor:", initialvalue=default_url)
        if not url: return None
        
        username = simpledialog.askstring("Configuración", "Usuario CDO:")
        if not username: return None
        
        password = simpledialog.askstring("Configuración", "Contraseña CDO:", show='*')
        if not password: return None
        
        root.destroy()
    else:
        url = input(f"URL del Servidor (Enter para '{default_url}'): ").strip()
        if not url: url = default_url
        
        username = input("Usuario CDO: ").strip()
        password = getpass("Contraseña CDO: ").strip()
    
    if not url.startswith("http"): url = "http://" + url
    url = url.rstrip("/")
    
    print(f"Conectando a: {url}")
    
    # Verify connection
    print("Verificando credenciales...")
    try:
        auth = (username, password)
        # Try to ping or poll to verify auth
        res = requests.get(f"{url}/tasks/poll", auth=auth, timeout=10)
        
        if res.status_code == 200:
            print("✅ Conexión exitosa!")
            return save_config(url, username, password)
        elif res.status_code == 401:
            print("❌ Error de autenticación: Usuario o contraseña incorrectos.")
            return None
        elif res.status_code == 404:
            print("❌ Error: No se encontró el servicio del agente en esa URL (404).")
            print("Asegúrese de incluir el puerto si es necesario (ej: :8000)")
            return None
        else:
            print(f"❌ Error inesperado: {res.status_code}")
            return None
    except Exception as e:
        print(f"❌ Error de conexión: {e}")
        return None

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

def launch_browser(url=None, headless=False):
    if abrir_navegador_inicial is None:
        return {"error": "Bot Zeus module not found"}
    
    try:
        # We need to ensure this runs in the main thread or compatible way
        # Since we are in a loop, it should be fine.
        # However, bot_zeus uses streamlit session state usually.
        # We might need to mock it or adjust bot_zeus to work standalone.
        # The modified bot_zeus checks for 'st' but handles it being None.
        
        driver = obtener_driver()
        if driver:
            if url:
                driver.get(url)
            return {"status": "success", "message": "Browser launched"}
        else:
             # Try initializing
            success, msg = abrir_navegador_inicial()
            if success and url:
                driver = obtener_driver()
                if driver: driver.get(url)
            return {"status": "success" if success else "error", "message": msg}
            
    except Exception as e:
        return {"error": str(e)}

def process_task(task):
    command = task.get("command")
    params = task.get("params", {})
    
    print(f"[{time.strftime('%H:%M:%S')}] Ejecutando comando: {command}")
    
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
            try:
                import tkinter as tk
                from tkinter import filedialog
                
                root = tk.Tk()
                root.withdraw() # Hide main window
                root.attributes('-topmost', True) # Bring to front
                
                path = filedialog.askdirectory(title=title)
                root.destroy()
                
                result["result"] = {"path": path if path else None}
            except Exception as e:
                result["status"] = "ERROR"
                result["result"] = {"error": f"Tkinter error: {str(e)}"}

        elif command == "browse_file":
            title = params.get("title", "Seleccionar Archivo")
            file_types_list = params.get("file_types", [])
            
            # Convert list of lists to list of tuples for tkinter
            file_types = []
            if file_types_list:
                for ft in file_types_list:
                    if len(ft) >= 2:
                        file_types.append((ft[0], ft[1]))
            
            if not file_types:
                file_types = [("All Files", "*.*")]
                
            try:
                import tkinter as tk
                from tkinter import filedialog
                
                root = tk.Tk()
                root.withdraw()
                root.attributes('-topmost', True)
                
                path = filedialog.askopenfilename(title=title, filetypes=file_types)
                root.destroy()
                
                result["result"] = {"path": path if path else None}
            except Exception as e:
                result["status"] = "ERROR"
                result["result"] = {"error": f"Tkinter error: {str(e)}"}
            
        elif command == "launch_browser":
            url = params.get("url")
            res = launch_browser(url)
            result["result"] = res
            if "error" in res:
                result["status"] = "ERROR"
                
        # --- BOT ZEUS COMMANDS ---
        elif command == "bot_get_focused_element":
            if not bot_zeus:
                result["status"] = "ERROR"
                result["result"] = {"error": "Module bot_zeus not found"}
            else:
                driver = bot_zeus.obtener_driver(create_if_missing=False)
                if not driver:
                    result["status"] = "ERROR"
                    result["result"] = {"error": "Browser not open"}
                else:
                    # Use internal function to get xpath with frames
                    xpath, frames, error = bot_zeus._detectar_foco_con_frames(driver)
                    if error:
                        result["status"] = "ERROR"
                        result["result"] = {"error": error}
                    else:
                        result["result"] = {"xpath": xpath, "frames": frames}

        elif command == "bot_start_visual_selector":
            if not bot_zeus:
                result["status"] = "ERROR"
                result["result"] = {"error": "Module bot_zeus not found"}
            else:
                ok, msg = bot_zeus.iniciar_selector_visual()
                if ok:
                    result["result"] = {"message": msg}
                else:
                    result["status"] = "ERROR"
                    result["result"] = {"error": msg}

        elif command == "bot_get_visual_selection":
            if not bot_zeus:
                result["status"] = "ERROR"
                result["result"] = {"error": "Module bot_zeus not found"}
            else:
                ok, xpath = bot_zeus.obtener_seleccion_visual()
                if ok:
                    result["result"] = {"xpath": xpath}
                else:
                    # Not an error, just no selection yet
                    result["result"] = {"xpath": None}

        elif command == "bot_switch_window":
            if not bot_zeus:
                result["status"] = "ERROR"
                result["result"] = {"error": "Module bot_zeus not found"}
            else:
                driver = bot_zeus.obtener_driver(create_if_missing=False)
                if not driver:
                    result["status"] = "ERROR"
                    result["result"] = {"error": "Browser not open"}
                else:
                    try:
                        index = params.get("index", -1)
                        handles = driver.window_handles
                        target = None
                        if index == -1:
                            target = handles[-1]
                        elif 0 <= index < len(handles):
                            target = handles[index]
                        
                        if target:
                            driver.switch_to.window(target)
                            result["result"] = {"status": "success", "title": driver.title}
                        else:
                            result["status"] = "ERROR"
                            result["result"] = {"error": f"Index {index} out of range"}
                    except Exception as e:
                        result["status"] = "ERROR"
                        result["result"] = {"error": str(e)}

        elif command == "bot_run_sequence":
            if not bot_zeus or not pd:
                result["status"] = "ERROR"
                result["result"] = {"error": "Module bot_zeus or pandas not found"}
            else:
                steps = params.get("steps", [])
                data = params.get("data", []) # List of dicts
                
                # Setup
                bot_zeus.set_pasos(steps)
                
                # Create DataFrame
                if data:
                    df = pd.DataFrame(data)
                else:
                    df = pd.DataFrame()

                # Run
                logs = []
                try:
                    # Consume generator
                    for log in bot_zeus.ejecutar_secuencia(df):
                        print(f"[Bot] {log}")
                        logs.append(log)
                    
                    result["result"] = {"logs": logs}
                except Exception as e:
                    result["status"] = "ERROR"
                    result["result"] = {"error": str(e), "logs": logs}

        else:
            result["status"] = "ERROR"
            result["result"] = {"error": f"Unknown command: {command}"}
            
    except Exception as e:
        result["status"] = "ERROR"
        result["result"] = {"error": str(e)}
        print(f"Error executing {command}: {e}")
        
    return result

def main():
    print("Iniciando Agente CDO...")
    
    # Ensure we are in the directory of the script to find config
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    
    config = load_config()
    if not config:
        config = setup()
        
    if not config:
        print("No se pudo configurar el agente. Saliendo.")
        input("Presione Enter para cerrar...")
        return

    url = config.get("server_url", config.get("url"))
    if not url:
        print("Error: Configuración incompleta (URL no encontrada).")
        # Try setup again? Or exit?
        config = setup()
        if config:
            url = config.get("server_url", config.get("url"))
        
        if not url:
            print("No se pudo obtener la URL del servidor.")
            input("Presione Enter para cerrar...")
            return

    username = config.get("username")
    password = config.get("password")
    auth = (username, password)
    
    print(f"Agente conectado a {url} como {username}")
    print("Esperando comandos... (Presione Ctrl+C para salir)")
    
    error_count = 0
    
    while True:
        try:
            res = requests.get(f"{url}/tasks/poll", auth=auth, timeout=30)
            
            if res.status_code == 200:
                error_count = 0
                data = res.json()
                tasks = data.get("tasks", [])
                
                if tasks:
                    print(f"Recibidas {len(tasks)} tareas.")
                
                for task in tasks:
                    res_data = process_task(task)
                    
                    # Submit result
                    try:
                        post_res = requests.post(
                            f"{url}/tasks/{task['id']}/result",
                            json=res_data,
                            auth=auth,
                            timeout=30
                        )
                        if post_res.status_code != 200:
                            print(f"Error enviando resultado: {post_res.text}")
                    except Exception as e:
                        print(f"Error enviando resultado: {e}")
                        
            elif res.status_code == 401:
                print("Error de autenticación. Credenciales inválidas.")
                # Delete config to force re-login
                if os.path.exists(CONFIG_FILE):
                    os.remove(CONFIG_FILE)
                print("Reinicie el agente para configurar nuevamente.")
                break
            else:
                print(f"Error del servidor: {res.status_code}")
                error_count += 1
                
        except Exception as e:
            print(f"Error de conexión: {e}")
            error_count += 1
            
        # Exponential backoff for errors
        sleep_time = 2 if error_count == 0 else min(30, 2 * error_count)
        time.sleep(sleep_time)

if __name__ == "__main__":
    main()
