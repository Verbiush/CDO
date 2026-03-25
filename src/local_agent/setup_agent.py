import os
import sys
import shutil
import winreg
import subprocess
import time
import json
import tkinter as tk
from tkinter import messagebox, ttk, filedialog
import threading
import requests

class AgentInstaller(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Instalador Agente Local CDO")
        self.geometry("550x550")
        self.resizable(False, False)
        
        # Center window
        self.eval('tk::PlaceWindow . center')
        
        # Estilo
        style = ttk.Style()
        style.theme_use('clam')
        
        # Header
        header_frame = tk.Frame(self, bg="#0056b3", height=60)
        header_frame.pack(fill="x")
        tk.Label(header_frame, text="Instalador Agente Local CDO", font=("Segoe UI", 16, "bold"), bg="#0056b3", fg="white").pack(pady=15)
        
        # Content
        main_frame = tk.Frame(self, padx=20, pady=20)
        main_frame.pack(fill="both", expand=True)
        
        self.info_label = tk.Label(main_frame, text="Este asistente instalará y configurará el Agente Local\npara permitir la gestión de archivos desde la nube.", justify="center", font=("Segoe UI", 10))
        self.info_label.pack(pady=5)
        
        # Paths Frame
        paths_frame = tk.LabelFrame(main_frame, text="Ubicaciones", padx=10, pady=5)
        paths_frame.pack(fill="x", pady=5)
        
        # 1. Source Path (Origen)
        tk.Label(paths_frame, text="Origen (Archivos de instalación):").pack(anchor="w")
        
        # Check specific user path first for source
        default_src = ""
        user_specific_path = r"G:\JUAN\local_agent"
        
        # Logic: If running from temp (installer), try to find source.
        if getattr(sys, 'frozen', False):
            default_src = sys._MEIPASS
        else:
            default_src = os.path.dirname(os.path.abspath(__file__))

        # Override if user path exists and has files (prioritize existing structure if valid)
        if os.path.exists(user_specific_path) and (os.path.exists(os.path.join(user_specific_path, "main.py")) or os.path.exists(os.path.join(user_specific_path, "CDO_Agente.exe"))):
             # If we are running FROM here, default_src is already correct.
             # If we are running from elsewhere, we might want to use this as source if it's an update.
             pass

        self.src_path_var = tk.StringVar(value=default_src)
        src_row = tk.Frame(paths_frame)
        src_row.pack(fill="x", pady=2)
        tk.Entry(src_row, textvariable=self.src_path_var).pack(side="left", fill="x", expand=True)
        tk.Button(src_row, text="...", command=self.browse_source, width=3).pack(side="right", padx=(5,0))

        # 2. Destination Path (Destino)
        tk.Label(paths_frame, text="Destino (Carpeta de instalación):").pack(anchor="w", pady=(5,0))
        
        # Default destination: AppData or G:\JUAN\local_agent if it exists
        local_appdata = os.getenv('LOCALAPPDATA', os.path.expanduser("~"))
        default_dest = os.path.join(local_appdata, "CDO_Organizer")
        
        # Prioritize G:\JUAN\local_agent if it exists
        if os.path.exists(user_specific_path):
            default_dest = user_specific_path
            
        self.dest_path_var = tk.StringVar(value=default_dest)
        dest_row = tk.Frame(paths_frame)
        dest_row.pack(fill="x", pady=2)
        tk.Entry(dest_row, textvariable=self.dest_path_var).pack(side="left", fill="x", expand=True)
        tk.Button(dest_row, text="...", command=self.browse_dest, width=3).pack(side="right", padx=(5,0))

        # Credentials Inputs
        cred_frame = tk.LabelFrame(main_frame, text="Credenciales de Acceso", padx=10, pady=10)
        cred_frame.pack(fill="x", pady=10)
        
        tk.Label(cred_frame, text="Usuario:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        self.user_entry = tk.Entry(cred_frame, width=30)
        self.user_entry.insert(0, "admin")
        self.user_entry.grid(row=0, column=1, padx=5, pady=5)
        
        tk.Label(cred_frame, text="Contraseña:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
        self.pass_entry = tk.Entry(cred_frame, show="*", width=30)
        self.pass_entry.grid(row=1, column=1, padx=5, pady=5)
        
        # Progress
        self.progress = ttk.Progressbar(main_frame, orient="horizontal", length=400, mode="determinate")
        self.progress.pack(pady=15)
        
        self.status_label = tk.Label(main_frame, text="Listo para instalar...", fg="gray", font=("Segoe UI", 9))
        self.status_label.pack(pady=5)
        
        # Buttons
        btn_frame = tk.Frame(self, pady=10)
        btn_frame.pack(fill="x", side="bottom")
        
        self.install_btn = tk.Button(btn_frame, text="Instalar y Ejecutar", command=self.start_install, bg="#28a745", fg="white", font=("Segoe UI", 10, "bold"), width=20, relief="flat")
        self.install_btn.pack(side="right", padx=20)
        
        self.exit_btn = tk.Button(btn_frame, text="Salir", command=self.quit, width=10, relief="groove")
        self.exit_btn.pack(side="right", padx=10)
        
    def browse_source(self):
        d = filedialog.askdirectory(title="Seleccionar carpeta origen")
        if d: self.src_path_var.set(d)

    def browse_dest(self):
        d = filedialog.askdirectory(title="Seleccionar carpeta destino")
        if d: self.dest_path_var.set(d)
        
    def log(self, message):
        self.status_label.config(text=message)
        self.update_idletasks()
        
    def start_install(self):
        self.username = self.user_entry.get().strip()
        self.password = self.pass_entry.get().strip()
        if not self.username or not self.password:
             messagebox.showerror("Error", "Debe ingresar usuario y contraseña.")
             return
        self.install_btn.config(state="disabled")
        self.exit_btn.config(state="disabled")
        
        self.src_dir = self.src_path_var.get().strip()
        self.dest_dir = self.dest_path_var.get().strip()
        
        threading.Thread(target=self.run_installation, daemon=True).start()
        
    def run_installation(self):
        try:
            # 0. Validate Credentials
            self.log("Validando credenciales con el servidor...")
            self.progress['value'] = 10
            try:
                auth = (self.username, self.password)
                url = "http://3.138.135.181:8000"
                res = requests.get(f"{url}/auth/verify", auth=auth, timeout=10)
                if res.status_code == 200:
                    self.log("✅ Credenciales válidas.")
                elif res.status_code == 401:
                    raise Exception("Usuario o contraseña incorrectos.")
                else:
                    self.log(f"⚠️ Advertencia: Validación parcial (HTTP {res.status_code}).")
            except Exception as e:
                resp = messagebox.askyesno("Advertencia de Conexión", f"No se pudo conectar al servidor para validar credenciales:\n{e}\n\n¿Desea continuar de todos modos?")
                if not resp:
                    self.reset_ui()
                    return

            self.progress['value'] = 20
            self.log("Verificando archivos...")
            
            # 1. Definir origen y archivos
            agent_exe_name = "CDO_Agente.exe"
            agent_script_name = "main.py"
            
            source_file = None
            is_script = False
            
            # Check source dir
            if os.path.exists(os.path.join(self.src_dir, agent_exe_name)):
                source_file = os.path.join(self.src_dir, agent_exe_name)
            elif os.path.exists(os.path.join(self.src_dir, agent_script_name)):
                source_file = os.path.join(self.src_dir, agent_script_name)
                is_script = True
            
            # If not found in user-provided source, try PyInstaller bundle or CWD
            if not source_file:
                if getattr(sys, 'frozen', False):
                    bundle_dir = sys._MEIPASS
                    if os.path.exists(os.path.join(bundle_dir, agent_exe_name)):
                        source_file = os.path.join(bundle_dir, agent_exe_name)
                        is_script = False
                else:
                    cwd = os.path.dirname(os.path.abspath(__file__))
                    if os.path.exists(os.path.join(cwd, agent_script_name)):
                        source_file = os.path.join(cwd, agent_script_name)
                        is_script = True

            if not source_file:
                raise Exception(f"No se encontró {agent_exe_name} ni {agent_script_name} en el origen.")

            self.progress['value'] = 40
            
            # 2. Preparar destino
            os.makedirs(self.dest_dir, exist_ok=True)
            dest_file = os.path.join(self.dest_dir, agent_exe_name if not is_script else agent_script_name)
            
            # Check if source == dest (In-place installation/update)
            is_inplace = os.path.abspath(source_file) == os.path.abspath(dest_file)
            
            self.log(f"Instalando en {self.dest_dir}...")
            
            # Stop existing process
            proc_name = agent_exe_name if not is_script else "python.exe" 
            if not is_script: # Only kill if exe, or if we are sure which python process (hard to know)
                subprocess.run(f'taskkill /F /IM "{agent_exe_name}"', shell=True, stderr=subprocess.DEVNULL, creationflags=subprocess.CREATE_NO_WINDOW if os.name=='nt' else 0)
            time.sleep(1)
            
            # Copy file (if not in-place)
            if not is_inplace:
                self.log("Copiando archivos...")
                shutil.copy2(source_file, dest_file)
                
                # If script, copy requirements too
                if is_script:
                    req_src = os.path.join(os.path.dirname(source_file), "requirements.txt")
                    if os.path.exists(req_src):
                        shutil.copy2(req_src, os.path.join(self.dest_dir, "requirements.txt"))
            else:
                self.log("Instalación en el lugar (sin copia)...")
            
            # Install dependencies if script
            if is_script:
                req_file = os.path.join(self.dest_dir, "requirements.txt")
                if os.path.exists(req_file):
                    self.log("Instalando dependencias...")
                    try:
                        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", req_file], 
                                            creationflags=subprocess.CREATE_NO_WINDOW if os.name=='nt' else 0)
                    except Exception as e:
                        self.log(f"Advertencia: Error instalando dependencias: {e}")

            self.progress['value'] = 70
            self.log("Creando configuración...")
            
            # Write config
            config_file = os.path.join(self.dest_dir, "agent_config.json")
            config = {
                "server_url": "http://3.138.135.181:8000",
                "username": self.username,
                "password": self.password,
                "task_url": "http://3.138.135.181:8000/tasks/poll",
                "result_url": "http://3.138.135.181:8000/tasks"
            }
            with open(config_file, "w") as f:
                json.dump(config, f, indent=4)
            
            self.progress['value'] = 80
            self.log("Configurando inicio automático...")
            
            # Registry Auto-Start
            try:
                key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
                cmd = f'"{dest_file}"'
                if is_script:
                    # Use pythonw to run without window
                    python_exe = sys.executable.replace("python.exe", "pythonw.exe")
                    cmd = f'"{python_exe}" "{dest_file}"'
                    
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE) as key:
                    winreg.SetValueEx(key, "CDO_Agente_Local", 0, winreg.REG_SZ, cmd)
            except Exception as e:
                print(f"Registry error: {e}")
                self.log("Advertencia: No se pudo configurar inicio automático en Registro.")
            
            self.progress['value'] = 90
            self.log("Iniciando agente...")
            
            if is_script:
                subprocess.Popen([sys.executable, dest_file], cwd=self.dest_dir, creationflags=subprocess.CREATE_NO_WINDOW if os.name=='nt' else 0)
            else:
                subprocess.Popen([dest_file], cwd=self.dest_dir, creationflags=subprocess.CREATE_NO_WINDOW if os.name=='nt' else 0)
            
            self.progress['value'] = 100
            self.log("¡Instalación Completada!")
            
            messagebox.showinfo("Éxito", f"El Agente Local se ha configurado en:\n{self.dest_dir}\n\nSe ejecutará automáticamente al iniciar sesión.")
            self.quit()
            
        except Exception as e:
            messagebox.showerror("Error Crítico", f"Ocurrió un error:\n{str(e)}")
            self.reset_ui()

    def reset_ui(self):
        self.install_btn.config(state="normal")
        self.exit_btn.config(state="normal")
        self.status_label.config(text="Error en la instalación.")
        self.progress['value'] = 0

if __name__ == "__main__":
    app = AgentInstaller()
    app.mainloop()
