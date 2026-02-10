
import tkinter as tk
from tkinter import ttk, messagebox
import os
import sys
import subprocess
import threading
import time
import shutil
import winreg

# --- CONFIG ---
# Determine if running as frozen (compiled exe) or script
if getattr(sys, 'frozen', False):
    BUNDLE_DIR = sys._MEIPASS
else:
    BUNDLE_DIR = os.path.dirname(os.path.abspath(__file__))

# Assets (bundled)
ASSETS_DIR = os.path.join(BUNDLE_DIR, "assets", "images")
LOGO_PATH = os.path.join(ASSETS_DIR, "CDO_logo.png")
SRC_SOURCE_DIR = os.path.join(BUNDLE_DIR, "src")
REQ_FILE_SOURCE = os.path.join(BUNDLE_DIR, "requirements.txt")

# Install Target (where files will be extracted)
# Default to LocalAppData/CDO_Organizer
INSTALL_DIR = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "CDO_Organizer")
VENV_DIR = os.path.join(INSTALL_DIR, "venv")
SRC_DEST_DIR = os.path.join(INSTALL_DIR, "src")

class SetupApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Instalador Cliente CDO")
        self.root.geometry("600x650") # Increased height
        self.root.resizable(False, False)
        
        # Style
        style = ttk.Style()
        style.theme_use('clam')
        
        # --- HEADER ---
        header_frame = tk.Frame(root, bg="white", height=100)
        header_frame.pack(side="top", fill="x")
        
        # Try to load Logo
        try:
            self.logo_img = tk.PhotoImage(file=LOGO_PATH)
            self.logo_img = self.logo_img.subsample(2, 2)
            lbl_logo = tk.Label(header_frame, image=self.logo_img, bg="white")
            lbl_logo.pack(pady=10)
        except Exception as e:
            lbl_title = tk.Label(header_frame, text="CDO ORGANIZER", font=("Arial", 20, "bold"), bg="white", fg="#333")
            lbl_title.pack(pady=30)

        # --- BUTTONS (Pack first to ensure visibility at bottom) ---
        btn_frame = tk.Frame(root, bg="#f0f0f0", height=60)
        btn_frame.pack(side="bottom", fill="x")
        
        self.btn_install = ttk.Button(btn_frame, text="Instalar y Configurar", command=self.start_install)
        self.btn_install.pack(side="right", padx=10, pady=15)
        
        self.btn_exit = ttk.Button(btn_frame, text="Salir", command=root.quit)
        self.btn_exit.pack(side="right", padx=10, pady=15)

        # --- CONTENT ---
        content_frame = tk.Frame(root, padx=40, pady=20)
        content_frame.pack(side="top", fill="both", expand=True)

        lbl_welcome = tk.Label(content_frame, text="Instalar CDO en PC Local", font=("Segoe UI", 16, "bold"))
        lbl_welcome.pack(anchor="w", pady=(0, 10))
        
        lbl_desc = tk.Label(content_frame, text=(
            "Esta instalación permite que el navegador acceda a sus carpetas locales\n"
            "para realizar búsquedas, conversiones y organizar archivos.\n"
        ), justify="left", font=("Segoe UI", 10))
        lbl_desc.pack(anchor="w", pady=(0, 10))
        
        # --- OPCIONES DE INSTALACIÓN ---
        options_frame = tk.LabelFrame(content_frame, text="Opciones de Configuración", padx=10, pady=10)
        options_frame.pack(fill="x", pady=10)
        
        # Email Input
        tk.Label(options_frame, text="Correo Electrónico (Para evitar alertas):").pack(anchor="w")
        self.email_var = tk.StringVar()
        entry_email = tk.Entry(options_frame, textvariable=self.email_var, width=40)
        entry_email.pack(anchor="w", pady=(0, 10))
        
        # Checkboxes
        self.install_python_var = tk.BooleanVar(value=False)
        self.install_docker_var = tk.BooleanVar(value=False)
        
        # Only show these if NOT running as frozen (standalone) to simplify UI for end users
        if not getattr(sys, 'frozen', False):
            chk_python = tk.Checkbutton(options_frame, text="Instalar Python (Requerido para ejecución manual)", variable=self.install_python_var)
            chk_python.pack(anchor="w")
            
            chk_docker = tk.Checkbutton(options_frame, text="Instalar Docker Desktop (Requerido para FEVRIPS)", variable=self.install_docker_var)
            chk_docker.pack(anchor="w")
        
        # --- PROGRESS ---
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(content_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill="x", pady=20)
        
        self.status_lbl = tk.Label(content_frame, text="Listo para instalar...", font=("Segoe UI", 9), fg="#666")
        self.status_lbl.pack(anchor="w")

    def log(self, text):
        self.root.after(0, lambda: self.status_lbl.config(text=text))

    def update_progress(self, val):
        self.root.after(0, lambda: self.progress_bar.config(value=val))

    def show_error(self, title, message):
        self.root.after(0, lambda: messagebox.showerror(title, message))
        
    def enable_buttons(self):
        self.root.after(0, lambda: self.btn_install.config(state="normal"))
        self.root.after(0, lambda: self.btn_exit.config(state="normal"))

    def start_install(self):
        # Validate Email
        if not self.email_var.get().strip():
            messagebox.showwarning("Falta Correo", "Por favor ingrese un correo electrónico para configurar la aplicación.")
            return

        self.btn_install.config(state="disabled")
        self.btn_exit.config(state="disabled")
        
        thread = threading.Thread(target=self.run_setup)
        thread.start()

    def run_setup(self):
        try:
            # 0. Prepare Directory
            self.log("Preparando directorio de instalación...")
            self.update_progress(5)
            os.makedirs(INSTALL_DIR, exist_ok=True)
            
            # --- GUARDAR EMAIL (credentials.toml) ---
            self.log("Configurando credenciales...")
            streamlit_conf_dir = os.path.join(INSTALL_DIR, ".streamlit")
            os.makedirs(streamlit_conf_dir, exist_ok=True)
            creds_path = os.path.join(streamlit_conf_dir, "credentials.toml")
            
            with open(creds_path, "w") as f:
                f.write('[general]\n')
                f.write(f'email = "{self.email_var.get().strip()}"\n')
            
            # --- INSTALAR PYTHON (Si se seleccionó) ---
            if self.install_python_var.get():
                self.log("Instalando Python (via Winget)...")
                try:
                    # Usamos winget para instalar Python 3.11
                    cmd_py = 'winget install -e --id Python.Python.3.11 --accept-source-agreements --accept-package-agreements'
                    subprocess.run(cmd_py, shell=True, check=True)
                    self.log("Python instalado correctamente.")
                except Exception as e:
                    print(f"Error instalando Python: {e}")
                    self.log("No se pudo instalar Python automáticamante.")
                
                # Instalar Dependencias (streamlit)
                self.log("Instalando dependencias (streamlit)...")
                try:
                    # Asumimos que python está en path tras instalar, o usamos 'py' launcher
                    subprocess.run("pip install streamlit pandas openpyxl requests google-generativeai watchdog pdfplumber pypdf2", shell=True)
                except Exception as e:
                    print(f"Error pip install: {e}")

            # --- INSTALAR DOCKER (Si se seleccionó) ---
            if self.install_docker_var.get():
                self.log("Verificando prerrequisitos de Docker (WSL2)...")
                
                # 1. Habilitar Características de Windows (WSL2)
                # Esto soluciona errores comunes como 0x800f081f al habilitar Hyper-V/WSL
                wsl_cmds = [
                    'dism.exe /online /enable-feature /featurename:Microsoft-Windows-Subsystem-Linux /all /norestart',
                    'dism.exe /online /enable-feature /featurename:VirtualMachinePlatform /all /norestart'
                ]
                
                features_enabled = True
                try:
                    for cmd in wsl_cmds:
                        # Usar PowerShell para ejecutar DISM con elevación implícita si el instalador ya es admin
                        subprocess.run(f"powershell -Command \"Start-Process -Verb RunAs -FilePath cmd.exe -ArgumentList '/c {cmd}' -Wait\"", shell=True, check=True)
                except Exception as wsl_err:
                    print(f"Advertencia WSL: {wsl_err}")
                    features_enabled = False
                    
                self.log("Instalando Docker Desktop (via Winget)...")
                try:
                    cmd_docker = 'winget install -e --id Docker.DockerDesktop --accept-source-agreements --accept-package-agreements'
                    subprocess.run(cmd_docker, shell=True, check=True)
                    self.log("Docker instalado. (Requiere reinicio)")
                except Exception as e:
                    err_msg = str(e)
                    print(f"Error instalando Docker: {err_msg}")
                    
                    if "0x800f081f" in err_msg or not features_enabled:
                        msg = "No se pudo instalar Docker debido a componentes de Windows faltantes.\n\nSolución:\n1. Ejecute 'Símbolo del sistema' como Administrador.\n2. Escriba: DISM /Online /Cleanup-Image /RestoreHealth\n3. Reinicie e intente de nuevo."
                        self.show_error("Error de Componentes Windows", msg)
                    else:
                        self.log("No se pudo instalar Docker automáticamente.")

            # Check for bundled Standalone EXE (CDO_Cliente.exe)
            bundled_client = os.path.join(BUNDLE_DIR, "CDO_Cliente.exe")
            bundled_agent = os.path.join(BUNDLE_DIR, "CDO_Agente.exe")
            use_standalone = os.path.exists(bundled_client)
            
            if use_standalone:
                self.log("Instalando Cliente Nativo...")
                self.update_progress(50)
                
                # Copy EXE
                target_exe = os.path.join(INSTALL_DIR, "CDO_Cliente.exe")
                shutil.copy2(bundled_client, target_exe)

                # Copy Agent EXE if exists
                if os.path.exists(bundled_agent):
                    self.log("Instalando Agente Local...")
                    target_agent = os.path.join(INSTALL_DIR, "CDO_Agente.exe")
                    shutil.copy2(bundled_agent, target_agent)
                
                # Copy Assets (Optional, but good for external access if needed)
                assets_dest = os.path.join(INSTALL_DIR, "assets")
                if os.path.exists(assets_dest): shutil.rmtree(assets_dest)
                shutil.copytree(os.path.join(BUNDLE_DIR, "assets"), assets_dest)
                
                self.update_progress(90)
                
                # Create Shortcut
                self.log("Creando accesos directos...")
                self.create_shortcuts(target_exe, use_bat=False)
                
                # Setup Agent (Registry - Primary Method)
                agent_exe = os.path.join(INSTALL_DIR, "CDO_Agente.exe")
                if os.path.exists(agent_exe):
                    self.log("Configurando Agente Local (Registro)...")
                    self.setup_agent_startup(agent_exe)
                    
                    # Start Agent Immediately
                    try:
                        self.log("Iniciando servicio de Agente...")
                        subprocess.Popen([agent_exe], cwd=INSTALL_DIR, creationflags=subprocess.CREATE_NO_WINDOW if os.name=='nt' else 0)
                    except Exception as e:
                        print(f"Error starting agent: {e}")
                
            else:
                # Fallback or Error?
                # If user just wanted Python/Docker, maybe we don't need the EXE?
                # But the app needs the EXE to run the GUI logic wrapped in Streamlit.
                # Assuming we proceed if EXE exists.
                self.show_error("Error Crítico", "No se encontró el archivo ejecutable 'CDO_Cliente.exe'.")
                self.enable_buttons()
                return

            self.update_progress(100)
            self.log("¡Instalación completada con éxito!")
            
            def finish_install():
                msg = f"El cliente local se ha instalado en:\n{INSTALL_DIR}\n\nSe ha creado un acceso directo 'CDO Organizer' en su Escritorio."
                if self.install_docker_var.get():
                    msg += "\n\nNOTA: Docker requiere cerrar sesión o reiniciar para finalizar su instalación."
                
                if messagebox.askyesno("Instalación Exitosa", msg + "\n\n¿Desea iniciar la aplicación ahora?"):
                    # Launch Client (Agent already started)
                    try:
                        target = os.path.join(INSTALL_DIR, "CDO_Cliente.exe") if use_standalone else os.path.join(INSTALL_DIR, "INICIAR_CDO.bat")
                        subprocess.Popen([target], shell=True, cwd=INSTALL_DIR)
                    except Exception as launch_err:
                        messagebox.showerror("Error al iniciar", f"No se pudo iniciar automáticamente: {launch_err}")
                self.root.quit()

            self.root.after(0, finish_install)
            
        except Exception as e:
            self.show_error("Error", f"Ocurrió un error:\n{str(e)}")
            self.log("Error en la instalación.")
            self.enable_buttons()

    def setup_agent_startup(self, agent_path):
        """Sets up the agent to run at startup via Registry (HKCU) and Shortcut."""
        # 1. Registry (Preferred for reliability)
        try:
            key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "CDO_Agente_Local", 0, winreg.REG_SZ, f'"{agent_path}"')
            print(f"Registry key set for {agent_path}")
            self.log("Inicio automático configurado (Registro).")
        except Exception as e:
            print(f"Error setting registry key: {e}")
            self.log(f"Advertencia: Falló registro ({e}), intentando acceso directo...")

        # 2. Startup Shortcut (Fallback/Redundant)
        try:
            startup_folder = os.path.join(os.getenv('APPDATA'), 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup')
            if not os.path.exists(startup_folder):
                os.makedirs(startup_folder, exist_ok=True)
                
            lnk_path = os.path.join(startup_folder, "CDO_Agente.lnk")
            working_dir = os.path.dirname(agent_path)
            
            ps_script = f"""
            $s=(New-Object -COM WScript.Shell).CreateShortcut('{lnk_path}');
            $s.TargetPath='{agent_path}';
            $s.WorkingDirectory='{working_dir}';
            $s.Save()
            """
            subprocess.run(["powershell", "-Command", ps_script], capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW if os.name=='nt' else 0)
        except Exception as e:
            print(f"Error creating Agent startup shortcut: {e}")
            self.log(f"Error: No se pudo crear inicio automático ({e})")

    def create_shortcuts(self, target_path, use_bat=False):
        # Create VBS script for silent launch if BAT
        if use_bat:
            vbs_path = os.path.join(INSTALL_DIR, "launch.vbs")
            with open(vbs_path, "w") as f:
                f.write('Set WshShell = CreateObject("WScript.Shell")\n')
                f.write(f'WshShell.Run chr(34) & "{target_path}" & Chr(34), 0\n')
                f.write('Set WshShell = Nothing\n')
        
        # Desktop Shortcut
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        if not os.path.exists(desktop):
            desktop = os.path.join(os.path.expanduser("~"), "Escritorio")
        
        if os.path.exists(desktop):
            lnk_path = os.path.join(desktop, "CDO Organizer.lnk")
            
            # PowerShell command to create shortcut
            # target_path is absolute
            icon_path = target_path if not use_bat else os.path.join(INSTALL_DIR, "assets", "CDO_logo.ico") # Assuming exe has icon
            
            ps_script = f"""
            $s=(New-Object -COM WScript.Shell).CreateShortcut('{lnk_path}');
            $s.TargetPath='{target_path}';
            $s.WorkingDirectory='{INSTALL_DIR}';
            $s.Save()
            """
            try:
                subprocess.run(["powershell", "-Command", ps_script], capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW if os.name=='nt' else 0)
            except Exception as e:
                print(f"Error creating LNK: {e}")
                # Fallback to BAT copy if LNK fails
                if use_bat:
                    shutil.copy2(target_path, os.path.join(desktop, "CDO Organizer.bat"))
        else:
            print("No Desktop found")


if __name__ == "__main__":
    root = tk.Tk()
    app = SetupApp(root)
    root.mainloop()
