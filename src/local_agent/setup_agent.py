import os
import sys
import shutil
import winreg
import subprocess
import time
import json
import tkinter as tk
from tkinter import messagebox, ttk
import threading

class AgentInstaller(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Instalador Agente Local CDO")
        self.geometry("400x300")
        self.resizable(False, False)
        
        # Center window
        self.eval('tk::PlaceWindow . center')
        
        self.label = tk.Label(self, text="Instalador Agente Local CDO", font=("Arial", 14, "bold"))
        self.label.pack(pady=10)
        
        self.info_label = tk.Label(self, text="Este asistente instalará el Agente Local\nnecesario para la conexión con archivos.", justify="center")
        self.info_label.pack(pady=5)
        
        self.progress = ttk.Progressbar(self, orient="horizontal", length=300, mode="determinate")
        self.progress.pack(pady=20)
        
        self.status_label = tk.Label(self, text="Listo para instalar...", fg="gray")
        self.status_label.pack(pady=5)
        
        # Credentials Inputs
        self.cred_frame = tk.Frame(self)
        self.cred_frame.pack(pady=10)
        
        tk.Label(self.cred_frame, text="Usuario:").grid(row=0, column=0, sticky="e")
        self.user_entry = tk.Entry(self.cred_frame)
        self.user_entry.insert(0, "admin")
        self.user_entry.grid(row=0, column=1, padx=5)
        
        tk.Label(self.cred_frame, text="Contraseña:").grid(row=1, column=0, sticky="e")
        self.pass_entry = tk.Entry(self.cred_frame, show="*")
        self.pass_entry.grid(row=1, column=1, padx=5)
        
        self.btn_frame = tk.Frame(self)
        self.btn_frame.pack(pady=20)
        
        self.install_btn = tk.Button(self.btn_frame, text="Instalar", command=self.start_install, bg="#007bff", fg="white", width=15)
        self.install_btn.pack(side="left", padx=5)
        
        self.exit_btn = tk.Button(self.btn_frame, text="Salir", command=self.quit, width=10)
        self.exit_btn.pack(side="right", padx=5)
        
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
        threading.Thread(target=self.run_installation, daemon=True).start()
        
    def run_installation(self):
        try:
            # 0. Validate Credentials
            self.log("Validando credenciales...")
            try:
                import requests
                auth = (self.username, self.password)
                url = "http://3.142.164.128:8000"
                res = requests.get(f"{url}/auth/verify", auth=auth, timeout=10)
                if res.status_code == 200:
                    self.log("✅ Credenciales válidas.")
                elif res.status_code == 401:
                    self.log("❌ Error: Usuario o contraseña incorrectos.")
                    messagebox.showerror("Error de Autenticación", "Usuario o contraseña incorrectos.")
                    self.install_btn.config(state="normal")
                    self.exit_btn.config(state="normal")
                    return
                else:
                    self.log(f"⚠️ Advertencia: No se pudo validar credenciales (HTTP {res.status_code}). Continuando...")
            except Exception as e:
                self.log(f"⚠️ Advertencia: Error conectando al servidor: {e}")

            self.progress['value'] = 0
            self.log("Buscando archivos...")
            
            # 1. Definir rutas
            if getattr(sys, 'frozen', False):
                current_dir = sys._MEIPASS
            else:
                current_dir = os.path.dirname(os.path.abspath(__file__))
                
            agent_exe_name = "CDO_Agente.exe"
            agent_source = os.path.join(current_dir, agent_exe_name)
            
            # Fallback dev check
            if not os.path.exists(agent_source):
                local_dir = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, 'frozen', False) else __file__))
                local_source = os.path.join(local_dir, agent_exe_name)
                if os.path.exists(local_source):
                    agent_source = local_source
            
            if not os.path.exists(agent_source):
                messagebox.showerror("Error", f"No se encontró el archivo {agent_exe_name}.")
                self.reset_ui()
                return

            self.progress['value'] = 20
            
            # Ruta de destino: %LOCALAPPDATA%\CDO_Organizer (Better than APPDATA/CDO_Agente for consistency)
            # Using LOCALAPPDATA to match the log fix I just made
            local_appdata = os.getenv('LOCALAPPDATA', os.path.expanduser("~"))
            dest_dir = os.path.join(local_appdata, "CDO_Organizer")
            dest_exe = os.path.join(dest_dir, agent_exe_name)
            
            self.log(f"Instalando en {dest_dir}...")
            os.makedirs(dest_dir, exist_ok=True)
            
            # Write config with user input
            config_file = os.path.join(dest_dir, "agent_config.json")
            config = {
                "server_url": "http://3.142.164.128:8000",
                "username": getattr(self, "username", "admin"),
                "password": getattr(self, "password", "password"),
                "task_url": "http://3.142.164.128:8000/tasks/poll",
                "result_url": "http://3.142.164.128:8000/tasks"
            }
            try:
                with open(config_file, "w") as f:
                    json.dump(config, f, indent=4)
                self.log("Configuración creada...")
            except Exception as e:
                print(f"Config write error: {e}")
            
            self.progress['value'] = 40
            
            # Stop existing
            subprocess.run(f'taskkill /F /IM "{agent_exe_name}"', shell=True, stderr=subprocess.DEVNULL, creationflags=subprocess.CREATE_NO_WINDOW if os.name=='nt' else 0)
            time.sleep(1)
            
            self.log("Copiando archivos...")
            shutil.copy2(agent_source, dest_exe)
            
            self.progress['value'] = 60
            self.log("Configurando inicio automático...")
            
            # Registry Auto-Start
            try:
                key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE) as key:
                    winreg.SetValueEx(key, "CDO_Agente_Local", 0, winreg.REG_SZ, f'"{dest_exe}"')
            except Exception as e:
                print(f"Registry error: {e}")
                self.log("Advertencia: No se pudo configurar inicio automático en Registro.")
            
            self.progress['value'] = 80
            self.log("Iniciando agente...")
            
            subprocess.Popen([dest_exe], cwd=dest_dir, creationflags=subprocess.CREATE_NO_WINDOW if os.name=='nt' else 0)
            
            self.progress['value'] = 100
            self.log("¡Instalación Completada!")
            
            messagebox.showinfo("Éxito", "El Agente Local se ha instalado y ejecutado correctamente.\n\nYa puede cerrar esta ventana.")
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
