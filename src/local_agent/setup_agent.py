import os
import sys
import shutil
import winreg
import subprocess
import time
from pathlib import Path

def install_agent():
    print("Iniciando instalación del Agente Local CDO...")
    
    # 1. Definir rutas
    # El ejecutable del agente debe estar en la misma carpeta que este instalador o empaquetado (MEIPASS)
    if getattr(sys, 'frozen', False):
        # Si corre como EXE (PyInstaller)
        current_dir = sys._MEIPASS
    else:
        # Si corre como script
        current_dir = os.path.dirname(os.path.abspath(__file__))
        
    agent_exe_name = "CDO_Agente.exe"
    agent_source = os.path.join(current_dir, agent_exe_name)
    
    # Fallback: Check local folder if not in MEIPASS (for dev/mixed scenarios)
    if not os.path.exists(agent_source):
        local_dir = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, 'frozen', False) else __file__))
        local_source = os.path.join(local_dir, agent_exe_name)
        if os.path.exists(local_source):
            agent_source = local_source
    
    if not os.path.exists(agent_source):
        print(f"Error: No se encontró {agent_exe_name} en {current_dir}")
        input("Presione ENTER para salir...")
        return

    # Ruta de destino: %APPDATA%\CDO_Agente
    appdata = os.environ.get("APPDATA")
    dest_dir = os.path.join(appdata, "CDO_Agente")
    dest_exe = os.path.join(dest_dir, agent_exe_name)
    
    print(f"Instalando en: {dest_dir}")
    
    try:
        os.makedirs(dest_dir, exist_ok=True)
        
        # Detener si ya está corriendo
        subprocess.run(f'taskkill /F /IM "{agent_exe_name}"', shell=True, stderr=subprocess.DEVNULL)
        time.sleep(1)
        
        # Copiar archivo
        shutil.copy2(agent_source, dest_exe)
        print("Archivos copiados.")
        
        # 2. Configurar inicio automático (Registry)
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
        print("Configurando inicio automático...")
        
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE) as key:
            winreg.SetValueEx(key, "CDO_Agente_Local", 0, winreg.REG_SZ, f'"{dest_exe}"')
            
        print("Registro actualizado.")
        
        # 3. Iniciar agente
        print("Iniciando agente...")
        subprocess.Popen([dest_exe], cwd=dest_dir, shell=False)
        
        # 4. Verificar inicio
        print("Verificando conexión...")
        time.sleep(2)
        import urllib.request
        try:
            with urllib.request.urlopen("http://localhost:8989/health", timeout=2) as response:
                if response.status == 200:
                    print("✅ AGENTE INICIADO Y RESPONDIENDO CORRECTAMENTE.")
                else:
                    print("⚠️ El agente inició pero respondió con código inesperado.")
        except Exception as e:
            print(f"⚠️ El agente se inició pero no responde al ping (puede tardar unos segundos). Error: {e}")

        print("\n¡Instalación completada con éxito!")
        print("El agente se está ejecutando en segundo plano (Puerto 8989).")
        print("Ya puede cerrar esta ventana.")
        
    except Exception as e:
        print(f"\n❌ Error durante la instalación: {e}")
        input("Presione ENTER para salir...")

if __name__ == "__main__":
    install_agent()
    time.sleep(5)
