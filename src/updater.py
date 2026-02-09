import os
import sys
import requests
import subprocess
import time
import json
try:
    from src.version import VERSION as CURRENT_VERSION
except ImportError:
    try:
        from version import VERSION as CURRENT_VERSION
    except ImportError:
        # Fallback if both fail (unlikely but safe)
        CURRENT_VERSION = "0.0.0"

def check_for_updates(server_url):
    """
    Checks for updates at the given server URL.
    Expected server structure:
    GET server_url/version.json -> {"version": "1.0.1", "url": "http://..."}
    """
    try:
        if not server_url:
            return False, None, "URL no configurada"
            
        if not server_url.endswith("/"): server_url += "/"
        
        target_url = server_url + "version.json"
        print(f"Checking update at: {target_url}")
        
        response = requests.get(target_url, timeout=5)
        response.raise_for_status()
        
        data = response.json()
        latest_version = data.get("version")
        download_url = data.get("url")
        
        print(f"Current: {CURRENT_VERSION}, Latest: {latest_version}")
        
        if latest_version != CURRENT_VERSION:
            return True, latest_version, download_url
            
        return False, latest_version, "Ya tienes la última versión."
        
    except requests.exceptions.ConnectionError:
        return False, None, f"No se pudo conectar al servidor de actualizaciones ({server_url}). Verifique que esté en ejecución."
    except Exception as e:
        return False, None, f"Error al buscar actualizaciones: {str(e)}"

def download_and_install(download_url):
    """
    Downloads the new executable and replaces the current one.
    """
    try:
        if not getattr(sys, 'frozen', False):
            return "No se puede actualizar en modo desarrollo (script python)."
            
        # 1. Download new exe
        print(f"Downloading from: {download_url}")
        r = requests.get(download_url, stream=True)
        r.raise_for_status()
        
        new_exe_path = "CDO_Cliente_New.exe"
        with open(new_exe_path, 'wb') as f:
            for chunk in r.iter_content(chunk_size=8192):
                f.write(chunk)
                
        print("Download complete. Preparing to update...")
        
        # 2. Create updater script (BAT)
        current_exe = sys.executable
        updater_script = "update_launcher.bat"
        
        # Batch script content
        # It waits 3 seconds, tries to delete the old exe (looping if locked),
        # moves the new exe, and restarts it.
        bat_content = f"""
@echo off
timeout /t 3 /nobreak > NUL
:loop
del "{current_exe}"
if exist "{current_exe}" (
    timeout /t 1 /nobreak > NUL
    goto loop
)
move "{new_exe_path}" "{current_exe}"
start "" "{current_exe}"
del "%~f0"
"""
        with open(updater_script, "w") as f:
            f.write(bat_content)
        
        # 3. Launch updater and exit
        subprocess.Popen([updater_script], shell=True)
        
        # Return specific code to signal main app to exit
        return "UPDATE_INITIATED"
        
    except Exception as e:
        return str(e)
