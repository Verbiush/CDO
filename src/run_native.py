import os
import sys
import socket
import traceback
import shutil
from datetime import datetime
import ctypes

# Setup logging function
def log_debug(msg):
    try:
        log_dir = os.path.join(os.path.expanduser("~"), "AppData", "Local", "CDO_Organizer")
        os.makedirs(log_dir, exist_ok=True)
        log_file = os.path.join(log_dir, "debug_log.txt")
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(log_file, "a") as f:
            f.write(f"[{timestamp}] {msg}\n")
    except:
        pass

def show_error(title, msg):
    try:
        ctypes.windll.user32.MessageBoxW(0, msg, title, 0x10)
    except:
        pass

try:
    from streamlit.web import cli as stcli
    log_debug("Successfully imported streamlit.web.cli")
except Exception as e:
    err_msg = f"Failed to import streamlit: {e}\n{traceback.format_exc()}"
    log_debug(err_msg)
    show_error("Error de Inicio", f"No se pudo iniciar el motor de la aplicación.\n\n{e}")
    sys.exit(1)

def get_version(dir_path):
    try:
        v_path = os.path.join(dir_path, "version.py")
        if not os.path.exists(v_path):
            return None
        with open(v_path, "r", encoding="utf-8") as f:
            content = f.read()
            import re
            m = re.search(r'VERSION\s*=\s*["\']([^"\']+)["\']', content)
            if m:
                return m.group(1)
    except:
        pass
    return None

def deploy_to_appdata(src_dir):
    """
    Copies the application source code to a persistent location in AppData.
    This ensures standard file paths, writability for session files, and correct relative imports.
    """
    app_data_dir = os.path.join(os.path.expanduser("~"), "AppData", "Local", "CDO_Organizer", "app")
    
    log_debug(f"Deploying app to: {app_data_dir}")
    
    if not os.path.exists(app_data_dir):
        os.makedirs(app_data_dir)

    # --- VERSION CHECK OPTIMIZATION ---
    # If the destination exists and has the same version, skip copying to save time and preserve state.
    src_ver = get_version(src_dir)
    dst_ver = get_version(app_data_dir)
    
    log_debug(f"Source Version: {src_ver}, Dest Version: {dst_ver}")

    if src_ver and dst_ver and src_ver == dst_ver:
        # Verify critical file exists
        if os.path.exists(os.path.join(app_data_dir, "app_web.py")):
            log_debug("Versions match and app exists. Skipping copy.")
            return app_data_dir
            
    log_debug("Versions differ or destination missing. Proceeding with copy...")

    # Copy all files from src_dir to app_data_dir
    # We overwrite .py files to ensure we run the version from the EXE
    # We preserve .json files (config, session) if they exist
    
    try:
        for item in os.listdir(src_dir):
            s = os.path.join(src_dir, item)
            d = os.path.join(app_data_dir, item)
            
            if os.path.isdir(s):
                # For directories (like 'pages' or 'assets'), we copy recursively
                # For simplicity, we can remove and recopy, OR merge.
                # Let's remove and recopy code folders to ensure updates.
                # But 'temp_sessions' should be preserved or cleaned.
                if item in ["__pycache__", "temp_sessions", "temp_uploads"]:
                    continue
                    
                if os.path.exists(d):
                    # Check if it's a data folder? Assuming code folders for now.
                    shutil.rmtree(d)
                shutil.copytree(s, d)
            else:
                # File
                # Don't overwrite config/session files
                if item.endswith(".json") and os.path.exists(d):
                    continue
                shutil.copy2(s, d)
                
        log_debug("Deployment successful")
        return app_data_dir
    except Exception as e:
        log_debug(f"Error deploying files: {e}")
        # Fallback to running in place if copy fails
        return src_dir

if __name__ == '__main__':
    try:
        log_debug("Starting CDO Client (Persistent Mode)...")
        
        # Determine source directory (where the code is NOW)
        if getattr(sys, 'frozen', False):
            # Running from PyInstaller bundle
            # The src folder was added via --add-data "clean_src_dir;src"
            # So it should be at sys._MEIPASS/src
            base_src_dir = os.path.join(sys._MEIPASS, "src")
        else:
            # Running from source
            base_src_dir = os.path.dirname(os.path.abspath(__file__))

        log_debug(f"Base Source Dir: {base_src_dir}")
        
        # Deploy to AppData (Persistent Location)
        # This solves the "404" (standard paths) and "Persistence" (writable dir)
        working_dir = deploy_to_appdata(base_src_dir)
        
        # Target App File
        app_path = os.path.join(working_dir, "app_web.py")
        
        if not os.path.exists(app_path):
             log_debug(f"CRITICAL: app_web.py not found at {app_path}")
             show_error("Error Fatal", f"No se encontró el archivo principal:\n{app_path}")
             sys.exit(1)

        # Change CWD to the persistent app directory
        os.chdir(working_dir)
        log_debug(f"Changed CWD to: {working_dir}")
        
        # Add working_dir to sys.path so imports work
        sys.path.insert(0, working_dir)

        # --- FIX: Create .streamlit/config.toml to suppress Email prompt ---
        streamlit_config_dir = os.path.join(working_dir, ".streamlit")
        os.makedirs(streamlit_config_dir, exist_ok=True)
        config_toml_path = os.path.join(streamlit_config_dir, "config.toml")
        credentials_toml_path = os.path.join(streamlit_config_dir, "credentials.toml")
        
        # Crear config.toml si no existe
        if not os.path.exists(config_toml_path):
            try:
                with open(config_toml_path, "w") as f:
                    f.write('[browser]\ngatherUsageStats = false\n\n')
                    f.write('[server]\nfileWatcherType = "none"\nrunOnSave = false\nheadless = true\n')
                    # No forzamos el tema para que el usuario pueda elegirlo
                    # f.write('[theme]\nbase = "dark"\n')
                log_debug(f"Created config.toml at {config_toml_path}")
            except Exception as e:
                log_debug(f"Failed to create config.toml: {e}")
        else:
            log_debug(f"config.toml already exists at {config_toml_path}")

        # Create credentials.toml ONLY if it doesn't exist to respect installer settings
        try:
            if not os.path.exists(credentials_toml_path):
                with open(credentials_toml_path, "w") as f:
                    f.write('[general]\nemail = ""\n')
                log_debug(f"Created credentials.toml at {credentials_toml_path}")
            else:
                log_debug(f"credentials.toml exists at {credentials_toml_path}, skipping overwrite.")
        except Exception as e:
            log_debug(f"Error creating credentials: {e}")

        # --- FIX: Set Environment Variables as backup ---
        os.environ["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"
        os.environ["STREAMLIT_SERVER_FILE_WATCHER_TYPE"] = "none"
        # Force usage of OUR config file, not the bundled one
        os.environ["STREAMLIT_CONFIG_FILE"] = config_toml_path
        os.environ["STREAMLIT_CREDENTIALS_FILE"] = credentials_toml_path

        # Smart headless mode
        # If frozen, we usually want the browser to open (headless=false). 
        # BUT if we are wrapping it in a webview later, or if we want to suppress the "new tab", we check.
        # For this app, we want the default browser to open.
        is_frozen = getattr(sys, 'frozen', False)
        headless = "false" if is_frozen else "true"

        # Dynamic Port Selection to allow multiple instances
        # We start at a random port in range to minimize collisions
        def find_free_port(start_port=8501, max_port=8600):
            import socket
            import random
            
            # Try random ports first to avoid sequential collisions
            attempts = list(range(start_port, max_port))
            random.shuffle(attempts)
            
            for p in attempts:
                with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                    try:
                        s.bind(("localhost", p))
                        return p
                    except OSError:
                        continue
            return start_port # Fallback

        port = find_free_port(8501)
            
        log_debug(f"Selected port: {port}")
        
        # --- PATH FIX: Ensure Streamlit can find its static assets ---
        if getattr(sys, 'frozen', False):
            import streamlit
            st_dir = os.path.dirname(streamlit.__file__)
            log_debug(f"Streamlit imported from: {st_dir}")
            
            # Check for static files
            static_path = os.path.join(st_dir, "static")
            if not os.path.exists(static_path):
                log_debug(f"WARNING: Static path not found at {static_path}")
                # Try to locate it in MEIPASS
                meipass_static = os.path.join(sys._MEIPASS, "streamlit", "static")
                if os.path.exists(meipass_static):
                     log_debug(f"Found static files in MEIPASS: {meipass_static}")
                     # Hack: If Streamlit expects it relative to __file__, and __file__ is in a zip...
                     # We can try to set an env var if Streamlit supports it, but it doesn't.
                     # However, if we copied the whole streamlit package to AppData, it would work.
                     # For now, let's assume the 404 was due to CWD/AppPath issues which we fixed.
        # ------------------------------------------------------------

        # Construct arguments for streamlit
        # NOTE: We use "run app_web.py" relative to CWD, which is now working_dir.
        sys.argv = [
            "streamlit",
            "run",
            "app_web.py", # Relative path, since we chdir'd
            "--global.developmentMode=false",
            f"--server.headless={headless}",
            "--browser.gatherUsageStats=false",
            "--server.address=127.0.0.1",
            f"--server.port={port}",
            "--server.fileWatcherType=none",
            "--server.runOnSave=false",
            "--theme.base=dark" 
        ]

        log_debug(f"Sys.argv: {sys.argv}")
        
        # Redirect stdout/stderr
        sys.stdout = open(os.path.join(os.path.expanduser("~"), "AppData", "Local", "CDO_Organizer", "console_output.txt"), "a")
        sys.stderr = sys.stdout
        
        # Run Streamlit
        log_debug("Calling stcli.main()...")
        sys.stdout.flush()
        
        try:
            stcli.main()
        except SystemExit as se:
            log_debug(f"Streamlit exited with code: {se.code}")
            if se.code != 0:
                # Try to read the console output to show the error
                try:
                    console_log = os.path.join(os.path.expanduser("~"), "AppData", "Local", "CDO_Organizer", "console_output.txt")
                    if os.path.exists(console_log):
                        with open(console_log, "r", encoding="utf-8", errors="ignore") as f:
                            lines = f.readlines()
                            last_lines = "".join(lines[-20:]) # Get last 20 lines
                            show_error("Error de Ejecución", f"Streamlit se cerró con código {se.code}.\n\nÚltimos logs:\n{last_lines}")
                except:
                    pass
        except BaseException as e:
            log_debug(f"Streamlit crashed: {e}\n{traceback.format_exc()}")
            
    except Exception as e:
        err_msg = f"CRITICAL ERROR MAIN: {e}\n{traceback.format_exc()}"
        log_debug(err_msg)
        show_error("Error Crítico", f"La aplicación se cerró inesperadamente.\n\n{e}")
        sys.exit(1)
