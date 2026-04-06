import streamlit as st
import os
import time
import shutil

# --- AUTO CLEANUP TEMPORARY FILES (BACKGROUND) ---
# Se ejecuta al cargar la aplicación para limpiar basura antigua de otras sesiones
def auto_cleanup_temp_dirs(max_age_minutes=60):
    now = time.time()
    for folder in ["temp_uploads", "temp_downloads"]:
        if os.path.exists(folder):
            for item in os.listdir(folder):
                item_path = os.path.join(folder, item)
                try:
                    # If older than max_age_minutes
                    if os.path.getmtime(item_path) < now - (max_age_minutes * 60):
                        if os.path.isdir(item_path):
                            shutil.rmtree(item_path, ignore_errors=True)
                        else:
                            os.remove(item_path)
                except Exception:
                    pass

# Limpiar archivos más antiguos de 60 minutos silenciosamente al iniciar
auto_cleanup_temp_dirs(60)

# --- CONFIGURACIÓN INICIAL DEL ESTADO ---
# En AWS queremos que se comporte como la APP NATIVA (usando el Agente para diálogos)
if "force_native_mode" not in st.session_state:
    st.session_state["force_native_mode"] = True

import pandas as pd
import datetime
import time
import os
import sys
import io
import zipfile
import shutil
import warnings
# Suppress Google Generative AI FutureWarning
warnings.filterwarnings("ignore", category=FutureWarning, module="google.generativeai")
import unicodedata
import json
import base64
import threading
try:
    from task_manager import render_task_center, show_task_notifications, render_task_items
except ImportError:
    from src.task_manager import render_task_center, show_task_notifications, render_task_items

try:
    from gui_utils import render_path_selector
except ImportError:
    try:
        from src.gui_utils import render_path_selector
    except ImportError:
        def render_path_selector(label, key, default_path=None, help_text=None, omit_checkbox=False):
            st.error("render_path_selector no disponible")
            return default_path

try:
    from tabs import tab_bot_zeus, tab_ai_assistant, tab_search_actions, tab_automated_actions
    from tabs import tab_conversion, tab_visor, tab_rips, tab_validator_fevrips, tab_user_validation, tab_gestion_documental, tab_admin, tab_user_management
except ImportError as e:
    # Only fallback if the error is about finding 'tabs' module itself
    if e.name == 'tabs' or "No module named 'tabs'" in str(e):
        from src.tabs import tab_bot_zeus, tab_ai_assistant, tab_search_actions, tab_automated_actions
        from src.tabs import tab_conversion, tab_visor, tab_rips, tab_validator_fevrips, tab_user_validation, tab_gestion_documental, tab_admin, tab_user_management
    else:
        raise e




# --- AGENT HELPERS ---
def create_standalone_agent_zip():
    """
    Creates a ZIP file containing the standalone agent installer or executable.
    Returns (zip_bytes, found, type_found)
    type_found can be 'installer' or 'portable' or None
    """
    import zipfile
    import io
    import os

    output = io.BytesIO()
    found = False
    type_found = None
    
    # Priority 1: Installer (Instalar_Agente.exe)
    possible_installers = [
        "Instalar_Agente.exe",
        os.path.join("build_release_output", "build_temp", "Instalar_Agente.exe"),
        os.path.join("..", "Instalar_Agente.exe"),
        os.path.join(os.path.dirname(__file__), "..", "Instalar_Agente.exe"),
        # Absolute paths for dev env
        os.path.join(os.path.dirname(__file__), "..", "..", "build_release_output", "build_temp", "Instalar_Agente.exe")
    ]
    
    installer_path = None
    for p in possible_installers:
        if os.path.exists(p):
            installer_path = p
            break
            
    if installer_path:
        with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as z:
            z.write(installer_path, "Instalar_Agente.exe")
            z.writestr("LEEME.txt", "Ejecute Instalar_Agente.exe para configurar el servicio local.")
        found = True
        type_found = "installer"
        return output.getvalue(), found, type_found

    # Priority 2: Portable Agent (CDO_Agente.exe)
    possible_exes = [
        "CDO_Agente.exe",
        os.path.join("build_release_output", "build_agent", "CDO_Agente.exe"),
        os.path.join("src", "local_agent", "CDO_Agente.exe"),
        os.path.join(os.path.dirname(__file__), "local_agent", "CDO_Agente.exe"),
        os.path.join(os.path.dirname(__file__), "..", "..", "build_release_output", "build_agent", "CDO_Agente.exe")
    ]
    
    exe_path = None
    for p in possible_exes:
        if os.path.exists(p):
            exe_path = p
            break
            
    if exe_path:
        with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as z:
            z.write(exe_path, "CDO_Agente.exe")
            z.writestr("iniciar_agente.bat", 'start "" "CDO_Agente.exe"')
            z.writestr("LEEME.txt", "Este es el agente portable. Ejecute iniciar_agente.bat.")
        found = True
        type_found = "portable"
        return output.getvalue(), found, type_found

    return None, False, None

def create_cdo_agent_zip():
    """
    Creates a full installer ZIP for the CDO Local Agent.
    """
    import zipfile
    import io
    import os
    
    output = io.BytesIO()
    base_dir = os.path.dirname(os.path.abspath(__file__)) # src/
    
    # Define files to include
    # (source_path_rel_to_src, archive_path)
    files_to_zip = [
        ("local_agent/install_agent.ps1", "install_agent.ps1"),
        ("local_agent/main.py", "src/local_agent/main.py"),
        ("local_agent/cert_gen.py", "src/local_agent/cert_gen.py"),
        ("local_agent/README.md", "src/local_agent/README.md"),
        ("bot_zeus.py", "src/bot_zeus.py")
    ]
    
    with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as z:
        for src_rel, archive_path in files_to_zip:
            full_path = os.path.join(base_dir, src_rel)
            if not os.path.exists(full_path):
                # Fallback for dev environment where src might be parent
                # If app_web.py is in src, base_dir is src.
                # src_rel is local_agent/...
                pass
                
            if os.path.exists(full_path):
                z.write(full_path, archive_path)
            else:
                # Try absolute path fallback if base_dir is wrong
                pass
                
        # Add a README at root
        readme_root = """# Instalación Agente CDO

1. Descomprima este archivo ZIP en una carpeta permanente (ej: C:\CDO_Agent).
2. Abra PowerShell como Administrador.
3. Navegue a la carpeta descomprimida:
   cd C:\CDO_Agent
4. Ejecute el instalador:
   Set-ExecutionPolicy Bypass -Scope Process -Force; ./install_agent.ps1

El agente se instalará como un servicio de inicio y se ejecutará en segundo plano.
"""
        z.writestr("LEEME.txt", readme_root)

    output.seek(0)
    return output.getvalue(), "CDO_Agente_Installer.zip"

def create_lightweight_agent_zip():
    # Legacy wrapper
    content, name = create_cdo_agent_zip()
    return content

# --- DATABASE / USER MANAGEMENT ---
try:
    import database as db
except ImportError:
    from src import database as db

# --- STREAMLIT PAGE CONFIG ---
try:
    st.set_page_config(
        page_title="CDO Clinical Document Organizer", 
        layout="wide", 
        page_icon="📂",
        initial_sidebar_state="expanded"
    )
    # Check for background task notifications immediately after page load
    show_task_notifications()
except Exception as e:
    # Log startup error if set_page_config fails
    try:
        with open(os.path.join(os.path.expanduser("~"), "AppData", "Local", "CDO_Organizer", "startup_error.txt"), "w") as f:
            f.write(f"Startup Error: {str(e)}")
    except:
        pass
    # Don't raise here to allow script to continue if possible, or handle gracefully
    # raise e

# --- NOTIFICATION BELL ---
# Layout for Header with Bell
col_header_title, col_header_bell = st.columns([0.95, 0.05])
with col_header_bell:
    if st.session_state.get("logged_in", False):
        # Use popover if available (Streamlit 1.33+)
        try:
            popover = st.popover("🔔", help="Centro de Tareas")
            popover.markdown("### 🔔 Centro de Tareas")
            render_task_center(popover)
        except AttributeError:
            # Fallback for older Streamlit
            if st.button("🔔"):
                st.toast("Revisa el Centro de Tareas en la barra lateral.")

# --- ADMIN PANEL ---
# Admin Panel logic moved to src/tabs/tab_admin.py

# --- USER PREFERENCES ---
# Use absolute path to ensure persistence across different working directories
# Check if /data volume is available and writable (Docker/AWS environment)
if os.path.exists("/data") and os.access("/data", os.W_OK):
    USER_PREFS_FILE = "/data/user_preferences.json"
else:
    USER_PREFS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "user_preferences.json")

def load_user_prefs():
    if os.path.exists(USER_PREFS_FILE):
        try:
            with open(USER_PREFS_FILE, "r") as f:
                return json.load(f)
        except:
            return {}
    return {}

def save_user_prefs(data):
    try:
        current = load_user_prefs()
        current.update(data)
        with open(USER_PREFS_FILE, "w") as f:
            json.dump(current, f)
        return True
    except Exception as e:
        print(f"Error saving prefs: {e}")
        return False

# --- AUTHENTICATION ---
if "logged_in" not in st.session_state:
    # Try to load from prefs first
    prefs = load_user_prefs()
    
    st.session_state.logged_in = False
    st.session_state.username = None
    st.session_state.role = None
    st.session_state.permissions = {}
    
    # Load saved mode preference
    saved_mode = prefs.get("force_native_mode", True)
    st.session_state.force_native_mode = saved_mode

def login_page():
    if st.session_state.get("logged_in", False):
        st.rerun()
        return

    # Use a container to hold the login page content
    login_container = st.container()
    
    with login_container:
        st.title("🔐 Iniciar Sesión")
        st.markdown("Versión Nube v2.7 (Modo Nativo)")
        
        col1, col2 = st.columns([1, 2])
        with col1:
            # Resolve logo path relative to this script
            base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            logo_path = os.path.join(base_dir, "assets", "images", "CDO_logo.png")
            if os.path.exists(logo_path):
                st.image(logo_path, width=280)
            else:
                st.warning("Logo no encontrado")
        with col2:
            prefs = load_user_prefs()
            last_user = prefs.get("last_username", "") if prefs.get("remember_me", False) else ""
            
            username = st.text_input("Usuario", value=last_user)
            password = st.text_input("Contraseña", type="password")
            
            # Mode Selection at Login
            st.markdown("##### Configuración de Entorno")
            mode_options = ["Nativo (Local)", "Web (Nube)"]
            default_idx = 0 if st.session_state.get("force_native_mode", True) else 1
            
            selected_mode = st.radio(
               "Modo de Operación:",
               mode_options,
               index=default_idx,
               horizontal=True,
               help="Seleccione 'Nativo' si está ejecutando la aplicación en su propia máquina. 'Web' deshabilita diálogos del sistema."
            )
            # selected_mode = "Web (Nube)"
            
            remember = st.checkbox("Recordar mi usuario", value=prefs.get("remember_me", False))
            
            if st.button("Ingresar", type="primary", key="btn_login_submit"):
                user_data = db.check_login(username, password)
                if user_data:
                    st.session_state.logged_in = True
                    
                    # Fallback mechanism if database module is cached and returns bool
                    if isinstance(user_data, bool):
                        user_data = db.get_user(username)

                    st.session_state.username = username
                    st.session_state.role = user_data.get("role", "user")
                    st.session_state.permissions = user_data.get("permissions", {})
                    
                    # Load App Config
                    user_config = user_data.get("config", {})
                    st.session_state.app_config = user_config.get("app_config", {})
                    
                    # Update Session Mode
                    is_native = (selected_mode == "Nativo (Local)")
                    st.session_state.force_native_mode = is_native
                    
                    # Save Prefs
                    save_data = {
                        "remember_me": remember,
                        "force_native_mode": is_native
                    }
                    if remember:
                        save_data["last_username"] = username
                    else:
                        save_data["last_username"] = ""
                        
                    save_user_prefs(save_data)
                    
                    # Clear the login container explicitly before rerun
                    login_container.empty()
                    st.rerun()
                else:
                    st.error("Usuario o contraseña incorrectos")

def logout():
    # Clear all session state to ensure a clean logout
    st.session_state.clear()
    st.rerun()

# --- MAIN APP LOGIC ---
if not st.session_state.logged_in:
    login_page()
else:
    # Ensure app_config is loaded for existing sessions
    if "app_config" not in st.session_state:
        st.session_state.app_config = {}
        if st.session_state.username:
            try:
                user_data = db.get_user(st.session_state.username)
                if user_data:
                    user_config = user_data.get("config", {})
                    st.session_state.app_config = user_config.get("app_config", {})
            except:
                pass

    # Sidebar Info
    with st.sidebar:
        # Resolve logo path
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        logo_path = os.path.join(base_dir, "assets", "images", "CDO_logo.png")
        if os.path.exists(logo_path):
            st.image(logo_path, width=100)
        st.write(f"👤 **{st.session_state.username}** ({st.session_state.role})")
        
        if st.button("Cerrar Sesión"):
            logout()
            
        st.divider()
        
        # Configuración General
        with st.expander("⚙️ Configuración General"):
            tab_gen, tab_proc, tab_ui, tab_sys = st.tabs(["General", "Procesamiento", "Interfaz", "Sistema"])
            
            with tab_gen:
                st.markdown("### Parámetros de Ejecución")
                
                # Mode Selection
                if "force_native_mode" not in st.session_state:
                    st.session_state.force_native_mode = True # Default to Native if not set
                
                def on_mode_change():
                    val = st.session_state.sb_mode_selector
                    new_mode = (val == "Nativo (Local)")
                    st.session_state.force_native_mode = new_mode
                    print(f"DEBUG: Mode changed to {'Native' if new_mode else 'Web'}")
                    
                    # Save preference immediately
                    if save_user_prefs({"force_native_mode": new_mode}):
                         msg = "Modo Nativo Activado" if new_mode else "Modo Web Activado"
                         st.toast(msg, icon="✅")
                    else:
                         st.toast("Advertencia: No se pudo guardar la preferencia en disco.", icon="⚠️")
                    
                    # Force rerun to ensure all UI components update their state immediately
                    # This is crucial for switching between Native and Web rendering in tabs
                    time.sleep(0.5)
                    st.rerun()
                
                current_mode_idx = 0 if st.session_state.force_native_mode else 1
                mode_options = ["Nativo (Local)", "Web (Nube)"]
                
                st.radio(
                    "Modo de Operación:",
                    mode_options,
                    index=current_mode_idx,
                    key="sb_mode_selector",
                    on_change=on_mode_change,
                    help="Seleccione 'Nativo' si está ejecutando la aplicación en su propia máquina."
                )
                
                if st.session_state.force_native_mode:
                    st.success("✅ Modo Nativo Activo")
                else:
                    st.info("☁️ Modo Web Activo")
                    # st.markdown("##### 🔌 Agente Local")
                    # st.caption("Para habilitar funciones nativas en Modo Web, instale el Agente Local.")
                    try:
                        # Ocultar descarga del agente en modo web por peticion del usuario
                        pass
                        # zip_bytes, filename = create_cdo_agent_zip()
                        # st.download_button(
                        #    label="⬇️ Descargar Instalador",
                        #    data=zip_bytes,
                        #    file_name=filename,
                        #    mime="application/zip",
                        #    help="Descargue el instalador para conectar su PC.",
                        #    key="btn_download_agent_sidebar"
                        # )
                    except Exception as e:
                        pass
                        # st.error(f"Error generando instalador: {e}")
                
                st.markdown("---")
                st.markdown("### Configuración de IA (Gemini)")
                
                # Check current API Key
                current_api_key = st.session_state.app_config.get("gemini_api_key", "")
                
                # Model Selection
                current_model = st.session_state.app_config.get("gemini_model", "gemini-flash-latest")
                model_options = ["gemini-flash-latest", "gemini-1.5-flash", "gemini-2.0-flash", "gemini-pro"]
                model_labels = {
                    "gemini-flash-latest": "Gemini Flash (Versión Estable - Recomendado)",
                    "gemini-1.5-flash": "Gemini 1.5 Flash (Gratuito/Rápido)",
                    "gemini-2.0-flash": "Gemini 2.0 Flash (Experimental)",
                    "gemini-pro": "Gemini 1.0 Pro (Estándar)"
                }
                
                selected_model = st.selectbox(
                    "Modelo de IA",
                    model_options,
                    index=model_options.index(current_model) if current_model in model_options else 0,
                    format_func=lambda x: model_labels.get(x, x),
                    help="Seleccione el modelo. 'Flash' es ideal para la capa gratuita."
                )

                new_api_key = st.text_input(
                    "Google Gemini API Key", 
                    value=current_api_key, 
                    type="password",
                    help="Clave de API para el asistente de IA."
                )
                
                # Check changes
                key_changed = (new_api_key != current_api_key)
                model_changed = (selected_model != current_model)
                
                if key_changed or model_changed:
                    if st.button("Guardar Configuración IA"):
                        st.session_state.app_config["gemini_api_key"] = new_api_key
                        st.session_state.app_config["gemini_model"] = selected_model
                        
                        try:
                            # Fetch full user config first to avoid overwriting other settings
                            user_data = db.get_user(st.session_state.username)
                            current_config = user_data.get("config", {}) if user_data else {}
                            
                            if "app_config" not in current_config:
                                current_config["app_config"] = {}
                                
                            current_config["app_config"]["gemini_api_key"] = new_api_key
                            current_config["app_config"]["gemini_model"] = selected_model
                            
                            db.update_user_config(st.session_state.username, current_config)
                            st.success("✅ Configuración de IA guardada exitosamente!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error guardando configuración: {e}")

            with tab_proc:
                st.caption("Configuraciones de procesamiento de archivos.")
                st.checkbox("Validación estricta de RIPS", value=True, key="cfg_strict_rips")
                
                col_p1, col_p2 = st.columns(2)
                with col_p1:
                    st.text_input("Prefijo PDF Unificado", value="Unificado", key="cfg_pdf_prefix")
                    st.text_input("Prefijo FEOV", value="FEOV", key="cfg_feov_prefix")
                    st.text_input("Nombre DOCX Salida", value="Documento", key="cfg_docx_name")
                with col_p2:
                    st.number_input("DPI Imágenes", value=150, step=10, key="cfg_dpi")
                    st.selectbox("Compresión PDF", ["Baja", "Media", "Alta"], index=1, key="cfg_compression")
                
                st.text_area("Patrones de Exclusión (separados por coma)", value="Thumbs.db, .DS_Store", key="cfg_exclusion_patterns", help="Archivos a ignorar durante el procesamiento.")
                
                st.number_input("Timeout de Bots (segundos)", value=30, min_value=10, max_value=300, key="cfg_bot_timeout")
                render_path_selector(
                    label="Carpeta de Descargas (Opcional)",
                    key="cfg_download_path",
                    help_text="Ruta local para guardar reportes automáticamente."
                )

            with tab_ui:
                st.caption("Personalización de la interfaz.")
                st.checkbox("Mostrar notificaciones de escritorio", value=False, key="cfg_desktop_notif")
                st.selectbox("Tema Visual", ["Claro", "Oscuro", "Sistema"], index=2, key="cfg_theme")
                
            with tab_sys:
                st.caption("Información del Sistema")
                st.text_input("Ruta Base de Datos", value=os.path.join(os.path.dirname(os.path.abspath(__file__)), "users.db"), disabled=True)
                st.text_input("Directorio de Trabajo", value=os.getcwd(), disabled=True)
                if st.button("Limpiar Caché de Aplicación"):
                    st.cache_data.clear()
                    st.cache_resource.clear()
                    
                    # Cleanup temp folders
                    import shutil
                    import time
                    try:
                        # Prevent deletion of recent files (e.g. less than 5 min old) to avoid breaking active sessions
                        now = time.time()
                        def clean_old_files(folder_path):
                            if os.path.exists(folder_path):
                                for item in os.listdir(folder_path):
                                    item_path = os.path.join(folder_path, item)
                                    try:
                                        if os.path.getmtime(item_path) < now - 300: # 5 minutes old
                                            if os.path.isdir(item_path):
                                                shutil.rmtree(item_path)
                                            else:
                                                os.remove(item_path)
                                    except Exception as e:
                                        print(f"Error borrando {item_path}: {e}")
                        
                        clean_old_files("temp_uploads")
                        clean_old_files("temp_downloads")
                        
                        os.makedirs("temp_uploads", exist_ok=True)
                        os.makedirs("temp_downloads", exist_ok=True)
                        st.success("Caché y archivos temporales antiguos limpiados exitosamente.")
                    except Exception as e:
                        st.error(f"Error limpiando archivos temporales: {e}")

        # Task Center in Sidebar
        with st.expander("📋 Centro de Tareas", expanded=False):
             render_task_center(st, key_prefix="sidebar_exp_")
    
    # Check permissions for Tabs
    user_perms = st.session_state.permissions
    allowed_tabs_config = user_perms.get("allowed_tabs", ["*"])
    
    # Define all possible tabs and their render functions
    # Using a list of tuples to maintain order: (Tab Name, Render Function)
    
    all_tabs_map = [
        ("🔎 Búsqueda y Acciones", tab_search_actions.render if 'tab_search_actions' in globals() else None),
        ("⚙️ Acciones Automatizadas", tab_automated_actions.render if 'tab_automated_actions' in globals() else None),
        ("🔄 Conversión de Archivos", tab_conversion.render if 'tab_conversion' in globals() else None),
        ("📄 Visor (JSON/XML)", tab_visor.render if 'tab_visor' in globals() else None),
        ("RIPS", tab_rips.render if 'tab_rips' in globals() else None),
        ("✅ Validador FevRips", tab_validator_fevrips.render if 'tab_validator_fevrips' in globals() else None),
        ("📂 Gestión Documental", tab_gestion_documental.render if 'tab_gestion_documental' in globals() else None),
        ("👤 Validación Usuario", tab_user_validation.render if 'tab_user_validation' in globals() else None),
        ("🤖 Bot Zeus Salud", tab_bot_zeus.render if 'tab_bot_zeus' in globals() else None),
        ("🤖 Asistente IA (Gemini)", tab_ai_assistant.render if 'tab_ai_assistant' in globals() else None),
        ("📊 Gestión de Información", tab_admin.render if 'tab_admin' in globals() else None),
        ("👥 Gestión de Usuarios", tab_user_management.render if 'tab_user_management' in globals() else None),
    ]
    
    # Filter tabs based on permissions
    visible_tabs = []
    for name, render_func in all_tabs_map:
        # Extra security check for Admin tabs
        if name in ["📊 Gestión de Información", "👥 Gestión de Usuarios"]:
            if st.session_state.role != "admin":
                continue

        if "*" in allowed_tabs_config or name in allowed_tabs_config:
            if render_func: # Only add if module is loaded
                visible_tabs.append((name, render_func))
    
    if not visible_tabs:
        st.error("No tienes acceso a ninguna pestaña. Contacta al administrador.")
    else:
        # Render Tabs
        tab_names = [t[0] for t in visible_tabs]
        tabs = st.tabs(tab_names)
        
        for i, tab in enumerate(tabs):
            with tab:
                # Call the render function for the tab
                try:
                    name = visible_tabs[i][0]
                    render_func = visible_tabs[i][1]
                    
                    if name == "🔎 Búsqueda y Acciones":
                        render_func(tab)
                    elif name == "⚙️ Acciones Automatizadas" or name == "📂 Gestión Documental":
                        render_func()
                    else:
                        # Bot Zeus and AI Assistant take tab_container
                        render_func(tab)
                except Exception as e:
                    st.error(f"Error cargando la pestaña {tab_names[i]}: {e}")
                    # Log detail
                    import traceback
                    st.code(traceback.format_exc())

# --- FOOTER ---
st.markdown("---")
st.caption("CDO Clinical Document Organizer v2.5 | © 2025")
