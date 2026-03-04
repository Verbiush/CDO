import streamlit as st
import os
import time
import zipfile
import shutil
from pathlib import Path

try:
    import tkinter as tk
    from tkinter import filedialog
    TKINTER_AVAILABLE = True
except ImportError:
    TKINTER_AVAILABLE = False

def abrir_dialogo_carpeta_nativo(title="Seleccionar Carpeta", initial_dir=None):
    """
    Abre un diálogo de selección de carpeta nativo usando Tkinter.
    Retorna la ruta seleccionada o None si se cancela.
    """
    if not st.session_state.get("force_native_mode", True):
        # En modo Web, no usamos Tkinter nativo
        return None

    if not TKINTER_AVAILABLE:
        st.error("Error: Tkinter no está disponible en este entorno.")
        return None

    try:
        # Verificar si estamos en un entorno compatible (local)
        # En Streamlit Cloud esto no funcionará, pero asumimos entorno local Windows
        root = tk.Tk()
        root.withdraw()  # Ocultar la ventana principal
        root.wm_attributes('-topmost', 1)  # Mantener siempre encima
        
        if not initial_dir:
            initial_dir = os.getcwd()
            
        folder_path = filedialog.askdirectory(title=title, initialdir=initial_dir)
        
        root.destroy()
        return folder_path if folder_path else None
    except Exception as e:
        # En modo Web, esto puede pasar si force_native_mode no se detecta correctamente
        if "DISPLAY" in str(e) or "screen" in str(e):
            st.error("Error: No se puede abrir ventana nativa en modo Web. Use la entrada manual.")
        else:
            st.error(f"Error al abrir selector nativo: {e}")
        return None

def abrir_dialogo_archivo_nativo(title="Seleccionar Archivo", initial_dir=None, file_types=None):
    """
    Abre un diálogo de selección de archivo nativo usando Tkinter.
    Retorna la ruta seleccionada o None si se cancela.
    """
    if not st.session_state.get("force_native_mode", True):
        return None
    
    if not TKINTER_AVAILABLE:
        st.error("Error: Tkinter no está disponible en este entorno.")
        return None

    try:
        root = tk.Tk()
        root.withdraw()
        root.wm_attributes('-topmost', 1)
        
        if not initial_dir:
            initial_dir = os.getcwd()
            
        if not file_types:
            file_types = [("Todos los archivos", "*.*")]
            
        file_path = filedialog.askopenfilename(title=title, initialdir=initial_dir, filetypes=file_types)
        
        root.destroy()
        return file_path if file_path else None
    except Exception as e:
        if "DISPLAY" in str(e) or "screen" in str(e):
            st.error("Error: No se puede abrir ventana nativa en modo Web. Use la entrada manual.")
        else:
            st.error(f"Error al abrir selector de archivo nativo: {e}")
        return None

def update_path_key(key, new_path, widget_key=None):
    """
    Actualiza una clave en session_state con la nueva ruta.
    Opcionalmente actualiza también la clave del widget de texto asociado.
    """
    if new_path:
        st.session_state[key] = new_path
        if widget_key:
            st.session_state[widget_key] = new_path
        # No llamamos a st.rerun() aquí para evitar bucles, el cambio de estado disparará la actualización reactiva si es necesario
        # o el usuario puede continuar.

def extract_uploaded_zip(uploaded_file):
    """
    Extracts an uploaded ZIP file to a temporary directory.
    Returns the absolute path to the extracted folder.
    """
    if not uploaded_file:
        return None
        
    # Create temp structure: temp_uploads/{session_id}/{filename}
    # We use a simplified session ID or timestamp if session_id not available
    session_id = getattr(st.session_state, "session_id", "default_session")
    timestamp = int(time.time())
    
    # Safe filename
    safe_name = "".join([c for c in uploaded_file.name if c.isalnum() or c in ('-', '_', '.')])
    extract_dir = os.path.join(os.getcwd(), "temp_uploads", f"{timestamp}_{safe_name}")
    
    if os.path.exists(extract_dir):
        # Already extracted? 
        # For now, we assume if it exists it's the same. 
        # Or we can just return it.
        return extract_dir
        
    os.makedirs(extract_dir, exist_ok=True)
    
    try:
        with zipfile.ZipFile(uploaded_file) as z:
            z.extractall(extract_dir)
            
        # Check if the zip contains a single top-level folder
        # If so, return that folder instead of the wrapper
        items = os.listdir(extract_dir)
        if len(items) == 1 and os.path.isdir(os.path.join(extract_dir, items[0])):
            return os.path.join(extract_dir, items[0])
            
        return extract_dir
    except Exception as e:
        st.error(f"Error al extraer ZIP: {e}")
        return None

def render_path_selector(label, key, default_path=None, help_text=None, omit_checkbox=False):
    """
    Renderiza un selector de ruta estandarizado.
    Soporta:
    - Modo Nativo: Diálogo de carpeta Tkinter.
    - Modo Web:
        - Subida de ZIP (descomprime y usa esa ruta).
        - Entrada Manual.
        - Agente Local (si está disponible).
    """
    if default_path is None:
        default_path = st.session_state.get("current_path", os.getcwd())

    # Checkbox state for "Use Custom"
    cb_key = f"cb_use_custom_{key}"
    if omit_checkbox:
        use_custom = True
    else:
        use_custom = st.checkbox(f"Modificar ruta: {label}", value=False, key=cb_key)

    # Determine target path
    if use_custom:
        if key not in st.session_state:
            st.session_state[key] = default_path
        target_path = st.session_state[key]
    else:
        target_path = default_path
        st.session_state[key] = target_path

    # --- RENDER UI ---
    is_native = st.session_state.get("force_native_mode", True)
    
    if is_native:
        # --- NATIVE MODE ---
        col1, col2 = st.columns([0.8, 0.2])
        with col1:
            input_key = f"input_{key}"
            if use_custom:
                st.text_input(label, value=target_path, key=input_key, help=help_text,
                              on_change=lambda: st.session_state.update({key: st.session_state[input_key]}))
            else:
                st.text_input(label, value=target_path, key=f"{input_key}_disabled", disabled=True, help=help_text)

        with col2:
            st.markdown('<div style="margin-top: 28px;"></div>', unsafe_allow_html=True)
            btn_key = f"btn_{key}"
            st.button("📁", key=btn_key, help="Seleccionar Carpeta", disabled=not use_custom,
                  on_click=lambda: update_path_key(key, abrir_dialogo_carpeta_nativo(initial_dir=target_path), widget_key=input_key))
    else:
        # --- WEB MODE ---
        st.markdown(f"**{label}**")
        
        # Options: Upload ZIP vs Manual/Agent
        method = st.radio("Método de Selección:", ["Subir Archivos (ZIP)", "Ruta del Servidor / Manual", "Agente Local"], 
                          key=f"method_{key}", horizontal=True, label_visibility="collapsed")
        
        if method == "Subir Archivos (ZIP)":
            uploaded = st.file_uploader(f"Subir ZIP con archivos para '{label}'", type="zip", key=f"upload_{key}")
            if uploaded:
                path = extract_uploaded_zip(uploaded)
                if path:
                    st.success(f"✅ Archivos extraídos en: {path}")
                    st.session_state[key] = path
                    target_path = path
                    # Show preview of extracted files?
                    # st.caption(f"Contenido: {os.listdir(path)[:5]}...")
            else:
                st.info("Sube un ZIP para trabajar con sus carpetas/archivos.")
                
        elif method == "Ruta del Servidor / Manual":
            input_key = f"input_man_{key}"
            st.text_input("Ruta en el Servidor", value=target_path, key=input_key, help=help_text,
                          on_change=lambda: st.session_state.update({key: st.session_state[input_key]}))
            
        elif method == "Agente Local":
            col1, col2 = st.columns([0.8, 0.2])
            with col1:
                st.text_input("Ruta (desde Agente)", value=target_path, disabled=True, key=f"disp_agent_{key}")
            with col2:
                st.markdown('<div style="margin-top: 28px;"></div>', unsafe_allow_html=True)
                btn_key = f"btn_agent_{key}"
                if st.button("🖥️", key=btn_key, help="Solicitar selección al Agente Local"):
                     username = st.session_state.get("username", "admin")
                     success, task_id = database.create_task(username, "SELECT_FOLDER")
                     
                     if success and task_id:
                         progress_text = "Esperando agente..."
                         status_area = st.empty()
                         status_area.info(progress_text)
                         
                         found = False
                         for _ in range(60):
                             time.sleep(1)
                             result = database.get_task_result(task_id)
                             if result and result.get("status") == "COMPLETED":
                                 res_data = result.get("result", {})
                                 if res_data and res_data.get("success"):
                                     path = res_data.get("data")
                                     if path:
                                         update_path_key(key, path)
                                         status_area.success("Seleccionado!")
                                         found = True
                                         time.sleep(0.5)
                                         st.rerun()
                                 break
                             elif result and result.get("status") == "ERROR":
                                 status_area.error(f"Error: {result.get('result', {}).get('error')}")
                                 found = True
                                 break
                         
                         if not found:
                             status_area.warning("Tiempo agotado.")

    return target_path

def render_file_selector(label, key, default_path=None, help_text=None, file_types=None, omit_checkbox=False):
    """
    Renderiza un selector de archivo estandarizado con checkbox 'Escoger archivo diferente'.
    Si omit_checkbox es True, el selector siempre está activo y no muestra el checkbox.
    Retorna la ruta seleccionada.
    """
    if default_path is None:
        default_path = st.session_state.get("current_path", os.getcwd())

    # Checkbox state
    cb_key = f"cb_use_custom_file_{key}"
    if omit_checkbox:
        use_custom = True
    else:
        use_custom = st.checkbox("Escoger archivo diferente", value=False, key=cb_key)

    # Determine target path
    if use_custom:
        # Initialize if not set
        if key not in st.session_state:
            st.session_state[key] = default_path
        target_path = st.session_state[key]
    else:
        target_path = default_path
        # Sync key if needed
        st.session_state[key] = target_path

    col1, col2 = st.columns([0.8, 0.2])
    
    with col1:
        input_key = f"input_{key}"
        if use_custom:
            # Active input with on_change sync
            st.text_input(label, value=target_path, key=input_key, help=help_text,
                          on_change=lambda: st.session_state.update({key: st.session_state[input_key]}))
        else:
            # Disabled input
            st.text_input(label, value=target_path, key=f"{input_key}_disabled", disabled=True, help=help_text)

    with col2:
        st.markdown('<div style="margin-top: 28px;"></div>', unsafe_allow_html=True)
        # Button to open dialog
        btn_key = f"btn_{key}"
        
        is_native = st.session_state.get("force_native_mode", True)
        
        if is_native:
            st.button("📄", key=btn_key, help="Seleccionar Archivo", disabled=not use_custom,
                  on_click=lambda: update_path_key(key, abrir_dialogo_archivo_nativo(initial_dir=os.path.dirname(target_path) if os.path.isfile(target_path) else target_path, file_types=file_types), widget_key=input_key))
        else:
             if st.button("🖥️", key=btn_key, help="Solicitar selección de archivo al Agente Local", disabled=not use_custom):
                 username = st.session_state.get("username", "admin")
                 success, task_id = database.create_task(username, "SELECT_FILE")
                 
                 if success and task_id:
                     progress_text = "Esperando que el agente seleccione el archivo... (Revise su PC local)"
                     status_area = st.empty()
                     status_area.info(progress_text)
                     
                     found = False
                     for _ in range(60):
                         time.sleep(1)
                         result = database.get_task_result(task_id)
                         if result and result.get("status") == "COMPLETED":
                             res_data = result.get("result", {})
                             if res_data and res_data.get("success"):
                                 path = res_data.get("data")
                                 if path:
                                     update_path_key(key, path, widget_key=input_key)
                                     status_area.success(f"Archivo seleccionado: {path}")
                                     found = True
                                     time.sleep(1)
                                     st.rerun()
                             else:
                                 status_area.warning("Selección cancelada o fallida.")
                                 found = True
                             break
                         elif result and result.get("status") == "ERROR":
                             status_area.error(f"Error del agente: {result.get('result', {}).get('error')}")
                             found = True
                             break
                     
                     if not found:
                         status_area.warning("Tiempo de espera agotado. Asegúrese de que el agente esté ejecutándose.")
                 else:
                     st.error("Error al crear la tarea para el agente.")

    return target_path

def render_download_button(folder_path, key, label="📦 Descargar ZIP"):
    """
    Renderiza un botón para descargar el contenido de una carpeta como ZIP,
    o un archivo individual directamente.
    """
    if not folder_path or not os.path.exists(folder_path):
        return
        
    is_file = os.path.isfile(folder_path)
    
    # Check if folder is empty (only if directory)
    if not is_file:
        try:
            if not os.listdir(folder_path):
                # st.warning("Carpeta vacía, nada que descargar.")
                return
        except Exception:
            return

    st.markdown("### Descargar Resultados")
    
    col1, col2 = st.columns([0.6, 0.4])
    with col1:
        st.info(f"Ruta: {folder_path}")
        
    with col2:
        if is_file:
            # Direct download for single file
            try:
                with open(folder_path, "rb") as f:
                    file_name = os.path.basename(folder_path)
                    # Use provided label if it's not the default generic one, otherwise make it specific
                    btn_label = label if label != "📦 Descargar ZIP" else f"📥 Descargar {file_name}"
                    
                    st.download_button(
                        label=btn_label,
                        data=f,
                        file_name=file_name,
                        mime="application/octet-stream",
                        key=f"dl_btn_file_{key}"
                    )
            except Exception as e:
                st.error(f"Error al leer archivo: {e}")
        else:
            # ZIP logic for folders
            gen_key = f"gen_zip_{key}"
            
            if st.button("Preparar Descarga (ZIP)", key=gen_key):
                with st.spinner("Comprimiendo archivos..."):
                    try:
                        # Create zip in temp
                        timestamp = int(time.time())
                        zip_name = f"download_{timestamp}" # make_archive adds .zip
                        
                        # Ensure temp dir exists
                        temp_dl_dir = os.path.join(os.getcwd(), "temp_downloads")
                        os.makedirs(temp_dl_dir, exist_ok=True)
                        
                        zip_base_path = os.path.join(temp_dl_dir, zip_name)
                        
                        zip_path = shutil.make_archive(zip_base_path, 'zip', folder_path)
                        
                        st.session_state[f"ready_zip_{key}"] = zip_path
                        st.success("✅ Archivo listo.")
                    except Exception as e:
                        st.error(f"Error al comprimir: {e}")
                
            # If ready, show download
            ready_zip = st.session_state.get(f"ready_zip_{key}")
            if ready_zip and os.path.exists(ready_zip):
                with open(ready_zip, "rb") as fp:
                    st.download_button(
                        label=label,
                        data=fp,
                        file_name=os.path.basename(ready_zip),
                        mime="application/zip",
                        key=f"dl_btn_{key}"
                    )
