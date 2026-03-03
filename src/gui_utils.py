import streamlit as st
import os
import time

try:
    import tkinter as tk
    from tkinter import filedialog
    TKINTER_AVAILABLE = True
except ImportError:
    TKINTER_AVAILABLE = False

try:
    import src.database as database
except ImportError:
    import database # Fallback for different execution contexts

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

def render_path_selector(label, key, default_path=None, help_text=None, omit_checkbox=False):
    """
    Renderiza un selector de ruta estandarizado con checkbox 'Escoger ruta diferente'.
    Si omit_checkbox es True, el selector siempre está activo y no muestra el checkbox.
    Retorna la ruta seleccionada.
    """
    if default_path is None:
        default_path = st.session_state.get("current_path", os.getcwd())

    # Checkbox state
    cb_key = f"cb_use_custom_{key}"
    if omit_checkbox:
        use_custom = True
    else:
        use_custom = st.checkbox("Escoger ruta diferente a la inicial", value=False, key=cb_key)

    # Determine target path
    if use_custom:
        # Initialize if not set
        if key not in st.session_state:
            st.session_state[key] = default_path
        target_path = st.session_state[key]
    else:
        target_path = default_path
        # Sync key if needed (optional, but good for consistency)
        st.session_state[key] = target_path

    col1, col2 = st.columns([0.8, 0.2])
    
    with col1:
        input_key = f"input_{key}"
        if use_custom:
            # Active input with on_change sync
            # Note: lambda inside on_change captures variables, but here input_key is local string, which is fine.
            # However, st.session_state[input_key] needs to be accessed dynamically.
            st.text_input(label, value=target_path, key=input_key, help=help_text,
                          on_change=lambda: st.session_state.update({key: st.session_state[input_key]}))
        else:
            # Disabled input
            st.text_input(label, value=target_path, key=f"{input_key}_disabled", disabled=True, help=help_text)

    with col2:
        st.markdown('<div style="margin-top: 28px;"></div>', unsafe_allow_html=True)
        # Button to open dialog
        btn_key = f"btn_{key}"
        
        # Determine button behavior based on mode
        is_native = st.session_state.get("force_native_mode", True)
        
        if is_native:
             st.button("📁", key=btn_key, help="Seleccionar Carpeta", disabled=not use_custom,
                  on_click=lambda: update_path_key(key, abrir_dialogo_carpeta_nativo(initial_dir=target_path), widget_key=input_key))
        else:
             # In Web Mode, we cannot open local dialog easily from server.
             # Trigger an agent task.
             if st.button("🖥️", key=btn_key, help="Solicitar selección al Agente Local", disabled=not use_custom):
                 username = st.session_state.get("username", "admin")
                 success, task_id = database.create_task(username, "SELECT_FOLDER")
                 
                 if success and task_id:
                     progress_text = "Esperando que el agente seleccione la carpeta... (Revise su PC local)"
                     status_area = st.empty()
                     status_area.info(progress_text)
                     
                     # Poll for result (max 60 seconds)
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
                                     status_area.success(f"Carpeta seleccionada: {path}")
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
