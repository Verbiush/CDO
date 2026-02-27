import streamlit as st
import os
import tkinter as tk
from tkinter import filedialog
import time

def abrir_dialogo_carpeta_nativo(title="Seleccionar Carpeta", initial_dir=None):
    """
    Abre un diálogo de selección de carpeta nativo usando Tkinter.
    Retorna la ruta seleccionada o None si se cancela.
    """
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
        st.error(f"Error al abrir selector nativo: {e}")
        return None

def abrir_dialogo_archivo_nativo(title="Seleccionar Archivo", initial_dir=None, file_types=None):
    """
    Abre un diálogo de selección de archivo nativo usando Tkinter.
    Retorna la ruta seleccionada o None si se cancela.
    """
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
        st.button("📁", key=btn_key, help="Seleccionar Carpeta", disabled=not use_custom,
                  on_click=lambda: update_path_key(key, abrir_dialogo_carpeta_nativo(initial_dir=target_path), widget_key=input_key))

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
        st.button("📄", key=btn_key, help="Seleccionar Archivo", disabled=not use_custom,
                  on_click=lambda: update_path_key(key, abrir_dialogo_archivo_nativo(initial_dir=os.path.dirname(target_path) if os.path.isfile(target_path) else target_path, file_types=file_types), widget_key=input_key))

    return target_path
