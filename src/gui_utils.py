import streamlit as st
import os
import time
import zipfile
import shutil
from pathlib import Path

try:
    import database
except ImportError:
    try:
        from src import database
    except ImportError:
        database = None

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
    
    # Cleanup old temp files (older than 1 hour) to prevent "No space left on device"
    try:
        base_temp = os.path.join(os.getcwd(), "temp_uploads")
        if os.path.exists(base_temp):
            for item in os.listdir(base_temp):
                item_path = os.path.join(base_temp, item)
                # If item is older than 1 hour (3600 seconds)
                if os.path.getmtime(item_path) < time.time() - 3600:
                    if os.path.isdir(item_path):
                        shutil.rmtree(item_path, ignore_errors=True)
                    else:
                        os.remove(item_path)
    except Exception as e:
        # Just log error, don't stop execution
        print(f"Warning: Cleanup failed: {e}")

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
        # Filter out system files like __MACOSX, .DS_Store, etc.
        items = [i for i in os.listdir(extract_dir) if i not in ['__MACOSX', '.DS_Store', 'Thumbs.db'] and not i.startswith('._')]
        
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
        # Evitar ruta predeterminada para obligar a la selección explícita
        default_path = ""

    # Checkbox state for "Use Custom"
    cb_key = f"cb_use_custom_{key}"
    if omit_checkbox:
        use_custom = True
    else:
        # Si la ruta está vacía, activamos "Modificar ruta" por defecto para que el usuario vea el input activo
        default_check = True if not default_path else False
        use_custom = st.checkbox(f"Modificar ruta: {label}", value=default_check, key=cb_key)

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
        
        # En modo Web, forzamos la carga de ZIP (sin opciones manuales ni agente)
        st.write("Sube un ZIP para trabajar con sus carpetas/archivos.")
        uploaded = st.file_uploader(f"Subir ZIP con archivos para '{label}'", type="zip", key=f"upload_{key}", label_visibility="collapsed")
        
        if uploaded:
            path = extract_uploaded_zip(uploaded)
            if path:
                st.success(f"✅ Archivos extraídos en: {path}")
                st.session_state[key] = path
                target_path = path
                # Show preview of extracted files?
                # st.caption(f"Contenido: {os.listdir(path)[:5]}...")
        else:
            st.info("Esperando archivo ZIP...")

    return target_path

def render_file_selector(label, key, default_path=None, help_text=None, file_types=None, omit_checkbox=False):
    """
    Renderiza un selector de archivo estandarizado con checkbox 'Escoger archivo diferente'.
    Si omit_checkbox es True, el selector siempre está activo y no muestra el checkbox.
    Retorna la ruta seleccionada.
    """
    if default_path is None:
        # Evitar ruta predeterminada para obligar a la selección explícita
        default_path = ""

    # Checkbox state
    cb_key = f"cb_use_custom_file_{key}"
    if omit_checkbox:
        use_custom = True
    else:
        # Si no hay ruta, activamos la edición por defecto
        default_check = True if not default_path else False
        use_custom = st.checkbox("Escoger archivo diferente", value=default_check, key=cb_key)

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

    is_native = st.session_state.get("force_native_mode", True)

    if is_native:
        # --- NATIVE MODE ---
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
    else:
        # --- WEB MODE ---
        st.markdown(f"**{label}**")
        
        # En modo Web, simplificamos a solo subida de archivo
        # Map file_types to extensions for uploader
        allowed_exts = None
        if file_types:
            # file_types is usually list of tuples [("Excel", "*.xlsx"), ...]
            allowed_exts = []
            for _, pat in file_types:
                if pat == "*.*": continue
                allowed_exts.append(pat.replace("*.", ""))
        
        uploaded = st.file_uploader(f"Subir archivo para '{label}'", type=allowed_exts, key=f"upload_file_{key}", label_visibility="collapsed")
        
        if uploaded:
            # Save to temp
            try:
                timestamp = int(time.time())
                safe_name = "".join([c for c in uploaded.name if c.isalnum() or c in ('-', '_', '.')])
                temp_dir = os.path.join(os.getcwd(), "temp_uploads", f"{timestamp}_file")
                os.makedirs(temp_dir, exist_ok=True)
                file_path = os.path.join(temp_dir, safe_name)
                
                # Only write if not exists or if we want to overwrite
                with open(file_path, "wb") as f:
                    f.write(uploaded.getbuffer())
                    
                st.success(f"✅ Archivo cargado: {safe_name}")
                st.session_state[key] = file_path
                target_path = file_path
            except Exception as e:
                st.error(f"Error guardando archivo: {e}")
        else:
            st.info("Esperando archivo...")

    return target_path

def render_download_button(folder_path, key, label="📦 Descargar ZIP"):
    """
    Renderiza un botón para descargar el contenido de una carpeta como ZIP,
    o un archivo individual directamente.
    Soporta modo Nativo (Guardar Como) y Web (Descarga navegador).
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
        is_native = st.session_state.get("force_native_mode", True)
        
        if is_native:
            # --- NATIVE MODE: SAVE AS DIALOG ---
            btn_label = label.replace("Descargar", "Guardar en") if "Descargar" in label else f"Guardar {label}"
            if st.button(f"💾 {btn_label}", key=f"native_save_{key}"):
                try:
                    if is_file:
                        # Save File
                        initial_file = os.path.basename(folder_path)
                        save_path = filedialog.asksaveasfilename(
                            title="Guardar archivo como...",
                            initialfile=initial_file,
                            defaultextension=os.path.splitext(initial_file)[1]
                        )
                        if save_path:
                            shutil.copy2(folder_path, save_path)
                            st.success(f"✅ Archivo guardado en: {save_path}")
                    else:
                        # Save ZIP
                        save_path = filedialog.asksaveasfilename(
                            title="Guardar ZIP como...",
                            initialfile=f"backup_{int(time.time())}.zip",
                            defaultextension=".zip",
                            filetypes=[("Zip files", "*.zip")]
                        )
                        if save_path:
                            with st.spinner("Comprimiendo y guardando..."):
                                shutil.make_archive(os.path.splitext(save_path)[0], 'zip', folder_path)
                                st.success(f"✅ ZIP guardado en: {save_path}")
                except Exception as e:
                    st.error(f"Error al guardar: {e}")
        else:
            # --- WEB MODE: BROWSER DOWNLOAD ---
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
                # ZIP logic for folders (In-Memory for Web Mode)
                gen_key = f"gen_zip_{key}"
                
                # Check folder size to decide between Memory or Disk?
                # For AWS/Cloud, Memory is preferred if small, but Disk (Temp) is safer for large.
                # User requested "no se guarden los archivos en el servidor".
                # We'll use Memory (BytesIO) to avoid persistent disk usage.
                
                if st.button("Preparar Descarga (ZIP)", key=gen_key):
                    with st.spinner("Comprimiendo en memoria..."):
                        try:
                            import io
                            mem_zip = io.BytesIO()
                            
                            with zipfile.ZipFile(mem_zip, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
                                root_len = len(folder_path) + 1
                                for root, dirs, files in os.walk(folder_path):
                                    for file in files:
                                        file_path = os.path.join(root, file)
                                        archive_name = file_path[root_len:]
                                        zf.write(file_path, archive_name)
                            
                            st.session_state[f"ready_zip_data_{key}"] = mem_zip.getvalue()
                            st.session_state[f"ready_zip_name_{key}"] = f"download_{int(time.time())}.zip"
                            st.success("✅ Archivo listo para descargar (En Memoria).")
                            
                        except Exception as e:
                            st.error(f"Error al comprimir: {e}")
                
                # If ready, show download
                zip_data = st.session_state.get(f"ready_zip_data_{key}")
                zip_name = st.session_state.get(f"ready_zip_name_{key}")
                
                if zip_data:
                    st.download_button(
                        label=label,
                        data=zip_data,
                        file_name=zip_name,
                        mime="application/zip",
                        key=f"dl_btn_{key}"
                    )
