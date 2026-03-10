import streamlit as st
import os
import time
import zipfile
import shutil
import platform
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
    # Intentar Tkinter primero si estamos en modo nativo
    # Optimizacion: Solo usar Tkinter si estamos en Windows (asumimos que es local)
    # En Linux (AWS) pasamos directo al Agente para evitar errores/delays
    is_windows = platform.system() == "Windows"
    if st.session_state.get("force_native_mode", True) and TKINTER_AVAILABLE and is_windows:
        try:
            # Verificar si estamos en un entorno compatible (local)
            root = tk.Tk()
            root.withdraw()  # Ocultar la ventana principal
            root.wm_attributes('-topmost', 1)  # Mantener siempre encima
            
            if not initial_dir:
                initial_dir = os.getcwd()
                
            folder_path = filedialog.askdirectory(title=title, initialdir=initial_dir)
            
            root.destroy()
            return folder_path if folder_path else None
        except Exception as e:
            print(f"DEBUG: Tkinter failed ({e}). Trying Agent fallback...")
    
    # Fallback al Agente Local (funciona en Web y Nativo si Tkinter falla)
    try:
        # Lazy import
        try: import agent_client
        except ImportError: from src import agent_client
        
        username = st.session_state.get("username", "admin")
        
        # Solo intentar si parece que estamos en un entorno donde podría haber un agente
        # O si el usuario explícitamente quiere usar el agente (podríamos añadir un flag)
        st.toast("Solicitando al Agente Local...", icon="🔌")
        
        # Usar el agente para abrir el diálogo en el PC del usuario
        return agent_client.select_folder(username, title=title)
    except Exception as agent_e:
        if st.session_state.get("force_native_mode", True):
             st.error(f"⚠️ No se pudo abrir la ventana en su PC via Agente. Verifique que el Agente Local esté conectado. (Error: {agent_e})")
        else:
             # En web es normal que falle si no hay agente, no mostramos error intrusivo
             print(f"Agent fallback failed: {agent_e}")
        return None

def abrir_dialogo_archivo_nativo(title="Seleccionar Archivo", initial_dir=None, file_types=None):
    """
    Abre un diálogo de selección de archivo nativo usando Tkinter.
    Retorna la ruta seleccionada o None si se cancela.
    """
    # Intentar Tkinter primero si estamos en modo nativo
    is_windows = platform.system() == "Windows"
    if st.session_state.get("force_native_mode", True) and TKINTER_AVAILABLE and is_windows:
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
            print(f"DEBUG: Tkinter failed ({e}). Trying Agent fallback...")
            
    # Fallback al Agente Local
    try:
        # Lazy import
        try: import agent_client
        except ImportError: from src import agent_client
        
        username = st.session_state.get("username", "admin")
        st.toast("Solicitando al Agente Local...", icon="🔌")
        
        # Usar el agente para abrir el diálogo en el PC del usuario
        return agent_client.select_file(username, title=title, file_types=file_types)
    except Exception as agent_e:
        if st.session_state.get("force_native_mode", True):
             st.error(f"⚠️ No se pudo abrir selección de archivo via Agente. Verifique conexión. (Error: {agent_e})")
        else:
             print(f"Agent fallback failed: {agent_e}")
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
    
    # Cleanup old temp files (older than 10 minutes) to prevent "No space left on device"
    try:
        base_temp = os.path.join(os.getcwd(), "temp_uploads")
        if os.path.exists(base_temp):
            for item in os.listdir(base_temp):
                item_path = os.path.join(base_temp, item)
                try:
                    # If item is older than 10 minutes (600 seconds)
                    if os.path.getmtime(item_path) < time.time() - 600:
                        if os.path.isdir(item_path):
                            shutil.rmtree(item_path, ignore_errors=True)
                        else:
                            os.remove(item_path)
                except Exception:
                    pass
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

    # En modo nativo, forzamos omit_checkbox para que siempre esté activo y no muestre el checkbox
    is_native = st.session_state.get("force_native_mode", True)
    if is_native:
        omit_checkbox = True

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
            # Use on_click to trigger dialog and update state
            def on_click_folder():
                # Get current path from state or default
                current = st.session_state.get(key, target_path)
                selected = abrir_dialogo_carpeta_nativo(title=label, initial_dir=current)
                if selected:
                    st.session_state[key] = selected
                    # Also update the text input key to reflect change immediately
                    st.session_state[f"input_{key}"] = selected
            
            st.button("📁", key=btn_key, help="Seleccionar Carpeta", disabled=not use_custom, on_click=on_click_folder)

        # Return the current path stored in session state
        return st.session_state.get(key, target_path)

    else:
        # --- WEB MODE ---
        st.markdown(f"**{label}**")
        
        # Opcion: Usar Agente Local
        use_agent_key = f"use_agent_folder_{key}"
        
        # Check if agent is likely available (optional heuristic)
        agent_available = True 
        
        col_agent_check, col_agent_status = st.columns([0.6, 0.4])
        with col_agent_check:
            use_agent = st.checkbox("🔌 Usar Agente Local", key=use_agent_key, help="Conectar con el agente instalado en tu PC para seleccionar carpetas locales.")
        
        if use_agent:
            username = st.session_state.get("username", "admin")
            # Show connection info
            with col_agent_status:
                st.caption(f"Usuario: {username}")

            col1, col2 = st.columns([0.7, 0.3])
            
            # Display current selected path (or empty)
            current_val = st.session_state.get(key, "")
            with col1:
                st.text_input("Ruta remota:", value=current_val, key=f"remote_path_display_{key}", disabled=True, label_visibility="collapsed")
            
            with col2:
                # Button to trigger agent
                # Usamos un icono de carpeta para que sea familiar
                if st.button("📁 Examinar PC", key=f"btn_agent_folder_{key}", type="primary"):
                    try:
                        # Lazy import to avoid circular dependency
                        try:
                            import agent_client
                        except ImportError:
                            from src import agent_client
                            
                        # Call agent (blocking with spinner)
                        with st.spinner(f"Esperando selección en el Agente ({username})..."):
                            selected = agent_client.select_folder(username, title=label)
                        
                        if selected:
                            st.session_state[key] = selected
                            target_path = selected
                            st.success(f"Ruta: {selected}")
                            st.rerun()
                        else:
                            st.warning("Cancelado o sin respuesta.")
                    except Exception as e:
                        st.error(f"Error: {e}")
            
            # If path is set, return it
            return st.session_state.get(key, target_path)

        # En modo Web (sin agente), forzamos la carga de ZIP
        st.write("--- O ---")
        st.write("Sube un ZIP para trabajar con sus carpetas/archivos.")
        uploaded = st.file_uploader(f"Subir ZIP con archivos para '{label}'", type="zip", key=f"upload_{key}", label_visibility="collapsed")
        
        if uploaded:
            path = extract_uploaded_zip(uploaded)
            if path:
                st.success(f"✅ Archivos extraídos en: {path}")
                st.session_state[key] = path
                target_path = path
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
        
        # Opcion: Usar Agente Local
        use_agent_key = f"use_agent_file_{key}"
        use_agent = st.checkbox("🔌 Usar Agente Local (PC)", key=use_agent_key, help="Seleccionar archivo en tu PC usando el Agente instalado.")
        
        if use_agent:
            username = st.session_state.get("username", "admin")
            col1, col2 = st.columns([0.7, 0.3])
            
            # Display current selected path (or empty)
            current_val = st.session_state.get(key, "")
            with col1:
                st.text_input("Ruta remota:", value=current_val, key=f"remote_path_file_display_{key}", disabled=True)
            
            with col2:
                # Button to trigger agent
                if st.button("📄 Explorar", key=f"btn_agent_file_{key}"):
                    try:
                        # Lazy import to avoid circular dependency
                        try:
                            import agent_client
                        except ImportError:
                            from src import agent_client
                            
                        # Call agent (blocking with spinner)
                        selected = agent_client.select_file(username, title=label, file_types=file_types)
                        
                        if selected:
                            st.session_state[key] = selected
                            target_path = selected
                            st.success(f"Seleccionado: {selected}")
                            st.rerun()
                        else:
                            st.warning("No se seleccionó ningún archivo o el agente no respondió.")
                    except Exception as e:
                        st.error(f"Error comunicando con agente: {e}")
            
            # If path is set, return it
            return st.session_state.get(key, target_path)
        
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

def render_download_button(folder_path, key, label="📦 Descargar ZIP", cleanup=False):
    """
    Renderiza un botón para descargar el contenido de una carpeta como ZIP,
    o un archivo individual directamente.
    Soporta modo Nativo (Guardar Como) y Web (Descarga navegador).
    Si cleanup=True y estamos en modo Web (ZIP en memoria), borra la carpeta tras comprimir.
    """
    # Check if we already have the zip prepared in session state
    zip_data_key = f"ready_zip_data_{key}"
    zip_name_key = f"ready_zip_name_{key}"
    
    has_ready_zip = zip_data_key in st.session_state and st.session_state[zip_data_key] is not None
    
    if not has_ready_zip and (not folder_path or not os.path.exists(folder_path)):
        return
        
    is_file = os.path.isfile(folder_path) if folder_path and os.path.exists(folder_path) else False
    
    # Check if folder is empty (only if directory and exists)
    if not has_ready_zip and not is_file and folder_path and os.path.exists(folder_path):
        try:
            if not os.listdir(folder_path):
                # st.warning("Carpeta vacía, nada que descargar.")
                return
        except Exception:
            return

    st.markdown("### Descargar Resultados")
    
    col1, col2 = st.columns([0.6, 0.4])
    with col1:
        if folder_path and os.path.exists(folder_path):
            st.info(f"Ruta: {folder_path}")
        elif has_ready_zip:
             st.info(f"Archivo listo para descarga (En memoria)")
        
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
                        # Tkinter dialogs need main thread, usually fine in local Streamlit
                        try:
                            root = tk.Tk()
                            root.withdraw()
                            root.wm_attributes('-topmost', 1)
                            save_path = filedialog.asksaveasfilename(
                                title="Guardar archivo como...",
                                initialfile=initial_file,
                                defaultextension=os.path.splitext(initial_file)[1]
                            )
                            root.destroy()
                        except:
                            save_path = None
                            
                        if save_path:
                            shutil.copy2(folder_path, save_path)
                            st.success(f"✅ Archivo guardado en: {save_path}")
                    else:
                        # Save ZIP
                        try:
                            root = tk.Tk()
                            root.withdraw()
                            root.wm_attributes('-topmost', 1)
                            save_path = filedialog.asksaveasfilename(
                                title="Guardar ZIP como...",
                                initialfile=f"backup_{int(time.time())}.zip",
                                defaultextension=".zip",
                                filetypes=[("Zip files", "*.zip")]
                            )
                            root.destroy()
                        except:
                            save_path = None
                            
                        if save_path:
                            with st.spinner("Comprimiendo y guardando..."):
                                base_name = os.path.splitext(save_path)[0]
                                shutil.make_archive(base_name, 'zip', folder_path)
                                st.success(f"✅ ZIP guardado en: {save_path}")
                except Exception as e:
                    st.error(f"Error al guardar: {e}")
        else:
            # --- WEB MODE: BROWSER DOWNLOAD ---
            if is_file:
                # Direct download for single file
                try:
                    # Read into memory to allow cleanup
                    with open(folder_path, "rb") as f:
                        file_content = f.read()
                    
                    # Cleanup if requested (BEFORE showing button, since we have content in memory)
                    if cleanup:
                        try:
                            os.remove(folder_path)
                            # Optional: Try to remove parent dir if it looks like a temp dir and is empty
                            parent_dir = os.path.dirname(folder_path)
                            if "temp" in parent_dir.lower() and not os.listdir(parent_dir):
                                os.rmdir(parent_dir)
                        except Exception as e:
                            print(f"Error cleaning up file {folder_path}: {e}")
                            
                    file_name = os.path.basename(folder_path)
                    # Use provided label if it's not the default generic one, otherwise make it specific
                    btn_label = label if label != "📦 Descargar ZIP" else f"📥 Descargar {file_name}"
                    
                    st.download_button(
                        label=btn_label,
                        data=file_content,
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
                            
                            # Cleanup if requested
                            if cleanup:
                                try:
                                    shutil.rmtree(folder_path)
                                    # st.info("Archivos temporales del servidor eliminados.")
                                except Exception as e:
                                    print(f"Error cleaning up {folder_path}: {e}")
                            
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
