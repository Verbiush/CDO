import os
import time
import streamlit as st
import tempfile
import shutil
import zipfile

# Global timestamp to prevent rapid-fire dialog openings (debounce)
_last_dialog_time = 0

def abrir_dialogo_carpeta_nativo(title="Seleccionar Carpeta", initial_dir=None):
    """
    Función BLOQUEANTE que abre el diálogo nativo de Tkinter.
    Debe ser llamada solo desde callbacks o cuando se sabe que es seguro bloquear.
    NO llamar directamente en el bucle de renderizado de Streamlit sin un botón previo.
    """
    global _last_dialog_time
    
    # Debounce check (2 seconds cooldown)
    current_time = time.time()
    if current_time - _last_dialog_time < 2.0:
        print("Ignorando llamada a diálogo nativo (debounce activo)")
        return None
        
    _last_dialog_time = current_time

    # 0. Validación básica de entorno
    try:
        if os.environ.get("STREAMLIT_SERVER_HEADLESS") == "true":
             return None
    except Exception:
        pass

    # 1. Intentar usar Agente Local
    try:
        agent_client = None
        try:
            import agent_client
        except ImportError:
            try:
                from src import agent_client
            except ImportError:
                pass
        
        if agent_client and hasattr(agent_client, 'is_agent_available') and agent_client.is_agent_available():
            folder = agent_client.select_folder()
            return folder
    except Exception as e:
        print(f"Advertencia: No se pudo contactar al agente local: {e}")

    # 2. Intentar usar Tkinter (Nativo)
    try:
        import tkinter as tk
        from tkinter import filedialog
        
        try:
            root = tk.Tk()
            root.withdraw() # Ocultar ventana principal
            root.attributes('-topmost', True) # Forzar al frente
            
            # Asegurar que initial_dir sea válido
            if initial_dir and isinstance(initial_dir, str) and not os.path.isdir(initial_dir):
                initial_dir = None
            
            # Abrir diálogo
            folder = filedialog.askdirectory(
                master=root, 
                title=title, 
                initialdir=initial_dir
            )
            
            # Destruir root
            try:
                root.destroy()
            except:
                pass
            
            if not folder:
                return None
            
            return os.path.normpath(folder)

        except Exception as e:
            print(f"Error Tkinter: {e}")
            return None
            
    except ImportError:
        print("Tkinter no instalado.")
        return None

def seleccionar_carpeta_nativa(title="Seleccionar Carpeta", initial_dir=None, key=None):
    """
    Componente UI de Streamlit para seleccionar carpeta.
    En Modo Web: Muestra un uploader para crear un entorno de trabajo temporal y permite descargar resultados.
    En Modo Nativo: Muestra la ruta actual y un botón 'Examinar' que abre el diálogo.
    Retorna la ruta seleccionada (o temporal).
    """
    # Verificar modo nativo (Configuración General)
    is_native = st.session_state.get("force_native_mode", True)
    
    # Generar key única si no existe
    # IMPORTANTE: Usar una key estable basada en el título para mantener el estado entre reruns.
    safe_title = "".join(c for c in title if c.isalnum() or c in ('_', '-')).strip()
    if not key:
        key = f"folder_selector_{safe_title}"
    
    # Inicializar estado si es necesario
    last_initial_key = f"last_initial_{key}"
    
    if key not in st.session_state:
        st.session_state[key] = initial_dir if initial_dir else os.getcwd()
        if initial_dir:
            st.session_state[last_initial_key] = initial_dir
    else:
        # Lógica inteligente para actualizar la ruta si el initial_dir cambia globalmente
        # (ej: usuario cambia carpeta en pestaña Búsqueda y queremos que se propague aquí)
        
        # 1. Recuperar último initial_dir conocido para este widget
        last_known = st.session_state.get(last_initial_key)
        
        # 2. Si initial_dir cambió respecto a la última vez, actualizamos el widget
        if initial_dir and initial_dir != last_known:
            st.session_state[key] = initial_dir
            st.session_state[last_initial_key] = initial_dir
        
        # 3. Fallback: Heurística eliminada para evitar sobrescrituras accidentales
        # Si last_known no existe, confiamos en el valor actual de st.session_state[key]
        # para evitar que el input del usuario sea reemplazado por initial_dir.

    if not is_native:
        # --- MODO WEB LOCAL: Entrada de Texto Directa ---
        # El usuario prefiere escribir/pegar la ruta localmente en lugar de usar un entorno aislado.
        
        # Sincronizar con cambios externos (search bar)
        if key not in st.session_state:
             st.session_state[key] = initial_dir if initial_dir else os.getcwd()

        # Renderizar input de texto
        # Usamos key=key para que Streamlit gestione el estado, pero permitimos actualizaciones externas
        # Si initial_dir cambió (detectado arriba), st.session_state[key] ya está actualizado.
        
        val = st.text_input(f"📂 {title}", value=st.session_state[key], key=f"input_{key}")
        
        # Sincronización bidireccional
        if val != st.session_state[key]:
            st.session_state[key] = val
            # Forzar actualización si es necesario (generalmente no lo es con key propia)
            
        return val

    # --- MODO NATIVO: UI Compuesta ---
    col1, col2 = st.columns([0.85, 0.15])
    
    with col1:
        # Mostrar ruta actual como input editable
        current_path = st.session_state.get(key, "")
        val = st.text_input(title, value=current_path, key=f"input_{key}")
        
        # Sincronización si se edita manualmente
        if val != current_path:
            st.session_state[key] = val
        
    with col2:
        st.write("")
        st.write("")
        # Botón para abrir diálogo
        if st.button("📂", key=f"btn_browse_{key}", help="Examinar carpeta..."):
            # Usar el valor actual del input como inicio si es válido
            start_dir = val if val and os.path.isdir(val) else (current_path if current_path else None)
            new_path = abrir_dialogo_carpeta_nativo(title, start_dir)
            if new_path:
                st.session_state[key] = new_path
                st.rerun()
                
    return st.session_state.get(key, "")
