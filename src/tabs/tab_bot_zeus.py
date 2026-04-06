import streamlit as st
import pandas as pd
import time
import json
import threading
import sys
import os
import io
import zipfile

# Try importing render_path_selector
try:
    from src.gui_utils import render_path_selector
except ImportError:
    try:
        # Fallback if src is not in path or running as script
        sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '../..')))
        from src.gui_utils import render_path_selector
    except ImportError:
        render_path_selector = None

@st.cache_data(show_spinner=False, max_entries=5)
def _get_excel_sheet_names(file_bytes):
    import pandas as pd
    import io
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    return xls.sheet_names

@st.cache_data(show_spinner=False, max_entries=10)
def _get_excel_preview(file_bytes, sheet_name, nrows=None):
    import pandas as pd
    import io
    if nrows:
        return pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, nrows=nrows)
    return pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name)

def render(tab_container):
    """
    Renderiza el contenido de la pestaña 'Bot Zeus Salud'.
    """
    # Lazy import de bot_zeus para mejorar el tiempo de carga inicial de la app
    try:
        import bot_zeus
    except ImportError as e1:
        # Intentar agregar el directorio padre al path si falla la importación directa
        current_dir = os.path.dirname(os.path.abspath(__file__))
        parent_dir = os.path.dirname(current_dir)
        if parent_dir not in sys.path:
            sys.path.insert(0, parent_dir)
        try:
            import bot_zeus
        except ImportError as e2:
            # Fallback para estructura src.bot_zeus
            try:
                from src import bot_zeus
            except ImportError as e3:
                st.error(f"❌ Error cargando módulo 'Bot Zeus'.\n\nDetalles:\n1. Import directo: {e1}\n2. Import path: {e2}\n3. Import src: {e3}\n\nRuta actual: {sys.path}")
                return
            except Exception as e4:
                st.error(f"❌ Error inesperado (src.bot_zeus): {e4}")
                return
        except Exception as e5:
             st.error(f"❌ Error inesperado (bot_zeus): {e5}")
             return
    except Exception as e6:
        st.error(f"❌ Error inesperado general: {e6}")
        return

    with tab_container:
        st.header("🤖 Bot Automatización Zeus Salud")
    
        # Use 'permissions' key as defined in app_web.py
        # Fallback to 'full' if key is missing or not set
        perms = st.session_state.get("permissions", {})
        bot_perm = perms.get("bot_zeus", "full")
    
        # 0. Verificación de Acceso Total
        if bot_perm == "none":
            st.error("⛔ Acceso Denegado: No tiene permisos para acceder al módulo Bot Zeus.")
            st.stop()

        # Configuración de Agente Local
        if "use_local_agent_bot" not in st.session_state:
            st.session_state.use_local_agent_bot = False
            
        use_agent = st.checkbox("🔌 Usar Agente Local (PC)", value=st.session_state.use_local_agent_bot, key="chk_use_agent_bot", help="Ejecutar el navegador en su PC local a través del Agente CDO.")
        st.session_state.use_local_agent_bot = use_agent
        
        if use_agent:
            st.info("ℹ️ Modo Agente: Las acciones se ejecutarán en su PC. Asegúrese de que el Agente CDO esté corriendo.")

        # Badge de Rol

        # Badge de Rol
        perm_labels = {
            "full": "✅ Control Total (Crear/Editar/Ejecutar)", 
            "edit": "✏️ Edición (Crear/Editar/Ejecutar)", 
            "execute": "🚀 Solo Ejecución",
            "none": "⛔ Sin Acceso"
        }
        st.caption(f"🛡️ Nivel de Acceso: **{perm_labels.get(bot_perm, bot_perm)}**")

        st.info("Automatización de ingreso de documentos desde Excel. Defina una secuencia de pasos (clicks, escritura, teclas) y ejecútela para cada fila.")
    
        col_bot1, col_bot2 = st.columns([1, 1])
    
        excel_cols = []
        df_bot = pd.DataFrame()
    

        with col_bot1:
            st.subheader("1. Conexión y Configuración")
            
            # Configuración de Carpeta de Descargas
            is_native = st.session_state.get("force_native_mode", True)
            if is_native:
                if render_path_selector is None:
                     st.error("Error: render_path_selector no pudo ser importado. Verifique src.gui_utils.")
                else:
                     dl_path = render_path_selector(
                         "Carpeta de Descargas (Bot)",
                         "bot_dl_path_sel",
                         default_path=st.session_state.get("current_path", os.path.join(os.getcwd(), "downloads")),
                         omit_checkbox=True
                     )
                     st.session_state.bot_download_dir = dl_path
            else:
                # Web Mode: Use temp folder
                if "bot_download_dir" not in st.session_state or not st.session_state.bot_download_dir.startswith(os.path.join(os.getcwd(), "temp_downloads")):
                    timestamp = int(time.time())
                    session_id = st.session_state.get("session_id", "default")
                    temp_dl = os.path.join(os.getcwd(), "temp_downloads", f"bot_session_{session_id}_{timestamp}")
                    os.makedirs(temp_dl, exist_ok=True)
                    st.session_state.bot_download_dir = temp_dl
                
                st.info(f"📂 Descargas en entorno temporal: {st.session_state.bot_download_dir}")

            if st.button("🚀 Abrir Navegador / Conectar", use_container_width=True):
                if use_agent:
                    try:
                        try:
                            import agent_client
                        except ImportError:
                            from src import agent_client
                        
                        username = st.session_state.get("username", "admin")
                        # Send command to agent
                        task_id = agent_client.send_command(username, "launch_browser", {"url": "https://ovidazs.siesacloud.com/ZeusSalud/ips/iniciando.php"})
                        if task_id:
                            with st.spinner("Esperando al agente..."):
                                res = agent_client.wait_for_result(task_id, timeout=10) # Reduced timeout for immediate return commands
                                # Consider missing "error" key or specific success status as successful
                                # If it timed out, but the agent logs show it launched, we assume it's running
                                if res and (res.get("status") == "success" or "error" not in res):
                                    st.success("✅ Navegador abierto en el Agente Local.")
                                elif res is None:
                                     # A timeout might just mean the agent didn't return a JSON properly or is blocking, 
                                     # but the browser likely opened.
                                     st.success("✅ Comando enviado al Agente Local. Verifique su pantalla.")
                                else:
                                    err_txt = res.get('error', res.get('message', 'Sin respuesta')) if isinstance(res, dict) else 'Sin respuesta'
                                    if isinstance(err_txt, str) and 'timeout waiting for agent response' in err_txt.lower():
                                        st.success("✅ Comando enviado al Agente Local. Verifique su pantalla.")
                                    else:
                                        st.error(f"Error del agente: {err_txt}")
                        else:
                            st.error("No se pudo conectar con el servidor para enviar la tarea.")
                    except Exception as e:
                        st.error(f"Error comunicando con agente: {e}")
                else:
                    success, msg = bot_zeus.abrir_navegador_inicial()
                    if success:
                        if not is_native:
                             st.success(f"{msg} (Modo Headless/Segundo Plano)")
                             st.info("ℹ️ En modo Web, el navegador se ejecuta en el servidor y no es visible. Use las funciones de captura de pantalla o logs para depurar si es necesario.")
                        else:
                            st.success(msg)
                    else:
                        st.error(msg)
            
            # Botón para descargar resultados (especialmente útil en Web Mode)
            if "bot_download_dir" in st.session_state and os.path.exists(st.session_state.bot_download_dir):
                files_in_dir = os.listdir(st.session_state.bot_download_dir)
                if files_in_dir:
                    st.markdown("---")
                    st.caption(f"Archivos en carpeta de descarga: {len(files_in_dir)}")
                    
                    if not is_native:
                        # Zip and Download for Web
                        import shutil
                        
                        if st.button("📦 Preparar ZIP de Descargas", key="btn_zip_bot_dl"):
                            zip_buffer = io.BytesIO()
                            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                                for root, dirs, files in os.walk(st.session_state.bot_download_dir):
                                    for file in files:
                                        abs_path = os.path.join(root, file)
                                        rel_path = os.path.relpath(abs_path, st.session_state.bot_download_dir)
                                        zf.write(abs_path, rel_path)
                            
                            zip_buffer.seek(0)
                            st.download_button(
                                label="📥 Descargar ZIP",
                                data=zip_buffer,
                                file_name=f"Bot_Descargas_{int(time.time())}.zip",
                                mime="application/zip"
                            )
                            
                            # Optional cleanup logic could go here
                    else:
                        # Native: Open folder
                        if st.button("📂 Abrir Carpeta", key="btn_open_bot_dl"):
                            import subprocess
                            try:
                                os.startfile(st.session_state.bot_download_dir)
                            except:
                                subprocess.Popen(["explorer", st.session_state.bot_download_dir])

            st.divider()
        
            st.subheader("2. Carga de Datos")
            uploaded_bot = st.file_uploader("Archivo Excel (.xlsx)", type=["xlsx", "xls"], key="bot_excel_uploader")
        
            if uploaded_bot:
                try:
                    import io
                    file_bytes = uploaded_bot.getvalue()
                    sheet_names = _get_excel_sheet_names(file_bytes)
                    sheet_bot = st.selectbox("Seleccione la Hoja", sheet_names, key="bot_sheet_sel")
                    
                    df_bot = _get_excel_preview(file_bytes, sheet_bot)
                    st.dataframe(df_bot.head(3), height=100)
                    excel_cols = df_bot.columns.tolist()
                    st.success(f"✅ {len(df_bot)} registros cargados.")
                except Exception as e:
                    st.error(f"Error leyendo Excel: {e}")


        with col_bot2:
            st.subheader("3. Definición de Pasos")

            # --- BOTONES DE SESIÓN ---
            if bot_perm in ["full", "edit"]:
                col_ses1, col_ses2 = st.columns(2)
                with col_ses1:
                    if st.button("🔄 Cargar Última Sesión", use_container_width=True, help="Recupera los pasos y flujos guardados automáticamente."):
                        ok, msg = bot_zeus.cargar_sesion()
                        if ok: 
                            st.toast(msg, icon="✅")
                            time.sleep(0.5)
                            # st.rerun()
                        else: 
                            st.error(msg)
                with col_ses2:
                    if st.button("💾 Guardar Sesión (Manual)", use_container_width=True, help="Fuerza el guardado de la configuración actual (aunque se guarda automáticamente al editar)."):
                        ok, msg = bot_zeus.guardar_sesion()
                        if ok: st.toast(msg, icon="💾")
                        else: st.error(msg)
            # -------------------------
        
            if bot_perm in ["full", "edit"]:
                st.caption("Configure la secuencia que el robot repetirá por cada fila.")
            
                # Selector de posición de inserción
                num_pasos = len(bot_zeus.get_pasos())
            
                # Opciones para inserción: [Final, 1, 2, ..., N]
                # Usamos un índice visual 1-based para el usuario, pero interno 0-based
                opts_insercion = ["Final"] + [str(i+1) for i in range(num_pasos)]
            
                # Guardar en session_state para persistencia
                if "pos_insercion" not in st.session_state:
                    st.session_state.pos_insercion = "Final"
                
                # UI para seleccionar donde insertar
                col_ins1, col_ins2 = st.columns([2, 1])
                with col_ins1:
                    st.info("💡 Los nuevos pasos se agregarán en la posición seleccionada.")
                with col_ins2:
                    st.session_state.pos_insercion = st.selectbox(
                        "Posición de inserción:", 
                        opts_insercion, 
                        index=0,
                        help="Seleccione 'Final' para agregar al final, o un número para insertar ANTES de ese paso."
                    )
            
                # Determinar índice real para pasar a las funciones
                indice_real = None
                if st.session_state.pos_insercion != "Final":
                    indice_real = int(st.session_state.pos_insercion) - 1 # Convertir a 0-based
            
                tab_click, tab_write, tab_key, tab_wait, tab_text, tab_alert, tab_scroll = st.tabs(["🖱️ Click", "✍️ Escribir", "⌨️ Tecla", "⏳ Espera", "🔤 Texto", "⚠️ Alerta", "📜 Scroll"])
            
                with tab_click:
                    st.write("Haga clic en el elemento en el navegador y presione:")
                    
                    saltar_click = st.checkbox("⏩ Saltar al Final si Éxito", key="saltar_click_foco", help="Si este paso se ejecuta correctamente, el bot saltará inmediatamente al último paso.")
                    
                    if st.button("Grabar Foco (Click)", use_container_width=True):
                        if use_agent:
                            try:
                                try: import agent_client
                                except ImportError: from src import agent_client
                                username = st.session_state.get("username", "admin")
                                task_id = agent_client.send_command(username, "bot_get_focused_element")
                                if task_id:
                                    with st.spinner("Obteniendo foco del agente..."):
                                        res = agent_client.wait_for_result(task_id)
                                        if res and "xpath" in res:
                                            xpath = res["xpath"]
                                            frames = res.get("frames", [])
                                            
                                            # Construct step manually
                                            paso = {
                                                "accion": "click",
                                                "xpath": xpath,
                                                "frames": frames,
                                                "descripcion": f"Click en {xpath}" + (f" (dentro de {len(frames)} frames)" if frames else "")
                                            }
                                            if saltar_click:
                                                paso["saltar_al_final"] = True
                                                paso["descripcion"] += " [⏩ SALTAR AL FINAL]"
                                            
                                            bot_zeus._insertar_paso(paso, indice_real)
                                            st.success(f"Paso agregado (Agente): {paso['descripcion']}")
                                        else:
                                            st.error(f"Error obteniendo foco: {res.get('error') if res else 'Sin respuesta'}")
                            except Exception as e:
                                st.error(f"Excepción agente: {e}")
                        else:
                            ok, msg = bot_zeus.agregar_paso_foco("click", indice_insercion=indice_real, saltar_al_final=saltar_click)
                            if ok: st.success(msg)
                            else: st.error(msg)
                
                    st.divider()
                    if st.button("🔄 Cambiar Ventana/Pestaña (Foco)", use_container_width=True, help="Cambia el foco del driver a la última ventana abierta o alterna entre ellas."):
                        if use_agent:
                            try:
                                try: import agent_client
                                except ImportError: from src import agent_client
                                username = st.session_state.get("username", "admin")
                                task_id = agent_client.send_command(username, "bot_switch_window", {"index": -1})
                                if task_id:
                                    with st.spinner("Cambiando ventana en agente..."):
                                        res = agent_client.wait_for_result(task_id)
                                        if res and res.get("status") == "success":
                                            # Add step locally for sequence recording
                                            ok, msg = bot_zeus.agregar_paso_cambiar_ventana(-1, indice_insercion=indice_real)
                                            if ok: st.success(f"{msg} (En Agente)")
                                            else: st.warning(f"Ventana cambiada en agente, pero error local: {msg}")
                                        else:
                                            st.error(f"Error cambiando ventana en agente: {res.get('error') if res else 'Sin respuesta'}")
                            except Exception as e:
                                st.error(f"Excepción agente: {e}")
                        else:
                             # Por defecto cambiamos a la última (-1). Si hay muchas, podría necesitarse un selector.
                             ok, msg = bot_zeus.agregar_paso_cambiar_ventana(-1, indice_insercion=indice_real) # No soporta saltar al final aun, o si? Verifiquemos bot_zeus
                             if ok: st.success(msg)
                             else: st.error(msg)

                with tab_write:
                    st.markdown("##### ✍️ Escribir Dato")
                    col_sel = st.selectbox("Columna Excel:", excel_cols, disabled=len(excel_cols)==0)
                    
                    saltar_write = st.checkbox("⏩ Saltar al Final si Éxito", key="saltar_write", help="Si este paso se ejecuta correctamente, el bot saltará inmediatamente al último paso.")
                    
                    if st.button("Grabar Foco (Escribir)", use_container_width=True, disabled=len(excel_cols)==0):
                        if use_agent:
                            try:
                                try: import agent_client
                                except ImportError: from src import agent_client
                                username = st.session_state.get("username", "admin")
                                task_id = agent_client.send_command(username, "bot_get_focused_element")
                                if task_id:
                                    with st.spinner("Obteniendo foco del agente..."):
                                        res = agent_client.wait_for_result(task_id)
                                        if res and "xpath" in res:
                                            xpath = res["xpath"]
                                            frames = res.get("frames", [])
                                            
                                            # Construct step manually
                                            paso = {
                                                "accion": "escribir",
                                                "columna": col_sel,
                                                "xpath": xpath,
                                                "frames": frames,
                                                "descripcion": f"Escribir [{col_sel}] en {xpath}" + (f" (dentro de {len(frames)} frames)" if frames else "")
                                            }
                                            if saltar_write:
                                                paso["saltar_al_final"] = True
                                                paso["descripcion"] += " [⏩ SALTAR AL FINAL]"
                                            
                                            bot_zeus._insertar_paso(paso, indice_real)
                                            st.success(f"Paso agregado (Agente): {paso['descripcion']}")
                                        else:
                                            st.error(f"Error obteniendo foco: {res.get('error') if res else 'Sin respuesta'}")
                            except Exception as e:
                                st.error(f"Excepción agente: {e}")
                        else:
                            ok, msg = bot_zeus.agregar_paso_foco("escribir", columna=col_sel, indice_insercion=indice_real, saltar_al_final=saltar_write)
                            if ok: st.success(msg)
                            else: st.error(msg)
                
                    st.divider()
                    st.markdown("##### 📅 Escribir Fecha")
                    st.caption("Toma una fecha del Excel, la formatea y la escribe en el campo seleccionado.")
                    col_date, col_fmt = st.columns([2, 1])
                    with col_date:
                        col_date_sel = st.selectbox("Columna Fecha:", excel_cols, disabled=len(excel_cols)==0, key="sel_col_date")
                    with col_fmt:
                        fmt_date = st.text_input("Formato:", value="%d/%m/%Y", help="Ej: %d/%m/%Y (01/12/2023) o %Y-%m-%d (2023-12-01)")
                    
                    saltar_date = st.checkbox("⏩ Saltar al Final si Éxito", key="saltar_date", help="Si este paso se ejecuta correctamente, el bot saltará inmediatamente al último paso.")

                    if st.button("Grabar Foco (Escribir Fecha)", use_container_width=True, disabled=len(excel_cols)==0):
                        if use_agent:
                            try:
                                try: import agent_client
                                except ImportError: from src import agent_client
                                username = st.session_state.get("username", "admin")
                                task_id = agent_client.send_command(username, "bot_get_focused_element")
                                if task_id:
                                    with st.spinner("Obteniendo foco del agente..."):
                                        res = agent_client.wait_for_result(task_id)
                                        if res and "xpath" in res:
                                            xpath = res["xpath"]
                                            frames = res.get("frames", [])
                                            
                                            # Construct step manually
                                            paso = {
                                                "accion": "escribir_fecha",
                                                "columna": col_date_sel,
                                                "formato": fmt_date,
                                                "xpath": xpath,
                                                "frames": frames,
                                                "descripcion": f"Escribir Fecha [{col_date_sel}] en {xpath} (Fmt: {fmt_date})" + (f" (dentro de {len(frames)} frames)" if frames else "")
                                            }
                                            if saltar_date:
                                                paso["saltar_al_final"] = True
                                                paso["descripcion"] += " [⏩ SALTAR AL FINAL]"
                                            
                                            bot_zeus._insertar_paso(paso, indice_real)
                                            st.success(f"Paso agregado (Agente): {paso['descripcion']}")
                                        else:
                                            st.error(f"Error obteniendo foco: {res.get('error') if res else 'Sin respuesta'}")
                            except Exception as e:
                                st.error(f"Excepción agente: {e}")
                        else:
                            ok, msg = bot_zeus.agregar_paso_foco("escribir_fecha", columna=col_date_sel, formato=fmt_date, indice_insercion=indice_real, saltar_al_final=saltar_date)
                            if ok: st.success(msg)
                            else: st.error(msg)

                    st.divider()
                    st.markdown("##### 🧹 Borrar Dato")
                    st.caption("Seleccione el campo en el navegador y presione el botón para vaciarlo.")
                    
                    saltar_clean = st.checkbox("⏩ Saltar al Final si Éxito", key="saltar_clean", help="Si este paso se ejecuta correctamente, el bot saltará inmediatamente al último paso.")

                    if st.button("Grabar Foco (Limpiar Campo)", use_container_width=True):
                         if use_agent:
                             try:
                                 try: import agent_client
                                 except ImportError: from src import agent_client
                                 username = st.session_state.get("username", "admin")
                                 task_id = agent_client.send_command(username, "bot_get_focused_element")
                                 if task_id:
                                     with st.spinner("Obteniendo foco del agente..."):
                                         res = agent_client.wait_for_result(task_id)
                                         if res and "xpath" in res:
                                             xpath = res["xpath"]
                                             frames = res.get("frames", [])
                                             
                                             # Construct step manually
                                             paso = {
                                                 "accion": "limpiar_campo",
                                                 "xpath": xpath,
                                                 "frames": frames,
                                                 "descripcion": f"Limpiar Campo {xpath}" + (f" (dentro de {len(frames)} frames)" if frames else "")
                                             }
                                             if saltar_clean:
                                                 paso["saltar_al_final"] = True
                                                 paso["descripcion"] += " [⏩ SALTAR AL FINAL]"
                                             
                                             bot_zeus._insertar_paso(paso, indice_real)
                                             st.success(f"Paso agregado (Agente): {paso['descripcion']}")
                                         else:
                                             st.error(f"Error obteniendo foco: {res.get('error') if res else 'Sin respuesta'}")
                             except Exception as e:
                                 st.error(f"Excepción agente: {e}")
                         else:
                             ok, msg = bot_zeus.agregar_paso_foco("limpiar_campo", indice_insercion=indice_real, saltar_al_final=saltar_clean)
                             if ok: st.success(msg)
                             else: st.error(msg)
            
                with tab_key:
                    key_sel = st.selectbox("Tecla Especial:", ["ENTER", "TAB", "ESCAPE", "DOWN", "UP"])
                    
                    saltar_key = st.checkbox("⏩ Saltar al Final si Éxito", key="saltar_key", help="Si este paso se ejecuta correctamente, el bot saltará inmediatamente al último paso.")
                    
                    if st.button("Agregar Tecla", use_container_width=True):
                        ok, msg = bot_zeus.agregar_paso_tecla(key_sel, indice_insercion=indice_real, saltar_al_final=saltar_key)
                        if ok: st.success(msg)
                        else: st.error(msg)
            
                with tab_wait:
                    sec = st.number_input("Segundos:", min_value=0.1, value=1.0, step=0.5)
                    
                    saltar_wait = st.checkbox("⏩ Saltar al Final si Éxito", key="saltar_wait", help="Si este paso se ejecuta correctamente, el bot saltará inmediatamente al último paso.")
                    
                    if st.button("Agregar Espera", use_container_width=True):
                        ok, msg = bot_zeus.agregar_paso_espera(sec, indice_insercion=indice_real, saltar_al_final=saltar_wait)
                        if ok: st.success(msg)
                        else: st.error(msg)
                
                    st.markdown("---")
                    st.write("🔧 Utilidades:")
                
                    # Mostrar conteo de ventanas activas para depuración
                    try:
                        drv = bot_zeus.obtener_driver(create_if_missing=False)
                        if drv:
                            n_wins = len(drv.window_handles)
                            st.caption(f"Ventanas detectadas por el sistema: {n_wins}")
                    except:
                        pass

                    col_win1, col_win2 = st.columns(2)
                    with col_win1:
                        if st.button("🔄 Ir a Popup (Última)", use_container_width=True):
                            if use_agent:
                                try:
                                    try: import agent_client
                                    except ImportError: from src import agent_client
                                    username = st.session_state.get("username", "admin")
                                    task_id = agent_client.send_command(username, "bot_switch_window", {"index": -1})
                                    if task_id:
                                        with st.spinner("Cambiando ventana en agente..."):
                                            res = agent_client.wait_for_result(task_id)
                                            if res and res.get("status") == "success":
                                                ok, msg = bot_zeus.agregar_paso_cambiar_ventana(-1, indice_insercion=indice_real)
                                                if ok: st.success(f"{msg} (En Agente)")
                                                else: st.warning(f"Ventana cambiada en agente, pero error local: {msg}")
                                            else:
                                                st.error(f"Error: {res.get('error') if res else 'Sin respuesta'}")
                                except Exception as e:
                                    st.error(f"Excepción agente: {e}")
                            else:
                                 ok, msg = bot_zeus.agregar_paso_cambiar_ventana(-1, indice_insercion=indice_real)
                                 if ok: st.success(msg)
                                 else: st.error(msg)
                    with col_win2:
                        if st.button("🏠 Ir a Principal (0)", use_container_width=True):
                            if use_agent:
                                try:
                                    try: import agent_client
                                    except ImportError: from src import agent_client
                                    username = st.session_state.get("username", "admin")
                                    task_id = agent_client.send_command(username, "bot_switch_window", {"index": 0})
                                    if task_id:
                                        with st.spinner("Cambiando ventana en agente..."):
                                            res = agent_client.wait_for_result(task_id)
                                            if res and res.get("status") == "success":
                                                ok, msg = bot_zeus.agregar_paso_cambiar_ventana(0, indice_insercion=indice_real)
                                                if ok: st.success(f"{msg} (En Agente)")
                                                else: st.warning(f"Ventana cambiada en agente, pero error local: {msg}")
                                            else:
                                                st.error(f"Error: {res.get('error') if res else 'Sin respuesta'}")
                                except Exception as e:
                                    st.error(f"Excepción agente: {e}")
                            else:
                                 ok, msg = bot_zeus.agregar_paso_cambiar_ventana(0, indice_insercion=indice_real)
                                 if ok: st.success(msg)
                                 else: st.error(msg)
                    
                with tab_text:
                    st.info("💡 Click en elementos por su contenido (Texto, ID, Título o Clase).")
                
                    tipo_txt = st.radio("Modo:", ["Texto Fijo", "Texto Dinámico (Desde Excel)", "Seleccionar Opción de Lista (Excel)", "Selector Personalizado (XPath)", "🎯 Selector Visual (Beta)"], horizontal=True)
                
                    txt_buscar = None
                    es_dinamico = False
                    tipo_seleccion = "texto"
                    tag_val = "*" # Default

                    if tipo_txt == "🎯 Selector Visual (Beta)":
                        st.markdown("""
                        <div style="background-color: #e8f5e9; padding: 10px; border-radius: 5px; border: 1px solid #4caf50;">
                            <b>Instrucciones:</b><br>
                            1. Presione "Activar Selector".<br>
                            2. Vaya al navegador. El cursor cambiará a una cruz.<br>
                            3. Haga <b>CLICK</b> en el elemento deseado (se pondrá verde).<br>
                            4. Regrese aquí y presione "Capturar Selección".
                        </div>
                        """, unsafe_allow_html=True)
                    
                        col_sel1, col_sel2 = st.columns(2)
                        with col_sel1:
                            if st.button("🚀 Activar Selector", use_container_width=True):
                                if use_agent:
                                    try:
                                        try: import agent_client
                                        except ImportError: from src import agent_client
                                        username = st.session_state.get("username", "admin")
                                        task_id = agent_client.send_command(username, "bot_start_visual_selector")
                                        if task_id:
                                            with st.spinner("Iniciando selector en agente..."):
                                                res = agent_client.wait_for_result(task_id)
                                                if res and "message" in res:
                                                    st.toast(res["message"], icon="🎯")
                                                    st.session_state.selector_activo = True
                                                else:
                                                    st.error(f"Error agente: {res.get('error') if res else 'Sin respuesta'}")
                                    except Exception as e:
                                        st.error(f"Error agente: {e}")
                                else:
                                    ok, msg = bot_zeus.iniciar_selector_visual()
                                    if ok: 
                                        st.toast(msg, icon="🎯")
                                        st.session_state.selector_activo = True
                                    else: st.error(msg)
                    
                        with col_sel2:
                            if st.button("✅ Capturar Selección", use_container_width=True):
                                if use_agent:
                                    try:
                                        try: import agent_client
                                        except ImportError: from src import agent_client
                                        username = st.session_state.get("username", "admin")
                                        task_id = agent_client.send_command(username, "bot_get_visual_selection")
                                        if task_id:
                                            with st.spinner("Obteniendo selección del agente..."):
                                                res = agent_client.wait_for_result(task_id)
                                                if res and "xpath" in res and res["xpath"]:
                                                    st.success(f"Elemento capturado: {res['xpath']}")
                                                    st.session_state.xpath_capturado = res["xpath"]
                                                else:
                                                    st.warning("No se detectó selección en el agente.")
                                    except Exception as e:
                                        st.error(f"Error agente: {e}")
                                else:
                                    ok, result = bot_zeus.obtener_seleccion_visual()
                                    if ok:
                                        # Autollenar campos para que el usuario guarde
                                        st.success(f"Elemento capturado: {result}")
                                        st.session_state.xpath_capturado = result
                                    else:
                                        st.warning("No se detectó selección o navegador desconectado.")

                        if "xpath_capturado" in st.session_state:
                            txt_buscar = st.text_input("XPath Capturado:", value=st.session_state.xpath_capturado)
                            tipo_seleccion = "xpath"
                            st.caption("Puede editar el XPath si es necesario antes de agregar el paso.")

                    elif tipo_txt == "Selector Personalizado (XPath)":
                        st.caption("Escriba el XPath exacto del elemento. Ej: `//div[@class='modal']//button`")
                        txt_buscar = st.text_input("XPath del elemento:")
                        tipo_seleccion = "xpath"
                    
                    elif tipo_txt == "Texto Fijo":
                        txt_buscar = st.text_input("Texto visible o Atributo (ej. 'Cerrar', 'Guardar', 'X'):")
                    
                        # Selector de Tipo de Elemento (Tag) con agrupación inteligente
                        tag_map = {
                            "Cualquiera (*)": "*",
                            "Botón / Enlace / Input": "*[self::button or self::a or self::input]",
                            "Icono / Imagen (i, svg, img, span)": "*[self::i or self::svg or self::img or self::span]",
                            "Texto / Etiqueta (div, p, label, h1-h6)": "*[self::div or self::p or self::label or self::span or self::h1 or self::h2 or self::h3 or self::h4 or self::h5 or self::h6]",
                            "--- Específicos ---": "*",
                            "Button": "button",
                            "A (Enlace)": "a",
                            "Input": "input",
                            "Div": "div",
                            "Span": "span",
                            "I (Icono)": "i",
                            "SVG": "svg"
                        }
                        tag_sel = st.selectbox("Tipo de Elemento (Ayuda a diferenciar):", list(tag_map.keys()))
                        tag_val = tag_map[tag_sel]

                    elif tipo_txt == "Seleccionar Opción de Lista (Excel)":
                        if excel_cols:
                            st.info("""
                            ℹ️ **Instrucciones:**
                            1. Agregue un paso previo de **Click** para abrir la lista desplegable (si no es nativa).
                            2. Use esta opción para seleccionar el texto exacto proveniente del Excel.
                            3. El bot resaltará visualmente (borde rojo) el elemento encontrado antes de hacer click.
                            """)
                            txt_buscar = st.selectbox("Columna Excel con la opción:", excel_cols, key="col_list_excel")
                            es_dinamico = True
                            # Permitimos cualquier tag, pero sugerimos buscar en todo (*)
                            tag_map = {
                                "Cualquiera (*)": "*",
                                "Opción (<option>)": "option", 
                                "Elemento de Lista (<li>)": "li",
                                "Div / Span": "*[self::div or self::span]",
                                "Enlace (<a>)": "a"
                            }
                            tag_sel = st.selectbox("Tipo de Elemento (Opcional):", list(tag_map.keys()), key="tag_list_excel")
                            tag_val = tag_map[tag_sel]
                        else:
                            st.warning("⚠️ Cargue un Excel primero para usar esta opción.")
                            txt_buscar = None
                    
                    else: # Texto Dinámico (Desde Excel)
                        if excel_cols:
                            txt_buscar = st.selectbox("Columna Excel con el texto:", excel_cols)
                            es_dinamico = True
                        else:
                            st.warning("⚠️ Cargue un Excel primero para usar esta opción.")
                            txt_buscar = None
                    
                        # Filtro de tag en dinámico
                        tag_map = {
                            "Cualquiera (*)": "*", 
                            "Botón / Enlace": "*[self::button or self::a or self::input]",
                            "Icono": "*[self::i or self::svg or self::img or self::span]"
                        }
                        tag_sel = st.selectbox("Tipo de Elemento:", list(tag_map.keys()), key="tag_dyn")
                        tag_val = tag_map[tag_sel]

                    # Checkbox de exactitud y case-insensitive
                    col_opts1, col_opts2 = st.columns(2)
                    with col_opts1:
                        default_exacto = True if tipo_txt == "Seleccionar Opción de Lista (Excel)" else False
                        exacto = st.checkbox("Búsqueda exacta (todo el texto debe coincidir)", value=default_exacto, disabled=(tipo_seleccion=="xpath"))
                    with col_opts2:
                        # Por defecto True para listas para ser más robusto, False para otros
                        default_ignore = True if tipo_txt == "Seleccionar Opción de Lista (Excel)" else False
                        ignore_case = st.checkbox("Ignorar Mayúsculas/Minúsculas", value=default_ignore, disabled=(tipo_seleccion=="xpath"))
                    
                    # --- NUEVO: RESTRICT TO CONTAINER (MULTI-SELECTION) ---
                    xpath_cont = None # Legacy support
                    lista_contenedores = st.session_state.get("multi_contenedores", [])

                    if tipo_seleccion != "xpath":
                        with st.expander("🎯 Restringir a Contenedores Visuales (Opcional - Avanzado)"):
                            st.caption("Úselo para limitar la búsqueda a una o varias zonas específicas. Puede agregar hasta 7 contenedores visuales. El bot buscará el texto en cualquiera de ellos.")
                            
                            col_c1, col_c2 = st.columns([1, 1])
                            with col_c1:
                                if st.button("👁️ Activar Selector Visual", key="btn_vis_cont"):
                                    try:
                                        ok, msg = bot_zeus.iniciar_selector_visual()
                                        if ok:
                                            st.toast(msg)
                                        else:
                                            st.error(msg)
                                    except Exception as e:
                                        st.error(f"Error activando selector: {e}")

                            with col_c2:
                                if st.button("📥 Capturar y Agregar a Lista", key="btn_cap_cont"):
                                    try:
                                        ok, result = bot_zeus.obtener_seleccion_visual()
                                        if ok:
                                            if "multi_contenedores" not in st.session_state:
                                                st.session_state["multi_contenedores"] = []
                                            
                                            if result not in st.session_state["multi_contenedores"]:
                                                st.session_state["multi_contenedores"].append(result)
                                                st.success("¡Contenedor agregado!")
                                            else:
                                                st.warning("Este contenedor ya está en la lista.")
                                        else:
                                            st.warning(f"No se detectó selección: {result}")
                                    except Exception as e:
                                        st.error(f"Error capturando: {e}")
                            
                            # Mostrar lista de contenedores
                            if "multi_contenedores" in st.session_state and st.session_state["multi_contenedores"]:
                                st.markdown("##### Contenedores Seleccionados:")
                                
                                def _del_cont(idx_to_del):
                                    st.session_state["multi_contenedores"].pop(idx_to_del)

                                for idx, xp in enumerate(st.session_state["multi_contenedores"]):
                                    c_idx, c_xp, c_del = st.columns([0.5, 4, 0.5])
                                    c_idx.text(f"#{idx+1}")
                                    c_xp.code(xp, language="text")
                                    c_del.button("🗑️", key=f"del_cont_{idx}", on_click=_del_cont, args=(idx,))
                                
                                def _clean_all_cont():
                                    st.session_state["multi_contenedores"] = []

                                st.button("Limpiar Todos los Contenedores", key="clean_all_cont", on_click=_clean_all_cont)
                                
                                lista_contenedores = st.session_state["multi_contenedores"]
                            else:
                                st.info("No hay contenedores seleccionados. Se buscará en toda la página.")
                            
                            # --- OPCIÓN: USAR COMO ÍNDICE ---
                            st.markdown("---")
                            usar_indice = st.checkbox("📍 Usar valor (Excel/Texto) como Índice de Contenedor", key="chk_usar_indice", help="Si se marca, el valor (ej: '1') indicará cual contenedor clickear de la lista (1º, 2º...). Si no se marca, se buscará el texto DENTRO de los contenedores.")
                            if usar_indice:
                                st.info("ℹ️ El bot leerá el número (1, 2, 3...) y hará click en el contenedor correspondiente de la lista de arriba.")


                    saltar_text = st.checkbox("⏩ Saltar al Final si Éxito", key="saltar_text", help="Si este paso se ejecuta correctamente, el bot saltará inmediatamente al último paso.")

                    if st.button("Agregar Click por Texto/Selector", use_container_width=True, disabled=not txt_buscar):
                        # Pasamos la lista completa como xpath_contenedor (el backend ya lo maneja)
                        usar_idx_val = st.session_state.get("chk_usar_indice", False) if tipo_seleccion != "xpath" else False
                        
                        ok, msg = bot_zeus.agregar_paso_click_texto(txt_buscar, exacto, es_dinamico, tag_val, tipo_seleccion, ignore_case, indice_insercion=indice_real, xpath_contenedor=lista_contenedores if lista_contenedores else None, usar_indice_contenedor=usar_idx_val, saltar_al_final=saltar_text)
                        if ok: 
                            st.success(msg)
                            # Limpiar lista tras agregar
                            if "multi_contenedores" in st.session_state:
                                del st.session_state["multi_contenedores"]
                            if "temp_xpath_contenedor" in st.session_state:
                                del st.session_state["temp_xpath_contenedor"]
                        else: st.error(msg)
            
                with tab_alert:
                    st.info("Úselo cuando aparezca una ventana emergente nativa (Aceptar/Cancelar).")
                    
                    saltar_alert = st.checkbox("⏩ Saltar al Final si Éxito", key="saltar_alert", help="Si este paso se ejecuta correctamente, el bot saltará inmediatamente al último paso.")
                
                    col_al1, col_al2 = st.columns(2)
                    with col_al1:
                        if st.button("✅ Aceptar Alerta (OK)", use_container_width=True):
                            ok, msg = bot_zeus.agregar_paso_alerta("aceptar", indice_insercion=indice_real, saltar_al_final=saltar_alert)
                            if ok: st.success(msg)
                            else: st.error(msg)
                    with col_al2:
                        if st.button("❌ Cancelar Alerta", use_container_width=True):
                            ok, msg = bot_zeus.agregar_paso_alerta("cancelar", indice_insercion=indice_real, saltar_al_final=saltar_alert)
                            if ok: st.success(msg)
                            else: st.error(msg)

                with tab_scroll:
                    st.info("Desplazarse por la página (útil si el elemento está oculto).")
                
                    tipo_scroll = st.radio("Tipo de movimiento:", ["Abajo (Pixels)", "Arriba (Pixels)", "Ir al Final (Bottom)", "Ir al Inicio (Top)"], horizontal=True)
                
                    cant_px = 0
                    if "Pixels" in tipo_scroll:
                        cant_px = st.number_input("Cantidad de Pixels:", min_value=10, max_value=5000, value=300, step=50)
                
                    saltar_scroll = st.checkbox("⏩ Saltar al Final si Éxito", key="saltar_scroll", help="Si este paso se ejecuta correctamente, el bot saltará inmediatamente al último paso.")

                    if st.button("Agregar Scroll", use_container_width=True):
                        # Mapear selección a parámetros
                        direccion = "abajo"
                        if "Arriba" in tipo_scroll: direccion = "arriba"
                        elif "Final" in tipo_scroll: direccion = "fin"
                        elif "Inicio" in tipo_scroll: direccion = "inicio"
                    
                        ok, msg = bot_zeus.agregar_paso_scroll(direccion, cant_px, indice_insercion=indice_real, saltar_al_final=saltar_scroll)
                        if ok: st.success(msg)
                        else: st.error(msg)
            else:
                st.info("🔒 **Modo Ejecución**: La edición y creación de pasos está deshabilitada para su rol. Puede cargar flujos existentes y ejecutarlos.")


            st.divider()
            st.subheader("Pasos Memorizados")
        
            # --- SECCIÓN GUARDAR/CARGAR ---
            col_io1, col_io2 = st.columns(2)
            with col_io1:
                # Botón de descarga
                pasos_actuales = bot_zeus.get_pasos()
                if pasos_actuales:
                    json_str = json.dumps(pasos_actuales, indent=4)
                
                    if bot_perm in ["full", "edit"]:
                        st.download_button(
                            label="💾 Guardar Flujo",
                            data=json_str,
                            file_name="flujo_bot_zeus.json",
                            mime="application/json",
                            use_container_width=True
                        )
                    else:
                        st.button("💾 Guardar Flujo (Deshabilitado)", disabled=True, use_container_width=True, help="Requiere permisos de Edición")
            with col_io2:
                # Carga de archivo
                uploaded_flow = st.file_uploader("📂 Cargar Flujo (.json)", type=["json"], label_visibility="collapsed")
                if uploaded_flow:
                    try:
                        # Usar un hash del contenido para evitar recargas infinitas
                        file_bytes = uploaded_flow.getvalue()
                        file_hash = hash(file_bytes)
                    
                        if st.session_state.get("last_flow_hash") != file_hash:
                            data = json.loads(file_bytes.decode("utf-8"))
                            ok, msg = bot_zeus.cargar_pasos_externos(data)
                            if ok:
                                st.session_state.last_flow_hash = file_hash
                                st.toast(msg, icon="✅")
                            else:
                                st.error(msg)
                    except Exception as e:
                        st.error(f"Error cargando flujo: {e}")
        
            # --- SECCIÓN FLUJO ALTERNATIVO (MULTIPLE) ---
            if bot_perm in ["full", "edit", "execute"]:
                with st.expander("🔀 Flujos Condicionales (Avanzado)"):
                    st.info("Configure flujos adicionales que se ejecutarán si se cumple una condición específica. Se evalúan en orden (1 -> 2 -> 3). Si ninguno cumple, se usa el Principal.")
                    
                    tabs_cond = st.tabs(["Alternativo 1", "Alternativo 2", "Alternativo 3"])
                    
                    for idx_tab, tab in enumerate(tabs_cond):
                        with tab:
                            # Recuperar estado actual
                            flujo_actual = bot_zeus.get_flujo_condicional(idx_tab)
                            
                            col_conf1, col_conf2 = st.columns(2)
                            
                            with col_conf1:
                                st.markdown(f"**Carga del Flujo {idx_tab + 1}**")
                                uploaded_alt = st.file_uploader(f"📂 Cargar JSON (Alt {idx_tab + 1})", type=["json"], key=f"uploader_alt_{idx_tab}")
                                if uploaded_alt:
                                    try:
                                        # Usar hash para evitar loop
                                        f_bytes = uploaded_alt.getvalue()
                                        f_hash = hash(f_bytes)
                                        
                                        if st.session_state.get(f"last_hash_alt_{idx_tab}") != f_hash:
                                            data_alt = json.loads(f_bytes.decode("utf-8"))
                                            # Validar estructura básica
                                            if isinstance(data_alt, list):
                                                # Guardar en memoria
                                                bot_zeus.update_flujo_condicional(idx_tab, pasos=data_alt)
                                                st.session_state[f"last_hash_alt_{idx_tab}"] = f_hash
                                                st.toast(f"Flujo Alternativo {idx_tab+1} cargado", icon="✅")
                                            else:
                                                st.error("Formato JSON inválido (debe ser una lista de pasos).")
                                    except Exception as e:
                                        st.error(f"Error: {e}")

                                # Mostrar estado actual
                                if flujo_actual and flujo_actual.get("pasos"):
                                    st.success(f"✅ {len(flujo_actual['pasos'])} pasos cargados.")
                                else:
                                    st.warning("⚠️ Sin pasos definidos.")

                            with col_conf2:
                                st.markdown(f"**Condición de Ejecución {idx_tab + 1}**")
                                st.caption("Defina cuándo debe ejecutarse este flujo en lugar del principal.")
                                
                                current_cond = flujo_actual.get("condicion", {}) if flujo_actual else {}
                                
                                tipo_cond = st.radio("Tipo de Condición:", 
                                                    ["Valor Excel (Simple)", "Valor Excel (Múltiple)", "Texto Visible en Pantalla"], 
                                                    key=f"tipo_cond_{idx_tab}")
                                
                                new_cond = {}
                                
                                if tipo_cond == "Texto Visible en Pantalla":
                                    txt = st.text_input("Texto a buscar:", 
                                                      value=current_cond.get("valor", "") if current_cond.get("tipo") == "texto" else "",
                                                      key=f"txt_cond_{idx_tab}")
                                    if txt:
                                        new_cond = {"tipo": "texto", "valor": txt}
                                
                                elif tipo_cond == "Valor Excel (Simple)":
                                    cols = df_bot.columns.tolist() if df_bot is not None else []
                                    curr_col = current_cond.get("columna", "")
                                    curr_val = current_cond.get("valor", "") if current_cond.get("tipo") == "excel" else ""
                                    
                                    if not cols: st.warning("Cargue Excel primero.")
                                    
                                    c_sel = st.selectbox("Columna", [""] + cols, 
                                                       index=cols.index(curr_col) + 1 if curr_col in cols else 0,
                                                       key=f"sel_col_{idx_tab}")
                                    v_sel = st.text_input("Valor(es) activador(es) (separar con |):",
                                                        value=curr_val,
                                                        placeholder="Ej: Si | Yes",
                                                        key=f"val_cond_{idx_tab}")
                                    
                                    if c_sel and v_sel:
                                        new_cond = {"tipo": "excel", "valor": v_sel, "columna": c_sel}

                                elif tipo_cond == "Valor Excel (Múltiple)":
                                    cols = df_bot.columns.tolist() if df_bot is not None else []
                                    if not cols: st.warning("Cargue Excel primero.")
                                    
                                    # Preparar datos para el editor
                                    raw_reglas = current_cond.get("reglas", []) if current_cond.get("tipo") == "excel_multi" else []
                                    # Convertir a formato UI (Capitalized keys)
                                    ui_reglas = [{"Columna": r.get("columna", ""), "Valor": r.get("valor", "")} for r in raw_reglas]
                                    if not ui_reglas: ui_reglas = [{"Columna": "", "Valor": ""}]
                                    
                                    df_reglas = pd.DataFrame(ui_reglas)
                                    
                                    column_config = {
                                        "Columna": st.column_config.SelectboxColumn(
                                            "Columna Excel",
                                            help="Seleccione la columna a validar",
                                            width="medium",
                                            options=cols,
                                            required=True
                                        ),
                                        "Valor": st.column_config.TextColumn(
                                            "Valor(es) Trigger",
                                            help="Separe valores con |",
                                            width="medium",
                                            required=True
                                        )
                                    }
                                    
                                    st.caption("Defina múltiples reglas. TODAS deben cumplirse (AND).")
                                    edited_df = st.data_editor(
                                        df_reglas,
                                        column_config=column_config,
                                        num_rows="dynamic",
                                        key=f"editor_reglas_{idx_tab}",
                                        use_container_width=True,
                                        hide_index=True
                                    )
                                    
                                    # Procesar resultado
                                    reglas_final = []
                                    for _, row in edited_df.iterrows():
                                        c = row.get("Columna")
                                        v = row.get("Valor")
                                        if c and v:
                                            reglas_final.append({"columna": c, "valor": v})
                                    
                                    if reglas_final:
                                        new_cond = {"tipo": "excel_multi", "reglas": reglas_final}

                                # Botón guardar condicion
                                if st.button("💾 Actualizar Condición", key=f"btn_save_cond_{idx_tab}"):
                                    bot_zeus.update_flujo_condicional(idx_tab, condicion=new_cond)
                                    st.toast("Condición actualizada", icon="💾")
                                    st.rerun()

                                # Mostrar resumen condicion
                                if flujo_actual and flujo_actual.get("condicion"):
                                    c = flujo_actual["condicion"]
                                    if c.get("tipo") == "excel":
                                        st.info(f"Si Columna **{c.get('columna')}** es **{c.get('valor')}** -> Ejecuta este flujo.")
                                    elif c.get("tipo") == "excel_multi":
                                        r_txt = " Y ".join([f"[{r['columna']}='{r['valor']}']" for r in c.get('reglas', [])])
                                        st.info(f"Si **{r_txt}** -> Ejecuta este flujo.")
                                    elif c.get("tipo") == "texto":
                                        st.info(f"Si Texto **{c.get('valor')}** es visible -> Ejecuta este flujo.")
                                else:
                                    st.warning("Sin condición definida (Nunca se ejecutará).")

                            st.divider()
                            if st.button(f"🗑️ Limpiar Alternativo {idx_tab+1}", key=f"clean_{idx_tab}"):
                                bot_zeus.update_flujo_condicional(idx_tab, pasos=[], condicion={})
                                st.session_state.pop(f"last_hash_alt_{idx_tab}", None)
                                st.toast(f"Flujo Alternativo {idx_tab+1} limpiado", icon="🗑️")

            st.divider()

            # Display steps with management controls
            pasos = bot_zeus.get_pasos()
            if not pasos:
                st.info("No hay pasos grabados.")
            else:
                st.write("---")
                st.write("**Gestión de Pasos:**")
            
                # Header row
                if bot_perm in ["full", "edit"]:
                    h1, h2, h3, h4, h5 = st.columns([5, 1, 1, 1, 1])
                    h1.markdown("**Descripción**")
                    h2.markdown("**Opcional**")
                    h3.markdown("**Subir**")
                    h4.markdown("**Bajar**")
                    h5.markdown("**Borrar**")
                else:
                    h1, h2 = st.columns([5, 1])
                    h1.markdown("**Descripción**")
                    h2.markdown("**Estado**")
            
                # Helper callbacks for step management
                def _toggle_opt(idx):
                    bot_zeus.alternar_opcional_paso(idx)

                def _move_up(idx):
                    bot_zeus.mover_paso(idx, -1)

                def _move_down(idx):
                    bot_zeus.mover_paso(idx, 1)

                def _delete_step(idx):
                    bot_zeus.eliminar_paso_indice(idx)
                
                def _clear_all_steps():
                    bot_zeus.limpiar_pasos()

                for i, p in enumerate(pasos):
                    # Check status
                    es_opcional = p.get("opcional", False)
                    desc_texto = f"{i+1}. {p.get('descripcion', 'Paso')}"
                    if es_opcional:
                        desc_texto += " (OPCIONAL)"
                
                    if bot_perm in ["full", "edit"]:
                        c1, c2, c3, c4, c5 = st.columns([5, 1, 1, 1, 1])
                        # Description with index
                        c1.text(desc_texto)

                        # Optional Toggle
                        btn_label = "⚠️" if es_opcional else "✅"
                        help_text = "Click para marcar como Opcional" if not es_opcional else "Click para marcar como Obligatorio"
                        c2.button(btn_label, key=f"btn_opt_{i}", help=help_text, on_click=_toggle_opt, args=(i,))
                    
                        # Move Up
                        if i > 0: 
                            c3.button("⬆️", key=f"btn_up_{i}", on_click=_move_up, args=(i,))
                    
                        # Move Down
                        if i < len(pasos) - 1:
                            c4.button("⬇️", key=f"btn_down_{i}", on_click=_move_down, args=(i,))
                            
                        # Delete
                        c5.button("🗑️", key=f"btn_del_{i}", on_click=_delete_step, args=(i,))
                    else:
                        c1, c2 = st.columns([5, 1])
                        c1.text(desc_texto)
                        c2.text("⚠️ Opcional" if es_opcional else "✅ Obligatorio")
                
                # Removed manual rerun logic as on_click handles it
            
                st.divider()
            
                if bot_perm in ["full", "edit"]:
                    c1, c2 = st.columns(2)
                    with c2:
                        st.button("Borrar Todo", type="primary", on_click=_clear_all_steps)

        st.markdown("---")
        st.subheader("4. Ejecución")
    
        delay = st.slider("Velocidad (segundos entre pasos):", 0.0, 3.0, 0.5)
        
        # --- LÓGICA DE EJECUCIÓN EN HILO (PARA PERMITIR STOP) ---
        if "bot_running" not in st.session_state:
            st.session_state.bot_running = False
        if "bot_logs" not in st.session_state:
            st.session_state.bot_logs = []
            
        # Wrapper para el hilo
        def run_bot_thread(df, delay):
            if use_agent:
                try:
                    try: import agent_client
                    except ImportError: from src import agent_client
                    username = st.session_state.get("username", "admin")
                    
                    # Convert df to records
                    data = df.to_dict(orient="records")
                    # Get steps
                    steps = bot_zeus.get_pasos()
                    
                    if "bot_logs" in st.session_state:
                        st.session_state.bot_logs.append("🚀 Enviando secuencia al agente...")
                    
                    # Send command
                    # Note: delay is not passed currently to agent command, could add it
                    task_id = agent_client.send_command(username, "bot_run_sequence", {"steps": steps, "data": data})
                    
                    if task_id:
                        if "bot_logs" in st.session_state:
                            st.session_state.bot_logs.append("⏳ Ejecutando en agente (espere el resultado final)...")
                        
                        # Wait with long timeout (30 min)
                        res = agent_client.wait_for_result(task_id, timeout=1800)
                        
                        if "bot_logs" in st.session_state:
                            if res:
                                if "logs" in res:
                                    for log in res["logs"]:
                                        st.session_state.bot_logs.append(log)
                                
                                if "error" in res:
                                    st.session_state.bot_logs.append(f"❌ Error reportado por agente: {res['error']}")
                                else:
                                    st.session_state.bot_logs.append("✅ Secuencia finalizada en agente.")
                            else:
                                st.session_state.bot_logs.append("❌ Timeout o error de comunicación con agente.")
                    else:
                        if "bot_logs" in st.session_state:
                            st.session_state.bot_logs.append("❌ No se pudo enviar la tarea al agente.")

                except Exception as e:
                    if "bot_logs" in st.session_state:
                        st.session_state.bot_logs.append(f"❌ Error en hilo agente: {e}")
                finally:
                    if "bot_running" in st.session_state:
                        st.session_state.bot_running = False

            else:
                try:
                    for msg in bot_zeus.ejecutar_secuencia(df, delay_pasos=delay):
                        if "bot_logs" in st.session_state:
                            st.session_state.bot_logs.append(msg)
                        else:
                            # Fallback si se pierde el contexto
                            pass
                        time.sleep(0.01)
                except Exception as e:
                    if "bot_logs" in st.session_state:
                        st.session_state.bot_logs.append(f"❌ Error en hilo: {e}")
                finally:
                    if "bot_running" in st.session_state:
                        st.session_state.bot_running = False

        col_ctrl1, col_ctrl2 = st.columns([1, 1])
        
        with col_ctrl1:
            # Botón INICIAR
            if not st.session_state.bot_running:
                if st.button("▶️ Iniciar Secuencia Masiva", use_container_width=True, disabled=not (uploaded_bot and len(bot_zeus.get_pasos()) > 0)):
                    st.session_state.bot_running = True
                    st.session_state.bot_logs = []
                    st.session_state.bot_finished_shown = False # Reset flag de toast
                    
                    t = threading.Thread(target=run_bot_thread, args=(df_bot, delay))
                    try:
                        from streamlit.runtime.scriptrunner import add_script_run_ctx
                        add_script_run_ctx(t)
                    except:
                        pass
                    t.start()
                    # Rerun triggered automatically by loop below
            else:
                 st.info("🚀 Ejecutando... (Espere o Detenga)")

        with col_ctrl2:
            # Botón DETENER
            if st.session_state.bot_running:
                if st.button("⛔ DETENER EJECUCIÓN", type="primary", use_container_width=True):
                    bot_zeus.detener_ejecucion()
                    st.toast("Deteniendo...", icon="🛑")
        
        # Mostrar Logs
        if st.session_state.bot_logs:
            st.code("\n".join(st.session_state.bot_logs[-15:]), language="text")
        
        # Auto-refresh mientras corre
        if st.session_state.bot_running:
            time.sleep(1)
            st.rerun()
            
        # --- POST EJECUCIÓN (Popups/Mensajes) ---
        if not st.session_state.bot_running and st.session_state.bot_logs:
            last_log = st.session_state.bot_logs[-1]
            
            # Solo mostrar feedback si acabamos de terminar (usando flag bot_finished_shown)
            if not st.session_state.get("bot_finished_shown", False):
                
                # Caso ERROR
                last_error = bot_zeus.get_ultimo_error()
                if last_error:
                    st.error(f"❌ El proceso se detuvo con errores.")
                    st.toast(f"Error: {last_error}", icon="❌")
                    with st.expander("Ver detalle del error", expanded=True):
                        st.write(last_error)
                    st.session_state.bot_finished_shown = True
                    
                # Caso ÉXITO (o Detenido por usuario pero sin error crash)
                elif "Fin del proceso" in last_log or "interrumpido" in last_log:
                    if "interrumpido" in last_log:
                        st.warning("⚠️ Proceso detenido por el usuario.")
                        st.toast("Proceso Detenido", icon="🛑")
                    else:
                        st.success("✅ Proceso Completado Exitosamente.")
                        st.toast("Proceso Completado", icon="✅")
                        st.balloons()
                    st.session_state.bot_finished_shown = True
            
            # Botón de descarga siempre visible al final
            st.download_button("Descargar Log Completo", "\n".join(st.session_state.bot_logs), file_name="log_bot_zeus.txt")

        with st.expander("ℹ️ Guía Rápida"):
            st.markdown("""
            1. **Conexión**: Abra el navegador y conéctese a su aplicación web.
            2. **Datos**: Cargue el Excel que contiene las filas a procesar.
            3. **Definir Pasos**: Construya la secuencia de acciones que se repetirá por cada fila:
               - **Escribir**: Seleccione una columna del Excel, haga clic en el campo del navegador donde va el dato, y presione *Grabar Foco (Escribir)*.
               - **Click**: Haga clic en el botón o elemento del navegador, regrese aquí y presione *Grabar Foco (Click)*.
               - **Tecla**: Agregue pulsaciones como ENTER o TAB para navegar entre campos.
            4. **Verificar**: Revise la lista de "Pasos Memorizados". Use "Deshacer" si se equivoca.
            5. **Ejecutar**: Ajuste la velocidad y presione "Iniciar Secuencia Masiva".
            """)
