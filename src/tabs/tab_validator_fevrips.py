import streamlit as st
import requests
import json
import os
import pandas as pd
from datetime import datetime
import io
import subprocess
import socket
import urllib.parse
import shutil
import time

try:
    import database as db
except ImportError:
    from src import database as db

try:
    from gui_utils import abrir_dialogo_carpeta_nativo, abrir_dialogo_archivo_nativo, update_path_key, render_path_selector, render_file_selector, render_download_button
except ImportError:
    try:
        from src.gui_utils import abrir_dialogo_carpeta_nativo, abrir_dialogo_archivo_nativo, update_path_key, render_path_selector, render_file_selector, render_download_button
    except ImportError:
        abrir_dialogo_carpeta_nativo = None
        abrir_dialogo_archivo_nativo = None
        def update_path_key(key, new_path, widget_key=None):
             if new_path:
                 st.session_state[key] = new_path
                 if widget_key:
                     st.session_state[widget_key] = new_path
        
        def render_path_selector(label, key, default_path=None, help_text=None, omit_checkbox=False):
            st.warning("render_path_selector no disponible")
            return default_path

        def render_file_selector(label, key, default_path=None, help_text=None, file_types=None, omit_checkbox=False):
            st.warning("render_file_selector no disponible")
            return default_path
            
        def render_download_button(file_path, key, label="📥 Descargar"):
            st.warning("render_download_button no disponible")

# --- HELPERS ---

def check_process_running(process_name):
    """Verifica si un proceso está corriendo en Windows"""
    try:
        # Usar tasklist es más compatible que psutil si no está instalado
        # Aseguramos encoding utf-8 o similar
        output = subprocess.check_output('tasklist /FI "IMAGENAME eq {}"'.format(process_name), shell=True)
        # Convertir bytes a string de forma segura
        output_str = output.decode('latin-1', errors='ignore')
        return process_name.lower() in output_str.lower()
    except:
        return False

def check_docker_available():
    """Verifica si Docker está instalado y corriendo"""
    try:
        subprocess.check_output("docker --version", shell=True)
        # Check daemon status
        subprocess.check_output("docker info", shell=True)
        return True
    except:
        return False

def get_container_status(container_name="fevrips-api"):
    """Verifica estado del contenedor"""
    try:
        output = subprocess.check_output(f'docker ps -a --filter "name={container_name}" --format "{{{{.Status}}}}"', shell=True).decode().strip()
        if not output:
            return "not_found"
        if "Up" in output:
            return "running"
        return "stopped"
    except:
        return "error"


def find_active_port(ports=[5000, 5001, 9443, 8080, 443]):
    """Busca en qué puerto está escuchando el servicio local"""
    for port in ports:
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.settimeout(0.5)
                if s.connect_ex(('localhost', port)) == 0:
                    return port
        except:
            pass
    return None

def update_user_config_helper(key, value):
    """Helper seguro para actualizar configuración de usuario"""
    if st.session_state.get("username"):
        try:
            current_config = st.session_state.get("app_config", {})
            current_config[key] = value
            # Intentar guardar en DB si existe la función
            if hasattr(db, "update_user_config"):
                db.update_user_config(st.session_state.username, current_config)
            st.session_state.app_config = current_config
        except Exception as e:
            print(f"Error actualizando config: {e}")

def update_config_key(key, new_value, widget_key=None):
    if new_value:
        update_user_config_helper(key, new_value)
        if widget_key:
            st.session_state[widget_key] = new_value
        # st.rerun() # Removed to avoid callback error

def on_click_update_config(key, new_value):
    if new_value:
        update_user_config_helper(key, new_value)

@st.dialog("Generar CUV Masivo (FEVRIPS)")
def dialog_generar_cuv():
    st.write("Genera el Código Único de Validación (CUV) enviando los archivos RIPS a la API local.")
    
    is_native = st.session_state.get("force_native_mode", True)
    
    current_global_path = st.session_state.get("current_path", os.getcwd())
    if not current_global_path: current_global_path = os.getcwd()

    target_cuv_path = render_path_selector(
        label="Carpeta de Facturas (Raíz)",
        key="path_cuv",
        help_text="Se usará la carpeta seleccionada para buscar los archivos RIPS.",
        default_path=current_global_path
    )
    
    # --- NUEVA SECCIÓN: MODO SERVICIO WEB ---
    if st.session_state.get("is_web_service_mode", False):
        st.info("🌐 Modo Servicio Web Activo: El validador se ejecuta en segundo plano.")
    # ----------------------------------------
    
    st.markdown("---")
    st.write("Configuración API (FEVRIPS)")
    
    # Modo forzado a Docker AWS
    conn_mode = "Docker AWS (Interno)"
    
    st.success("☁️ Modo Docker AWS (Interno) Activo")
    
    # Usamos HTTP interno (puerto 5000) apuntando a la IP de red local del host de AWS para que Docker lo resuelva
    default_url = "http://172.17.0.1:5000/api/Validacion/ValidarArchivo"
    
    # Permitir editar la URL (útil si el host cambia o si se quiere probar localhost)
    api_url = st.text_input("URL Interna API:", value=default_url, help="URL del servicio FEVRIPS.")
    
    # Verificación de resolución de nombre para ayudar al usuario
    if "fevrips-api" in api_url:
        try:
            socket.gethostbyname("fevrips-api")
        except socket.gaierror:
            st.warning("⚠️ No se puede resolver el host 'fevrips-api'. Intente usar 'localhost' o la IP pública del servidor en su lugar.")
            st.info("💡 Sugerencia: Cambie 'fevrips-api' por 'localhost' en la caja de texto superior.")

    st.markdown("---")

    # Credenciales de Autenticación SISPRO
    st.write("🔐 Autenticación SISPRO (Opcional)")
    with st.expander("Configurar Credenciales de Login", expanded=False):
        # Inferir URL de auth basada en la API configurada
        default_auth_url = "https://172.17.0.1:9443/api/Auth/LoginSISPRO"
        if api_url and "/api/" in api_url:
            base_url = api_url.split("/api/")[0]
            # Si el usuario usa el puerto 5000 para validar, usamos el 9443 (https) para el login como indica el manual
            if ":5000" in base_url:
                base_url = base_url.replace("http://", "https://").replace(":5000", ":9443")
            default_auth_url = f"{base_url}/api/Auth/LoginSISPRO"
            
        # Cargar credenciales guardadas si existen
        saved_creds = st.session_state.app_config.get("sispro_creds", {})
        default_user = saved_creds.get("usuario", "")
        default_nit = saved_creds.get("nit", "")
        default_pass = saved_creds.get("clave", "")
        
        # Si no hay credenciales guardadas pero el usuario las proporcionó en el chat (caso actual)
        # las usamos como sugerencia inicial
        if not default_user: default_user = "31996431"
        if not default_nit: default_nit = "900438792"
        if not default_pass: default_pass = "Oportunidad2026*"
            
        auth_url = st.text_input("URL Login:", value=default_auth_url, placeholder="Ej: https://mi-servidor.com/api/Auth/LoginSISPRO")
        
        col_auth1, col_auth2 = st.columns(2)
        with col_auth1:
            usuario = st.text_input("Usuario (Cédula/Número):", value=default_user)
            tipo_usuario = st.selectbox("Tipo de Usuario:", ["RE", "PIN", "PINx", "PIE"], index=0, help="RE: Representante Entidad, PIN: Profesional Independiente, etc.")
        with col_auth2:
            clave = st.text_input("Contraseña:", type="password", value=default_pass)
            nit = st.text_input("NIT:", value=default_nit)
    
        save_creds = st.checkbox("Guardar credenciales para futuros inicios de sesión", value=True)
    
        # Auto-Login check
        auto_login_token = st.session_state.get("temp_token")
     
        # Si hay credenciales por defecto, mostramos el botón más destacado
        login_label = "🔑 Obtener Token (Login Rápido)" if (default_user and default_pass) else "🔑 Obtener Token"
     
        if st.button(login_label, type="primary" if (default_user and default_pass) else "secondary"):
            if not auth_url:
                st.warning("⚠️ La URL de autenticación no puede estar vacía.")
            elif not usuario or not clave or not nit:
                st.warning("Complete todos los campos de autenticación.")
            else:
                # Guardar credenciales si el usuario lo solicitó
                if save_creds:
                    creds = {
                        "usuario": usuario,
                        "nit": nit,
                        "clave": clave,
                        "tipo_usuario": tipo_usuario
                    }
                    update_user_config_helper("sispro_creds", creds)
                    
                try:
                    verify_ssl = True
                    if auth_url.startswith("https://localhost") or auth_url.startswith("https://127.0.0.1") or auth_url.startswith("https://172.17.0.1"):
                        verify_ssl = False
                        requests.packages.urllib3.disable_warnings()
                        
                    payload = {
                        "tipo": "CC", # Asumimos CC por defecto
                        "numero": usuario,
                        "clave": clave,
                        "nit": nit,
                        "tipoUsuario": tipo_usuario
                    }
                    
                    with st.spinner("Autenticando..."):
                        r = requests.post(auth_url, json=payload, verify=verify_ssl, timeout=10)
                        
                        if r.status_code == 200:
                            resp_json = r.json()
                            token_val = resp_json.get("token") or resp_json.get("Token")
                            if token_val:
                                st.session_state.temp_token = token_val
                                st.success("¡Autenticación Exitosa! Token obtenido.")
                            else:
                                st.error(f"No se encontró token en la respuesta: {resp_json}")
                        else:
                            st.error(f"Error Login ({r.status_code}): {r.text}")
                except requests.exceptions.ConnectionError:
                    st.error("❌ No se pudo conectar al servidor de Autenticación.")
                    st.warning("⚠️ Asegúrese de que el contenedor Docker FEV-RIPS esté corriendo.")
                except Exception as e:
                    st.error(f"Error de conexión: {e}")

    # Usar token obtenido o manual
    token_default = st.session_state.get("temp_token", "")
    token = st.text_input("Token de Autorización (Bearer):", value=token_default, type="password")
    
    if st.button("🔌 Probar Conexión"):
        if conn_mode in ["Local (Nativo)", "Docker Integrado", "Docker AWS (Interno)"]:
             try:
                 parsed_u = urllib.parse.urlparse(api_url)
                 check_host = parsed_u.hostname or "localhost"
                 check_port = parsed_u.port or 443
                 
                 sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                 sock.settimeout(1.0)
                 res = sock.connect_ex((check_host, check_port))
                 sock.close()
                 
                 if res == 0:
                     st.success(f"✅ Puerto {check_port} Detectado (Servicio Activo)")
                 else:
                     st.error(f"❌ Puerto {check_port} Cerrado (Servicio Inactivo)")
                     if conn_mode == "Local (Nativo)":
                         st.info("💡 **Solución (Sin Docker):**\n1. Asegúrese de tener instalado el aplicativo **'Mecanismo Único de Validación (Cliente-Servidor)'** de MinSalud.\n2. Ejecútelo antes de validar.\n3. Verifique que el puerto configurado coincida con el de la aplicación (usualmente 9443 o 5000).")
                     elif conn_mode == "Docker Integrado":
                         st.info("💡 **Solución (Docker):**\n1. Verifique que el contenedor 'fevrips-api' esté corriendo en la sección de arriba.\n2. Si está detenido, haga clic en 'Iniciar'.")
                     elif conn_mode == "Docker AWS (Interno)":
                         st.info("💡 **Solución (AWS):**\n1. Verifique que el contenedor 'fevrips-api' esté corriendo (docker-compose logs fevrips-api).\n2. Verifique que el puerto 5000 esté expuesto internamente.")
             except Exception as e:
                 st.error(f"Error diagnóstico: {e}")

    st.markdown("---")
    if st.button("🚀 Iniciar Validación Masiva", type="primary"):
        if not api_url:
            st.error("URL de API no configurada")
            return
            
        # --- AUTO LOGIN IF TOKEN IS EMPTY ---
        if not token:
            st.info("🔄 Obteniendo token de autenticación automáticamente...")
            try:
                login_payload = {
                    "tipo": "CC",
                    "numero": "31996431",
                    "clave": "Oportunidad2026*",
                    "nit": "900438792",
                    "tipoUsuario": "RE"
                }
                login_url = default_auth_url
                r = requests.post(login_url, json=login_payload, timeout=10)
                if r.status_code == 200:
                    resp_json = r.json()
                    token = resp_json.get("token") or resp_json.get("Token")
                    if token:
                        st.session_state.temp_token = token
                        st.success("✅ Token obtenido exitosamente.")
                    else:
                        st.error("Error: No se recibió un token en la respuesta de autenticación.")
                        return
                else:
                    st.error(f"Error de Autenticación automática ({r.status_code}): {r.text}")
                    return
            except Exception as e:
                st.error(f"Error conectando al servicio de autenticación: {e}")
                return
        # ------------------------------------
            
        # Lógica de procesamiento masivo
        path_input = target_cuv_path
        
        # Modificación para Modo Nativo: Permitir ruta aunque no exista en el entorno del servidor (Docker/Cloud)
        is_native = st.session_state.get("force_native_mode", True)
        
        if not path_input or (not is_native and not os.path.exists(path_input)):
            st.error("Ruta de archivos inválida.")
            return

        files_to_process = []
        try:
            # Si la ruta existe localmente (o estamos en el mismo entorno), listamos
            if os.path.exists(path_input):
                for f in os.listdir(path_input):
                     if f.lower().endswith(".json") and not f.startswith("Resultados") and not f.startswith("Resp_"):
                         files_to_process.append(os.path.join(path_input, f))
            elif is_native:
                # Si estamos en modo nativo y la ruta no es accesible, usamos el Agente Local
                try:
                    from src.agent_client import send_command, wait_for_result
                    username = st.session_state.get("username", "default")
                    
                    st.info("Conectando con Agente Local para validación remota...")
                    
                    params = {
                        "base_path": path_input,
                        "api_url": api_url,
                        "token": token,
                        "verify_ssl": True
                    }
                    # Ajuste de verify_ssl si api_url es localhost (aunque el agente lo maneja, lo pasamos explícito)
                    if api_url.startswith("https://localhost") or api_url.startswith("https://127.0.0.1"):
                        params["verify_ssl"] = False

                    task_id = send_command(username, "validate_rips", params)
                    
                    if task_id:
                        with st.spinner("El Agente Local está validando los archivos... Esto puede tomar varios minutos."):
                            res = wait_for_result(task_id, timeout=600)
                            
                        if "error" in res:
                            st.error(f"Error del Agente: {res['error']}")
                            return
                        else:
                            # Procesar resultados del agente
                            agent_results = res.get("results", [])
                            agent_gen_files = res.get("generated_files", [])
                            
                            st.success(f"✅ Validación completada por Agente. Procesados: {res.get('processed', 0)}")
                            
                            if agent_gen_files:
                                st.info(f"📂 Se han generado {len(agent_gen_files)} archivos de resultados directamente en la carpeta: {path_input}")
                            
                            # Mostrar DataFrame
                            df_res = pd.DataFrame(agent_results)
                            st.dataframe(df_res)
                            return
                    else:
                         st.error("No se pudo enviar la tarea al Agente Local.")
                         return

                except ImportError:
                    st.error("Librería de Agente no encontrada. No se puede procesar en Modo Nativo sin acceso directo a la carpeta.")
                    return
                except Exception as e:
                    st.error(f"Error inesperado con el Agente: {e}")
                    return
        except Exception as e:
            st.error(f"Error al listar archivos: {e}")
            return
        
        if not files_to_process:
            st.warning("No se encontraron archivos JSON para validar.")
            return

        st.info(f"Procesando {len(files_to_process)} archivos...")
        progress_bar = st.progress(0)
        results = []
        
        headers = {"Content-Type": "application/json"}
        if token:
            headers["Authorization"] = f"Bearer {token}"
            
        verify_ssl = True
        if api_url.startswith("https://localhost") or api_url.startswith("https://127.0.0.1"):
            verify_ssl = False
            requests.packages.urllib3.disable_warnings()

        generated_files = []
        for i, full_path in enumerate(files_to_process):
            fname = os.path.basename(full_path)
            res_row = {"Archivo": fname, "Estado": "Pendiente", "CUV": "", "Mensaje": ""}
            
            try:
                with open(full_path, "r", encoding="utf-8") as f_obj:
                    fdata = json.load(f_obj)
                
                r_val = requests.post(api_url, json=fdata, headers=headers, verify=verify_ssl, timeout=60)
                res_row["Estado"] = r_val.status_code
                
                try:
                    r_json = r_val.json()
                    res_row["CUV"] = r_json.get("cuv") or r_json.get("CUV") or ""
                    res_row["Mensaje"] = json.dumps(r_json, ensure_ascii=False)[:200]
                    
                    # Guardar Archivos Extra (Lógica recuperada)
                    try:
                        factura_num = fdata.get('numFactura', os.path.splitext(fname)[0])
                        provider_id = fdata.get('numDocumentoIdentificacionObligado', '999')
                        
                        # 1. ResultadosLocales
                        f_loc_name = f"ResultadosLocales_{factura_num}.json"
                        f_loc_path = os.path.join(path_input, f_loc_name)
                        with open(f_loc_path, "w", encoding="utf-8") as f_out:
                            json.dump(r_json, f_out, indent=2, ensure_ascii=False)
                        generated_files.append(f_loc_path)
                        
                        # 2. ResultadosMSPS
                        f_msps_name = f"ResultadosMSPS_{factura_num}_ID{provider_id}_R.json"
                        f_msps_path = os.path.join(path_input, f_msps_name)
                        with open(f_msps_path, "w", encoding="utf-8") as f_out:
                            json.dump(r_json, f_out, indent=2, ensure_ascii=False)
                        generated_files.append(f_msps_path)
                        
                    except Exception as e:
                        print(f"Error guardando extras: {e}")
                        
                except:
                    res_row["Mensaje"] = r_val.text[:200]
                    
            except Exception as e:
                res_row["Estado"] = "Error"
                res_row["Mensaje"] = str(e)
            
            results.append(res_row)
            progress_bar.progress((i + 1) / len(files_to_process))
            
        st.success("Validación completada.")
        
        # --- DESCARGA DE RESULTADOS ---
        if generated_files:
            try:
                # Prepare a clean download folder
                timestamp = int(time.time())
                session_id = st.session_state.get("session_id", "default")
                temp_dir = os.path.join("temp_downloads", f"fevrips_results_{session_id}_{timestamp}")
                os.makedirs(temp_dir, exist_ok=True)
                
                # Copy generated files
                count_copied = 0
                for f_path in generated_files:
                    if os.path.exists(f_path):
                        shutil.copy2(f_path, temp_dir)
                        count_copied += 1
                
                if count_copied > 0:
#                     render_download_button(temp_dir, "btn_download_fevrips", "📦 Descargar Resultados (ZIP)", cleanup=not is_native)
                    pass
                else:
                    st.warning("No se pudieron copiar los archivos generados para descarga.")
                    
            except Exception as e:
                st.error(f"Error preparando descarga: {e}")
        # ------------------------------
        
        df_res = pd.DataFrame(results)
        st.dataframe(df_res)

    if st.button("Cerrar"):
        st.session_state.active_fevrips_dialog = None
        st.rerun()

def render(container=None):
    if container is None:
        container = st.container()
        
    with container:
        st.header("🏥 Validador FEVRIPS - Generación de CUV")
        st.markdown("""
        Esta herramienta permite validar archivos RIPS y generar el Código Único de Validación (CUV) 
        utilizando el validador oficial FEVRIPS (local o remoto).
        """)
        
        col_val_1, col_val_2 = st.columns([1, 2])
        
        with col_val_1:
            st.markdown('<div class="group-box">', unsafe_allow_html=True)
            st.markdown('<div class="group-title-left">Acciones</div>', unsafe_allow_html=True)
            
            if st.button("🆔 Generar CUV (Masivo)", use_container_width=True):
                 st.session_state.active_fevrips_dialog = "generar_cuv"
                 st.rerun()
                 
            st.markdown('</div>', unsafe_allow_html=True)
            
        with col_val_2:
            st.info("ℹ️ Asegúrese de que el servicio Docker FEVRIPS esté en ejecución si utiliza el modo local.")

    active_dialog = st.session_state.get("active_fevrips_dialog")
    if active_dialog == "generar_cuv":
        dialog_generar_cuv()
