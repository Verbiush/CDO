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
        help_text="Se usará la carpeta seleccionada para buscar los archivos RIPS."
    )
    
    # --- NUEVA SECCIÓN: MODO SERVICIO WEB ---
    if st.session_state.get("is_web_service_mode", False):
        st.info("🌐 Modo Servicio Web Activo: El validador se ejecuta en segundo plano.")
    # ----------------------------------------
    
    st.markdown("---")
    st.write("Configuración API (FEVRIPS)")
    
    col_mode, col_url = st.columns([0.3, 0.7])
    
    with col_mode:
        conn_mode = st.radio("Modo de Conexión:", ["Docker Integrado", "Docker AWS (Interno)"], key="fevrips_mode", help="Seleccione el modo Docker para validación FEVRIPS.")
        
        # Checkbox para activar modo "Servicio Web" (oculta UI nativa)
        # if conn_mode == "Local (Nativo)": ... Removed
    
    api_url = ""
    if conn_mode == "Servidor Remoto":
        st.info("ℹ️ Ingrese la URL del servidor donde está alojado el API FEVRIPS.")

    with col_url:
        # --- MODO DOCKER INTEGRADO ---
        if conn_mode == "Docker Integrado":
            is_docker_ok = check_docker_available()
            if not is_docker_ok:
                st.error("❌ Docker Desktop no detectado.")
                st.caption("Asegúrese de instalar y ejecutar Docker Desktop.")
                api_url = ""
            else:
                status = get_container_status("fevrips-api")
                col_d_status, col_d_action = st.columns([0.6, 0.4])
                
                with col_d_status:
                    if status == "running":
                        st.success("✅ Contenedor 'fevrips-api' Activo")
                        api_url = "https://localhost:9443/api/Validacion/ValidarArchivo"
                        st.caption(f"URL: `{api_url}`")
                    elif status == "stopped":
                        st.warning("⚠️ Contenedor detenido")
                    else:
                        st.info("ℹ️ Contenedor no instalado")

                with col_d_action:
                    if status == "running":
                        if st.button("⏹️ Detener", key="stop_docker"):
                            subprocess.run("docker stop fevrips-api", shell=True)
                            # st.rerun()
                    elif status == "stopped":
                        if st.button("▶️ Iniciar", key="start_docker", type="primary"):
                            subprocess.run("docker start fevrips-api", shell=True)
                            st.toast("Iniciando contenedor...")
                            time.sleep(3)
                            st.rerun()
                    else: # not_found
                        if st.button("🛠️ Instalar", key="install_docker", type="primary"):
                            cmd = 'docker run -d --name fevrips-api -p 9443:5100 -e ASPNETCORE_ENVIRONMENT=Docker -e ASPNETCORE_URLS="https://+:5100;http://+:5000" fevripsacr.azurecr.io/minsalud.fevrips.apilocal:latest'
                            subprocess.run(cmd, shell=True)
                            st.toast("Descargando e iniciando (puede tardar)...")
                            time.sleep(10)
                            # st.rerun()
        
        # --- MODO DOCKER AWS (INTERNO) ---
        elif conn_mode == "Docker AWS (Interno)":
            st.success("☁️ Modo Nube Activo")
            
            # Usamos HTTP interno (puerto 5000) por defecto
            default_url = "http://fevrips-api:5000/api/Validacion/ValidarArchivo"
            
            # Permitir editar la URL (útil si el host cambia o si se quiere probar localhost)
            api_url = st.text_input("URL Interna API:", value=default_url, help="URL del servicio FEVRIPS.")
            
            # Verificación de resolución de nombre para ayudar al usuario
            if "fevrips-api" in api_url:
                try:
                    socket.gethostbyname("fevrips-api")
                except socket.gaierror:
                    st.warning("⚠️ No se puede resolver el host 'fevrips-api'.")
                    st.info("💡 Si está ejecutando la aplicación LOCALMENTE (fuera de la red Docker), por favor seleccione el modo **'Docker Integrado'** en su lugar.")

        # Modo Local: Permitir URL personalizada para flexibilidad (Docker o Nativo Windows)
        elif conn_mode == "Local (Nativo)":
            # Auto-detectar puerto si no se ha configurado
            current_url = st.session_state.app_config.get("local_api_url", "https://localhost:9443/api/Validacion/ValidarArchivo")
            
            detected_port = find_active_port()
            if detected_port:
                # Si detectamos un puerto diferente al que está en la URL, sugerimos el cambio
                url_port = "9443"
                if ":" in current_url.split("/")[-1]: # Check for port in hostname part? No, usually in netloc
                    pass 
                
                protocol = "https" if detected_port in [443, 9443, 5001] else "http"
                
                # Check if current URL matches detected port
                if str(detected_port) not in current_url:
                    new_url = f"{protocol}://localhost:{detected_port}/api/Validacion/ValidarArchivo"
                    current_url = new_url
                    st.toast(f"🔎 Servicio detectado en puerto {detected_port}. URL actualizada.", icon="🔗")
            
            # Opción para cambiar entre predefinidos y manual
            use_custom = st.checkbox("Editar URL del Servicio Local", value=False)
            
            if use_custom:
                api_url = st.text_input("URL del Endpoint Local:", value=current_url, help="Si usa el aplicativo nativo de MinSalud, verifique el puerto (ej: 9443, 5000, 443).")
            else:
                api_url = current_url
                st.info(f"Apuntando a: `{api_url}`")
                
            # Guardar URL si cambió
            if api_url != st.session_state.app_config.get("local_api_url"):
                update_user_config_helper("local_api_url", api_url)
                
            st.caption("ℹ️ Para validar sin Docker, debe tener ejecutándose el aplicativo 'Mecanismo Único de Validación (Cliente-Servidor)' de MinSalud.")
            
            # --- STARTUP AUTOMÁTICO VALIDADOR LOCAL ---
            if st.session_state.get("is_web_service_mode", False):
                st.markdown("#### ☁️ Servicio de Validación (Background)")
            else:
                st.markdown("#### ⚙️ Control de Servicio Local")
            
            # Default path inferido
            default_exe_path = r"C:\Program Files\FEVRIPS Validador Local\FVE.ValidadorLocal.exe"
            current_exe_path = st.session_state.app_config.get("fevrips_exe_path", default_exe_path)
            
            # Usar render_file_selector estandarizado
            # Usamos una key específica para el selector
            selector_key = "fevrips_exe_path_selector"
            
            # Asegurar que el selector inicie con el valor actual de config si no está en session
            if selector_key not in st.session_state:
                st.session_state[selector_key] = current_exe_path
                
            new_exe_path = render_file_selector(
                "Ruta del Ejecutable FEVRIPS:", 
                key=selector_key,
                help_text="Seleccione el ejecutable de FEVRIPS (FVE.ValidadorLocal.exe).",
                file_types=[("Ejecutables", "*.exe"), ("Todos", "*.*")]
            )
            
            # Sincronizar con config si cambió
            if new_exe_path != current_exe_path:
                update_user_config_helper("fevrips_exe_path", new_exe_path)
                current_exe_path = new_exe_path
            
            exe_path = current_exe_path
            exe_name = os.path.basename(exe_path)
            is_running = check_process_running(exe_name)
            
            # --- AUTO-START EN MODO SERVICIO WEB ---
            if st.session_state.get("is_web_service_mode", False) and not is_running:
                st.info("🔄 Iniciando servicio de validación en segundo plano...")
                if os.path.exists(exe_path):
                    try:
                        import subprocess
                        startupinfo = subprocess.STARTUPINFO()
                        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                        startupinfo.wShowWindow = 0 # SW_HIDE
                        subprocess.Popen([exe_path], startupinfo=startupinfo, cwd=os.path.dirname(exe_path), shell=False)
                        time.sleep(5)
                        st.rerun()
                    except:
                        pass
            # ----------------------------------------
            
            col_status, col_action = st.columns([0.6, 0.4])
            with col_status:
                if is_running:
                    st.success(f"✅ Servicio '{exe_name}' en ejecución")
                else:
                    st.warning(f"⚠️ Servicio '{exe_name}' detenido")
            
            with col_action:
                if not is_running:
                    # Opción para iniciar silenciosamente (tipo servicio web)
                    start_silent = st.checkbox("Iniciar en modo silencioso (Background)", value=True, help="Ejecuta el validador sin mostrar la ventana, como un servicio web.")
                    
                    if st.button("▶️ Iniciar Servicio Local", type="primary"):
                        if os.path.exists(exe_path):
                            try:
                                import subprocess
                                
                                # Configurar startup info para ocultar ventana si se solicita
                                startupinfo = None
                                if start_silent:
                                    startupinfo = subprocess.STARTUPINFO()
                                    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                                    startupinfo.wShowWindow = 0 # SW_HIDE
                                
                                # Lanzar proceso
                                subprocess.Popen([exe_path], startupinfo=startupinfo, cwd=os.path.dirname(exe_path), shell=False)
                                
                                st.toast("🚀 Iniciando servicio en segundo plano...", icon="⏳")
                                time.sleep(8) # Dar más tiempo para arranque completo
                                st.rerun()
                            except Exception as e:
                                st.error(f"Error al iniciar: {e}")
                        else:
                            st.error(f"❌ No se encuentra el archivo: {exe_path}")
                else:
                    if st.session_state.get("is_web_service_mode", False):
                        st.success("✅ Servicio Web Activo y Escuchando")
                    else:
                        st.caption("✅ Servicio activo.")
            # -------------------------------------------

        else:
            # En modo remoto, permitimos ingresar cualquier URL
            api_url = st.text_input("URL del Endpoint de Validación:", value="", placeholder="Ej: https://mi-servidor.com/api/Validacion/ValidarArchivo")
            
        st.caption("ℹ️ Tip: Puede verificar la disponibilidad en `/swagger/index.html` (ej: https://localhost:9443/swagger/index.html)")

        if conn_mode == "Local (Nativo)":
            with st.expander("🛠️ Ayuda para Modo Nativo (Sin Docker)", expanded=False):
                 st.markdown("""
                 **Cómo validar sin Docker:**
                 1. Debe tener instalado el **Mecanismo Único de Validación (Cliente-Servidor)** de MinSalud.
                 2. Puede descargarlo desde el [Micrositio de Facturación Electrónica (SISPRO)](https://www.sispro.gov.co/central-financiamiento/Pages/facturacion-electronica.aspx).
                 3. Ejecute la aplicación en su equipo antes de intentar generar el CUV.
                 4. Verifique el puerto en la configuración de la aplicación (usualmente 9443, 5000 o 443).
                 5. Si el puerto es diferente, habilite la opción **'Editar URL del Servicio Local'** arriba y actualice la dirección.
                 """)

        if conn_mode == "Docker Integrado":
            with st.expander("🛠️ Asistente de Configuración (Docker)", expanded=False):
                st.info("Siga estos pasos si desea usar el contenedor oficial de FEVRIPS.")
                
                # Paso 1: Generar Archivos
                st.markdown("#### 1. Generar Archivos de Configuración")
                if st.button("📄 Crear docker-compose-fevrips.yml"):
                    compose_content = """version: '3.4'
services:
  fevrips-api:
    image: fevripsacr.azurecr.io/minsalud.fevrips.apilocal:latest
    container_name: fevrips-api
    ports:
      - "9443:5100"
    environment:
      - ASPNETCORE_ENVIRONMENT=Docker
      - ASPNETCORE_URLS=https://+:5100;http://+:5000
      - ASPNETCORE_Kestrel__Certificates__Default__Password=fevrips2024*
      - ASPNETCORE_Kestrel__Certificates__Default__Path=/certificates/server.pfx
    volumes:
      - C:/Certificates:/certificates
"""
                    try:
                        with open("docker-compose-fevrips.yml", "w") as f:
                            f.write(compose_content)
                        st.success("Archivo 'docker-compose-fevrips.yml' creado en la carpeta del proyecto.")
                    except Exception as e:
                        st.error(f"Error creando archivo: {e}")
                
                # Paso 2: Certificados
                st.markdown("#### 2. Generar Certificados SSL")
                st.markdown("El contenedor requiere un certificado SSL para funcionar en HTTPS.")
                if st.button("🔐 Generar Certificados (Requiere OpenSSL)"):
                    try:
                        # Buscar script dinámicamente
                        base_dirs = [
                            os.getcwd(),
                            os.path.join(os.getcwd(), "OrganizadorArchivos"),
                            os.path.dirname(os.path.abspath(__file__)),
                            os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")
                        ]
                        
                        script_path = None
                        for d in base_dirs:
                            p = os.path.join(d, "scripts", "generar_certificados_auto.ps1")
                            if os.path.exists(p):
                                script_path = p
                                break
                        
                        if not script_path:
                            # Fallback: Crear el script si no existe
                            script_content = """
$certPath = "C:\\Certificates"
if (!(Test-Path -Path $certPath)) {
    New-Item -ItemType Directory -Force -Path $certPath
}
openssl req -x509 -newkey rsa:4096 -keyout "$certPath\\server.key" -out "$certPath\\server.crt" -days 365 -nodes -subj "/CN=localhost"
openssl pkcs12 -export -out "$certPath\\server.pfx" -inkey "$certPath\\server.key" -in "$certPath\\server.crt" -passout pass:fevrips2024*
Write-Host "Certificados generados en $certPath"
"""
                            try:
                                script_path = "generar_certificados_auto.ps1"
                                with open(script_path, "w") as f:
                                    f.write(script_content)
                            except:
                                pass

                        if not script_path or not os.path.exists(script_path):
                             st.error("❌ No se pudo crear ni encontrar el script de certificados.")
                        else:
                            cmd = ["powershell", "-ExecutionPolicy", "Bypass", "-File", script_path]
                            result = subprocess.run(cmd, capture_output=True, text=True)
                            
                            if result.returncode == 0:
                                st.success("✅ Certificados generados en C:\\Certificates")
                                st.code(result.stdout)
                            else:
                                st.error("❌ Error generando certificados")
                                st.text("Salida:")
                                st.code(result.stdout)
                                st.text("Error:")
                                st.code(result.stderr)
                                st.info("Asegúrese de tener OpenSSL instalado (Git Bash suele incluirlo).")
                    except Exception as e:
                        st.error(f"Error ejecutando script: {e}")

                # Paso 3: Comandos
                st.markdown("#### 3. Iniciar Contenedor")
                st.markdown("Ejecute estos comandos en su terminal (PowerShell):")
                st.code("""# 1. Login en Azure (Credenciales del Manual)
docker login -u puller -p v1GLVFn6pWoNrQWgEzmx7MYsf1r7TKJQo+kwadvffq+ACRA3mLxs fevripsacr.azurecr.io

# 2. Iniciar Servicio
docker-compose -f docker-compose-fevrips.yml up -d""", language="powershell")

                if st.button("🚀 Intentar Iniciar Docker Aquí"):
                    try:
                        # Intentar V2 primero (docker compose)
                        cmd_v2 = ["docker", "compose", "-f", "docker-compose-fevrips.yml", "up", "-d"]
                        try:
                            res = subprocess.run(cmd_v2, capture_output=True, text=True, check=False)
                            used_cmd = "docker compose"
                        except FileNotFoundError:
                            # Fallback a V1 (docker-compose)
                            cmd_v1 = ["docker-compose", "-f", "docker-compose-fevrips.yml", "up", "-d"]
                            res = subprocess.run(cmd_v1, capture_output=True, text=True, check=False)
                            used_cmd = "docker-compose"

                        if res.returncode == 0:
                            st.success(f"✅ Comando '{used_cmd}' ejecutado correctamente.")
                            st.text(res.stdout)
                        else:
                            st.error(f"❌ Error iniciando Docker ({used_cmd}).")
                            st.text(res.stderr)
                            if "FileNotFoundError" in str(res.stderr) or res.returncode == 2 or "The system cannot find the file specified" in str(res.stderr):
                                st.warning("Asegúrate de que Docker Desktop esté instalado y agregado al PATH del sistema.")
                                
                    except FileNotFoundError:
                         st.error("❌ No se encontró el ejecutable de Docker.")
                         st.warning("Asegúrate de instalar Docker Desktop y que los comandos 'docker' y 'docker-compose' funcionen en tu terminal.")
                    except Exception as e:
                        st.error(f"Error inesperado: {e}")

    # Credenciales de Autenticación SISPRO
    st.write("🔐 Autenticación SISPRO (Opcional)")
    with st.expander("Configurar Credenciales de Login", expanded=False):
        # Inferir URL de auth basada en la API configurada
        default_auth_url = "https://localhost:9443/api/Auth/LoginSISPRO"
        if api_url and "/api/" in api_url:
            base_url = api_url.split("/api/")[0]
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
                    if auth_url.startswith("https://localhost") or auth_url.startswith("https://127.0.0.1"):
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
                        "verify_ssl": verify_ssl if 'verify_ssl' in locals() else True
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
                temp_dir = os.path.join("temp_downloads", f"fevrips_results_{timestamp}")
                os.makedirs(temp_dir, exist_ok=True)
                
                # Copy generated files
                count_copied = 0
                for f_path in generated_files:
                    if os.path.exists(f_path):
                        shutil.copy2(f_path, temp_dir)
                        count_copied += 1
                
                if count_copied > 0:
                    render_download_button(temp_dir, "btn_download_fevrips", "📦 Descargar Resultados (ZIP)", cleanup=not is_native)
                else:
                    st.warning("No se pudieron copiar los archivos generados para descarga.")
                    
            except Exception as e:
                st.error(f"Error preparando descarga: {e}")
        # ------------------------------
        
        df_res = pd.DataFrame(results)
        st.dataframe(df_res)


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
                 dialog_generar_cuv()
                 
            st.markdown('</div>', unsafe_allow_html=True)
            
        with col_val_2:
            st.info("ℹ️ Asegúrese de que el servicio Docker FEVRIPS esté en ejecución si utiliza el modo local.")
