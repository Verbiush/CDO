import streamlit as st
import pandas as pd
import os
import time
import io
try:
    from modules.adres_validator import ValidatorAdres, ValidatorAdresWeb
    from modules.registraduria_validator import ValidatorRegistraduria
except ImportError:
    from src.modules.adres_validator import ValidatorAdres, ValidatorAdresWeb
    from src.modules.registraduria_validator import ValidatorRegistraduria



@st.cache_data(show_spinner=False, max_entries=5)
def _load_excel(file_bytes):
    import pandas as pd
    import io
    return pd.read_excel(io.BytesIO(file_bytes))

def render(container=None):
    if container is None:
        container = st.container()
    
    with container:
        st.header("👤 Validación de Usuarios (ADRES / Registraduría)")
        st.markdown("Valide el estado de afiliación y documentos de identidad.")
        
        tab1, tab2 = st.tabs(["Validación Individual", "Validación Masiva (Excel)"])
        
        # --- TAB 1: INDIVIDUAL ---
        with tab1:
            col_in, col_res = st.columns([0.4, 0.6])
            
            with col_in:
                st.subheader("Consulta")
                tipo_validacion = st.radio("Tipo de Validación", ["ADRES (API)", "ADRES (Web)", "Registraduría (Defunción)"])
                cedula = st.text_input("Número de Documento", key="val_ind_cedula")
                
                if st.button("🔍 Consultar Individual", type="primary"):
                    if not cedula:
                        st.warning("Ingrese un número de documento.")
                    else:
                        with st.spinner("Consultando..."):
                            try:
                                if tipo_validacion == "ADRES (API)":
                                    val = ValidatorAdres()
                                    res = val.validate_cedula(cedula)
                                    st.session_state.val_result = res
                                elif tipo_validacion == "ADRES (Web)":
                                    val = ValidatorAdresWeb(headless=False)
                                    res = val.validate_cedula(cedula)
                                    st.session_state.val_result = res
                                    val.close_driver()
                                elif tipo_validacion == "Registraduría (Defunción)":
                                    val = ValidatorRegistraduria(headless=True)
                                    res = val.validate_cedula(cedula)
                                    st.session_state.val_result = res
                                    val.close_driver()
                            except Exception as e:
                                st.error(f"Error: {e}")

            with col_res:
                st.subheader("Resultados")
                if "val_result" in st.session_state and st.session_state.val_result:
                    res = st.session_state.val_result
                    st.json(res)
                    
                    # Formato tarjeta
                    if isinstance(res, dict):
                        st.divider()
                        c1, c2 = st.columns(2)
                        c1.metric("Estado", res.get("Estado", "Desconocido"))
                        c2.metric("Entidad", res.get("Entidad", "N/A"))
                        st.write(f"**Nombre:** {res.get('Nombres', '')} {res.get('Apellidos', '')}")

        # --- TAB 2: MASIVA ---
        with tab2:
            st.markdown("#### Cargar Excel con Cédulas")
            uploaded_file = st.file_uploader("Seleccione archivo Excel (.xlsx)", type=["xlsx"], key="up_user_val")
            
            if uploaded_file:
                file_bytes = uploaded_file.getvalue()
                df = _load_excel(file_bytes)
                st.write("Vista previa:", df.head())
                
                cols = df.columns.tolist()
                col_cedula = st.selectbox("Seleccione columna de Cédulas", cols)
                
                col_tipo_doc = st.selectbox("Seleccione columna de Tipo de Documento (Opcional)", ["Ninguna"] + cols)
                tipo_doc_column = None if col_tipo_doc == "Ninguna" else col_tipo_doc
                
                tipo_masivo = st.radio("Servicio Masivo", ["ADRES (API)", "ADRES (Web)", "Registraduría", "FOMAG (Certificados)"], horizontal=True)
                
                if st.button("🚀 Iniciar Procesamiento Masivo"):
                    # Wrapper functions for task manager
                    try:
                        from tabs.tab_automated_actions import worker_registraduria_masiva, worker_adres_api_masiva, worker_adres_web_massive
                    except ImportError:
                        from src.tabs.tab_automated_actions import worker_registraduria_masiva, worker_adres_api_masiva, worker_adres_web_massive
                    
                    if tipo_masivo == "Registraduría":
                        with st.spinner("Procesando validación masiva Registraduría..."):
                            try:
                                result = worker_registraduria_masiva(df, col_cedula, headless=True, silent_mode=False)
                                if "error" in result:
                                    st.error(f"Error: {result['error']}")
                                else:
                                    file_info = result["files"][0]
                                    st.success(f"Validación completada. {result['message']}")
                                    st.download_button("Descargar Resultados", file_info["data"], file_name=file_info["name"])
                            except Exception as e:
                                st.error(f"Excepción: {e}")
                    elif tipo_masivo == "ADRES (API)":
                        # ADRES (API as it does not require Chrome/Selenium on the server)
                        # Pass tipo_doc_column to the worker
                        with st.spinner("Procesando validación masiva ADRES (API)..."):
                            try:
                                result = worker_adres_api_masiva(df, col_cedula, col_tipo_doc=tipo_doc_column, silent_mode=False)
                                if "error" in result:
                                    st.error(f"Error: {result['error']}")
                                else:
                                    file_info = result["files"][0]
                                    st.success(f"Validación completada. {result['message']}")
                                    st.download_button("Descargar Resultados", file_info["data"], file_name=file_info["name"])
                            except Exception as e:
                                st.error(f"Excepción: {e}")
                    elif tipo_masivo == "ADRES (Web)":
                        with st.spinner("Enviando validación masiva ADRES (Web) al Agente Local..."):
                            try:
                                import base64
                                import io
                                import time
                                from src import agent_client
                                
                                output = io.BytesIO()
                                df.to_excel(output, index=False)
                                file_data_b64 = base64.b64encode(output.getvalue()).decode("utf-8")
                                
                                username = st.session_state.get("username", "admin")
                                if not agent_client.is_agent_active(username):
                                    st.error("⚠️ El Agente Local no está activo. Ejecute la aplicación localmente e inicie sesión.")
                                    st.stop()
                                
                                task_id = agent_client.send_command(username, "adres_web_massive", {
                                    "file_data": file_data_b64,
                                    "col_cedula": col_cedula,
                                    "col_tipo_doc": tipo_doc_column
                                })
                                
                                if task_id:
                                    st.info(f"Tarea enviada al agente (ID: {task_id}). Esperando resultado (esto puede tardar varios minutos)...")
                                    result = agent_client.wait_for_result(task_id, timeout=600) # 10 minutes timeout for massive tasks
                                    
                                    if result and "error" not in result:
                                        if "files" in result and result["files"]:
                                            file_info = result["files"][0]
                                            # Agent might return base64 string for file data
                                            file_data = base64.b64decode(file_info["data"]) if isinstance(file_info["data"], str) else file_info["data"]
                                            st.success("Validación masiva completada por el Agente Local.")
                                            st.download_button("Descargar Resultados", file_data, file_name=f"Resultados_ADRES_WEB_{int(time.time())}.xlsx")
                                        else:
                                            st.warning("El agente terminó la tarea pero no devolvió un archivo.")
                                    else:
                                        err_msg = result.get("error", "Error desconocido") if result else "No se obtuvo respuesta"
                                        st.error(f"Error del Agente: {err_msg}")
                                else:
                                    st.error("No se pudo crear la tarea en el servidor.")
                                    
                            except Exception as e:
                                st.error(f"Excepción: {e}")
                    elif tipo_masivo == "FOMAG (Certificados)":
                        with st.spinner("Enviando tarea de FOMAG al Agente Local..."):
                            try:
                                import base64
                                import io
                                import time
                                from src import agent_client
                                
                                output = io.BytesIO()
                                df.to_excel(output, index=False)
                                file_data_b64 = base64.b64encode(output.getvalue()).decode("utf-8")
                                
                                username = st.session_state.get("username", "admin")
                                if not agent_client.is_agent_active(username):
                                    st.error("⚠️ El Agente Local no está activo. Ejecute la aplicación localmente e inicie sesión.")
                                    st.stop()
                                
                                task_id = agent_client.send_command(username, "fomag_cert_massive", {
                                    "file_data": file_data_b64,
                                    "col_cedula": col_cedula
                                })
                                
                                if task_id:
                                    st.info(f"Tarea enviada al agente (ID: {task_id}). Se abrirá Chrome en su PC. Por favor INICIE SESIÓN manualmente, el agente descargará los certificados después.")
                                    result = agent_client.wait_for_result(task_id, timeout=3600) # 1 hora de timeout (login + proceso)
                                    
                                    if result and "error" not in result:
                                        if "files" in result and result["files"]:
                                            file_info = result["files"][0]
                                            file_data = base64.b64decode(file_info["data"]) if isinstance(file_info["data"], str) else file_info["data"]
                                            st.success("Descarga de certificados FOMAG completada por el Agente Local.")
                                            st.download_button("Descargar ZIP con Certificados", file_data, file_name=f"Certificados_FOMAG_{int(time.time())}.zip", mime="application/zip")
                                        else:
                                            st.warning("El agente terminó la tarea pero no devolvió un archivo ZIP.")
                                    else:
                                        err_msg = result.get("error", "Error desconocido") if result else "No se obtuvo respuesta del Agente (Tiempo agotado)."
                                        st.error(f"Error del Agente: {err_msg}")
                                else:
                                    st.error("No se pudo crear la tarea en el servidor.")
                                    
                            except Exception as e:
                                st.error(f"Excepción: {e}")

