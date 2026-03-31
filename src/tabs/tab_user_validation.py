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
                tipo_validacion = st.radio("Tipo de Validación", ["ADRES (Web)", "Registraduría (Defunción)"])
                cedula = st.text_input("Número de Documento", key="val_ind_cedula")
                
                if st.button("🔍 Consultar Individual", type="primary"):
                    if not cedula:
                        st.warning("Ingrese un número de documento.")
                    else:
                        with st.spinner("Consultando..."):
                            try:
                                if tipo_validacion == "ADRES (Web)":
                                    st.info("Iniciando navegador para CAPTCHA manual...")
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
                
                tipo_masivo = st.radio("Servicio Masivo", ["ADRES (API/Web)", "Registraduría"], horizontal=True)
                
                if st.button("🚀 Iniciar Procesamiento Masivo"):
                    # Wrapper functions for task manager
                    try:
                        from tabs.tab_automated_actions import worker_registraduria_masiva, worker_adres_api_masiva, worker_adres_web_massive
                    except ImportError:
                        from src.tabs.tab_automated_actions import worker_registraduria_masiva, worker_adres_api_masiva, worker_adres_web_massive
                    
                    if tipo_masivo == "Registraduría":
                        with st.spinner("Procesando validación masiva Registraduría..."):
                            try:
                                res_bytes, msg = worker_registraduria_masiva(df, col_cedula, headless=True, silent_mode=False)
                                if res_bytes:
                                    st.success(f"Validación completada. {msg}")
                                    st.download_button("Descargar Resultados", res_bytes, file_name="validacion_registraduria.xlsx")
                                else:
                                    st.error(f"Error: {msg}")
                            except Exception as e:
                                st.error(f"Excepción: {e}")
                    else:
                        # ADRES (Default to Web as per original intent, but could be API if specified)
                        # Pass tipo_doc_column to the worker
                        with st.spinner("Procesando validación masiva ADRES..."):
                            try:
                                res_bytes, msg = worker_adres_web_massive(df, col_cedula, col_tipo_doc=tipo_doc_column, silent_mode=False)
                                if res_bytes:
                                    st.success(f"Validación completada. {msg}")
                                    st.download_button("Descargar Resultados", res_bytes, file_name="validacion_adres.xlsx")
                                else:
                                    st.error(f"Error: {msg}")
                            except Exception as e:
                                st.error(f"Excepción: {e}")

