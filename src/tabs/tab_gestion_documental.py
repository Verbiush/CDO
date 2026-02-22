
import streamlit as st
import pandas as pd
import os
import shutil
import requests
from datetime import datetime
import time
from docx import Document
from io import BytesIO

try:
    from src import db_gestion
except ImportError:
    import db_gestion

def render():
    st.header("📂 Gestión Documental (Base de Datos)")
    
    tab_import, tab_view, tab_actions = st.tabs(["📥 Importar Excel", "👁️ Ver Registros", "⚙️ Acciones Masivas"])
    
    # --- TAB IMPORTAR ---
    with tab_import:
        st.subheader("Cargar Base de Datos desde Excel")
        uploaded_file = st.file_uploader("Seleccionar archivo Excel (.xlsx)", type=["xlsx", "xls"])
        
        if uploaded_file:
            try:
                df = pd.read_excel(uploaded_file)
                st.write("Vista previa de los datos:")
                st.dataframe(df.head())
                
                st.info(f"Se encontraron {len(df)} filas y {len(df.columns)} columnas.")
                
                # Mapeo de columnas (Intento de auto-detección)
                expected_cols = {
                    "nro_estudio": ["Nro ESTUDIO", "ESTUDIO", "NUMERO ESTUDIO"],
                    "descripcion": ["DESCRIPCION", "DESCRIPCION DEL SERVICIO", "CONCEPTO"],
                    "eps": ["EPS", "ENTIDAD", "ASEGURADORA"],
                    "tipo_doc": ["TIPO DOC", "TIPO DOCUMENTO"],
                    "no_doc": ["No DOC", "NUMERO DOCUMENTO", "DOCUMENTO"],
                    "nombre_completo": ["NOMBRE COMPLETO", "PACIENTE", "NOMBRE"],
                    "nombre_tercero": ["NOMBRE DEL TERCERO", "TERCERO"],
                    "fecha_ingreso": ["FECHA DE INGRESO", "INGRESO"],
                    "fecha_salida": ["FECHA DE SALIDA", "SALIDA"],
                    "autorizacion": ["AUTORIZACION", "NUMERO AUTORIZACION"],
                    "no_factura": ["No FACTURA", "FACTURA", "NUMERO FACTURA"],
                    "fecha_factura": ["FECHA FACTURA", "FECHA DE FACTURA"],
                    "tipo_pago": ["TIPO DE PAGO", "PAGO"],
                    "valor_servicio": ["VALOR SERVICIO", "VALOR"],
                    "copago": ["COPAGO", "CUOTA MODERADORA", "COPAGO / CUOTA MODERADORA"],
                    "total": ["TOTAL", "VALOR TOTAL"],
                    "regimen": ["REGIMEN", "TIPO USUARIO", "RÉGIMEN"]
                }
                
                col_mapping = {}
                st.markdown("#### Mapeo de Columnas")
                st.caption("Verifique que las columnas del Excel coincidan con los campos de la base de datos.")
                
                cols = st.columns(3)
                idx = 0
                
                for db_field, candidates in expected_cols.items():
                    # Find best match
                    default_idx = 0
                    options = ["(Ignorar)"] + list(df.columns)
                    
                    for i, col in enumerate(options):
                        if col in candidates:
                            default_idx = i
                            break
                        # Fuzzy match simple
                        if i > 0 and col.upper() in [c.upper() for c in candidates]:
                            default_idx = i
                            break
                    
                    with cols[idx % 3]:
                        selected = st.selectbox(f"Campo: {db_field}", options, index=default_idx, key=f"map_{db_field}")
                        if selected != "(Ignorar)":
                            col_mapping[db_field] = selected
                    idx += 1
                
                if st.button("💾 Guardar en Base de Datos", type="primary"):
                    success_count = 0
                    error_count = 0
                    
                    progress_bar = st.progress(0)
                    
                    for i, row in df.iterrows():
                        data = {}
                        for db_field, excel_col in col_mapping.items():
                            val = row[excel_col]
                            # Clean data
                            if pd.isna(val):
                                val = ""
                            else:
                                val = str(val).strip()
                            data[db_field] = val
                            
                        # Insert
                        try:
                            db_gestion.insert_document_record(data)
                            success_count += 1
                        except Exception as e:
                            error_count += 1
                            print(f"Error inserting row {i}: {e}")
                            
                        progress_bar.progress((i + 1) / len(df))
                        
                    st.success(f"✅ Importación completada: {success_count} registros guardados.")
                    if error_count > 0:
                        st.warning(f"⚠️ {error_count} registros fallaron.")
                        
            except Exception as e:
                st.error(f"Error leyendo el archivo: {e}")

    # --- TAB VER REGISTROS ---
    with tab_view:
        st.subheader("Registros en Base de Datos")
        
        if st.button("🔄 Actualizar Tabla"):
            st.rerun()
            
        records = db_gestion.get_all_document_records()
        
        if records:
            df_records = pd.DataFrame(records)
            st.dataframe(df_records, use_container_width=True)
            
            st.divider()
            st.caption("Seleccione un ID para eliminar (si es necesario):")
            id_to_delete = st.number_input("ID a eliminar", min_value=0, step=1)
            if st.button("🗑️ Eliminar Registro"):
                if db_gestion.delete_document_record(id_to_delete):
                    st.success(f"Registro {id_to_delete} eliminado.")
                    time.sleep(1)
                    st.rerun()
                else:
                    st.error("No se pudo eliminar el registro (ID no encontrado).")
        else:
            st.info("No hay registros en la base de datos. Importe un Excel primero.")

    # --- TAB ACCIONES ---
    with tab_actions:
        st.subheader("Generación de Carpetas y Documentos")
        
        records = db_gestion.get_all_document_records()
        if not records:
            st.warning("Primero debe importar datos.")
        else:
            st.markdown("### 1. Configuración de Ruta")
            
            # Load path from session state or default
            if "gd_base_path" not in st.session_state:
                st.session_state.gd_base_path = os.path.join(os.path.expanduser("~"), "Documents", "GestionDocumental")
            
            col_path1, col_path2 = st.columns([0.85, 0.15])
            with col_path1:
                base_path = st.text_input("Carpeta Raíz de Salida", value=st.session_state.gd_base_path, key="input_base_path")
            with col_path2:
                # Mock folder selector (in web mode strictly text input, native could use dialog)
                if st.button("📁", help="Actualizar Ruta"):
                    st.session_state.gd_base_path = base_path
                    st.rerun()

            st.markdown("### 2. Estructura de Carpetas")
            st.caption("Use los nombres de columnas entre llaves para definir la estructura. Ej: `{eps}/{fecha_ingreso}/{nombre_completo}`")
            
            folder_structure = st.text_input("Patrón de Carpetas", value="{eps}/{no_factura} - {nombre_completo}")
            
            st.markdown("### 3. Filtros de Procesamiento")
            
            # Extract unique EPS
            all_eps = sorted(list(set([r.get("eps", "") for r in records if r.get("eps")])))
            selected_eps = st.multiselect("Filtrar por EPS (Dejar vacío para todas)", all_eps)
            
            # Date Range
            col_d1, col_d2 = st.columns(2)
            with col_d1:
                date_start = st.date_input("Fecha Inicio (Ingreso)", value=None)
            with col_d2:
                date_end = st.date_input("Fecha Fin (Ingreso)", value=None)
                
            st.markdown("### 4. Vista Previa")
            
            # Apply Filters logic for preview and execution
            filtered_records = []
            for r in records:
                # EPS Filter
                if selected_eps and r.get("eps") not in selected_eps:
                    continue
                
                # Date Filter (parsing fecha_ingreso)
                if date_start and date_end:
                    r_date_str = str(r.get("fecha_ingreso", "")).split(" ")[0] # Handle potential timestamps
                    try:
                        # Attempt multiple formats
                        r_date = None
                        for fmt in ["%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d"]:
                            try:
                                r_date = datetime.strptime(r_date_str, fmt).date()
                                break
                            except:
                                pass
                        
                        if r_date:
                            if not (date_start <= r_date <= date_end):
                                continue
                    except:
                        pass # Keep if date parse fails or strict logic? Let's keep for now to avoid excluding valid but weird formats
                
                filtered_records.append(r)
            
            st.info(f"Registros seleccionados: {len(filtered_records)} de {len(records)}")

            st.markdown("### 4. Acciones a Ejecutar")
            
            col_row1_1, col_row1_2 = st.columns(2)
            with col_row1_1:
                do_create_folders = st.checkbox("1. Crear Carpetas", value=True, help="Crea la estructura de carpetas. Si ya existe, añade consecutivo (1), (2)...")
            with col_row1_2:
                do_modify_template = st.checkbox("2. Generar Documento (Plantilla)", value=True, help="Copia y rellena variables {campo} en el .docx.")
            
            col_row2_1, col_row2_2 = st.columns(2)
            with col_row2_1:
                do_download_sign = st.checkbox("3. Descargar Firma (Web)", value=True, help="Descarga la firma digital usando un patrón de URL.")
            with col_row2_2:
                do_create_signature = st.checkbox("4. Crear Firma (Archivo Local)", value=True, help="Busca firma local y la renombra/mueve a la carpeta.")

            # Template Config
            template_file = None
            if do_modify_template:
                st.caption("Configuración Plantilla (.docx)")
                st.info("El sistema buscará variables como `{nombre_completo}`, `{no_doc}`, `{regimen}` en el documento y las reemplazará.")
                template_file = st.file_uploader("Subir Plantilla", type=["docx"], key="uploader_template")
                
            # Signature Download Config
            sign_url_pattern = ""
            if do_download_sign:
                st.caption("Configuración Descarga Firma")
                st.info("Use las variables entre llaves para construir la URL. Ej: `{no_doc}`, `{nro_estudio}`.")
                sign_url_pattern = st.text_input("URL Patrón", value="https://oportunidaddevida.com/opvcitas/admisionescall/firmas/{no_doc}.png", key="input_sign_url")
            
            # Signature Create Config
            signatures_source_path = ""
            if do_create_signature:
                st.caption("Configuración Crear Firma (Local)")
                st.info("Busca un archivo con el NOMBRE DEL USUARIO en la carpeta indicada. Si el documento es TI o RC, busca por NOMBRE DEL TERCERO.")
                signatures_source_path = st.text_input("Carpeta Origen de Firmas", value="", key="input_sign_source_path")

            if filtered_records:
                example_record = filtered_records[0]
                try:
                    # Clean keys for formatting
                    safe_record = {k: str(v).replace("/", "-").replace("\\", "-").strip() for k, v in example_record.items()}
                    preview_path = folder_structure.format(**safe_record)
                    # FORCE UPPERCASE
                    preview_path = preview_path.upper()
                    
                    full_preview = os.path.join(base_path, preview_path)
                    st.markdown("**Vista Previa Ruta:**")
                    st.code(full_preview)
                    
                    if do_download_sign:
                        preview_url = sign_url_pattern.format(**safe_record)
                        st.markdown("**Vista Previa URL Firma (Descarga):**")
                        st.code(preview_url)
                    
                    if do_create_signature:
                        is_minor = safe_record.get('tipo_doc', '').upper() in ['TI', 'RC']
                        search_name = safe_record.get('nombre_tercero', '') if is_minor else safe_record.get('nombre_completo', '')
                        st.markdown(f"**Vista Previa Búsqueda Firma Local:**\nBuscará archivo que contenga: `{search_name}`")
                        
                except Exception as e:
                    st.error(f"Error en el patrón: {e}")
            
            st.divider()
            
            if st.button("🚀 Ejecutar Procesamiento (Filtrados)", type="primary"):
                count_created = 0
                count_template = 0
                count_sign_download = 0
                count_sign_create = 0
                count_error = 0
                
                status_text = st.empty()
                progress_bar = st.progress(0)
                
                # Pre-read template if needed
                template_bytes = None
                if do_modify_template and template_file:
                    template_bytes = template_file.getvalue()
                elif do_modify_template and not template_file:
                    st.error("Debe subir una plantilla para usar la opción de generación.")
                    st.stop()
                    
                # Check signature source path
                if do_create_signature and not os.path.exists(signatures_source_path):
                    st.error("La carpeta origen de firmas no existe.")
                    st.stop()
                
                # Pre-load signature files list for performance if needed
                signature_files = []
                if do_create_signature:
                    signature_files = os.listdir(signatures_source_path)
                
                for i, record in enumerate(filtered_records):
                    try:
                        # Prepare data for formatting
                        safe_record = {k: str(v).replace("/", "-").replace("\\", "-").strip() for k, v in record.items()}
                        
                        # Format path
                        rel_path = folder_structure.format(**safe_record)
                        # FORCE UPPERCASE
                        rel_path = rel_path.upper()
                        
                        full_path = os.path.join(base_path, rel_path)
                        
                        # 1. Create Folder (Consecutive Logic)
                        if do_create_folders:
                            if os.path.exists(full_path):
                                # Logic to find next consecutive
                                counter = 1
                                base_full_path = full_path
                                while os.path.exists(f"{base_full_path} ({counter})"):
                                    counter += 1
                                full_path = f"{base_full_path} ({counter})"
                            
                            os.makedirs(full_path, exist_ok=True)
                            count_created += 1
                        
                        # Verify path exists before trying to copy files
                        if not os.path.exists(full_path) and (do_modify_template or do_download_sign or do_create_signature):
                            pass

                        if os.path.exists(full_path):
                            # 2. Generate Document from Template
                            if do_modify_template and template_bytes:
                                try:
                                    # Load docx from bytes
                                    doc = Document(BytesIO(template_bytes))
                                    
                                    # Replace in Paragraphs
                                    for p in doc.paragraphs:
                                        for key, val in safe_record.items():
                                            placeholder = f"{{{key}}}"
                                            if placeholder in p.text:
                                                p.text = p.text.replace(placeholder, val)
                                    
                                    # Replace in Tables
                                    for table in doc.tables:
                                        for row in table.rows:
                                            for cell in row.cells:
                                                for p in cell.paragraphs:
                                                    for key, val in safe_record.items():
                                                        placeholder = f"{{{key}}}"
                                                        if placeholder in p.text:
                                                            p.text = p.text.replace(placeholder, val)
                                    
                                    dest_doc = os.path.join(full_path, "documento_generado.docx")
                                    doc.save(dest_doc)
                                    count_template += 1
                                except Exception as e:
                                    print(f"Error generating document: {e}")

                            # 3. Download Signature
                            if do_download_sign and sign_url_pattern:
                                try:
                                    # Construct URL
                                    sign_url = sign_url_pattern.format(**safe_record)
                                    
                                    # Download
                                    response = requests.get(sign_url, verify=False, timeout=5)
                                    if response.status_code == 200:
                                        # Determine extension or default to png
                                        ext = ".png"
                                        if "image/jpeg" in response.headers.get("Content-Type", ""):
                                            ext = ".jpg"
                                        elif "application/pdf" in response.headers.get("Content-Type", ""):
                                            ext = ".pdf"
                                            
                                        dest_sign = os.path.join(full_path, f"firma_descargada{ext}")
                                        with open(dest_sign, "wb") as f:
                                            f.write(response.content)
                                        count_sign_download += 1
                                    else:
                                        print(f"Error downloading sign {sign_url}: {response.status_code}")
                                except Exception as e:
                                    print(f"Error processing sign download: {e}")
                                    
                            # 4. Create Signature (Local Search)
                            if do_create_signature and signature_files:
                                try:
                                    # Determine search name
                                    tipo_doc = safe_record.get('tipo_doc', '').upper()
                                    nombre_completo = safe_record.get('nombre_completo', '').strip()
                                    nombre_tercero = safe_record.get('nombre_tercero', '').strip()
                                    
                                    search_name = nombre_completo
                                    if tipo_doc in ['TI', 'RC'] and nombre_tercero:
                                        search_name = nombre_tercero
                                    
                                    # Find matching file
                                    found_file = None
                                    for f_name in signature_files:
                                        if search_name.lower() in f_name.lower():
                                            found_file = f_name
                                            break
                                    
                                    if found_file:
                                        src_file_path = os.path.join(signatures_source_path, found_file)
                                        # Get extension
                                        _, ext = os.path.splitext(found_file)
                                        dest_file_path = os.path.join(full_path, f"firma_creada{ext}")
                                        
                                        shutil.copy2(src_file_path, dest_file_path)
                                        count_sign_create += 1
                                    else:
                                        print(f"Signature not found for {search_name}")
                                        
                                except Exception as e:
                                    print(f"Error creating signature: {e}")
                        
                        status_text.text(f"Procesando: {rel_path}")
                    except Exception as e:
                        count_error += 1
                        print(f"Error processing record {record.get('id')}: {e}")
                    
                    progress_bar.progress((i + 1) / len(filtered_records))
                
                msg = f"✅ Proceso finalizado.\n"
                if do_create_folders: msg += f"- Carpetas Creadas: {count_created}\n"
                if do_modify_template: msg += f"- Documentos Generados: {count_template}\n"
                if do_download_sign: msg += f"- Firmas Descargadas: {count_sign_download}\n"
                if do_create_signature: msg += f"- Firmas Creadas (Local): {count_sign_create}\n"
                
                st.success(msg)
                if count_error > 0:
                    st.warning(f"{count_error} errores durante el proceso.")

            st.divider()
            st.subheader("📂 Distribución (Mover Archivos Existentes)")
            st.caption("Mover archivos desde una carpeta origen a las carpetas creadas, buscando coincidencias en el nombre.")
            
            source_path = st.text_input("Carpeta Origen (donde están los archivos desordenados)")
            match_field = st.selectbox("Campo Clave para Coincidencia", ["no_factura", "no_doc", "nro_estudio", "autorizacion"])
            
            if st.button("🔍 Analizar y Mover Archivos"):
                if not os.path.exists(source_path):
                    st.error("La carpeta origen no existe.")
                else:
                    moved_count = 0
                    files = os.listdir(source_path)
                    st.info(f"Analizando {len(files)} archivos...")
                    
                    progress_bar_move = st.progress(0)
                    status_text_move = st.empty()
                    
                    for idx, filename in enumerate(files):
                        file_path = os.path.join(source_path, filename)
                        if not os.path.isfile(file_path):
                            continue
                            
                        # Check match with records
                        matched = False
                        for record in records:
                            key_value = str(record.get(match_field, "")).strip()
                            if not key_value:
                                continue
                                
                            if key_value in filename:
                                # Match found! Determine dest path
                                try:
                                    safe_record = {k: str(v).replace("/", "-").replace("\\", "-").strip() for k, v in record.items()}
                                    rel_path = folder_structure.format(**safe_record)
                                    dest_dir = os.path.join(base_path, rel_path)
                                    
                                    if not os.path.exists(dest_dir):
                                        os.makedirs(dest_dir, exist_ok=True)
                                        
                                    dest_path = os.path.join(dest_dir, filename)
                                    shutil.move(file_path, dest_path)
                                    
                                    moved_count += 1
                                    status_text_move.text(f"Movido: {filename} -> {rel_path}")
                                    matched = True
                                    break # Stop checking other records for this file
                                except Exception as e:
                                    print(f"Error moving {filename}: {e}")
                        
                        progress_bar_move.progress((idx + 1) / len(files))
                    
                    st.success(f"Se movieron {moved_count} archivos exitosamente.")

