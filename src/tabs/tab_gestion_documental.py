
import streamlit as st
import pandas as pd
import os
import shutil
import requests
from datetime import datetime
import time
try:
    from docx import Document
    from docx.text.paragraph import Paragraph
    from docx.oxml.ns import qn
except ImportError:
    Document = None
    Paragraph = None
    qn = None
import re
from io import BytesIO
import unicodedata
from PIL import Image, ImageDraw, ImageFont

import zipfile
import importlib
import base64
import urllib.parse
import time

try:
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.chrome.service import Service
    from webdriver_manager.chrome import ChromeDriverManager
except ImportError:
    webdriver = None

try:
    from src import db_gestion
    from src.gui_utils import abrir_dialogo_carpeta_nativo, update_path_key as update_path_key_gui, render_path_selector
except ImportError:
    import db_gestion
    from gui_utils import abrir_dialogo_carpeta_nativo, update_path_key as update_path_key_gui, render_path_selector

# Force reload to ensure latest DB logic is used
importlib.reload(db_gestion)

def get_google_font(font_name, size):
    """Download and load a Google Font."""
    font_dir = "assets/fonts"
    os.makedirs(font_dir, exist_ok=True)
    
    font_urls = {
        "Pacifico": "https://github.com/google/fonts/raw/main/ofl/pacifico/Pacifico-Regular.ttf",
        # "Dancing Script": "https://github.com/google/fonts/raw/main/ofl/dancingscript/DancingScript-VariableFont_wght.ttf",
    }
    
    font_path = os.path.join(font_dir, f"{font_name}.ttf")
    
    # Download if not exists
    if font_name == "My Ugly Handwriting":
        if not os.path.exists(font_path):
            try:
                # Direct download from dafont (simulated as browser)
                url = "https://dl.dafont.com/dl/?f=my_ugly_handwriting"
                headers = {
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
                }
                r = requests.get(url, headers=headers, timeout=15)
                if r.status_code == 200:
                    with zipfile.ZipFile(BytesIO(r.content)) as z:
                        for name in z.namelist():
                            if name.lower().endswith(".otf") or name.lower().endswith(".ttf"):
                                with open(font_path, "wb") as f:
                                    f.write(z.read(name))
                                break
            except Exception as e:
                print(f"Error downloading My Ugly Handwriting: {e}")
                
    elif font_name in font_urls and not os.path.exists(font_path):
        try:
            r = requests.get(font_urls[font_name])
            with open(font_path, "wb") as f:
                f.write(r.content)
        except Exception as e:
            print(f"Error downloading font: {e}")
            
    try:
        return ImageFont.truetype(font_path, size)
    except:
        try:
            return ImageFont.truetype("arial.ttf", size)
        except:
            return ImageFont.load_default()

def generate_signature_image(text, font_name="Pacifico", size=70, width=500, height=200):
    """Generate a signature image from text."""
    img = Image.new('RGB', (width, height), color=(255, 255, 255))
    d = ImageDraw.Draw(img)
    
    font = get_google_font(font_name, size)
    
    # Center text
    try:
        # getbbox is cleaner in newer Pillow, textbbox in even newer
        left, top, right, bottom = d.textbbox((0, 0), text, font=font)
        text_w = right - left
        text_h = bottom - top
    except:
        # Fallback for older Pillow
        text_w, text_h = d.textsize(text, font=font)
        
    x = (width - text_w) / 2
    y = (height - text_h) / 2
    
    d.text((x, y), text, font=font, fill=(0, 0, 0))
    return img

def normalize_text(text):
    """Normalize text to remove accents and convert to lowercase."""
    if not isinstance(text, str):
        return str(text)
    return ''.join(c for c in unicodedata.normalize('NFD', text) if unicodedata.category(c) != 'Mn').lower()

def match_signature_file(search_name, filenames):
    """
    Find a matching filename for a given name.
    Checks if all parts of the search_name are present in the filename.
    """
    if not search_name:
        return None
        
    search_parts = normalize_text(search_name).split()
    
    for fname in filenames:
        norm_fname = normalize_text(fname)
        # Check if all parts of the name are in the filename
        if all(part in norm_fname for part in search_parts):
            return fname
    return None

def resolve_unique_paths(records, folder_structure):
    """
    Generates unique folder paths for each record to avoid collisions.
    If multiple records map to the same path, appends a consecutive suffix (_1, _2, etc.)
    sorted by 'no_factura' to ensure determinism.
    
    Returns:
        dict: {record_id: unique_relative_path}
    """
    # 1. Group records by their raw path
    path_groups = {}
    
    for record in records:
        safe_record = {k: str(v).replace("/", "-").replace("\\", "-").strip() for k, v in record.items()}
        try:
            raw_path = folder_structure.format(**safe_record).upper()
        except KeyError:
            # Fallback if key missing
            raw_path = f"ERROR_KEY_MISSING_{record.get('id')}"
            
        if raw_path not in path_groups:
            path_groups[raw_path] = []
        path_groups[raw_path].append(record)
    
    # 2. Resolve collisions
    record_id_to_path = {}
    
    for raw_path, group in path_groups.items():
        if len(group) == 1:
            record_id_to_path[group[0]['id']] = raw_path
        else:
            # Sort by invoice number (or ID if invoice is same/missing) to be deterministic
            # We assume 'no_factura' exists, else 'id'
            group.sort(key=lambda x: (str(x.get('no_factura', '')), x.get('id', 0)))
            
            for i, record in enumerate(group):
                # Append consecutive: Path_1, Path_2, etc.
                # Or maybe Path (1), Path (2)? Let's use _1, _2 as requested "consecutivo"
                unique_path = f"{raw_path}_{i+1}"
                record_id_to_path[record['id']] = unique_path
                
    return record_id_to_path



def worker_descargar_historias_ovida(records, download_path, resolved_paths):
    """
    Descarga historias clínicas de OVIDA para los registros dados.
    Usa Selenium para inicio de sesión y descarga.
    """
    if webdriver is None:
        return "Error: Selenium no está instalado o no se puede cargar."

    if not os.path.isdir(download_path):
        return "Error: Carpeta de descarga inválida."

    driver = None
    try:
        options = webdriver.ChromeOptions()
        prefs = {
            "download.default_directory": download_path,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True
        }
        options.add_experimental_option("prefs", prefs)
        
        # Open visible browser for login
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        
        driver.get("https://ovidazs.siesacloud.com/ZeusSalud/ips/iniciando.php")
        
        st.warning("⚠️ Se abrió una ventana de Chrome. INICIE SESIÓN en OVIDA manualmente. El proceso continuará automáticamente cuando detecte el ingreso.")
        
        # Wait for login (detect change to main page or timeout)
        # Timeout 5 minutes
        timeout = 300 
        start_time = time.time()
        while time.time() - start_time < timeout:
             try:
                 # Check if URL changed from login page
                 if "iniciando.php" not in driver.current_url and "login" not in driver.current_url.lower():
                     break
             except:
                 pass
             time.sleep(1)
        
        if time.time() - start_time >= timeout:
            driver.quit()
            return "Error: Tiempo de espera de inicio de sesión agotado."
            
        # Check logged in state more robustly
        logged_in = False
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                if "App/Vistas" in driver.current_url:
                    logged_in = True
                    break
            except: pass
            time.sleep(2)
            
        if not logged_in:
            return "Error: No se detectó inicio de sesión en 5 minutos."

        st.info("Inicio de sesión detectado. Comenzando descargas...")

        descargados = 0
        errores = 0
        conflictos = 0
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        total = len(records)
        
        for i, record in enumerate(records):
            progress_bar.progress((i + 1) / total)
            
            try:
                estudio = str(record.get('nro_estudio', '')).strip()
                # Clean estudio (remove .0 if float)
                if estudio.endswith(".0"): estudio = estudio[:-2]
                
                if not estudio or estudio == "nan" or estudio == "None":
                    errores += 1
                    continue
                    
                # Parse dates
                try:
                    f_ing_raw = record.get('fecha_ingreso', '')
                    f_egr_raw = record.get('fecha_salida', '')
                    
                    if not f_ing_raw or not f_egr_raw:
                        errores += 1
                        continue
                        
                    f_ing = pd.to_datetime(f_ing_raw).strftime('%Y/%m/%d')
                    f_egr = pd.to_datetime(f_egr_raw).strftime('%Y/%m/%d')
                except:
                    errores += 1
                    continue
                
                # Determine folder
                rel_path = resolved_paths.get(record['id'])
                if not rel_path:
                    # Fallback (shouldn't happen if resolved_paths is complete)
                    continue 

                dest_dir = os.path.join(download_path, rel_path)
                os.makedirs(dest_dir, exist_ok=True)
                
                final_path = os.path.join(dest_dir, f"HC_{estudio}.pdf")
                
                if os.path.exists(final_path):
                    conflictos += 1
                    continue
                    
                status_text.text(f"Descargando Estudio: {estudio}")

                # URL construction
                base_url = "https://ovidazs.siesacloud.com/ZeusSalud/Reportes/Cliente//html/reporte_historia_general.php"
                params = {
                    'estudio': estudio, 'fecha_inicio': f_ing, 'fecha_fin': f_egr,
                    'verHC': 1, 'verEvo': 1, 'verPar': 1, 'ImprimirOrdenamiento': 1,
                    'ImprimirNotasPcte': 0, 'ImprimirSolOrdenesExt': 1, 'ImprimirGraficasHC': 1,
                    'ImprimirFormatos': 1, 'ImprimirRegistroAdmon': 1, 'ImprimirNovedad': 0,
                    'ImprimirRecomendaciones': 0, 'ImprimirDescripcionQX': 0, 'ImprimirNotasEnfermeria': 1,
                    'ImprimirSignosVitales': 0, 'ImprimirLog': 0, 'ImprimirEpicrisisSinHC': 0
                }
                full_url = f"{base_url}?{urllib.parse.urlencode(params)}"
                
                driver.get(full_url)
                time.sleep(2) # Wait for render
                
                pdf_b64 = driver.execute_cdp_cmd("Page.printToPDF", {
                    "landscape": False, "printBackground": True,
                    "paperWidth": 8.5, "paperHeight": 11,
                    "marginTop": 0.4, "marginBottom": 0.4, "marginLeft": 0.4, "marginRight": 0.4
                })
                
                pdf_data = base64.b64decode(pdf_b64['data'])
                with open(final_path, 'wb') as f:
                    f.write(pdf_data)
                
                descargados += 1
                
            except Exception as e:
                errores += 1
                # print(f"Error en estudio {estudio}: {e}")
                
        return f"Finalizado. Descargados: {descargados}, Errores: {errores}, Conflictos: {conflictos}."

    except Exception as e:
        return f"Error crítico: {e}"
    finally:
        if driver: driver.quit()

def render():
    st.markdown("## 📂 Gestión Documental (Relacional)")
    
    # --- SYNC LOGIC ---
    # Sincronizar automáticamente con la ruta del panel principal si esta cambia
    if "current_path" not in st.session_state:
         st.session_state.current_path = "" # Inicialmente vacío para cumplir requerimiento
         
    current_main_path = st.session_state.get("current_path")
    last_synced = st.session_state.get("gd_last_synced_path")
    
    # Logic: Sync if path changed globally OR if we haven't synced yet
    # We remove the widget check to avoid overriding manual user changes in this tab 
    # UNLESS the global path has explicitly changed since last sync.
    if current_main_path and current_main_path != last_synced:
        st.session_state.gd_base_path = current_main_path
        st.session_state.input_base_path_struct = current_main_path
        st.session_state.input_base_path_content = current_main_path
        st.session_state.gd_source_path = current_main_path
        st.session_state.gd_source_path_mov = current_main_path # Sync Move Tool Source
        st.session_state.input_base_path_struct_mov = current_main_path # Sync Move Tool Dest
        st.session_state.input_gd_source_path = current_main_path
        st.session_state.input_gd_source_path_mov = current_main_path
        st.session_state.gd_last_synced_path = current_main_path
        # Force rerun to update widgets immediately
        st.rerun()
    
    # Ensure keys exist to prevent KeyErrors or empty fields on first load
    if "input_base_path_struct" not in st.session_state:
        st.session_state.input_base_path_struct = st.session_state.get("gd_base_path", current_main_path)
    if "input_base_path_content" not in st.session_state:
        st.session_state.input_base_path_content = st.session_state.get("gd_base_path", current_main_path)
    
    # Updated Tabs Structure: Flattened for better visibility
    tab_import, tab_view, tab_structure, tab_organize, tab_content = st.tabs([
        "📥 Importar Excel", 
        "👁️ Ver Registros", 
        "📁 Crear Carpetas", 
        "🚀 Organizar (Renombrar/Mover)",
        "📝 Generar Docs/Firmas"
    ])
    
    # Common Data Loading (Available across tabs)
    records = db_gestion.get_all_document_records()

    # --- TAB IMPORTAR ---
    with tab_import:
        st.subheader("Cargar Base de Datos desde Excel")
        uploaded_file = st.file_uploader("Seleccionar archivo Excel (.xlsx)", type=["xlsx", "xls"])
        
        if uploaded_file:
            try:
                df = pd.read_excel(uploaded_file)
                # Normalize column names (strip whitespace)
                df.columns = df.columns.str.strip()
                
                st.write("Vista previa de los datos:")
                st.dataframe(df.head())
                
                st.info(f"Se encontraron {len(df)} filas y {len(df.columns)} columnas.")
                
                # Mapeo de columnas (Intento de auto-detección)
                expected_cols = {
                    # --- PACIENTE ---
                    "tipo_doc": ["TIPO DOC", "TIPO DOCUMENTO", "TIPO DE IDENTIFICACION"],
                    "no_doc": ["No DOC", "NUMERO DOCUMENTO", "DOCUMENTO", "IDENTIFICACION"],
                    "nombre_completo": ["NOMBRE COMPLETO", "PACIENTE", "NOMBRE"],
                    "nombre_tercero": ["NOMBRE DEL TERCERO", "TERCERO"],
                    "eps": ["EPS", "ENTIDAD", "ASEGURADORA"],
                    "regimen": ["REGIMEN", "TIPO USUARIO", "RÉGIMEN"],
                    "categoria": ["CATEGORIA", "NIVEL", "CATEGORÍA"],

                    # --- ATENCION ---
                    "nro_estudio": ["Nro ESTUDIO", "ESTUDIO", "NUMERO ESTUDIO", "ADMISION"],
                    "descripcion": ["DESCRIPCION DEL CUPS", "DESCRIPCION", "CONCEPTO", "SERVICIO"],
                    "fecha_ingreso": ["FECHA DE INGRESO", "INGRESO", "FECHA INGRESO"],
                    "fecha_salida": ["FECHA DE SALIDA", "SALIDA", "FECHA SALIDA"],
                    "autorizacion": ["AUTORIZACION", "NUMERO AUTORIZACION", "NO AUTORIZACION"],

                    # --- FACTURA ---
                    "no_factura": ["No FACTURA", "FACTURA", "NUMERO FACTURA"],
                    "fecha_factura": ["FECHA FACTURA", "FECHA DE FACTURA"],
                    "tipo_pago": ["TIPO DE PAGO", "PAGO", "FORMA PAGO"],
                    "valor_servicio": ["VALOR SERVICIO", "VALOR", "VALOR UNITARIO"],
                    "copago": ["COPAGO / CUOTA MODERADORA", "COPAGO", "CUOTA MODERADORA"],
                    "radicado": ["RADICADO", "NUMERO RADICADO", "NO RADICADO"],
                    "total": ["TOTAL", "VALOR TOTAL", "VALOR NETO"],
                    "tipo_servicio": ["TIPO DE SERVICIO", "TIPO SERVICIO", "SERVICIO TIPO"],
                    "fecha_radicado": ["FECHA RADICADO", "FECHA DE RADICADO", "F. RADICADO"]
                }
                
                col_mapping = {}
                st.markdown("#### Mapeo de Columnas")
                st.caption("Verifique que las columnas del Excel coincidan con los campos de la base de datos.")
                st.info("💡 Si el número de factura ya existe, se actualizarán los campos proporcionados.")
                
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
                                # Handle float to int if applicable (e.g. 1.0 -> 1)
                                if isinstance(val, float) and val.is_integer():
                                    val = int(val)
                                val = str(val).strip()
                            data[db_field] = val
                            
                        # Insert
                        try:
                            res = db_gestion.insert_document_record(data)
                            if res and res[0]:
                                success_count += 1
                            else:
                                error_count += 1
                                err_msg = res[1] if res else "Unknown error"
                                print(f"Error importing row {i}: {err_msg}")
                        except Exception as e:
                            error_count += 1
                            print(f"Error inserting row {i}: {e}")
                            
                        progress_bar.progress((i + 1) / len(df))
                        
                    st.success(f"✅ Importación completada: {success_count} registros guardados.")
                    if error_count > 0:
                        st.warning(f"⚠️ {error_count} registros fallaron.")
                    
                    # Force rerun to show updated data immediately
                    time.sleep(1)
                    st.rerun()
                        
            except Exception as e:
                st.error(f"Error leyendo el archivo: {e}")

    # --- TAB VER REGISTROS ---
    with tab_view:
        st.subheader("Registros en Base de Datos")
        
        col_act, col_del = st.columns([0.8, 0.2])
        with col_act:
            if st.button("🔄 Actualizar Tabla"):
                st.rerun()
            
        if records:
            # 1. Selection for Editing
            st.divider()
            st.markdown("#### ✏️ Editar Registro")
            
            # Create a label for selection
            record_options = {f"{r['no_factura']} - {r['nombre_completo']} (ID: {r['id']})": r for r in records}
            selected_label = st.selectbox("Seleccionar Registro para Editar:", options=list(record_options.keys()), key="sel_edit_record")
            
            if selected_label:
                record = record_options[selected_label]
                
                with st.expander("📝 Formulario de Edición", expanded=True):
                    # Group fields
                    c1, c2, c3 = st.columns(3)
                    
                    new_values = {}
                    
                    with c1:
                        st.caption("Paciente")
                        new_values['tipo_doc'] = st.text_input("Tipo Doc", value=record.get('tipo_doc', ''), key=f"edit_td_{record['id']}")
                        new_values['no_doc'] = st.text_input("No. Documento", value=record.get('no_doc', ''), key=f"edit_doc_{record['id']}")
                        new_values['nombre_completo'] = st.text_input("Nombre Completo", value=record.get('nombre_completo', ''), key=f"edit_nom_{record['id']}")
                        new_values['nombre_tercero'] = st.text_input("Nombre Tercero", value=record.get('nombre_tercero', ''), key=f"edit_ter_{record['id']}")
                        new_values['eps'] = st.text_input("EPS", value=record.get('eps', ''), key=f"edit_eps_{record['id']}")
                        new_values['regimen'] = st.text_input("Régimen", value=record.get('regimen', ''), key=f"edit_reg_{record['id']}")
                        new_values['categoria'] = st.text_input("Categoría", value=record.get('categoria', ''), key=f"edit_cat_{record['id']}")
                        
                    with c2:
                        st.caption("Atención")
                        new_values['nro_estudio'] = st.text_input("Nro. Estudio", value=record.get('nro_estudio', ''), key=f"edit_est_{record['id']}")
                        new_values['descripcion'] = st.text_input("Descripción", value=record.get('descripcion', ''), key=f"edit_desc_{record['id']}")
                        new_values['fecha_ingreso'] = st.text_input("F. Ingreso", value=record.get('fecha_ingreso', ''), key=f"edit_fi_{record['id']}")
                        new_values['fecha_salida'] = st.text_input("F. Salida", value=record.get('fecha_salida', ''), key=f"edit_fs_{record['id']}")
                        new_values['autorizacion'] = st.text_input("Autorización", value=record.get('autorizacion', ''), key=f"edit_auth_{record['id']}")

                    with c3:
                        st.caption("Factura")
                        new_values['no_factura'] = st.text_input("No. Factura", value=record.get('no_factura', ''), key=f"edit_fac_{record['id']}")
                        new_values['fecha_factura'] = st.text_input("F. Factura", value=record.get('fecha_factura', ''), key=f"edit_ff_{record['id']}")
                        new_values['valor_servicio'] = st.text_input("Valor", value=str(record.get('valor_servicio', '')), key=f"edit_val_{record['id']}")
                        new_values['copago'] = st.text_input("Copago", value=str(record.get('copago', '')), key=f"edit_cop_{record['id']}")
                        new_values['total'] = st.text_input("Total", value=str(record.get('total', '')), key=f"edit_tot_{record['id']}")
                        new_values['radicado'] = st.text_input("Radicado", value=str(record.get('radicado', '')), key=f"edit_rad_{record['id']}")
                        new_values['fecha_radicado'] = st.text_input("F. Radicado", value=str(record.get('fecha_radicado', '')), key=f"edit_frad_{record['id']}")
                        new_values['tipo_pago'] = st.text_input("Tipo Pago", value=str(record.get('tipo_pago', '')), key=f"edit_tp_{record['id']}")
                        new_values['tipo_servicio'] = st.text_input("Tipo Servicio", value=str(record.get('tipo_servicio', '')), key=f"edit_ts_{record['id']}")
                        new_values['status'] = st.text_input("Estado", value=str(record.get('status', '')), key=f"edit_st_{record['id']}")
                    
                    if st.button("💾 Guardar Cambios", type="primary", key="btn_save_changes"):
                        updates_count = 0
                        errors = []
                        
                        for field, new_val in new_values.items():
                            old_val = str(record.get(field, ''))
                            # Only update if changed
                            if new_val != old_val:
                                success, msg = db_gestion.update_document_field(record['id'], field, new_val)
                                if success:
                                    updates_count += 1
                                else:
                                    errors.append(f"{field}: {msg}")
                        
                        if errors:
                            st.error(f"Errores al actualizar: {'; '.join(errors)}")
                        elif updates_count > 0:
                            st.success(f"✅ Se actualizaron {updates_count} campos correctamente.")
                            time.sleep(1)
                            st.rerun()
                        else:
                            st.info("No se detectaron cambios.")

            st.divider()
            df_records = pd.DataFrame(records)
            st.dataframe(df_records, use_container_width=True)
            
            st.caption("ℹ️ Para eliminar registros, utilice la pestaña 'Gestión de Información'.")
        else:
            st.info("No hay registros en la base de datos. Importe un Excel primero.")

    # --- SHARED FILTERING LOGIC (Helper for Action Tabs) ---
    def get_filtered_records(records, key_suffix=""):
        if not records:
            return []
            
        # --- FILTERS ---
        with st.expander("🔍 Filtros de Procesamiento", expanded=True):
            # Extract unique EPS
            all_eps = sorted(list(set([r.get("eps", "") for r in records if r.get("eps")])))
            selected_eps = st.multiselect(
                "Filtrar por EPS (Dejar vacío para todas)", 
                all_eps, 
                key=f"filter_eps_{key_suffix}"
            )
            
            # Extract unique Regimen
            all_regimen = sorted(list(set([r.get("regimen", "") for r in records if r.get("regimen")])))
            selected_regimen = st.multiselect(
                "Filtrar por Régimen (Dejar vacío para todos)", 
                all_regimen, 
                key=f"filter_regimen_{key_suffix}"
            )
            
            # Date Range (Ingreso)
            col_d1, col_d2 = st.columns(2)
            with col_d1:
                date_start = st.date_input(
                    "Fecha Inicio (Ingreso)", 
                    value=None, 
                    key=f"filter_d1_{key_suffix}"
                )
            with col_d2:
                date_end = st.date_input(
                    "Fecha Fin (Ingreso)", 
                    value=None, 
                    key=f"filter_d2_{key_suffix}"
                )

            # Date Range (Factura)
            col_f1, col_f2 = st.columns(2)
            with col_f1:
                date_fact_start = st.date_input(
                    "Fecha Inicio (Factura)", 
                    value=None, 
                    key=f"filter_f1_{key_suffix}"
                )
            with col_f2:
                date_fact_end = st.date_input(
                    "Fecha Fin (Factura)", 
                    value=None, 
                    key=f"filter_f2_{key_suffix}"
                )
                
        filtered = []
        for r in records:
            # EPS Filter
            if selected_eps and r.get("eps") not in selected_eps:
                continue
            
            # Regimen Filter
            if selected_regimen and r.get("regimen") not in selected_regimen:
                continue
            
            # Date Filter (Ingreso)
            if date_start and date_end:
                r_date_str = str(r.get("fecha_ingreso", "")).split(" ")[0] 
                try:
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
                    pass

            # Date Filter (Factura)
            if date_fact_start and date_fact_end:
                r_date_str = str(r.get("fecha_factura", "")).split(" ")[0] 
                try:
                    r_date = None
                    for fmt in ["%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d"]:
                        try:
                            r_date = datetime.strptime(r_date_str, fmt).date()
                            break
                        except:
                            pass
                    
                    if r_date:
                        if not (date_fact_start <= r_date <= date_fact_end):
                            continue
                except:
                    pass
            
            filtered.append(r)
        return filtered

    # --- TAB ESTRUCTURA (Former Subtab 1) ---
    with tab_structure:
        st.subheader("📁 Estructura de Carpetas")
        
        # Common Config - ALWAYS VISIBLE
        st.markdown("### Configuración de Rutas")
        
        default_hardcoded = os.path.join(os.path.expanduser("~"), "Documents", "GestionDocumental")
        
        # Initialize fallback if somehow missing
        if "gd_base_path" not in st.session_state:
             st.session_state.gd_base_path = default_hardcoded
             st.session_state.input_base_path_struct = default_hardcoded
        
        # Selector de Ruta Estandarizado
        base_path = render_path_selector(
            label="Carpeta Raíz de Salida",
            key="gd_base_path_struct",
            default_path=st.session_state.get("gd_base_path", st.session_state.get("current_path", os.getcwd()))
        )
        st.session_state.input_base_path_struct = base_path

        st.caption("Patrón de Carpetas: Use `{eps}`, `{fecha_ingreso}`, `{nombre_completo}`, etc.")
        folder_structure = st.text_input("Patrón de Carpetas", value="{eps}/{no_factura} - {nombre_completo}", key="input_folder_structure_struct")
        
        # --- NEW LOGIC: ALWAYS SHOW ACTIONS ---
        st.divider()
        st.markdown("### Acciones")
        do_create_folders = st.checkbox("Activar Creación de Carpetas", value=True, help="Crea la estructura de carpetas.")
        
        disabled_struct = False
        filtered_records = []
        
        if not records:
             st.warning("⚠️ Para habilitar las acciones, primero debe importar datos en la pestaña 'Importar Excel'.")
             disabled_struct = True
        else:
             filtered_records = get_filtered_records(records, key_suffix="struct")
             st.info(f"Registros seleccionados: {len(filtered_records)} de {len(records)}")
             if not filtered_records:
                 st.warning("No hay registros que coincidan con los filtros.")
                 disabled_struct = True

        if st.button("🚀 Crear Carpetas", type="primary", disabled=disabled_struct):
            count_created = 0
            progress_bar = st.progress(0)
            
            # --- Resolve unique paths ---
            resolved_paths = resolve_unique_paths(filtered_records, folder_structure)
            
            for i, record in enumerate(filtered_records):
                try:
                    rel_path = resolved_paths.get(record['id'])
                    full_path = os.path.join(base_path, rel_path)
                    
                    if do_create_folders:
                        os.makedirs(full_path, exist_ok=True)
                        count_created += 1
                        
                except Exception as e:
                    print(f"Error creating folder: {e}")
                
                progress_bar.progress((i + 1) / len(filtered_records))
            
            st.success(f"✅ Proceso finalizado. Carpetas Creadas: {count_created}")

    # --- TAB ORGANIZAR (Renombrar/Mover) ---
    with tab_organize:
        st.subheader("🚀 Organizar Archivos (Renombrar y Mover)")
        st.caption("Utilice los datos de la base de datos para renombrar y organizar archivos físicos.")
        
        # Configuración común
        # --- SYNC LOGIC: Auto-update from Search Tab ---
        if "current_path" in st.session_state and os.path.exists(st.session_state.current_path):
             # Check if we should sync (either first run or path changed)
             # We use a tracking variable to know if current_path changed since we last looked
             if "last_synced_search_path" not in st.session_state:
                 st.session_state.last_synced_search_path = None
             
             if st.session_state.current_path != st.session_state.last_synced_search_path:
                 # Sync!
                 st.session_state.gd_source_path = st.session_state.current_path
                 st.session_state.gd_source_path_mov = st.session_state.current_path # Sync Move Tool Source
                 st.session_state.input_base_path_struct_mov = st.session_state.current_path # Sync Move Tool Dest
                 st.session_state.last_synced_search_path = st.session_state.current_path
                 # Force update the widget key to reflect the change immediately
                 st.session_state.input_gd_source_path = st.session_state.current_path
                 st.session_state.input_gd_source_path_mov = st.session_state.current_path

        if "gd_source_path" not in st.session_state:
             # Intenta heredar la ruta de Búsqueda y Acciones si existe
             if "current_path" in st.session_state and os.path.exists(st.session_state.current_path):
                 st.session_state.gd_source_path = st.session_state.current_path
             else:
                 st.session_state.gd_source_path = os.path.join(os.path.expanduser("~"), "Documents")
             
        # Determinar ruta a usar
        current_global_path = st.session_state.get("current_path", os.getcwd())
        if not current_global_path: current_global_path = os.getcwd()

        source_path = render_path_selector(
            label="Carpeta Origen de Archivos",
            key="gd_source_path",
            default_path=current_global_path,
            omit_checkbox=True
        )
            
        if not records:
             st.warning("⚠️ No hay registros en la base de datos. Importe un Excel primero.")
        else:
            st.divider()
            col_ren, col_mov = st.columns(2)
            
            # --- RENOMBRAR CARPETAS ---
            with col_ren:
                st.markdown("#### 1. Renombrar Carpetas")
                st.info("Renombra carpetas en el origen usando datos de la BD.")
                
                # Get all available keys from the first record
                available_keys = list(records[0].keys()) if records else []
                # Common useful defaults
                default_keys = ["no_doc", "nombre_completo", "no_factura", "nro_estudio"]
                # Sort available keys for better UX, putting defaults first if they exist
                sorted_keys = sorted(list(set(available_keys) - set(default_keys)))
                display_keys = [k for k in default_keys if k in available_keys] + sorted_keys

                col_cur_name = st.selectbox("Columna Nombre Actual (Carpeta)", display_keys, key="sel_col_cur_name_folder")
                col_new_name = st.selectbox("Columna Nuevo Nombre (Carpeta)", display_keys, key="sel_col_new_name_folder")
                
                if st.button("Ejecutar Renombrado Carpetas", key="btn_run_ren_folder"):
                    count_ren = 0
                    errors_ren = 0
                    
                    # Map: current_name_value -> new_name_value
                    # Be careful with duplicates: last one wins or we skip?
                    # Let's assume unique mapping for now or log duplicates
                    map_ren = {}
                    for r in records:
                        curr = str(r.get(col_cur_name, "")).strip()
                        new_n = str(r.get(col_new_name, "")).strip()
                        if curr and new_n:
                            # Sanitize new name
                            new_n = "".join([c for c in new_n if c.isalnum() or c in (' ', '-', '_', '.')]).strip()
                            map_ren[curr] = new_n
                            
                    try:
                        dirs = [d for d in os.listdir(source_path) if os.path.isdir(os.path.join(source_path, d))]
                        progress_bar = st.progress(0)
                        
                        for i, dirname in enumerate(dirs):
                            dir_path = os.path.join(source_path, dirname)
                            
                            # Check match
                            # We look for exact match or loose match? User said "escogiendo la columna donde esta el nombre actual"
                            # Usually this implies exact match or contains. Let's try loose match for flexibility but prioritize exact.
                            
                            matched_new_name = None
                            
                            # Strategy: Direct lookup first
                            if dirname in map_ren:
                                matched_new_name = map_ren[dirname]
                            else:
                                # Loose match: dirname contains key? Or key contains dirname?
                                # If key is "12345" and folder is "12345 - Juan", loose match might be tricky if we want to rename TO "Juan".
                                # Let's stick to: Does dirname CONTAIN the value of Columna Nombre Actual?
                                for curr_val, new_val in map_ren.items():
                                    if curr_val in dirname:
                                        matched_new_name = new_val
                                        break
                            
                            if matched_new_name and matched_new_name != dirname:
                                try:
                                    new_path = os.path.join(source_path, matched_new_name)
                                    if os.path.exists(new_path):
                                        print(f"Skipping rename {dirname} -> {matched_new_name}: Dest exists")
                                        errors_ren += 1
                                    else:
                                        os.rename(dir_path, new_path)
                                        count_ren += 1
                                except Exception as e:
                                    errors_ren += 1
                                    print(f"Error renaming folder {dirname}: {e}")
                                    
                            progress_bar.progress((i + 1) / len(dirs))
                            
                        st.success(f"Carpetas Renombradas: {count_ren}. Errores/Omis: {errors_ren}")
                    except Exception as e:
                        st.error(f"Error accediendo a la carpeta: {e}")

            # --- MOVER ---
            with col_mov:
                st.markdown("#### 2. Mover a Estructura")
                st.info("Mueve archivos/carpetas de Origen -> Carpeta Destino.")
                
                source_path_mov = render_path_selector(
                    label="Origen (Mover)",
                    key="gd_source_path_mov",
                    default_path=st.session_state.get("current_path", os.getcwd())
                )

                # Reuse keys from above
                col_src_match = st.selectbox("Columna Coincidencia Origen", display_keys, key="sel_col_src_match_mov")
                col_dst_name = st.selectbox("Columna Nombre Destino", display_keys, key="sel_col_dst_name_mov")
                
                if "input_base_path_struct_mov" not in st.session_state:
                     st.session_state.input_base_path_struct_mov = st.session_state.get("input_base_path_struct", st.session_state.gd_base_path)

                # Selector de Ruta Estandarizado
                base_dest = render_path_selector(
                    label="Destino Base",
                    key="gd_dest_path_mov_selector",
                    default_path=st.session_state.get("input_base_path_struct_mov", st.session_state.get("current_path", os.getcwd()))
                )
                st.session_state.input_base_path_struct_mov = base_dest
                
                if st.button("Ejecutar Movimiento", key="btn_run_mov"):
                    count_mov = 0
                    errors_mov = 0
                    
                    # Map: src_value -> dst_value
                    map_mov = {}
                    for r in records:
                        src = str(r.get(col_src_match, "")).strip()
                        dst = str(r.get(col_dst_name, "")).strip()
                        if src and dst:
                            dst = "".join([c for c in dst if c.isalnum() or c in (' ', '-', '_', '.')]).strip()
                            map_mov[src] = dst

                    try:
                        items = os.listdir(source_path_mov) # Files and Folders
                        progress_bar = st.progress(0)
                        
                        for i, item_name in enumerate(items):
                            item_path = os.path.join(source_path_mov, item_name)
                            
                            # Find match
                            matched_dst_name = None
                            for src_val, dst_val in map_mov.items():
                                if src_val in item_name:
                                    matched_dst_name = dst_val
                                    break
                            
                            if matched_dst_name:
                                try:
                                    # Dest folder: Base + Dst_Name
                                    dest_folder_path = os.path.join(base_dest, matched_dst_name)
                                    os.makedirs(dest_folder_path, exist_ok=True)
                                    
                                    # Move the item into that folder
                                    final_item_path = os.path.join(dest_folder_path, item_name)
                                    shutil.move(item_path, final_item_path)
                                    count_mov += 1
                                except Exception as e:
                                    errors_mov += 1
                                    print(f"Error moving {item_name}: {e}")
                                    
                            progress_bar.progress((i + 1) / len(items))
                            
                        st.success(f"Items Movidos: {count_mov}. Errores: {errors_mov}")
                    except Exception as e:
                         st.error(f"Error accediendo a la carpeta: {e}")

            # --- AÑADIR SUFIJO ---
            st.divider()
            st.markdown("#### 3. Añadir Sufijo desde BD")
            st.info("Esta herramienta buscará carpetas o archivos que coincidan con la columna de búsqueda seleccionada y añadirá el sufijo especificado.")
            
            c_match, c_suf = st.columns(2)
            
            with c_match:
                # Select match column (default to col_cur_name if possible, else 0)
                default_idx = 0
                if col_cur_name in display_keys:
                    default_idx = display_keys.index(col_cur_name)
                match_col = st.selectbox("Columna para buscar Carpeta/Archivo", display_keys, index=default_idx, key="sel_match_col_suf")
                
            with c_suf:
                suffix_field = st.selectbox("Columna para Sufijo", display_keys, key="sel_field_suf_val")
                
            separator_suf = st.text_input("Separador", value="_", key="input_sep_suf")

            if st.button("Ejecutar Añadir Sufijo", key="btn_run_suf"):
                count_suf = 0
                errors_suf = 0
                
                try:
                    # Pre-fetch files and folders in root
                    root_files = []
                    root_folders = []
                    if os.path.exists(source_path):
                        try:
                            all_items = os.listdir(source_path)
                            root_files = [f for f in all_items if os.path.isfile(os.path.join(source_path, f))]
                            root_folders = [d for d in all_items if os.path.isdir(os.path.join(source_path, d))]
                        except Exception as e:
                            print(f"Error listing root items: {e}")
                    
                    st.toast(f"📂 Buscando en: {os.path.basename(source_path)} ({len(root_files)} archivos, {len(root_folders)} carpetas)", icon="🔍")
                    
                    # Iterate records to find target folders or files
                    progress_bar = st.progress(0)
                    total_records = len(records)
                    
                    for i, r in enumerate(records):
                        # 1. Identify Key from Match Config
                        match_key = str(r.get(match_col, "")).strip()
                        
                        if not match_key:
                            continue
                            
                        # Get suffix value
                        suffix_val = str(r.get(suffix_field, "")).strip()
                        # Sanitize suffix
                        suffix_val = "".join([c for c in suffix_val if c.isalnum() or c in (' ', '-', '_')]).strip()
                        
                        if not suffix_val:
                            continue

                        # A) FOLDER MODE: Check if folders match key OR key_N
                        # Manual check for robustness with consecutive numbering (e.g. Folder, Folder_1, Folder_2)
                        
                        matching_folders = []
                        normalized_key = match_key.lower()
                        
                        for d in root_folders:
                            d_lower = d.lower()
                            
                            # 1. Exact match
                            if d_lower == normalized_key:
                                matching_folders.append(d)
                                continue
                            
                            # 2. Consecutive match (Starts with key + "_")
                            # Check if it starts with key followed by underscore
                            # Also robustly handle if there are other suffixes if strict number check fails?
                            # User requirement: "que sufijo iria en las carpetas que fueron creadas con el _1"
                            # These are strictly Name_N.
                            
                            if d_lower.startswith(normalized_key):
                                remainder = d_lower[len(normalized_key):]
                                # Expecting "_1", "_2", etc.
                                if remainder.startswith("_") and len(remainder) > 1 and remainder[1:].isdigit():
                                    matching_folders.append(d)
                        
                        for folder_name in matching_folders:
                            folder_path = os.path.join(source_path, folder_name)
                            
                            # --- PROCESS FOLDER ---
                            if os.path.isdir(folder_path):
                                try:
                                    files = os.listdir(folder_path)
                                    for filename in files:
                                        file_full_path = os.path.join(folder_path, filename)
                                        if os.path.isfile(file_full_path):
                                            base_name, ext = os.path.splitext(filename)
                                            
                                            # Check if already ends with suffix
                                            if not base_name.endswith(f"{separator_suf}{suffix_val}"):
                                                new_name = f"{base_name}{separator_suf}{suffix_val}{ext}"
                                                new_file_path = os.path.join(folder_path, new_name)
                                                
                                                os.rename(file_full_path, new_file_path)
                                                count_suf += 1
                                except Exception as e:
                                    print(f"Error iterating folder {folder_name}: {e}")
                                    errors_suf += 1
                        
                        # --- PROCESS FLAT FILES (Root) ---
                        # Look for files in root that CONTAIN the match_key
                        # Only if we didn't process it as a folder? Or both? 
                        # Usually user has either folders OR flat files. Doing both is safe.
                        for filename in root_files:
                            # Loose match: filename contains match_key
                            if match_key in filename:
                                try:
                                    file_full_path = os.path.join(source_path, filename)
                                    if not os.path.exists(file_full_path): continue # Moved or renamed already
                                    
                                    base_name, ext = os.path.splitext(filename)
                                    
                                    # Check if already ends with suffix
                                    if not base_name.endswith(f"{separator_suf}{suffix_val}"):
                                        new_name = f"{base_name}{separator_suf}{suffix_val}{ext}"
                                        new_file_path = os.path.join(source_path, new_name)
                                        
                                        # Avoid collision
                                        if os.path.exists(new_file_path):
                                            # Skip or timestamp? Skip to be safe
                                            continue
                                            
                                        os.rename(file_full_path, new_file_path)
                                        count_suf += 1
                                        
                                        # Update root_files list effectively? 
                                        # We can't modify the list we are iterating easily, 
                                        # but since we check os.path.exists, it's okay.
                                        # However, we might rename a file that matches 2 records.
                                        # e.g. "123" and "1234". "File_1234" matches "123".
                                        # This is a risk. But for now acceptable.
                                except Exception as e:
                                    print(f"Error renaming file {filename}: {e}")
                                    errors_suf += 1
                        
                        progress_bar.progress((i + 1) / total_records)
                        
                    st.success(f"Archivos renombrados con sufijo: {count_suf}. Errores/Omitidos: {errors_suf}")
                except Exception as e:
                     st.error(f"Error general: {e}")

    # --- TAB CONTENIDO (Former Subtab 2) ---
    with tab_content:
        st.subheader("📝 Generación de Contenido (Docs y Firmas)")
        
        default_hardcoded = os.path.join(os.path.expanduser("~"), "Documents", "GestionDocumental")
        
        # 1. Config (Always Visible)
        st.markdown("### 1. Configuración General")
        
        current_global_path = st.session_state.get("current_path", os.getcwd())
        if not current_global_path: current_global_path = os.getcwd()
        
        base_path_content = render_path_selector(
            label="Carpeta Raíz",
            key="input_base_path_content",
            default_path=st.session_state.get("gd_base_path", default_hardcoded)
        )
        
        st.caption("Patrón de Carpetas (Debe coincidir con la estructura creada)")
        folder_structure_content = st.text_input("Patrón de Carpetas", value="{eps}/{no_factura} - {nombre_completo}", key="input_folder_structure_content")

        # 2. Action Config (Always Visible)
        st.divider()
        st.markdown("### 2. Configuración de Acciones")
        
        col_conf1, col_conf2 = st.columns(2)
        
        with col_conf1:
            st.markdown("#### Documentos (.docx)")
            template_file = st.file_uploader("Subir Plantilla Base (.docx)", type=["docx"], key="uploader_template_gen")
            template_bytes = None
            if template_file:
                template_bytes = template_file.getvalue()
            st.caption("Si no sube plantilla, buscará 'plantilla.docx' en cada carpeta.")

        with col_conf2:
            st.markdown("#### Firmas")
            st.caption("Firma Web (URL)")
            sign_url_pattern = st.text_input("URL Patrón", value="https://oportunidaddevida.com/opvcitas/admisionescall/firmas/{no_doc}.png", key="input_sign_url_gen")
            
            st.caption("Firma Digital (Generada)")
            font_name = st.radio("Fuente", ["Pacifico", "My Ugly Handwriting"], index=0, horizontal=True)
            font_size = st.number_input("Tamaño Fuente", value=70, min_value=20, max_value=200, step=5)
            use_natural_style = st.checkbox("Estilo Natural (Rotación)", value=True)

        st.divider()
        st.markdown("### 3. Ejecutar Acciones")

        disabled_state = False
        filtered_records_content = []
        
        if not records:
            st.warning("⚠️ Para ejecutar acciones, primero debe importar datos en la pestaña 'Importar Excel'.")
            disabled_state = True
        else:
            filtered_records_content = get_filtered_records(records, key_suffix="content")
            st.info(f"Registros a procesar: {len(filtered_records_content)}")
        
        col_act1, col_act2, col_act3, col_act4 = st.columns(4)
        
        # ACTION 1: DISTRIBUTE BASE DOC (Copy Template Only)
        with col_act1:
            if st.button("� Distribuir Base", use_container_width=True, disabled=disabled_state):
                count_dist = 0
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                if not template_bytes:
                    st.error("❌ Debe subir una plantilla base primero.")
                else:
                    # --- Resolve unique paths ---
                    resolved_paths_content = resolve_unique_paths(filtered_records_content, folder_structure_content)

                    for i, record in enumerate(filtered_records_content):
                        try:
                            rel_path = resolved_paths_content.get(record['id'])
                            if not rel_path:
                                safe_record = {k: str(v).replace("/", "-").replace("\\", "-").strip() for k, v in record.items()}
                                rel_path = folder_structure_content.format(**safe_record).upper()
                                
                            full_path = os.path.join(base_path_content, rel_path)
                            os.makedirs(full_path, exist_ok=True)
                            
                            # Save base template
                            dest_base = os.path.join(full_path, "documento_base.docx")
                            with open(dest_base, "wb") as f:
                                f.write(template_bytes)
                            count_dist += 1
                            
                        except Exception as e:
                            print(f"Error distributing doc: {e}")
                        
                        progress_bar.progress((i + 1) / len(filtered_records_content))
                        status_text.text(f"Distribuyendo {i+1}/{len(filtered_records_content)}...")
                    
                    st.success(f"✅ Base Distribuida: {count_dist}")

        # ACTION 2: MODIFY DOC (Fill Data)
        with col_act2:
            # Logic is always enabled by default (hidden from UI)
            use_hardcoded_logic = True 
            if st.button("📝 Llenar Datos", use_container_width=True, disabled=disabled_state):
                if Document is None:
                    st.error("❌ La librería 'python-docx' no está instalada. No se pueden modificar documentos.")
                else:
                    count_template = 0
                    errors_template = []
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    # --- Resolve unique paths ---
                    resolved_paths_content = resolve_unique_paths(filtered_records_content, folder_structure_content)
                    
                    # Show available keys for debugging
                    if filtered_records_content:
                        sample_keys = list(filtered_records_content[0].keys())
                        st.info(f"🔑 Claves disponibles para plantilla: {', '.join([f'{{{k}}}' for k in sample_keys])}")

                    for i, record in enumerate(filtered_records_content):
                        try:
                            safe_record = {k: str(v).replace("/", "-").replace("\\", "-").strip() for k, v in record.items()}
                            rel_path = resolved_paths_content.get(record['id'])
                            if not rel_path:
                                rel_path = folder_structure_content.format(**safe_record).upper()
                                
                            full_path = os.path.join(base_path_content, rel_path)
                            os.makedirs(full_path, exist_ok=True)
                            
                            # Logic: Prefer 'documento_base.docx', then 'plantilla.docx', then any docx
                            doc_to_process = None
                            doc_path = None # Initialize to avoid UnboundLocalError
                            
                            # 1. Try documento_base.docx (Distributed)
                            base_path_file = os.path.join(full_path, "documento_base.docx")
                            if os.path.exists(base_path_file):
                                doc_to_process = Document(base_path_file)
                                doc_path = base_path_file
                            
                            # 2. Try template_bytes if uploaded (Direct mode)
                            elif template_bytes:
                                doc_to_process = Document(BytesIO(template_bytes))
                                
                            # 3. Fallback to existing files
                            else:
                                # Allow 'documento_generado' but prefer 'plantilla.docx'
                                local_candidates = [f for f in os.listdir(full_path) if f.lower().endswith(".docx") and not f.startswith("~")]
                                
                                if local_candidates:
                                    if "plantilla.docx" in local_candidates:
                                        doc_path = os.path.join(full_path, "plantilla.docx")
                                    elif "documento_generado.docx" in local_candidates:
                                        doc_path = os.path.join(full_path, "documento_generado.docx")
                                    else:
                                        doc_path = os.path.join(full_path, local_candidates[0])
                                    
                                    if doc_path:
                                        doc_to_process = Document(doc_path)
                            
                            if doc_to_process:
                                # --- CUSTOM MAPPING LOGIC (User Request) ---
                                # Helper to format dates without time
                                def fmt_date(val):
                                    if not val: return ""
                                    s = str(val).split(" ")[0] # Remove time part if exists
                                    return s

                                custom_map = {
                                    "Nombre carpeta": safe_record.get("nombre_completo", ""),
                                    "Nombre completo": safe_record.get("nombre_completo", ""),
                                    "Ciudad y Fecha": fmt_date(safe_record.get("fecha_salida", "")),
                                    "Tipo Documento": safe_record.get("tipo_doc", ""),
                                    "Numero Documento": safe_record.get("no_doc", ""),
                                    "Servicio": safe_record.get("descripcion", ""), # descripcion_cups aliased as descripcion
                                    "EPS": safe_record.get("eps", ""),
                                    "Tipo Servicio": safe_record.get("tipo_servicio", ""),
                                    "Regimen": safe_record.get("regimen", ""),
                                    "Categoria": safe_record.get("categoria", ""),
                                    "Valor Cuota Moderadora": safe_record.get("copago", ""),
                                    "Valor Cuota Moderador": safe_record.get("copago", ""), # Variation
                                    "Numero Autorizacion": safe_record.get("autorizacion", ""),
                                    "Fecha y Hora Atencion": fmt_date(safe_record.get("fecha_ingreso", "")),
                                    "Fecha Finalizacion": fmt_date(safe_record.get("fecha_salida", ""))
                                }
                                
                                # Merge custom map into safe_record so we can use both keys
                                # But we iterate over custom_map first to prioritize these specific placeholders
                                
                                replacements_made = 0

                                # --- Enhanced Replacement Logic with Regex and XML Support ---
                                
                                def replace_text_in_element(paragraph, mapping):
                                    """
                                    Replaces placeholders in a paragraph object using regex for flexibility.
                                    Returns the number of replacements made.
                                    """
                                    if not paragraph.text.strip():
                                        return 0
                                    
                                    # Optimization: Quick check for '{' or '<'
                                    text_has_braces = "{" in paragraph.text
                                    text_has_chevrons = "<<" in paragraph.text or "«" in paragraph.text
                                    
                                    if not (text_has_braces or text_has_chevrons):
                                        return 0

                                    original_text = paragraph.text
                                    current_text = original_text
                                    count = 0
                                    
                                    for key, val in mapping.items():
                                        # Clean key for variations (e.g. "Nombre Completo" -> "NombreCompleto")
                                        key_clean = key.replace(" ", "")
                                        
                                        # Build regex patterns
                                        # 1. Exact with spaces: { Key }
                                        p1 = r"\{\s*" + re.escape(key) + r"\s*\}"
                                        # 2. No spaces in key: { KeyClean }
                                        p2 = r"\{\s*" + re.escape(key_clean) + r"\s*\}"
                                        # 3. Underscores: { Key_With_Underscores }
                                        p3 = r"\{\s*" + re.escape(key.replace(" ", "_")) + r"\s*\}"
                                        
                                        # Chevrons variations
                                        c1 = r"(?:«|<<)\s*" + re.escape(key) + r"\s*(?:»|>>)"
                                        c2 = r"(?:«|<<)\s*" + re.escape(key_clean) + r"\s*(?:»|>>)"
                                        c3 = r"(?:«|<<)\s*" + re.escape(key.replace(" ", "_")) + r"\s*(?:»|>>)"
                                        
                                        # Combine patterns
                                        patterns = [p1, p2, p3] if text_has_braces else []
                                        if text_has_chevrons:
                                            patterns.extend([c1, c2, c3])
                                            
                                        for pat in patterns:
                                            if re.search(pat, current_text, re.IGNORECASE):
                                                current_text = re.sub(pat, str(val), current_text, flags=re.IGNORECASE)
                                                count += 1
                                    
                                    if count > 0 and current_text != original_text:
                                        paragraph.text = current_text
                                        return count
                                    return 0

                                # Helper to iterate all paragraphs including those in Text Boxes (Shapes)
                                def iter_all_paragraphs(doc_obj):
                                    # 1. Main Body Paragraphs
                                    yield from doc_obj.paragraphs
                                    
                                    # 2. Tables in Body
                                    for table in doc_obj.tables:
                                        for row in table.rows:
                                            for cell in row.cells:
                                                yield from cell.paragraphs
                                    
                                    # 3. Text Boxes in Body (via XML)
                                    # Find all <w:txbxContent> elements
                                    if doc_obj.element.body is not None:
                                        for txbx in doc_obj.element.body.iter(qn('w:txbxContent')):
                                            for p_element in txbx.iter(qn('w:p')):
                                                yield Paragraph(p_element, doc_obj)

                                    # 4. Headers and Footers (including their tables and text boxes)
                                    for section in doc_obj.sections:
                                        # Process all header/footer types
                                        headers = [section.header, section.first_page_header, section.even_page_header]
                                        footers = [section.footer, section.first_page_footer, section.even_page_footer]
                                        
                                        for header in headers:
                                            if header and not header.is_linked_to_previous:
                                                yield from header.paragraphs
                                                for table in header.tables:
                                                    for row in table.rows:
                                                        for cell in row.cells:
                                                            yield from cell.paragraphs
                                                # Text Boxes in Header
                                                for txbx in header._element.iter(qn('w:txbxContent')):
                                                    for p_element in txbx.iter(qn('w:p')):
                                                        yield Paragraph(p_element, header)

                                        for footer in footers:
                                            if footer and not footer.is_linked_to_previous:
                                                yield from footer.paragraphs
                                                for table in footer.tables:
                                                    for row in table.rows:
                                                        for cell in row.cells:
                                                            yield from cell.paragraphs
                                                # Text Boxes in Footer
                                                for txbx in footer._element.iter(qn('w:txbxContent')):
                                                    for p_element in txbx.iter(qn('w:p')):
                                                        yield Paragraph(p_element, footer)

                                # Execute Replacement
                                for p in iter_all_paragraphs(doc_to_process):
                                    replacements_made += replace_text_in_element(p, custom_map)
                                    replacements_made += replace_text_in_element(p, safe_record)

                                # --- LEGACY / HARDCODED LOGIC (Requested by User) ---
                                if use_hardcoded_logic:
                                    # Map safe_record to legacy keys
                                    legacy_data = {
                                        'date': fmt_date(safe_record.get("fecha_salida", "")), 
                                        'full_name': safe_record.get("nombre_completo", ""),
                                        'doc_type': safe_record.get("tipo_doc", ""),
                                        'doc_num': safe_record.get("no_doc", ""),
                                        'service': safe_record.get("descripcion", ""),
                                        'eps': safe_record.get("eps", ""),
                                        'tipo_servicio': safe_record.get("tipo_servicio", ""),
                                        'regimen': safe_record.get("regimen", ""),
                                        'categoria': safe_record.get("categoria", ""),
                                        'cuota': safe_record.get("copago", ""),
                                        'auth': safe_record.get("autorizacion", ""),
                                        'fecha_atencion': fmt_date(safe_record.get("fecha_ingreso", "")),
                                        'fecha_fin': fmt_date(safe_record.get("fecha_salida", ""))
                                    }
                                    
                                    # Iterate again for specific text patterns (Santiago, Yo..., EPS...)
                                    for p in iter_all_paragraphs(doc_to_process):
                                        txt = p.text
                                        if not txt.strip(): continue

                                        # 1. Santiago de Cali
                                        if "Santiago de Cali, " in txt: 
                                            if legacy_data['date'] and legacy_data['date'] not in txt:
                                                p.text = f"Santiago de Cali,  {legacy_data['date']}"
                                                replacements_made += 1
                                        
                                        # 2. Yo ... identificado con ...
                                        if "Yo " in txt and "identificado con" in txt:
                                            p.text = f"Yo {legacy_data['full_name']} identificado con {legacy_data['doc_type']}, Numero {legacy_data['doc_num']} en calidad de paciente, doy fé y acepto el servicio de {legacy_data['service']} brindado por la IPS OPORTUNIDAD DE VIDA S.A.S"
                                            replacements_made += 1
                                        
                                        # 3. Key-Value replacements (EPS:, TIPO SERVICIO:, etc.)
                                        legacy_replacements = {
                                            "EPS:": legacy_data['eps'], "TIPO SERVICIO:": legacy_data['tipo_servicio'],
                                            "REGIMEN:": legacy_data['regimen'], "CATEGORIA:": legacy_data['categoria'],
                                            "VALOR CUOTA MODERADORA:": legacy_data['cuota'], "AUTORIZACION:": legacy_data['auth'],
                                            "Fecha de Atención:": legacy_data['fecha_atencion'], "Fecha de Finalización:": legacy_data['fecha_fin']
                                        }
                                        
                                        for key, val in legacy_replacements.items():
                                            if key in txt:
                                                # Regex to replace value after key
                                                pattern = rf'({re.escape(key)})\s*.*'
                                                if re.search(pattern, txt):
                                                    new_text = re.sub(pattern, r'\1 ' + str(val), p.text, count=1)
                                                    if new_text != p.text:
                                                        p.text = new_text
                                                        replacements_made += 1

                                    # 4. Signature (Search for "FIRMA DE ACEPTACION")
                                    sig_idx = -1
                                    main_paragraphs = list(doc_to_process.paragraphs)
                                    for idx, p in enumerate(main_paragraphs):
                                        if "FIRMA DE ACEPTACION" in p.text.upper():
                                            sig_idx = idx
                                            break
                                    
                                    if sig_idx != -1 and sig_idx + 2 < len(main_paragraphs):
                                        main_paragraphs[sig_idx + 2].text = legacy_data['full_name'].upper()
                                        replacements_made += 1

                                # Save: Overwrite the source file (or default to documento_base.docx) to avoid duplication
                                dest_doc = doc_path if doc_path else os.path.join(full_path, "documento_base.docx")
                                doc_to_process.save(dest_doc)
                                
                                # Remove legacy 'plantilla.docx' if we are overwriting 'documento_base.docx' to clean up
                                if os.path.basename(dest_doc) == "documento_base.docx":
                                    legacy_plantilla = os.path.join(full_path, "plantilla.docx")
                                    if os.path.exists(legacy_plantilla):
                                        try:
                                            os.remove(legacy_plantilla)
                                        except: pass

                                count_template += 1
                                
                                # Optional: Log if no replacements were made
                                if replacements_made == 0:
                                    # Diagnostic: Get a sample of the text to show the user
                                    sample_text = []
                                    # Use the same iterator to see what the system sees (tables, boxes, etc.)
                                    diag_iter = iter_all_paragraphs(doc_to_process)
                                    for _ in range(10): # Check first 10 distinct text chunks
                                        try:
                                            p = next(diag_iter)
                                            if p.text.strip():
                                                sample_text.append(p.text.strip())
                                        except StopIteration:
                                            break
                                            
                                    sample_str = " | ".join(sample_text)[:250]
                                    
                                    errors_template.append(f"⚠️ {record.get('nombre_completo', 'Unknown')}: No se encontraron placeholders. (Texto detectado: {sample_str}...)")
                            else:
                                errors_template.append(f"No se encontró documento base o plantilla en: {rel_path}")
                                
                        except Exception as e:
                            errors_template.append(f"Error modificando docx para {record.get('nombre_completo', 'Unknown')}: {e}")
                            print(f"Error modifying docx: {e}")
                        
                        progress_bar.progress((i + 1) / len(filtered_records_content))
                        status_text.text(f"Procesando DOCX {i+1}/{len(filtered_records_content)}...")
                    
                    st.success(f"✅ Documentos Modificados: {count_template}")
                    if errors_template:
                        with st.expander(f"⚠️ Errores ({len(errors_template)})", expanded=True):
                            for err in errors_template:
                                st.error(err)

        # ACTION 3: DOWNLOAD SIGNATURES
        with col_act3:
            if st.button("⬇️ Descargar Firmas", use_container_width=True, disabled=disabled_state):
                count_sign_download = 0
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                # --- Resolve unique paths ---
                resolved_paths_content = resolve_unique_paths(filtered_records_content, folder_structure_content)

                for i, record in enumerate(filtered_records_content):
                    try:
                        safe_record = {k: str(v).replace("/", "-").replace("\\", "-").strip() for k, v in record.items()}
                        rel_path = resolved_paths_content.get(record['id'])
                        if not rel_path:
                             rel_path = folder_structure_content.format(**safe_record).upper()
                             
                        full_path = os.path.join(base_path_content, rel_path)
                        os.makedirs(full_path, exist_ok=True)
                        
                        if sign_url_pattern:
                            sign_url = sign_url_pattern.format(**safe_record)
                            try:
                                r = requests.get(sign_url, timeout=5)
                                if r.status_code == 200:
                                    with open(os.path.join(full_path, "firma_descargada.png"), "wb") as f:
                                        f.write(r.content)
                                    count_sign_download += 1
                            except Exception as e:
                                print(f"Error downloading sign: {e}")
                                
                    except Exception as e:
                        print(f"Error processing record {i}: {e}")
                    
                    progress_bar.progress((i + 1) / len(filtered_records_content))
                    status_text.text(f"Descargando Firmas {i+1}/{len(filtered_records_content)}...")
                    
                st.success(f"✅ Firmas Descargadas: {count_sign_download}")

        # ACTION 4: CREATE DIGITAL SIGNATURES
        with col_act4:
            if st.button("✍️ Crear Firmas Digitales", use_container_width=True, disabled=disabled_state):
                count_sign_create = 0
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                # --- Resolve unique paths ---
                resolved_paths_content = resolve_unique_paths(filtered_records_content, folder_structure_content)
                
                for i, record in enumerate(filtered_records_content):
                    try:
                        safe_record = {k: str(v).replace("/", "-").replace("\\", "-").strip() for k, v in record.items()}
                        rel_path = resolved_paths_content.get(record['id'])
                        if not rel_path:
                            rel_path = folder_structure_content.format(**safe_record).upper()
                            
                        full_path = os.path.join(base_path_content, rel_path)
                        os.makedirs(full_path, exist_ok=True)
                        
                        # Name Logic
                        final_name = ""
                        tipo_doc = str(record.get("tipo_doc", "")).upper()
                        nombre_tercero = str(record.get("nombre_tercero", "")).strip()
                        
                        if tipo_doc in ["TI", "RC"] and nombre_tercero:
                            full_name_source = nombre_tercero
                        else:
                            full_name_source = str(record.get("nombre_completo", "")).strip()
                        
                        # Parse First Name + First Surname
                        parts = full_name_source.split()
                        if len(parts) >= 3:
                            if len(parts) >= 4:
                                final_name = f"{parts[0]} {parts[2]}"
                            else:
                                final_name = f"{parts[0]} {parts[1]}"
                        elif len(parts) == 2:
                            final_name = f"{parts[0]} {parts[1]}"
                        elif len(parts) == 1:
                            final_name = parts[0]
                        
                        if final_name:
                            # Title Case
                            final_name = final_name.title()
                            
                            # Generate
                            img = generate_signature_image(final_name, font_name=font_name, size=font_size)
                            
                            if use_natural_style:
                                import random
                                angle = random.uniform(-2, 2)
                                img = img.rotate(angle, resample=Image.BICUBIC, expand=True, fillcolor="white")
                            
                            save_path = os.path.join(full_path, "firma_digital.jpg")
                            img.save(save_path)
                            count_sign_create += 1
                            
                    except Exception as e:
                        print(f"Error creating signature {i}: {e}")
                    
                    progress_bar.progress((i + 1) / len(filtered_records_content))
                    status_text.text(f"Creando Firmas {i+1}/{len(filtered_records_content)}...")
                    
                st.success(f"✅ Firmas Creadas: {count_sign_create}")

        # --- ACTION 5: DOWNLOAD OVIDA HISTORY ---
        st.divider()
        st.markdown("### 4. Integraciones Externas")

        # Configuración de Ruta Específica para OVIDA
        
        final_ovida_path = render_path_selector(
            label="Carpeta Base para Descargas OVIDA",
            key="ovida_base_path",
            default_path=base_path_content,
            help_text="Ruta donde se guardarán las historias clínicas. Si se deja vacía, usará la configuración general."
        )

        col_ovida_1, col_ovida_2 = st.columns([0.3, 0.7])
        
        with col_ovida_1:
             st.info(f"Descarga automática de historias clínicas desde OVIDA usando Selenium.\n\nRuta destino: {final_ovida_path}")
             
        with col_ovida_2:
             if st.button("🏥 Descargar Historias OVIDA", use_container_width=True, disabled=disabled_state):
                 if webdriver is None:
                     st.error("❌ Selenium no está instalado. No se puede ejecutar esta acción.")
                 elif not final_ovida_path or not os.path.isdir(final_ovida_path):
                     st.error(f"❌ La ruta de descarga no es válida: {final_ovida_path}")
                 else:
                     # --- Resolve unique paths ---
                     resolved_paths_content = resolve_unique_paths(filtered_records_content, folder_structure_content)
                     
                     st.info("⏳ Iniciando proceso... Se abrirá un navegador.")
                     
                     # Call worker
                     result_msg = worker_descargar_historias_ovida(
                         filtered_records_content, 
                         final_ovida_path, 
                         resolved_paths_content
                     )
                     
                     if "Error" in result_msg:
                         st.error(result_msg)
                     else:
                         st.success(result_msg)

    # --- TAB 5: ADMIN BD (MOVED TO TAB_ADMIN.PY) ---
    # Logic moved to main admin tab.

