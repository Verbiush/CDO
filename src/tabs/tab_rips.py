import streamlit as st
import os
import json
import pandas as pd
import io
import time
import shutil

try:
    from src.gui_utils import abrir_dialogo_carpeta_nativo, render_path_selector
except ImportError:
    abrir_dialogo_carpeta_nativo = None
    render_path_selector = None

# --- HELPERS ---

def get_val_ci(data_dict, key):
    if not isinstance(data_dict, dict): return None
    for k, v in data_dict.items():
        if k.lower() == key.lower():
            return v
    return None

def recursive_update_cups(data, old_val, new_val):
    count = 0
    if isinstance(data, dict):
        for k, v in data.items():
            if k == "codServicio" and str(v).strip() == str(old_val).strip():
                data[k] = new_val
                count += 1
            elif isinstance(v, (dict, list)):
                count += recursive_update_cups(v, old_val, new_val)
    elif isinstance(data, list):
        for item in data:
            count += recursive_update_cups(item, old_val, new_val)
    return count

def recursive_update_notes(data, target_text, new_note):
    count = 0
    if isinstance(data, dict):
        for k, v in data.items():
            if isinstance(v, str) and target_text in v:
                data[k] = new_note
                count += 1
            elif isinstance(v, (dict, list)):
                count += recursive_update_notes(v, target_text, new_note)
    elif isinstance(data, list):
        for item in data:
            count += recursive_update_notes(item, target_text, new_note)
    return count

def recursive_update_key(data, key, new_val):
    count = 0
    if isinstance(data, dict):
        for k, v in data.items():
            if k == key:
                data[k] = new_val
                count += 1
            elif isinstance(v, (dict, list)):
                count += recursive_update_key(v, key, new_val)
    elif isinstance(data, list):
        for item in data:
            count += recursive_update_key(item, key, new_val)
    return count

def recursive_strip(data):
    """Elimina espacios en blanco al inicio y final de strings en claves y valores recursivamente."""
    if isinstance(data, dict):
        return {k.strip(): recursive_strip(v) for k, v in data.items()}
    elif isinstance(data, list):
        return [recursive_strip(i) for i in data]
    elif isinstance(data, str):
        return data.strip()
    return data

def clean_df_for_json(df):
    """Limpia un DataFrame para conversión a JSON, manejando NaN y tipos."""
    df.columns = [str(c).strip() for c in df.columns]
    df = df.where(pd.notnull(df), None)
    numeric_fields = [
        "consecutivo", "consecutivo_usuario", "codservicio", "vrservicio", 
        "valorpagomoderador", "copago", "cuotamoderadora", 
        "numfevpagomoderador", "bonificacion", "valortotal", 
        "cantidad", "valorunitario"
    ]
    for col in df.columns:
        if col.lower() in numeric_fields:
            df[col] = pd.to_numeric(df[col], errors='coerce')
    return df

# --- WORKERS ---

def worker_json_a_xlsx_ind(file_obj):
    try:
        if hasattr(file_obj, 'seek'):
            file_obj.seek(0)
        data = json.load(file_obj)
        
        service_map = {
            "consultas": "Consultas", "procedimientos": "Procedimientos", "urgencias": "Urgencias",
            "hospitalizacion": "Hospitalizacion", "recienNacidos": "RecienNacidos",
            "medicamentos": "Medicamentos", "otrosServicios": "OtrosServicios"
        }
        
        header_info = {
            "numDocumentoIdObligado": data.get("numDocumentoIdObligado"),
            "numFactura": data.get("numFactura"),
            "tipoNota": data.get("tipoNota"),
            "numNota": data.get("numNota")
        }
        
        usuarios_rows = []
        all_services = {name: [] for name in service_map.values()}
        usuarios_lista = data.get("usuarios", []) if isinstance(data, dict) else []
        
        for usuario in usuarios_lista:
            u_info = {
                "tipoDocumentoIdentificacion": get_val_ci(usuario, "tipoDocumentoIdentificacion"),
                "numDocumentoIdentificacion": get_val_ci(usuario, "numDocumentoIdentificacion"),
                "tipoUsuario": get_val_ci(usuario, "tipoUsuario"),
                "fechaNacimiento": get_val_ci(usuario, "fechaNacimiento"),
                "codSexo": get_val_ci(usuario, "codSexo"),
                "codPaisResidencia": get_val_ci(usuario, "codPaisResidencia"),
                "codMunicipioResidencia": get_val_ci(usuario, "codMunicipioResidencia"),
                "codZonaTerritorialResidencia": get_val_ci(usuario, "codZonaTerritorialResidencia"),
                "incapacidad": get_val_ci(usuario, "incapacidad"),
                "consecutivo": get_val_ci(usuario, "consecutivo"),
                "codPaisOrigen": get_val_ci(usuario, "codPaisOrigen"),
            }
            usuarios_rows.append(u_info)
            servicios = usuario.get("servicios", {})
            for json_key, sheet_name in service_map.items():
                items = get_val_ci(servicios, json_key)
                if items and isinstance(items, list):
                    for item in items:
                        item_copy = item.copy()
                        item_copy["consecutivoUsuario"] = u_info["consecutivo"]
                        all_services[sheet_name].append(item_copy)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            pd.DataFrame([header_info]).to_excel(writer, sheet_name="Transaccion", index=False)
            if usuarios_rows:
                pd.DataFrame(usuarios_rows).to_excel(writer, sheet_name="Usuarios", index=False)
            for sheet_name, rows in all_services.items():
                if rows:
                    pd.DataFrame(rows).to_excel(writer, sheet_name=sheet_name, index=False)
        return output.getvalue(), None
    except Exception as e:
        return None, str(e)

def worker_xlsx_a_json_ind(file_obj):
    try:
        if hasattr(file_obj, 'seek'):
            file_obj.seek(0)
        xls = pd.ExcelFile(file_obj)
        service_map = {
            "Consultas": "consultas", "Procedimientos": "procedimientos", "Urgencias": "urgencias",
            "Hospitalizacion": "hospitalizacion", "RecienNacidos": "recienNacidos",
            "Medicamentos": "medicamentos", "OtrosServicios": "otrosServicios"
        }
        if "Transaccion" in xls.sheet_names and "Usuarios" in xls.sheet_names:
            df_t = pd.read_excel(xls, sheet_name="Transaccion")
            df_t = clean_df_for_json(df_t)
            transaccion_data = df_t.iloc[0].to_dict() if not df_t.empty else {}
            
            usuarios_map = {}
            df_u = pd.read_excel(xls, sheet_name="Usuarios")
            df_u = clean_df_for_json(df_u)
            for _, row in df_u.iterrows():
                u_obj = row.to_dict()
                u_obj["servicios"] = {k: [] for k in service_map.values()}
                usuarios_map[str(u_obj.get("consecutivo"))] = u_obj
            
            for sheet_name, json_key in service_map.items():
                if sheet_name in xls.sheet_names:
                    df_s = pd.read_excel(xls, sheet_name=sheet_name)
                    df_s = clean_df_for_json(df_s)
                    for _, row in df_s.iterrows():
                        s_obj = row.to_dict()
                        cons_u = str(s_obj.pop("consecutivoUsuario", None))
                        if cons_u in usuarios_map:
                            usuarios_map[cons_u]["servicios"][json_key].append(s_obj)
            
            final_json = transaccion_data
            final_json["usuarios"] = list(usuarios_map.values())
            return json.dumps(final_json, ensure_ascii=False, indent=4), None
        return None, "Formato Excel inválido (Faltan hojas Transaccion/Usuarios)"
    except Exception as e:
        return None, str(e)

def worker_consolidar_json_xlsx(folder_path):
    try:
        if not os.path.isdir(folder_path):
            return None, "La ruta proporcionada no es una carpeta válida."

        json_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.json')]
        if not json_files:
            return None, "No hay archivos JSON en la carpeta."
        
        master_header = []
        master_users = []
        master_services = {
            "Consultas": [], "Procedimientos": [], "Urgencias": [], 
            "Hospitalizacion": [], "RecienNacidos": [], "Medicamentos": [], "OtrosServicios": []
        }
        service_map = {
            "consultas": "Consultas", "procedimientos": "Procedimientos", "urgencias": "Urgencias",
            "hospitalizacion": "Hospitalizacion", "recienNacidos": "RecienNacidos",
            "medicamentos": "Medicamentos", "otrosServicios": "OtrosServicios"
        }
        
        progress_bar = st.progress(0, text="Consolidando...")
        total_files = len(json_files)

        for idx, fname in enumerate(json_files):
            progress_bar.progress((idx + 1) / total_files, text=f"Procesando {fname}")
            with open(os.path.join(folder_path, fname), 'r', encoding='utf-8') as f:
                data = json.load(f)
            h_info = {
                "archivo_origen": fname,
                "numDocumentoIdObligado": data.get("numDocumentoIdObligado"),
                "numFactura": data.get("numFactura")
            }
            master_header.append(h_info)
            usuarios = data.get("usuarios", [])
            for u in usuarios:
                u_clean = {k: v for k, v in u.items() if k != "servicios"}
                u_clean["archivo_origen"] = fname
                master_users.append(u_clean)
                servicios = u.get("servicios", {})
                for j_key, s_name in service_map.items():
                    items = get_val_ci(servicios, j_key)
                    if items:
                        for item in items:
                            item_copy = item.copy()
                            item_copy["archivo_origen"] = fname
                            item_copy["consecutivoUsuario"] = u.get("consecutivo")
                            master_services[s_name].append(item_copy)
        
        progress_bar.empty()
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            pd.DataFrame(master_header).to_excel(writer, sheet_name="Transaccion", index=False)
            pd.DataFrame(master_users).to_excel(writer, sheet_name="Usuarios", index=False)
            for s_name, rows in master_services.items():
                if rows:
                    pd.DataFrame(rows).to_excel(writer, sheet_name=s_name, index=False)
        return output.getvalue(), f"Consolidados {len(json_files)} archivos."
    except Exception as e:
        return None, str(e)

def worker_desconsolidar_xlsx_json(file_obj, dest_folder):
    try:
        if hasattr(file_obj, 'seek'):
            file_obj.seek(0)
        xls = pd.ExcelFile(file_obj)
        if "Transaccion" not in xls.sheet_names:
            return False, "Falta hoja Transaccion"
        df_t = pd.read_excel(xls, sheet_name="Transaccion")
        if "archivo_origen" not in df_t.columns:
            return False, "Falta columna 'archivo_origen' en Transaccion"
            
        service_map = {
            "Consultas": "consultas", "Procedimientos": "procedimientos", "Urgencias": "urgencias",
            "Hospitalizacion": "hospitalizacion", "RecienNacidos": "recienNacidos",
            "Medicamentos": "medicamentos", "OtrosServicios": "otrosServicios"
        }
        df_t = clean_df_for_json(df_t)
        headers_by_file = {row["archivo_origen"]: row.to_dict() for _, row in df_t.iterrows()}
        
        users_by_file = {}
        if "Usuarios" in xls.sheet_names:
            df_u = clean_df_for_json(pd.read_excel(xls, sheet_name="Usuarios"))
            for _, row in df_u.iterrows():
                fname = row.get("archivo_origen")
                if fname not in users_by_file: users_by_file[fname] = []
                users_by_file[fname].append(row.to_dict())
                
        services_by_file = {}
        for s_name, j_key in service_map.items():
            if s_name in xls.sheet_names:
                df_s = clean_df_for_json(pd.read_excel(xls, sheet_name=s_name))
                for _, row in df_s.iterrows():
                    fname = row.get("archivo_origen")
                    if fname:
                        if fname not in services_by_file: services_by_file[fname] = {k: [] for k in service_map.values()}
                        services_by_file[fname][j_key].append(row.to_dict())

        count = 0
        progress_bar = st.progress(0, text="Desconsolidando...")
        total_files = len(headers_by_file)

        for idx, (fname, header) in enumerate(headers_by_file.items()):
            progress_bar.progress((idx + 1) / total_files, text=f"Generando {fname}")
            header_copy = header.copy()
            header_copy.pop("archivo_origen", None)
            final = header_copy
            users = []
            for u in users_by_file.get(fname, []):
                u_copy = u.copy()
                u_copy.pop("archivo_origen", None)
                u_cons = u_copy.get("consecutivo")
                u_copy["servicios"] = {k: [] for k in service_map.values()}
                if fname in services_by_file:
                    for s_key, items in services_by_file[fname].items():
                        for item in items:
                            if item.get("consecutivoUsuario") == u_cons:
                                i_clean = item.copy()
                                i_clean.pop("archivo_origen", None)
                                i_clean.pop("consecutivoUsuario", None)
                                u_copy["servicios"][s_key].append(i_clean)
                users.append(u_copy)
            final["usuarios"] = users
            with open(os.path.join(dest_folder, fname), 'w', encoding='utf-8') as f:
                json.dump(final, f, ensure_ascii=False, indent=4)
            count += 1
            
        progress_bar.empty()
        return True, f"Desconsolidados {count} archivos."
    except Exception as e:
        return False, str(e)

def worker_update_cups_masivo(folder_path, old_val, new_val):
    count_files = 0
    total_changes = 0
    errors = []
    
    if not os.path.isdir(folder_path):
        return 0, 0, ["Carpeta no válida"]

    progress_bar = st.progress(0, text="Iniciando actualización de CUPS...")
    
    files_to_process = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.lower().endswith('.json'):
                files_to_process.append(os.path.join(root, file))
    
    total = len(files_to_process)
    
    if total == 0:
        progress_bar.empty()
        return 0, 0, ["No se encontraron archivos .json en la carpeta o subcarpetas."]

    for i, file_path in enumerate(files_to_process):
        if i % 5 == 0: progress_bar.progress(min(i/total, 1.0), text=f"Procesando {i}/{total}")
        
        filename = os.path.basename(file_path)
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
        
            changes = recursive_update_cups(data, old_val, new_val)
        
            if changes > 0:
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(data, f, indent=4, ensure_ascii=False)
                count_files += 1
                total_changes += changes
        
        except Exception as e:
            errors.append(f"{filename}: {e}")
    
    progress_bar.progress(1.0, text="Finalizado.")
    time.sleep(0.5)
    progress_bar.empty()
    
    return count_files, total_changes, errors

def worker_update_key_masivo(folder_path, key_target, new_value):
    count_files = 0
    total_changes = 0
    errors = []
    
    if not os.path.isdir(folder_path):
        return 0, 0, ["Carpeta no válida"]

    progress_bar = st.progress(0, text=f"Iniciando actualización de {key_target}...")
    
    files_to_process = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.lower().endswith('.json'):
                files_to_process.append(os.path.join(root, file))
    
    total = len(files_to_process)
    
    if total == 0:
        progress_bar.empty()
        return 0, 0, ["No se encontraron archivos .json en la carpeta o subcarpetas."]

    for i, file_path in enumerate(files_to_process):
        if i % 5 == 0: progress_bar.progress(min(i/total, 1.0), text=f"Procesando {i}/{total}")
        
        filename = os.path.basename(file_path)
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
        
            changes = recursive_update_key(data, key_target, new_value)
        
            if changes > 0:
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(data, f, indent=4, ensure_ascii=False)
                count_files += 1
                total_changes += changes
        
        except Exception as e:
            errors.append(f"{filename}: {e}")
    
    progress_bar.progress(1.0, text="Finalizado.")
    time.sleep(0.5)
    progress_bar.empty()
    
    return count_files, total_changes, errors

def worker_update_notes_masivo(folder_path, target_text, new_note):
    count_files = 0
    total_changes = 0
    errors = []
    
    if not os.path.isdir(folder_path):
        return 0, 0, ["Carpeta no válida"]

    progress_bar = st.progress(0, text="Iniciando actualización de Notas...")
    
    files_to_process = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.lower().endswith('.json'):
                files_to_process.append(os.path.join(root, file))
    
    total = len(files_to_process)
    
    if total == 0:
        progress_bar.empty()
        return 0, 0, ["No se encontraron archivos .json en la carpeta o subcarpetas."]

    for i, file_path in enumerate(files_to_process):
        if i % 5 == 0: progress_bar.progress(min(i/total, 1.0), text=f"Procesando {i}/{total}")
        
        filename = os.path.basename(file_path)
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
        
            changes = recursive_update_notes(data, target_text, new_note)
        
            if changes > 0:
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(data, f, indent=4, ensure_ascii=False)
                count_files += 1
                total_changes += changes
        
        except Exception as e:
            errors.append(f"{filename}: {e}")
    
    progress_bar.progress(1.0, text="Finalizado.")
    time.sleep(0.5)
    progress_bar.empty()
    
    return count_files, total_changes, errors

def worker_flat_to_excel(uploaded_files):
    # 1. Identificar archivos por prefijo (US, AC, AP, etc.)
    files_map = {}
    prefixes = ["US", "AF", "AC", "AP", "AU", "AH", "AN", "AM", "AT", "CT"]
    
    for f in uploaded_files:
        name_upper = os.path.basename(f.name).upper()
        # Buscar si el nombre empieza con alguno de los prefijos
        # Ejemplo: US123.txt -> US
        found_p = None
        for p in prefixes:
            if name_upper.startswith(p):
                found_p = p
                break
        
        if found_p:
            files_map[found_p] = f
            # Resetear puntero del archivo por si acaso
            f.seek(0)
    
    if "US" not in files_map:
        return None, "❌ Error: Falta el archivo de Usuarios (US) para generar los consecutivos."

    # 2. Leer US y generar consecutivo
    try:
        # Leer US. Asumimos CSV sin cabecera, encoding latin-1 (común en RIPS)
        df_us = pd.read_csv(files_map["US"], header=None, dtype=str, encoding='latin-1')
        
        # Generar columna 'consecutivo' (1, 2, 3...)
        # Se inserta al final o inicio? El usuario lo necesita para relacionar.
        # Lo añadimos como columna nueva.
        df_us['consecutivo'] = range(1, len(df_us) + 1)
        
        # Crear mapa: (TipoDoc, NumDoc) -> consecutivo
        # Indices US: 0=TipoDoc, 1=NumDoc
        # Limpiar espacios en claves
        df_us[0] = df_us[0].str.strip()
        df_us[1] = df_us[1].str.strip()
        
        user_map = dict(zip(zip(df_us[0], df_us[1]), df_us['consecutivo']))
        
    except Exception as e:
        return None, f"❌ Error procesando archivo US: {e}. Verifique que sea un CSV/TXT válido."

    # 3. Procesar otros archivos y generar Excel
    output = io.BytesIO()
    
    sheet_names = {
        "US": "Usuarios",
        "AF": "Transaccion",
        "AC": "Consultas",
        "AP": "Procedimientos",
        "AU": "Urgencias",
        "AH": "Hospitalizacion",
        "AN": "RecienNacidos",
        "AM": "Medicamentos",
        "AT": "OtrosServicios",
        "CT": "Control"
    }
    
    # Indices de (TipoDoc, NumDoc) en otros archivos para cruzar con US
    # Estándar RIPS: 
    # AF: 2, 3 (CodPres, Razon, Tipo, Num)
    # AC, AP, AU, AH, AM, AT: 2, 3 (NumFac, CodPres, Tipo, Num)
    # AN: 2, 3 (NumFac, CodPres, TipoMadre, NumMadre) -> Cruza con Madre en US
    doc_idx_map = {
        "AF": (2, 3), 
        "AC": (2, 3), "AP": (2, 3), "AU": (2, 3), "AH": (2, 3),
        "AM": (2, 3), "AT": (2, 3), "AN": (2, 3)
    }

    try:
        # Usamos ExcelWriter. Si xlsxwriter no está, pandas usa openpyxl por defecto.
        with pd.ExcelWriter(output) as writer:
            # Escribir Usuarios
            df_us.to_excel(writer, sheet_name='Usuarios', index=False)
            
            for p, f in files_map.items():
                if p == "US": continue
                
                try:
                    df = pd.read_csv(f, header=None, dtype=str, encoding='latin-1')
                except:
                    # Si falla, intentar header=0 o diferente encoding?
                    # Por ahora reportamos error en una hoja
                    pd.DataFrame([f"Error leyendo {f.name}"]).to_excel(writer, sheet_name=f"Error_{p}")
                    continue
                
                # Intentar cruzar consecutivo
                if p in doc_idx_map and len(df.columns) > 3:
                    idx_t, idx_n = doc_idx_map[p]
                    
                    # Asegurar columnas limpias para cruce
                    # No modificamos original para no perder datos, usamos temporal
                    t_types = df[idx_t].str.strip()
                    t_nums = df[idx_n].str.strip()
                    
                    # Función lookup optimizada
                    def get_cons(t, n):
                        return user_map.get((t, n), "")
                    
                    # Aplicar map
                    # Vectorizado es difícil con tuplas, usamos apply o list comprehension
                    # List comprehension es más rápido que apply
                    cons_list = [user_map.get((t, n), "") for t, n in zip(t_types, t_nums)]
                    
                    df['consecutivoUsuario'] = cons_list
                
                # Escribir hoja
                sheet = sheet_names.get(p, p)
                df.to_excel(writer, sheet_name=sheet, index=False)
                
    except Exception as e:
        return None, f"❌ Error generando Excel: {e}"

    output.seek(0)
    return output, None

# --- RENDER ---

def render(container=None):
    if container is None:
        container = st.container()
        
    with container:
        st.header("💊 RIPS")
        st.info("Módulo completo para gestión y conversión de archivos RIPS (JSON/Excel).")
        
        tab_ops = st.tabs(["Convertidor", "Cambio de CUPS", "Notas de Ajuste", "Validación", "Cambio Tecnología", "Planos a Excel"])
        
        with tab_ops[5]:
            st.subheader("Convertidor Planos a Excel (Soporte Nueva Resolución)")
            st.markdown("""
            Sube tus archivos planos actuales (**US, AC, AP, AF, AH, AU, AN, AM, AT**) para generar un archivo Excel consolidado.
            
            **Características:**
            - Genera automáticamente la hoja **Usuarios** con el campo `consecutivo`.
            - Relaciona los demás archivos mediante `consecutivoUsuario`.
            - Organiza los datos en hojas separadas (Consultas, Procedimientos, etc.) listas para la conversión a JSON.
            """)
            
            is_native = st.session_state.get("force_native_mode", True)
            
            if is_native:
                path_flat = render_path_selector("Carpeta con Archivos Planos", key="path_rips_flat")
                
                if st.button("🔄 Convertir a Excel", key="btn_convert_flat_native"):
                    if path_flat and os.path.isdir(path_flat):
                        with st.spinner("Procesando archivos en carpeta..."):
                             files_to_close = []
                             try:
                                 files_list = []
                                 # List files
                                 for fname in os.listdir(path_flat):
                                     if fname.lower().endswith(('.txt', '.csv')):
                                         full_path = os.path.join(path_flat, fname)
                                         f_obj = open(full_path, 'rb')
                                         files_list.append(f_obj)
                                         files_to_close.append(f_obj)
                                 
                                 if not files_list:
                                     st.warning("No se encontraron archivos TXT/CSV en la carpeta.")
                                 else:
                                     excel_data, error_msg = worker_flat_to_excel(files_list)
                                     
                                     if excel_data:
                                         st.success("✅ Conversión completada exitosamente.")
                                         st.download_button(
                                             label="📥 Descargar Excel Consolidado",
                                             data=excel_data,
                                             file_name="RIPS_Consolidado_NuevaRes.xlsx",
                                             mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                         )
                                     else:
                                         st.error(error_msg)
                             except Exception as e:
                                 st.error(f"Error accediendo a archivos: {e}")
                             finally:
                                 for f in files_to_close:
                                     try: f.close()
                                     except: pass
                    else:
                        st.warning("Seleccione una carpeta válida.")
            else:
                uploaded_files = st.file_uploader("Seleccionar archivos planos (TXT/CSV):", accept_multiple_files=True, type=["txt", "csv"], key="rips_flat_files_org")
                
                if uploaded_files:
                    if st.button("🔄 Convertir a Excel", key="btn_convert_flat_org"):
                        with st.spinner("Procesando archivos y generando relaciones..."):
                            excel_data, error_msg = worker_flat_to_excel(uploaded_files)
                            
                            if excel_data:
                                st.success("✅ Conversión completada exitosamente.")
                                st.download_button(
                                    label="📥 Descargar Excel Consolidado",
                                    data=excel_data,
                                    file_name="RIPS_Consolidado_NuevaRes.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
                            else:
                                st.error(error_msg)

        with tab_ops[0]:
            st.subheader("Convertidor RIPS")
            
            with st.expander("JSON a XLSX (Individual)", expanded=False):
                uploaded_json = st.file_uploader("Subir JSON", type=["json"], key="rips_json_ind")
                if uploaded_json and st.button("Convertir a Excel", key="btn_json_xlsx"):
                    xlsx_data, err = worker_json_a_xlsx_ind(uploaded_json)
                    if xlsx_data:
                        st.success("Conversión exitosa.")
                        st.download_button("Descargar Excel", xlsx_data, 
                                           file_name=f"{os.path.splitext(uploaded_json.name)[0]}.xlsx",
                                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    else:
                        st.error(f"Error: {err}")

            with st.expander("XLSX a JSON (Individual)", expanded=False):
                uploaded_xlsx = st.file_uploader("Subir Excel (RIPS)", type=["xlsx"], key="rips_xlsx_ind")
                if uploaded_xlsx and st.button("Convertir a JSON", key="btn_xlsx_json"):
                    json_data, err = worker_xlsx_a_json_ind(uploaded_xlsx)
                    if json_data:
                        st.success("Conversión exitosa.")
                        st.download_button("Descargar JSON", json_data, 
                                           file_name=f"{os.path.splitext(uploaded_xlsx.name)[0]}.json",
                                           mime="application/json")
                    else:
                        st.error(f"Error: {err}")

            with st.expander("JSON Evento a XLSX (Masivo - Consolidar)", expanded=False):
                st.markdown("Consolida múltiples archivos JSON de una carpeta en un único Excel.")
                path_consol = render_path_selector("Carpeta con JSONs", key="path_rips_consol")
                if st.button("Consolidar", key="btn_consol_rips"):
                    if path_consol:
                        xlsx_data, msg = worker_consolidar_json_xlsx(path_consol)
                        if xlsx_data:
                            st.success(msg)
                            st.download_button("Descargar Consolidado", xlsx_data, 
                                               file_name="RIPS_Consolidado.xlsx",
                                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                        else:
                            st.error(msg)
                    else:
                        st.warning("Ingrese una ruta.")

            with st.expander("XLSX Evento a JSONs (Masivo - Desconsolidar)", expanded=False):
                st.markdown("Genera múltiples archivos JSON a partir de un Excel consolidado (requiere hoja Transaccion con 'archivo_origen').")
                uploaded_consol = st.file_uploader("Subir Excel Consolidado", type=["xlsx"], key="rips_xlsx_consol")
                path_desconsol = render_path_selector("Carpeta de Salida", key="path_rips_desconsol")
                if st.button("Desconsolidar", key="btn_desconsol_rips"):
                    if uploaded_consol and path_desconsol:
                        if not os.path.exists(path_desconsol):
                            try:
                                os.makedirs(path_desconsol)
                            except:
                                st.error("No se pudo crear la carpeta de salida.")
                        
                        if os.path.exists(path_desconsol):
                            ok, msg = worker_desconsolidar_xlsx_json(uploaded_consol, path_desconsol)
                            if ok: st.success(msg)
                            else: st.error(msg)
                    else:
                        st.warning("Faltan datos.")

        with tab_ops[1]:
            st.subheader("Cambio de CUPS Masivo")
            path_cups = render_path_selector("Carpeta Raíz (busca recursivamente JSONs)", key="path_cups_mass")
            col1, col2 = st.columns(2)
            with col1:
                old_cup = st.text_input("CUP Anterior (codServicio)", key="cup_old")
            with col2:
                new_cup = st.text_input("CUP Nuevo", key="cup_new")
            
            if st.button("Actualizar CUPS", key="btn_update_cups"):
                if path_cups and old_cup and new_cup:
                    count, changes, errors = worker_update_cups_masivo(path_cups, old_cup, new_cup)
                    st.success(f"Proceso finalizado. Archivos modificados: {count}. Total cambios: {changes}.")
                    if errors:
                        st.error(f"Errores en {len(errors)} archivos.")
                        st.expander("Ver Errores").write(errors)
                else:
                    st.warning("Complete todos los campos.")
            
        with tab_ops[2]:
            st.subheader("Notas de Ajuste (Actualización Masiva)")
            st.markdown("Actualiza recursivamente notas o textos en todos los archivos JSON de una carpeta.")
            
            path_notes = render_path_selector("Carpeta Raíz (busca recursivamente JSONs)", key="path_notes_mass")
            col1, col2 = st.columns(2)
            with col1:
                target_note = st.text_input("Texto a buscar (Coincidencia Parcial)", key="note_target")
            with col2:
                new_note = st.text_input("Nueva nota (Reemplazo completo)", key="note_new")
            
            if st.button("Actualizar Notas", key="btn_update_notes"):
                if path_notes and target_note:
                    # new_note puede ser vacío si se quiere borrar/limpiar
                    count, changes, errors = worker_update_notes_masivo(path_notes, target_note, new_note)
                    st.success(f"Proceso finalizado. Archivos modificados: {count}. Total cambios: {changes}.")
                    if errors:
                        st.error(f"Errores en {len(errors)} archivos.")
                        st.expander("Ver Errores").write(errors)
                else:
                    st.warning("Debe ingresar al menos la carpeta y el texto a buscar.")
            
        with tab_ops[3]:
            st.subheader("Validación y Limpieza")
            st.markdown("Elimina espacios en blanco innecesarios en claves y valores de todos los archivos JSON en una carpeta.")
            
            path_clean = render_path_selector("Carpeta a Limpiar", key="path_rips_clean")
            if st.button("Limpiar JSONs (Espacios)", key="btn_clean_json"):
                if path_clean:
                    count, errs = worker_limpiar_json_rips(path_clean)
                    st.success(f"Proceso finalizado. {count} archivos limpiados.")
                    if errs:
                        st.error(f"Errores en {len(errs)} archivos.")
                        st.expander("Ver Errores").write(errs)
                else:
                    st.warning("Seleccione una carpeta.")

            st.divider()
            st.info("Para validación FEVRIPS completa (Estructura, Reglas, CUV), use la pestaña dedicada 'Validación FEVRIPS'.")

        with tab_ops[4]:
            st.subheader("Cambio de Tecnología (finalidadTecnologiaSalud)")
            
            mode_tech = st.radio("Modo:", ["Individual (Archivo)", "Masivo (Carpeta)"], horizontal=True)
            
            if mode_tech == "Individual (Archivo)":
                st.info("Sube un archivo JSON para actualizar el valor de 'finalidadTecnologiaSalud'.")
                f_json = st.file_uploader("Sube JSON:", type="json", key="tech_json_ind")
                new_val = st.text_input("Nuevo Valor:", value="44", key="tech_val_ind")
                
                if f_json and new_val and st.button("🚀 Actualizar Archivo", key="btn_tech_ind"):
                    try:
                        content = json.load(f_json)
                        count_changes = recursive_update_key(content, "finalidadTecnologiaSalud", new_val)
                        
                        if count_changes > 0:
                            json_str = json.dumps(content, indent=4, ensure_ascii=False)
                            st.download_button("📥 Descargar JSON Actualizado", data=json_str, file_name=f"update_{f_json.name}", mime="application/json")
                            st.success(f"✅ Se actualizaron {count_changes} campos.")
                        else:
                            st.warning("⚠️ No se encontró el campo 'finalidadTecnologiaSalud' en el archivo.")
                    except Exception as e:
                        st.error(f"Error: {e}")
                        
            else: # Masivo
                st.info("Actualiza 'finalidadTecnologiaSalud' en todos los JSON de una carpeta (recursivo).")
                path_tech = render_path_selector("Carpeta Raíz:", key="path_tech_mass")
                new_val_mass = st.text_input("Nuevo Valor:", value="44", key="tech_val_mass")
                
                if st.button("🚀 Actualizar Masivamente", key="btn_tech_mass"):
                     if path_tech and new_val_mass:
                        count, changes, errors = worker_update_key_masivo(path_tech, "finalidadTecnologiaSalud", new_val_mass)
                        st.success(f"Proceso finalizado. Archivos modificados: {count}. Total cambios: {changes}.")
                        if errors:
                            st.error(f"Errores en {len(errors)} archivos.")
                            st.expander("Ver Errores").write(errors)
                     else:
                        st.warning("Ingrese carpeta y valor.")
