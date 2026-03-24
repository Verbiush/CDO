import streamlit as st
import os
import time
import shutil
import fitz  # PyMuPDF
from PIL import Image
import tempfile
import pandas as pd
import os as _os_env

try:
    from gui_utils import abrir_dialogo_carpeta_nativo, update_path_key, render_path_selector, render_download_button
except ImportError:
    try:
        from src.gui_utils import abrir_dialogo_carpeta_nativo, update_path_key, render_path_selector, render_download_button
    except ImportError:
        abrir_dialogo_carpeta_nativo = None
        def update_path_key(key, new_path, widget_key=None):
            if new_path:
                st.session_state[key] = new_path
                if widget_key:
                    st.session_state[widget_key] = new_path
        
        def render_path_selector(label, key, default_path=None, help_text=None, omit_checkbox=False):
            st.warning("render_path_selector no disponible")
            return default_path
            
        def render_download_button(folder_path, key, label="📦 Descargar ZIP"):
            pass

try:
    from pdf2docx import Converter
except ImportError:
    Converter = None

def _valid_exts_for(conv_type_code):
    v = []
    if conv_type_code.startswith("PDF"): v = [".pdf"]
    elif conv_type_code.startswith("JPG"): v = [".jpg", ".jpeg"]
    elif conv_type_code.startswith("DOCX"): v = [".docx"]
    elif conv_type_code.startswith("PNG"): v = [".png"]
    elif conv_type_code == "XLS2XLSX": v = [".xls"]
    elif conv_type_code.startswith("XLSX"): v = [".xlsx", ".xls"]
    elif conv_type_code == "PDF_GRAY": v = [".pdf"]
    return v

try:
    from docx2pdf import convert as convert_docx_to_pdf
    HAS_DOCX2PDF = True
except ImportError:
    HAS_DOCX2PDF = False

# --- WORKERS ---

def _pdf_a_docx(input_path, output_path):
    if Converter is None:
        raise ImportError("Librería pdf2docx no instalada.")
    cv = Converter(input_path)
    cv.convert(output_path, start=0, end=None)
    cv.close()

def _jpg_a_pdf(input_path, output_path):
    image = Image.open(input_path)
    pdf_bytes = image.convert('RGB')
    pdf_bytes.save(output_path)

def _docx_a_pdf(input_path, output_path):
    if HAS_DOCX2PDF:
        # docx2pdf requires absolute paths
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except ImportError:
            pass
            
        convert_docx_to_pdf(os.path.abspath(input_path), os.path.abspath(output_path))
    else:
        raise ImportError("Librería docx2pdf no disponible o no compatible (requiere MS Word en Windows).")

def _pdf_a_jpg(input_path, output_folder, base_name):
    doc = fitz.open(input_path)
    saved_files = []
    for i, page in enumerate(doc):
        pix = page.get_pixmap()
        out_name = f"{base_name}_page{i+1}.jpg"
        out_path = os.path.join(output_folder, out_name)
        pix.save(out_path)
        saved_files.append(out_path)
    doc.close()
    return saved_files

def _png_a_jpg(input_path, output_path):
    img = Image.open(input_path)
    rgb_img = img.convert('RGB')
    rgb_img.save(output_path, 'jpeg')

def _txt_a_json(input_path, output_path):
    # Simplemente cambia la extensión, lógica del original
    if input_path == output_path: return
    if not os.path.exists(output_path):
        os.rename(input_path, output_path)
    else:
        base, ext = os.path.splitext(output_path)
        new_out = f"{base}_{int(time.time())}.json"
        os.rename(input_path, new_out)

def _xlsx_a_txt(input_path, output_path, sep=','):
    try:
        # keep_default_na=False para preservar "NA" como texto y no como NaN
        df = pd.read_excel(input_path, keep_default_na=False)
        
        # header=False para eliminar la primera fila (los encabezados) del archivo resultante
        df.to_csv(output_path, sep=sep, index=False, header=False)
    except Exception as e:
        raise Exception(f"Error convirtiendo Excel a TXT: {e}")

def _xls_a_xlsx(input_path, output_path):
    try:
        # Verificar si es HTML camuflado como XLS (lógica del referente)
        with open(input_path, 'r', encoding='utf-8', errors='ignore') as f:
            contenido = f.read()
        
        if contenido.strip().lower().startswith('<!doctype html') or '<table' in contenido.lower():
            # Es una tabla HTML
            dfs = pd.read_html(input_path)
            if not dfs:
                raise Exception("No se encontraron tablas en el archivo HTML/XLS.")
            df = dfs[0]
        else:
            # Es un archivo XLS real (requiere xlrd)
            df = pd.read_excel(input_path, engine='xlrd')
            
        df.to_excel(output_path, index=False)
    except Exception as e:
        raise Exception(f"Error convirtiendo XLS a XLSX: {e}")

def _pdf_escala_grises(input_path, output_path):
    doc = fitz.open(input_path)
    doc_final = fitz.open()
    
    dpi = st.session_state.app_config.get("pdf_dpi", 300) if "app_config" in st.session_state else 300
    matrix_scale = dpi / 72.0
    mat = fitz.Matrix(matrix_scale, matrix_scale)
    
    for page in doc:
        pix = page.get_pixmap(matrix=mat, colorspace=fitz.csGRAY)
        new_page = doc_final.new_page(width=page.rect.width, height=page.rect.height)
        new_page.insert_image(new_page.rect, pixmap=pix)
    doc.close()
    
    compression = st.session_state.app_config.get("pdf_compression", 4) if "app_config" in st.session_state else 4
    doc_final.save(output_path, garbage=compression, deflate=True)
    doc_final.close()

def worker_convertir_archivo(file_path, tipo, output_folder=None, sep=',', force_local=False):
    is_native = st.session_state.get("force_native_mode", True)
    if _os_env.environ.get("CDO_AGENT_MODE") == "1" or force_local:
        is_native = False
    
    if is_native:
        try:
            from src.agent_client import send_command, wait_for_result
            username = st.session_state.get("username", "default")
            
            # Send task to agent
            task_id = send_command(username, "convert_file", {
                "file_path": file_path,
                "type": tipo,
                "output_folder": output_folder,
                "sep": sep
            })
            
            if task_id:
                # Wait for result
                res = wait_for_result(task_id, timeout=300)
                
                if isinstance(res, dict) and "error" in res:
                     return False, res["error"]
                # New shape from agent: {"ok": bool, "message": "..." }
                if isinstance(res, dict) and "ok" in res and "message" in res:
                    return res["ok"], res["message"]
                # Backward compatibility if agent returns list [ok, msg]
                if isinstance(res, list) and len(res) >= 2:
                    return res[0], res[1]
                return False, f"Respuesta inesperada del agente: {res}"
            else:
                return False, "No se pudo conectar con el Agente."
        except Exception as e:
            return False, f"Error Agente: {str(e)}"

    if not file_path or (not is_native and not os.path.exists(file_path)):
        return False, "Archivo no encontrado"
        
    folder = output_folder if output_folder else os.path.dirname(file_path)
    filename = os.path.basename(file_path)
    name_no_ext = os.path.splitext(filename)[0]
    
    try:
        if tipo == "PDF2DOCX":
            out = os.path.join(folder, f"{name_no_ext}.docx")
            _pdf_a_docx(file_path, out)
        elif tipo == "JPG2PDF":
            out = os.path.join(folder, f"{name_no_ext}.pdf")
            _jpg_a_pdf(file_path, out)
        elif tipo == "DOCX2PDF":
            out = os.path.join(folder, f"{name_no_ext}.pdf")
            _docx_a_pdf(file_path, out)
        elif tipo == "PDF2JPG":
            # PDF to JPG creates multiple files usually, handle output
            _pdf_a_jpg(file_path, folder, name_no_ext)
        elif tipo == "PNG2JPG":
            out = os.path.join(folder, f"{name_no_ext}.jpg")
            _png_a_jpg(file_path, out)
        elif tipo == "TXT2JSON":
            out = os.path.join(folder, f"{name_no_ext}.json")
            _txt_a_json(file_path, out)
        elif tipo == "XLSX2TXT":
            out = os.path.join(folder, f"{name_no_ext}.txt")
            _xlsx_a_txt(file_path, out, sep=sep)
        elif tipo == "XLS2XLSX":
            out = os.path.join(folder, f"{name_no_ext}.xlsx")
            _xls_a_xlsx(file_path, out)
        elif tipo == "PDF_GRAY":
            temp_out = os.path.join(folder, f"{name_no_ext}_temp_gray.pdf")
            _pdf_escala_grises(file_path, temp_out)
            # Reemplazo seguro que funciona entre unidades de disco
            if os.path.exists(temp_out):
                try:
                    # Intentar mover (shutil maneja cross-filesystem)
                    shutil.move(temp_out, file_path)
                except Exception:
                    # Si falla (ej. archivo destino existe y no se puede sobrescribir directamente en Windows con move)
                    try:
                        if os.path.exists(file_path):
                            os.remove(file_path)
                        shutil.move(temp_out, file_path)
                    except Exception as e:
                        # Último intento: copiar y borrar
                        shutil.copy2(temp_out, file_path)
                        os.remove(temp_out)
            
        return True, "Conversión exitosa"
    except Exception as e:
        return False, str(e)

def worker_convertir_masivo(folder_path, tipo, output_folder=None, sep=',', return_zip=False):
    is_native = st.session_state.get("force_native_mode", True)
    if _os_env.environ.get("CDO_AGENT_MODE") == "1":
        is_native = False
    _bare = _os_env.environ.get("CDO_AGENT_MODE") == "1"
    
    if is_native:
        try:
            from src.agent_client import send_command, wait_for_result
            username = st.session_state.get("username", "default")
            
            task_id = send_command(username, "convert_bulk", {
                "folder_path": folder_path,
                "type": tipo,
                "output_folder": output_folder,
                "sep": sep
            })
            
            if task_id:
                res = wait_for_result(task_id, timeout=600)
                
                if isinstance(res, dict):
                    if "error" in res and not isinstance(res["error"], list): 
                        # Agent error (e.g. timeout) returns {"error": "msg"}
                        msg = res["error"]
                        if return_zip:
                            return {"count": 0, "message": msg, "error": True}
                        return 0, msg
                        
                    count = res.get("count", 0)
                    msg = res.get("message", "Proceso finalizado")
                    errors = res.get("errors", [])
                    
                    if errors:
                        msg += f" (con {len(errors)} errores)"
                    
                    if return_zip:
                         # In native mode we don't return files to browser, just status
                         return {
                            "count": count,
                            "message": msg,
                            "files": [] 
                         }
                    return count, msg
                else:
                    err_msg = f"Respuesta inválida: {res}"
                    if return_zip:
                         return {"count": 0, "message": err_msg, "error": True}
                    return 0, err_msg
                    
        except Exception as e:
            err_msg = f"Error Agente: {str(e)}"
            if return_zip:
                 return {"count": 0, "message": err_msg, "error": True}
            return 0, err_msg

    if not folder_path or (not is_native and not os.path.exists(folder_path)):
        if return_zip:
             return {"count": 0, "message": "Carpeta no encontrada", "error": True}
        return 0, "Carpeta no encontrada"
    
    count = 0
    files_to_process = []
    
    # Búsqueda recursiva
    try:
        # Si es modo nativo y la ruta no existe localmente para python (ej: Cloud), os.walk fallará
        # Pero intentamos por si acaso (usuario local)
        if os.path.exists(folder_path):
            for r, d, f in os.walk(folder_path):
                for file in f:
                    files_to_process.append(os.path.join(r, file))
        elif is_native:
             return 0, "Modo Nativo: No se puede acceder a los archivos directamente desde la aplicación. (Requiere Agente)"
    except Exception as e:
        return 0, f"Error accediendo a carpeta: {str(e)}"

    total = len(files_to_process)
    if total == 0:
        if return_zip:
             return {"count": 0, "message": "La carpeta está vacía.", "error": True}
        return 0, "La carpeta está vacía (no se encontraron archivos)."

    progress_bar = None if _bare else st.progress(0, text="Convirtiendo...")
    
    for i, full_path in enumerate(files_to_process):
        if progress_bar and i % 5 == 0:
            progress_bar.progress(min(i/total, 1.0), text=f"Procesando {i}/{total}")
        
        f = os.path.basename(full_path)
        f_lower = f.lower()
        process = False
        
        if tipo == "PDF2DOCX" and f_lower.endswith(".pdf"): process = True
        elif tipo == "JPG2PDF" and (f_lower.endswith(".jpg") or f_lower.endswith(".jpeg")): process = True
        elif tipo == "DOCX2PDF" and f_lower.endswith(".docx") and not f.startswith("~$"): process = True
        elif tipo == "PDF2JPG" and f_lower.endswith(".pdf"): process = True
        elif tipo == "PNG2JPG" and f_lower.endswith(".png"): process = True
        elif tipo == "TXT2JSON" and f_lower.endswith(".txt"): process = True
        elif tipo == "XLSX2TXT" and (f_lower.endswith(".xlsx") or f_lower.endswith(".xls")): process = True
        elif tipo == "XLS2XLSX" and f_lower.endswith(".xls"): process = True
        elif tipo == "PDF_GRAY" and f_lower.endswith(".pdf"): process = True
        
        if process:
            # Pass output_folder to individual worker
            ok, msg = worker_convertir_archivo(full_path, tipo, output_folder=output_folder, sep=sep)
            if ok: count += 1
            else: print(f"Error convirtiendo {f}: {msg}")
            
    if progress_bar:
        progress_bar.progress(1.0, text="Finalizado.")
    
    if return_zip:
        import io
        import zipfile
        
        if count == 0:
            return {"count": 0, "message": "No se procesaron archivos.", "error": True}
            
        mem_zip = io.BytesIO()
        with zipfile.ZipFile(mem_zip, "w", zipfile.ZIP_DEFLATED) as zf:
            # Add files from output_folder to zip
            for root, dirs, files in os.walk(output_folder):
                for file in files:
                    abs_path = os.path.join(root, file)
                    rel_path = os.path.relpath(abs_path, output_folder)
                    zf.write(abs_path, rel_path)
        mem_zip.seek(0)
        
        return {
            "count": count,
            "message": f"Procesados {count} archivos.",
            "files": [{
                "name": f"Conversion_Masiva_{int(time.time())}.zip",
                "data": mem_zip.getvalue(),
                "label": "📦 Descargar Archivos Convertidos (ZIP)"
            }]
        }
    
    if count == 0:
        return 0, "No se encontraron archivos compatibles para el tipo de conversión seleccionado."
    return count, f"Procesados {count} archivos."

# --- RENDER ---

def render(container=None):
    if container is None:
        container = st.container()
        
    with container:
        st.header("🔄 Centro de Conversiones")
        
        # Mapeo de nombres amigables a códigos internos
        conversion_options = {
            "PDF → DOCX": "PDF2DOCX",
            "JPG → PDF": "JPG2PDF",
            "DOCX → PDF": "DOCX2PDF",
            "PDF → JPG": "PDF2JPG",
            "PNG → JPG": "PNG2JPG",
            "TXT → JSON": "TXT2JSON",
            "Excel → TXT": "XLSX2TXT",
            "XLS → XLSX": "XLS2XLSX",
            "PDF → PDF (Grises)": "PDF_GRAY"
        }

        tab_ind, tab_mass = st.tabs(["📄 Conversión Individual", "📂 Conversión Masiva"])
        
        # --- TAB 1: INDIVIDUAL ---
        with tab_ind:
            st.markdown("### Configuración de Conversión")
            
            # 1. Tipo de Conversión
            st.write("**Tipo de Conversión:**")
            conv_type_label = st.selectbox("Tipo de Conversión", list(conversion_options.keys()), label_visibility="collapsed", key="ind_conv_type")
            conv_type_code = conversion_options[conv_type_label]

            # Separator selection for XLSX2TXT
            sep = ','
            if conv_type_code == "XLSX2TXT":
                st.write("**Delimitador:**")
                sep_label = st.selectbox("Seleccione delimitador", ["Coma (,)", "Punto y coma (;)", "Pipe (|)"], key="ind_sep")
                sep_map = {"Coma (,)": ",", "Punto y coma (;)": ";", "Pipe (|)": "|"}
                sep = sep_map[sep_label]

            # 2. Subir archivo (Único modo)
            uploaded_file = st.file_uploader(f"Subir Archivo para {conv_type_label}", key="ind_uploader")
            file_to_process = None
            is_uploaded = False
            temp_dir_obj = None
            if uploaded_file:
                is_uploaded = True
                file_to_process = uploaded_file
                valid_exts = _valid_exts_for(conv_type_code)
                ext = os.path.splitext(uploaded_file.name)[1].lower()
                if valid_exts and ext not in valid_exts:
                    st.error("El archivo subido no corresponde al tipo de conversión seleccionado.")
                    file_to_process = None

            # 3. Resultado (se muestra botón de descarga tras convertir)
            st.markdown("### 3. Resultado")


            # 4. Action
            st.write("")
            if st.button("Ejecutar Conversión", key="btn_exec_ind"):
                if not file_to_process:
                    st.error("⚠️ Seleccione una carpeta de destino.")
                else:
                    # Prepare paths
                    timestamp = int(time.time())
                    actual_output_folder = os.path.join(os.getcwd(), "temp_downloads", f"ind_{timestamp}")
                    os.makedirs(actual_output_folder, exist_ok=True)

                    try:
                        # Handle Uploaded File
                        if is_uploaded:
                            # Save in actual_output_folder so it's shared between frontend and backend containers
                            temp_path = os.path.join(actual_output_folder, file_to_process.name)
                            with open(temp_path, "wb") as f:
                                f.write(file_to_process.getbuffer())
                            actual_input_path = temp_path
                        else:
                            actual_input_path = file_to_process

                        # Execute
                        with st.spinner("Procesando..."):
                            ok, msg = worker_convertir_archivo(actual_input_path, conv_type_code, actual_output_folder, sep=sep, force_local=is_uploaded)
                        
                        if ok:
                            st.info("Archivo convertido listo para descargar.")

                            
                            # --- Download Logic ---
                            filename = os.path.basename(actual_input_path)
                            name_no_ext = os.path.splitext(filename)[0]
                            
                            if conv_type_code == "PDF2DOCX": out_file_name = f"{name_no_ext}.docx"
                            elif conv_type_code == "JPG2PDF": out_file_name = f"{name_no_ext}.pdf"
                            elif conv_type_code == "DOCX2PDF": out_file_name = f"{name_no_ext}.pdf"
                            elif conv_type_code == "PNG2JPG": out_file_name = f"{name_no_ext}.jpg"
                            elif conv_type_code == "TXT2JSON": out_file_name = f"{name_no_ext}.json"
                            elif conv_type_code == "XLSX2TXT": out_file_name = f"{name_no_ext}.txt"
                            elif conv_type_code == "XLS2XLSX": out_file_name = f"{name_no_ext}.xlsx"
                            elif conv_type_code == "PDF_GRAY": out_file_name = filename # Overwrites or same name
                            
                            if out_file_name:
                                # For PDF_GRAY with uploaded file, the file is modified in place at actual_input_path
                                # For others, it is at actual_output_folder/out_file_name
                                
                                target_file_path = os.path.join(actual_output_folder, out_file_name)
                                
                                # Special case for PDF_GRAY which overwrites input
                                if conv_type_code == "PDF_GRAY":
                                    target_file_path = actual_input_path
                                
                                if os.path.exists(target_file_path):
                                    mime = "application/octet-stream"
                                    if target_file_path.lower().endswith(".pdf"): mime = "application/pdf"
                                    elif target_file_path.lower().endswith(".docx"): mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                                    elif target_file_path.lower().endswith(".jpg") or target_file_path.lower().endswith(".jpeg"): mime = "image/jpeg"
                                    elif target_file_path.lower().endswith(".json"): mime = "application/json"
                                    elif target_file_path.lower().endswith(".txt"): mime = "text/plain"
                                    with open(target_file_path, "rb") as f:
                                        st.download_button(
                                            "📥 Descargar archivo convertido",
                                            data=f.read(),
                                            file_name=os.path.basename(target_file_path),
                                            mime=mime,
                                            key=f"dl_ind_file_{int(time.time())}"
                                        )
                            
                            elif conv_type_code == "PDF2JPG":
                                # Find generated JPGs
                                import glob
                                pattern = os.path.join(actual_output_folder, f"{name_no_ext}_page*.jpg")
                                jpgs = glob.glob(pattern)
                                if jpgs:
                                    # Create a temp zip file
                                    try:
                                        import zipfile
                                        timestamp = int(time.time())
                                        temp_dl_dir = "temp_downloads"
                                        os.makedirs(temp_dl_dir, exist_ok=True)
                                        zip_path = os.path.join(temp_dl_dir, f"{name_no_ext}_images_{timestamp}.zip")
                                        
                                        with zipfile.ZipFile(zip_path, "w") as zf:
                                            for jpg in jpgs:
                                                zf.write(jpg, os.path.basename(jpg))
                                        
                                        with open(zip_path, "rb") as f:
                                            st.download_button(
                                                "📦 Descargar Imágenes (ZIP)",
                                                data=f.read(),
                                                file_name=os.path.basename(zip_path),
                                                mime="application/zip",
                                                key="dl_ind_jpgs"
                                            )
                                        
                                        # No borramos la carpeta temporal aún para permitir descargas repetidas
                                    except Exception as e:
                                        st.error(f"Error preparando ZIP: {e}")
                            # ----------------------

                        else:
                            st.error(f"❌ Error: {msg}")

                    except Exception as e:
                        st.error(f"Error inesperado: {e}")
                    finally:
                        if temp_dir_obj:
                            temp_dir_obj.cleanup()

        # --- TAB 2: MASSIVE ---
        with tab_mass:
            st.markdown("### Configuración de Conversión Masiva")
            
            st.write("**Tipo de Conversión:**")
            conv_type_mass_label = st.selectbox("Tipo de Conversión", list(conversion_options.keys()), label_visibility="collapsed", key="mass_conv_type")
            conv_type_mass = conversion_options[conv_type_mass_label]

            sep_mass = ','
            if conv_type_mass == "XLSX2TXT":
                st.write("**Delimitador:**")
                sep_label_mass = st.selectbox("Seleccione delimitador", ["Coma (,)", "Punto y coma (;)", "Pipe (|)"], key="mass_sep")
                sep_map = {"Coma (,)": ",", "Punto y coma (;)": ";", "Pipe (|)": "|"}
                sep_mass = sep_map[sep_label_mass]

            # Source
            source_path = render_path_selector(
                label="Carpeta Origen (Masivo)",
                key="conv_mass_source",
                default_path=st.session_state.get("current_path")
            )

            # Output
            st.markdown("### Carpeta de Salida")
            is_native = st.session_state.get("force_native_mode", True)
            
            if is_native:
                out_path = render_path_selector(
                    label="Carpeta Destino",
                    key="conv_mass_out",
                    default_path=st.session_state.get("current_path")
                )
            else:
                # Web Mode: Use temp folder
                timestamp = int(time.time())
                out_path = os.path.join(os.getcwd(), "temp_downloads", f"mass_{timestamp}")
                os.makedirs(out_path, exist_ok=True)
                st.info(f"📂 Procesando en entorno temporal: {out_path}")

            # Execute
            st.write("")
            if st.button("🚀 Ejecutar Conversión Masiva", key="btn_exec_mass"):
                if not out_path:
                    st.error("⚠️ Seleccione una carpeta de salida.")
                elif source_path and (is_native or os.path.exists(source_path)):
                    count, msg = worker_convertir_masivo(source_path, conv_type_mass, output_folder=out_path, sep=sep_mass)
                    if count > 0:
                        st.success(msg)
#                         render_download_button(out_path, "dl_mass_conv", "📦 Descargar Archivos Convertidos (ZIP)", cleanup=not is_native)
                    else:
                        st.warning(msg)
                else:
                    st.error("La carpeta objetivo no es válida.")
