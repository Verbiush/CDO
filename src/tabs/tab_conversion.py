import streamlit as st
import os
import time
import shutil
import fitz  # PyMuPDF
from PIL import Image
import tempfile
import pandas as pd

try:
    from gui_utils import seleccionar_carpeta_nativa
except ImportError:
    def seleccionar_carpeta_nativa(title="Seleccionar Carpeta", initial_dir=None, key=None):
        return st.text_input(f"{title}", value=initial_dir if initial_dir else "", key=key or f"fallback_{title}")

try:
    from pdf2docx import Converter
except ImportError:
    Converter = None

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

def worker_convertir_archivo(file_path, tipo, output_folder=None, sep=','):
    if not file_path or not os.path.exists(file_path):
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
        elif tipo == "PDF_GRAY":
            temp_out = os.path.join(folder, f"{name_no_ext}_temp_gray.pdf")
            _pdf_escala_grises(file_path, temp_out)
            # Replace original logic
            if os.path.exists(temp_out):
                try:
                    os.replace(temp_out, file_path)
                except OSError:
                    time.sleep(0.5)
                    if os.path.exists(file_path): os.remove(file_path)
                    os.rename(temp_out, file_path)
            
        return True, "Conversión exitosa"
    except Exception as e:
        return False, str(e)

def worker_convertir_masivo(folder_path, tipo, output_folder=None, sep=','):
    if not folder_path or not os.path.exists(folder_path):
        return 0, "Carpeta no encontrada"
    
    count = 0
    files_to_process = []
    
    # Búsqueda recursiva
    for r, d, f in os.walk(folder_path):
        for file in f:
            files_to_process.append(os.path.join(r, file))

    total = len(files_to_process)
    if total == 0:
        return 0, "La carpeta está vacía (no se encontraron archivos)."

    progress_bar = st.progress(0, text="Convirtiendo...")
    
    for i, full_path in enumerate(files_to_process):
        if i % 5 == 0: progress_bar.progress(min(i/total, 1.0), text=f"Procesando {i}/{total}")
        
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
        elif tipo == "PDF_GRAY" and f_lower.endswith(".pdf"): process = True
        
        if process:
            # Pass output_folder to individual worker
            ok, msg = worker_convertir_archivo(full_path, tipo, output_folder=output_folder, sep=sep)
            if ok: count += 1
            else: print(f"Error convirtiendo {f}: {msg}")
            
    progress_bar.progress(1.0, text="Finalizado.")
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

            # 2. Origen
            st.write("**Origen:**")
            origin_mode = st.radio("Origen", ["Subir", "Carpeta Actual"], label_visibility="collapsed", horizontal=True, key="ind_origin_mode")

            file_to_process = None
            is_uploaded = False
            temp_dir_obj = None

            if origin_mode == "Subir":
                uploaded_file = st.file_uploader(f"Subir Archivo para {conv_type_label}", key="ind_uploader")
                if uploaded_file:
                    is_uploaded = True
                    # Create temp file immediately to have a path if needed, or handle in button
                    # We will handle in button to avoid creating temp files unnecessarily
                    file_to_process = uploaded_file
            else:
                # Carpeta Actual
                current_path = st.session_state.get("current_path", os.getcwd())
                st.info(f"📁 Buscando archivos en: {current_path}")
                
                # Filter files based on conversion type
                valid_exts = []
                if conv_type_code.startswith("PDF"): valid_exts = [".pdf"]
                elif conv_type_code.startswith("JPG"): valid_exts = [".jpg", ".jpeg"]
                elif conv_type_code.startswith("DOCX"): valid_exts = [".docx"]
                elif conv_type_code.startswith("PNG"): valid_exts = [".png"]
                elif conv_type_code.startswith("TXT"): valid_exts = [".txt"]
                elif conv_type_code.startswith("XLSX"): valid_exts = [".xlsx", ".xls"]

                try:
                    files = [f for f in os.listdir(current_path) if any(f.lower().endswith(ext) for ext in valid_exts)]
                except Exception:
                    files = []
                
                if files:
                    selected_file = st.selectbox("Seleccionar Archivo:", files, key="ind_file_select")
                    file_to_process = os.path.join(current_path, selected_file)
                else:
                    st.warning(f"No se encontraron archivos compatibles con {conv_type_label} en la carpeta actual.")

            # 3. Guardar en
            st.write("**Ubicación de Salida:**")
            ind_save_option = st.radio("Ubicación de Salida", ["Carpeta Principal (Por defecto)", "Otra Carpeta"], label_visibility="collapsed", key="ind_save_mode")
            
            default_save = st.session_state.get("current_path", os.getcwd())
            if ind_save_option == "Otra Carpeta":
                output_folder = seleccionar_carpeta_nativa("Seleccionar Carpeta:", initial_dir=default_save, key="conv_ind_out")
            else:
                output_folder = default_save
                st.info(f"Se guardará en: {output_folder}")

            # 4. Action
            st.write("")
            if st.button("Ejecutar Conversión", key="btn_exec_ind"):
                if not file_to_process:
                    st.error("⚠️ Seleccione o suba un archivo para procesar.")
                else:
                    # Prepare paths
                    actual_input_path = ""
                    actual_output_folder = output_folder if output_folder else default_save
                    
                    if not os.path.exists(actual_output_folder):
                        try:
                            os.makedirs(actual_output_folder)
                        except:
                            st.error(f"No se pudo crear carpeta destino: {actual_output_folder}")
                            st.stop()

                    try:
                        # Handle Uploaded File
                        if is_uploaded:
                            # Create a temp dir that persists only for this block
                            temp_dir_obj = tempfile.TemporaryDirectory()
                            temp_path = os.path.join(temp_dir_obj.name, file_to_process.name)
                            with open(temp_path, "wb") as f:
                                f.write(file_to_process.getbuffer())
                            actual_input_path = temp_path
                        else:
                            actual_input_path = file_to_process

                        # Execute
                        with st.spinner("Procesando..."):
                            ok, msg = worker_convertir_archivo(actual_input_path, conv_type_code, actual_output_folder, sep=sep)
                        
                        if ok:
                            st.success(f"✅ {msg}")
                            if is_uploaded:
                                st.info(f"Archivo guardado en: {actual_output_folder}")
                        else:
                            st.error(f"❌ Error: {msg}")

                    except Exception as e:
                        st.error(f"Error inesperado: {e}")
                    finally:
                        if temp_dir_obj:
                            temp_dir_obj.cleanup()

        # --- TAB 2: MASIVA ---
        with tab_mass:
            st.markdown("### Configuración de Conversión Masiva")
            
            # 1. Carpeta Objetivo
            target_folder = seleccionar_carpeta_nativa("Carpeta Objetivo:", initial_dir=st.session_state.get("current_path", os.getcwd()), key="conv_mass_target")
            
            # 2. Tipo de Conversión Masiva
            st.write("**Tipo de Conversión Masiva:**")
            mass_conv_type_label = st.selectbox("Tipo de Conversión Masiva", list(conversion_options.keys()), label_visibility="collapsed", key="mass_conv_type")
            mass_conv_type_code = conversion_options[mass_conv_type_label]

            # Separator selection for XLSX2TXT
            mass_sep = ','
            if mass_conv_type_code == "XLSX2TXT":
                st.write("**Delimitador:**")
                mass_sep_label = st.selectbox("Seleccione delimitador", ["Coma (,)", "Punto y coma (;)", "Pipe (|)"], key="mass_sep")
                sep_map = {"Coma (,)": ",", "Punto y coma (;)": ";", "Pipe (|)": "|"}
                mass_sep = sep_map[mass_sep_label]

            # 3. Guardar en
            st.write("**Ubicación de Salida:**")
            mass_save_option = st.radio("Ubicación de Salida", ["Carpeta de Origen (Por defecto)", "Otra Carpeta"], label_visibility="collapsed", key="mass_save_mode")
            
            target_output = None
            if mass_save_option == "Otra Carpeta":
                default_save_mass = st.session_state.get("current_path", os.getcwd())
                target_output = seleccionar_carpeta_nativa("Seleccionar Carpeta:", initial_dir=default_save_mass, key="conv_mass_out")
            else:
                st.info("Los archivos convertidos se guardarán en la misma carpeta que los originales.")

            # 4. Execute
            st.write("")
            if st.button("🚀 Ejecutar Conversión Masiva", key="btn_exec_mass"):
                if target_folder and os.path.exists(target_folder):
                    # Si target_output está vacío, usamos None para que worker_convertir_archivo use carpeta origen
                    final_output = target_output if target_output.strip() else None
                    count, msg = worker_convertir_masivo(target_folder, mass_conv_type_code, output_folder=final_output, sep=mass_sep)
                    if count > 0:
                        st.success(msg)
                    else:
                        st.warning(msg)
                else:
                    st.error("La carpeta objetivo no es válida.")
