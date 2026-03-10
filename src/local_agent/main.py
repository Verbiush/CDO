import time
from datetime import datetime
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import json
import os
import platform
import sys
import threading
import queue
import tkinter as tk
from tkinter import simpledialog, messagebox, scrolledtext, ttk
from getpass import getpass
import re
import base64
from io import BytesIO
import pandas as pd
import io
import shutil
import zipfile

# --- CONVERSION LIBRARIES ---
try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

try:
    from PIL import Image
except ImportError:
    Image = None

try:
    from pdf2docx import Converter
except ImportError:
    Converter = None

try:
    from docx2pdf import convert as convert_docx_to_pdf
    HAS_DOCX2PDF = True
except ImportError:
    HAS_DOCX2PDF = False
    convert_docx_to_pdf = None


# Try to import send2trash
try:
    from send2trash import send2trash
except ImportError:
    send2trash = None

# Try to import docx, but don't fail if not present
try:
    from docx import Document
    from docx.text.paragraph import Paragraph
    from docx.oxml.ns import qn
except ImportError:
    Document = None
    Paragraph = None
    qn = None

# Add parent directory to path to allow imports from src (if running from source)
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)
if parent_dir not in sys.path:
    sys.path.insert(0, parent_dir)

CONFIG_FILE = "agent_config.json"

# --- FILE PROCESSING LOGIC ---

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

def process_update_cups(folder_path, old_val, new_val):
    count_files = 0
    total_changes = 0
    errors = []
    
    if not os.path.isdir(folder_path):
        return {"status": "error", "message": "Carpeta no válida"}

    files_to_process = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.lower().endswith('.json'):
                files_to_process.append(os.path.join(root, file))
    
    for file_path in files_to_process:
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
            errors.append(f"{os.path.basename(file_path)}: {str(e)}")
            
    return {
        "count_files": count_files,
        "total_changes": total_changes,
        "errors": errors
    }

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

def process_update_key(folder_path, key_target, new_value):
    count_files = 0
    total_changes = 0
    errors = []
    
    if not os.path.isdir(folder_path):
        return {"count_files": 0, "total_changes": 0, "errors": ["Carpeta no válida"]}

    files_to_process = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.lower().endswith('.json'):
                files_to_process.append(os.path.join(root, file))
    
    for file_path in files_to_process:
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
            errors.append(f"{os.path.basename(file_path)}: {str(e)}")
            
    return {
        "count_files": count_files,
        "total_changes": total_changes,
        "errors": errors
    }

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

def process_update_notes(folder_path, target_text, new_note):
    count_files = 0
    total_changes = 0
    errors = []
    
    if not os.path.isdir(folder_path):
        return {"count_files": 0, "total_changes": 0, "errors": ["Carpeta no válida"]}

    files_to_process = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.lower().endswith('.json'):
                files_to_process.append(os.path.join(root, file))
    
    for file_path in files_to_process:
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
            errors.append(f"{os.path.basename(file_path)}: {str(e)}")
            
    return {
        "count_files": count_files,
        "total_changes": total_changes,
        "errors": errors
    }

def recursive_strip(data):
    if isinstance(data, dict):
        return {k.strip(): recursive_strip(v) for k, v in data.items()}
    elif isinstance(data, list):
        return [recursive_strip(i) for i in data]
    elif isinstance(data, str):
        return data.strip()
    return data

def process_clean_json(folder_path):
    count_files = 0
    errors = []
    
    if not os.path.isdir(folder_path):
        return {"count_files": 0, "errors": ["Carpeta no válida"]}

    files_to_process = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.lower().endswith('.json'):
                files_to_process.append(os.path.join(root, file))
    
    for file_path in files_to_process:
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
        
            cleaned_data = recursive_strip(data)
            
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(cleaned_data, f, indent=4, ensure_ascii=False)
            
            count_files += 1
        except Exception as e:
            errors.append(f"{os.path.basename(file_path)}: {str(e)}")
            
    return {
        "count_files": count_files,
        "errors": errors
    }

def replace_text_in_element(paragraph, mapping):
    if not paragraph.text.strip():
        return 0
    
    text_has_braces = "{" in paragraph.text
    text_has_chevrons = "<<" in paragraph.text or "«" in paragraph.text
    
    if not (text_has_braces or text_has_chevrons):
        return 0

    original_text = paragraph.text
    current_text = original_text
    count = 0
    
    for key, val in mapping.items():
        key_clean = key.replace(" ", "")
        p1 = r"\{\s*" + re.escape(key) + r"\s*\}"
        p2 = r"\{\s*" + re.escape(key_clean) + r"\s*\}"
        p3 = r"\{\s*" + re.escape(key.replace(" ", "_")) + r"\s*\}"
        
        c1 = r"(?:«|<<)\s*" + re.escape(key) + r"\s*(?:»|>>)"
        c2 = r"(?:«|<<)\s*" + re.escape(key_clean) + r"\s*(?:»|>>)"
        c3 = r"(?:«|<<)\s*" + re.escape(key.replace(" ", "_")) + r"\s*(?:»|>>)"
        
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

def iter_all_paragraphs(doc_obj):
    # 1. Main Body Paragraphs
    yield from doc_obj.paragraphs
    
    # 2. Tables in Body
    for table in doc_obj.tables:
        for row in table.rows:
            for cell in row.cells:
                yield from cell.paragraphs
    
    # 3. Text Boxes in Body (via XML)
    if doc_obj.element.body is not None:
        for txbx in doc_obj.element.body.iter(qn('w:txbxContent')):
            for p_element in txbx.iter(qn('w:p')):
                yield Paragraph(p_element, doc_obj)

    # 4. Headers and Footers
    for section in doc_obj.sections:
        headers = [section.header, section.first_page_header, section.even_page_header]
        footers = [section.footer, section.first_page_footer, section.even_page_footer]
        
        for header in headers:
            if header and not header.is_linked_to_previous:
                yield from header.paragraphs
                for table in header.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            yield from cell.paragraphs
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
                for txbx in footer._element.iter(qn('w:txbxContent')):
                    for p_element in txbx.iter(qn('w:p')):
                        yield Paragraph(p_element, footer)

# --- CONVERSION HELPERS ---

def _pdf_a_docx(input_path, output_path):
    if Converter is None:
        raise ImportError("Librería pdf2docx no instalada.")
    cv = Converter(input_path)
    cv.convert(output_path, start=0, end=None)
    cv.close()

def _jpg_a_pdf(input_path, output_path):
    if Image is None:
        raise ImportError("Librería Pillow (PIL) no instalada.")
    image = Image.open(input_path)
    pdf_bytes = image.convert('RGB')
    pdf_bytes.save(output_path)

def _docx_a_pdf(input_path, output_path):
    if HAS_DOCX2PDF:
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except ImportError:
            pass
        convert_docx_to_pdf(os.path.abspath(input_path), os.path.abspath(output_path))
    else:
        raise ImportError("Librería docx2pdf no disponible (requiere MS Word en Windows).")

def _pdf_a_jpg(input_path, output_folder, base_name):
    if fitz is None:
        raise ImportError("Librería PyMuPDF (fitz) no instalada.")
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
    if Image is None:
        raise ImportError("Librería Pillow (PIL) no instalada.")
    img = Image.open(input_path)
    rgb_img = img.convert('RGB')
    rgb_img.save(output_path, 'jpeg')

def _txt_a_json(input_path, output_path):
    if input_path == output_path: return
    if not os.path.exists(output_path):
        os.rename(input_path, output_path)
    else:
        base, ext = os.path.splitext(output_path)
        new_out = f"{base}_{int(time.time())}.json"
        os.rename(input_path, new_out)

def _xlsx_a_txt(input_path, output_path, sep=','):
    try:
        df = pd.read_excel(input_path, keep_default_na=False)
        df.to_csv(output_path, sep=sep, index=False, header=False)
    except Exception as e:
        raise Exception(f"Error convirtiendo Excel a TXT: {e}")

def _xls_a_xlsx(input_path, output_path):
    try:
        with open(input_path, 'r', encoding='utf-8', errors='ignore') as f:
            contenido = f.read()
        
        if contenido.strip().lower().startswith('<!doctype html') or '<table' in contenido.lower():
            dfs = pd.read_html(input_path)
            if not dfs:
                raise Exception("No se encontraron tablas en el archivo HTML/XLS.")
            df = dfs[0]
        else:
            df = pd.read_excel(input_path, engine='xlrd')
            
        df.to_excel(output_path, index=False)
    except Exception as e:
        raise Exception(f"Error convirtiendo XLS a XLSX: {e}")

def _pdf_escala_grises(input_path, output_path):
    if fitz is None:
        raise ImportError("Librería PyMuPDF (fitz) no instalada.")
    doc = fitz.open(input_path)
    doc_final = fitz.open()
    
    # Default DPI 300
    dpi = 300 
    matrix_scale = dpi / 72.0
    mat = fitz.Matrix(matrix_scale, matrix_scale)
    
    for page in doc:
        pix = page.get_pixmap(matrix=mat, colorspace=fitz.csGRAY)
        new_page = doc_final.new_page(width=page.rect.width, height=page.rect.height)
        new_page.insert_image(new_page.rect, pixmap=pix)
    doc.close()
    
    doc_final.save(output_path, garbage=4, deflate=True)
    doc_final.close()

def process_convert_file(file_path, tipo, output_folder=None, sep=','):
    if not os.path.exists(file_path):
        return {"success": False, "message": "Archivo no encontrado"}
        
    folder = output_folder if output_folder else os.path.dirname(file_path)
    if not os.path.exists(folder):
        os.makedirs(folder, exist_ok=True)
        
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
            # Logic to replace original if needed, but for agent we might just return the new file path
            # But tab logic tries to replace. Let's replicate replacement logic if output_folder is same as input
            if output_folder is None or os.path.abspath(output_folder) == os.path.abspath(os.path.dirname(file_path)):
                 if os.path.exists(temp_out):
                    try:
                        shutil.move(temp_out, file_path)
                    except:
                        if os.path.exists(file_path): os.remove(file_path)
                        shutil.move(temp_out, file_path)
            
        return {"success": True, "message": "Conversión exitosa"}
    except Exception as e:
        return {"success": False, "message": str(e)}

def process_convert_bulk(folder_path, tipo, output_folder=None, sep=','):
    if not os.path.exists(folder_path):
        return {"count": 0, "message": "Carpeta no encontrada"}
        
    files_to_process = []
    for r, d, f in os.walk(folder_path):
        for file in f:
            files_to_process.append(os.path.join(r, file))
            
    count = 0
    errors = []
    
    for full_path in files_to_process:
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
            res = process_convert_file(full_path, tipo, output_folder, sep)
            if res["success"]:
                count += 1
            else:
                errors.append(f"{f}: {res['message']}")
                
    return {"count": count, "message": f"Procesados {count} archivos", "errors": errors}

def process_fill_docx(base_path, tasks, template_b64=None):
    count_success = 0
    errors = []
    
    if Document is None:
        return {"count": 0, "errors": ["Librería python-docx no instalada en Agente"]}

    template_bytes = None
    if template_b64:
        try:
            template_bytes = base64.b64decode(template_b64)
        except Exception as e:
            return {"count": 0, "errors": [f"Error decodificando plantilla: {e}"]}

    for i, task in enumerate(tasks):
        rel_path = task.get("rel_path")
        data = task.get("data", {})
        
        if not rel_path:
            errors.append(f"Task {i}: Falta ruta relativa")
            continue
            
        full_path = os.path.join(base_path, rel_path)
        
        try:
            os.makedirs(full_path, exist_ok=True)
            
            doc_to_process = None
            doc_path = None
            
            # 1. Try distributed base doc
            base_doc_path = os.path.join(full_path, "documento_base.docx")
            if os.path.exists(base_doc_path):
                doc_to_process = Document(base_doc_path)
                doc_path = base_doc_path
            # 2. Try template bytes
            elif template_bytes:
                doc_to_process = Document(BytesIO(template_bytes))
            # 3. Fallback to local files
            else:
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
                # Perform replacements
                replacements = 0
                regex_replacements = task.get("regex_replacements", [])
                
                for p in iter_all_paragraphs(doc_to_process):
                    # 1. Map-based replacements
                    replacements += replace_text_in_element(p, data)
                    
                    # 2. Regex-based replacements
                    for pat, repl in regex_replacements:
                        if p.text.strip() and re.search(pat, p.text):
                            try:
                                new_text = re.sub(pat, repl, p.text)
                                if new_text != p.text:
                                    p.text = new_text
                                    replacements += 1
                            except Exception as e:
                                pass # Ignore regex errors

                # Save
                dest_doc = doc_path if doc_path else os.path.join(full_path, "documento_base.docx")
                doc_to_process.save(dest_doc)
                count_success += 1
            else:
                errors.append(f"No se encontró plantilla para: {rel_path}")
                
        except Exception as e:
            errors.append(f"Error procesando {rel_path}: {str(e)}")
            
    return {"count": count_success, "errors": errors}

def process_create_folders(folder_list):
    count = 0
    errors = []
    for folder_path in folder_list:
        try:
            os.makedirs(folder_path, exist_ok=True)
            count += 1
        except Exception as e:
            errors.append(f"Error creating {folder_path}: {e}")
    return {"count": count, "errors": errors}

def process_write_file(file_path, content_b64):
    try:
        import base64
        content = base64.b64decode(content_b64)
        os.makedirs(os.path.dirname(file_path), exist_ok=True)
        with open(file_path, "wb") as f:
            f.write(content)
        return {"success": True}
    except Exception as e:
        return {"success": False, "error": str(e)}

def process_distribute_file(file_paths, content_b64):
    try:
        import base64
        content = base64.b64decode(content_b64)
    except Exception as e:
        return {"count": 0, "errors": [f"Error decoding: {e}"]}
        
    count = 0
    errors = []
    
    for path in file_paths:
        try:
            os.makedirs(os.path.dirname(path), exist_ok=True)
            with open(path, "wb") as f:
                f.write(content)
            count += 1
        except Exception as e:
            errors.append(f"Error writing {path}: {e}")
            
    return {"count": count, "errors": errors}

def process_write_files(files_list):
    count = 0
    errors = []
    import base64
    
    for item in files_list:
        path = item.get("path")
        content_b64 = item.get("content_b64")
        if not path or not content_b64:
            continue
            
        try:
            content = base64.b64decode(content_b64)
            os.makedirs(os.path.dirname(path), exist_ok=True)
            with open(path, "wb") as f:
                f.write(content)
            count += 1
        except Exception as e:
            errors.append(f"Error writing {path}: {e}")
            
    return {"count": count, "errors": errors}

def process_bulk_rename(source_path, items, separator):
    count_renamed = 0
    errors = []
    
    if not os.path.isdir(source_path):
        return {"count": 0, "errors": ["Carpeta no válida"]}

    try:
        root_items = os.listdir(source_path)
    except Exception as e:
        return {"count": 0, "errors": [f"Error listando directorio: {str(e)}"]}
        
    root_folders = [d for d in root_items if os.path.isdir(os.path.join(source_path, d))]
    root_files = [f for f in root_items if os.path.isfile(os.path.join(source_path, f))]
    
    for item in items:
        match_key = str(item.get("key", "")).strip()
        suffix_val = str(item.get("suffix", "")).strip()
        
        if not match_key or not suffix_val:
            continue
            
        normalized_key = match_key.lower()
        
        # 1. Matching Folders
        matching_folders = []
        for d in root_folders:
            d_lower = d.lower()
            # Exact match
            if d_lower == normalized_key:
                matching_folders.append(d)
            # Consecutive match (Key_1, Key_2)
            elif d_lower.startswith(normalized_key + "_") and d[len(normalized_key)+1:].isdigit():
                matching_folders.append(d)
            # Consecutive match (Key (1), Key (2))
            elif d_lower.startswith(normalized_key + " (") and d.endswith(")"):
                 matching_folders.append(d)
        
        for folder_name in matching_folders:
            folder_path = os.path.join(source_path, folder_name)
            target_path = folder_path # Default to original if not renamed
            
            # Check if folder name already has suffix
            # We construct the expected suffix pattern
            suffix_pattern = f"{separator}{suffix_val}"
            
            if not folder_name.endswith(suffix_pattern):
                new_folder_name = f"{folder_name}{suffix_pattern}"
                new_folder_path = os.path.join(source_path, new_folder_name)
                
                try:
                    os.rename(folder_path, new_folder_path)
                    count_renamed += 1
                    target_path = new_folder_path # Update target for file scanning
                except Exception as e:
                    errors.append(f"Error renombrando carpeta {folder_name}: {str(e)}")
                    # If rename fails, we keep target_path as original, 
                    # but maybe we shouldn't process files if folder rename failed?
                    # Let's try to process files anyway in the original folder.
            
            # Rename internal files
            try:
                if os.path.isdir(target_path):
                    for filename in os.listdir(target_path):
                        file_full_path = os.path.join(target_path, filename)
                        if os.path.isfile(file_full_path):
                            base_name, ext = os.path.splitext(filename)
                            if not base_name.endswith(suffix_pattern):
                                new_name = f"{base_name}{suffix_pattern}{ext}"
                                try:
                                    os.rename(file_full_path, os.path.join(target_path, new_name))
                                    count_renamed += 1
                                except: pass
            except Exception as e:
                errors.append(f"Error procesando archivos internos de {folder_name}: {str(e)}")

        # 2. Matching Files in Root
        for filename in root_files:
            if match_key in filename:
                file_full_path = os.path.join(source_path, filename)
                base_name, ext = os.path.splitext(filename)
                
                suffix_pattern = f"{separator}{suffix_val}"
                if not base_name.endswith(suffix_pattern):
                    new_name = f"{base_name}{suffix_pattern}{ext}"
                    try:
                        os.rename(file_full_path, os.path.join(source_path, new_name))
                        count_renamed += 1
                    except Exception as e:
                        errors.append(f"Error renombrando archivo {filename}: {str(e)}")

    return {"count": count_renamed, "errors": errors}

def process_validate_rips(base_path, api_url, token=None, verify_ssl=True):
    results = []
    generated_files = []
    
    if not os.path.isdir(base_path):
        return {"status": "error", "message": "Ruta inválida"}

    try:
        files = [f for f in os.listdir(base_path) if f.lower().endswith('.json') 
                 and not f.startswith("Resultados") and not f.startswith("Resp_")]
    except Exception as e:
        return {"status": "error", "message": f"Error listando archivos: {str(e)}"}
        
    if not files:
        return {"status": "warning", "message": "No se encontraron archivos JSON"}

    headers = {"Content-Type": "application/json"}
    if token:
        headers["Authorization"] = f"Bearer {token}"
        
    if not verify_ssl:
        requests.packages.urllib3.disable_warnings()

    count_processed = 0
    
    for fname in files:
        full_path = os.path.join(base_path, fname)
        res_row = {"Archivo": fname, "Estado": "Pendiente", "CUV": "", "Mensaje": ""}
        
        try:
            with open(full_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                
            try:
                r = requests.post(api_url, json=data, headers=headers, verify=verify_ssl, timeout=60)
                res_row["Estado"] = str(r.status_code)
                
                if r.status_code == 200:
                    r_json = r.json()
                    res_row["CUV"] = r_json.get("cuv") or r_json.get("CUV") or ""
                    res_row["Mensaje"] = "Validado Correctamente"
                    
                    # Save results locally
                    factura_num = data.get('numFactura', os.path.splitext(fname)[0])
                    provider_id = data.get('numDocumentoIdentificacionObligado', '999')
                    
                    # 1. ResultadosLocales
                    f_loc_name = f"ResultadosLocales_{factura_num}.json"
                    f_loc_path = os.path.join(base_path, f_loc_name)
                    with open(f_loc_path, "w", encoding="utf-8") as f_out:
                        json.dump(r_json, f_out, indent=2, ensure_ascii=False)
                    generated_files.append(f_loc_name)
                    
                    # 2. ResultadosMSPS
                    f_msps_name = f"ResultadosMSPS_{factura_num}_ID{provider_id}_R.json"
                    f_msps_path = os.path.join(base_path, f_msps_name)
                    with open(f_msps_path, "w", encoding="utf-8") as f_out:
                        json.dump(r_json, f_out, indent=2, ensure_ascii=False)
                    generated_files.append(f_msps_name)
                else:
                    res_row["Mensaje"] = r.text[:200]
            except Exception as e:
                 res_row["Estado"] = "Error Conexión"
                 res_row["Mensaje"] = str(e)
                 
        except Exception as e:
            res_row["Estado"] = "Error Lectura"
            res_row["Mensaje"] = str(e)
            
        results.append(res_row)
        count_processed += 1
        
    return {
        "status": "success",
        "processed": count_processed,
        "results": results,
        "generated_files": generated_files
    }

def process_copy_files(file_list, target_folder):
    if not os.path.exists(target_folder):
        try:
            os.makedirs(target_folder)
        except Exception as e:
            return {"count": 0, "errors": [f"Error creando carpeta destino: {e}"]}

    count = 0
    errors = []
    
    for item in file_list:
        src_path = item.get("Ruta completa")
        if not src_path or not os.path.exists(src_path):
            continue
            
        try:
            filename = os.path.basename(src_path)
            dest_path = os.path.join(target_folder, filename)
            
            # Handle collisions
            if os.path.exists(dest_path):
                base, ext = os.path.splitext(filename)
                dest_path = os.path.join(target_folder, f"{base}_{int(time.time())}{ext}")
            
            if os.path.isdir(src_path):
                shutil.copytree(src_path, dest_path)
            else:
                shutil.copy2(src_path, dest_path)
            count += 1
        except Exception as e:
            errors.append(f"Error copiando {src_path}: {e}")
            
    return {"count": count, "errors": errors}

def process_move_files(file_list, target_folder):
    if not os.path.exists(target_folder):
        try:
            os.makedirs(target_folder)
        except Exception as e:
            return {"count": 0, "errors": [f"Error creando carpeta destino: {e}"]}

    count = 0
    errors = []
    
    for item in file_list:
        src_path = item.get("Ruta completa")
        if not src_path or not os.path.exists(src_path):
            continue
            
        try:
            filename = os.path.basename(src_path)
            dest_path = os.path.join(target_folder, filename)
            
            # Handle collisions
            if os.path.exists(dest_path):
                base, ext = os.path.splitext(filename)
                dest_path = os.path.join(target_folder, f"{base}_{int(time.time())}{ext}")
            
            shutil.move(src_path, dest_path)
            count += 1
        except Exception as e:
            errors.append(f"Error moviendo {src_path}: {e}")
            
    return {"count": count, "errors": errors}

def process_delete_files(file_list, force_delete=False):
    count_del = 0
    errors = []
    
    for item in file_list:
        path = item.get("Ruta completa")
        if not path or not os.path.exists(path):
            continue
            
        try:
            safe_path = os.path.normpath(path)
            if force_delete:
                if os.path.isdir(safe_path):
                    shutil.rmtree(safe_path)
                else:
                    os.remove(safe_path)
                count_del += 1
            else:
                if send2trash:
                    send2trash(safe_path)
                    count_del += 1
                else:
                    errors.append(f"Send2Trash no disponible para {path} (use force_delete)")
        except Exception as e:
            errors.append(f"Error eliminando {path}: {e}")
            
    return {"count": count_del, "errors": errors}

def process_zip_files(file_list, target_zip_path):
    count = 0
    errors = []
    
    try:
        # Create parent dir if not exists
        os.makedirs(os.path.dirname(target_zip_path), exist_ok=True)
        
        with zipfile.ZipFile(target_zip_path, 'w', compression=zipfile.ZIP_DEFLATED) as zipf:
            for item in file_list:
                src_path = item.get("Ruta completa")
                if not src_path or not os.path.exists(src_path): continue
                
                try:
                    if os.path.isdir(src_path):
                        for root, dirs, files in os.walk(src_path):
                            for file in files:
                                f_path = os.path.join(root, file)
                                arcname = os.path.relpath(f_path, os.path.dirname(src_path))
                                zipf.write(f_path, arcname)
                    else:
                        zipf.write(src_path, os.path.basename(src_path))
                    count += 1
                except Exception as e:
                    errors.append(f"Error zipeando {src_path}: {e}")
                    
        return {"count": count, "errors": errors, "zip_path": target_zip_path}
    except Exception as e:
        return {"count": 0, "errors": [f"Error creando ZIP: {e}"]}

def process_zip_folders_individually(folder_list, target_folder):
    if target_folder and not os.path.exists(target_folder):
        try:
            os.makedirs(target_folder, exist_ok=True)
        except Exception as e:
            return {"count": 0, "errors": [f"Error creando carpeta destino: {e}"]}

    count = 0
    errors = []
    
    for item in folder_list:
        src_path = item.get("Ruta completa")
        if not src_path or not os.path.isdir(src_path): continue
        
        try:
            if target_folder:
                zip_base = os.path.join(target_folder, os.path.basename(src_path))
            else:
                zip_base = src_path # create zip next to folder

            shutil.make_archive(zip_base, 'zip', src_path)
            count += 1
        except Exception as e:
            errors.append(f"Error zipeando carpeta {src_path}: {e}")
            
    return {"count": count, "errors": errors}

def process_edit_text(file_list, search_text, replace_text):
    count = 0
    errors = []
    
    for item in file_list:
        file_path = item.get("Ruta completa")
        if not file_path or not os.path.exists(file_path): continue
        
        try:
            ext = os.path.splitext(file_path)[1].lower()
            modified = False
            
            # Plain text files
            if ext in ['.txt', '.json', '.xml', '.csv', '.html', '.md', '.log', '.py', '.js', '.css', '.bat', '.ps1']:
                try:
                    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                        content = f.read()
                    
                    if search_text in content:
                        new_content = content.replace(search_text, replace_text)
                        with open(file_path, 'w', encoding='utf-8') as f:
                            f.write(new_content)
                        modified = True
                except Exception as e:
                    errors.append(f"{os.path.basename(file_path)}: {e}")

            # Word documents
            elif ext == '.docx':
                if Document:
                    try:
                        doc = Document(file_path)
                        doc_modified = False
                        for p in iter_all_paragraphs(doc):
                            if search_text in p.text:
                                p.text = p.text.replace(search_text, replace_text)
                                doc_modified = True
                                
                        if doc_modified:
                            doc.save(file_path)
                            modified = True
                    except Exception as e:
                         errors.append(f"DOCX {os.path.basename(file_path)}: {e}")
                else:
                    errors.append(f"DOCX {os.path.basename(file_path)} skipped (no python-docx)")
            
            if modified:
                count += 1
                
        except Exception as e:
            errors.append(f"General error {os.path.basename(file_path)}: {e}")
            
    return {"count": count, "errors": errors}

def process_rename_folders_by_mapping(source_path, mapping):
    count = 0
    errors = []
    
    if not os.path.isdir(source_path):
        return {"count": 0, "errors": ["Carpeta no válida"]}

    try:
        dirs = [d for d in os.listdir(source_path) if os.path.isdir(os.path.join(source_path, d))]
    except Exception as e:
        return {"count": 0, "errors": [f"Error listando: {e}"]}
        
    for dirname in dirs:
        matched_new_name = None
        
        # 1. Exact match in mapping keys
        if dirname in mapping:
            matched_new_name = mapping[dirname]
        else:
            # 2. Loose match: dirname contains key
            for curr_val, new_val in mapping.items():
                if curr_val in dirname:
                    matched_new_name = new_val
                    break
        
        if matched_new_name and matched_new_name != dirname:
            src_full = os.path.join(source_path, dirname)
            dst_full = os.path.join(source_path, matched_new_name)
            
            try:
                # Handle collision with consecutive suffix
                if os.path.exists(dst_full):
                    counter = 1
                    while True:
                        new_name_suffix = f"{matched_new_name}_{counter}"
                        dst_full_suffix = os.path.join(source_path, new_name_suffix)
                        if not os.path.exists(dst_full_suffix):
                            dst_full = dst_full_suffix
                            break
                        counter += 1
                
                os.rename(src_full, dst_full)
                count += 1
            except Exception as e:
                errors.append(f"Error renombrando {dirname}: {e}")
                
    return {"count": count, "errors": errors}

def process_organize_files(source_path, base_dest, mapping):
    count = 0
    errors = []
    
    if not os.path.isdir(source_path):
        return {"count": 0, "errors": ["Carpeta Origen no válida"]}
        
    try:
        os.makedirs(base_dest, exist_ok=True)
        items = os.listdir(source_path)
    except Exception as e:
        return {"count": 0, "errors": [f"Error acceso: {e}"]}
        
    for item_name in items:
        item_path = os.path.join(source_path, item_name)
        
        matched_dst_name = None
        for src_val, dst_val in mapping.items():
            if src_val in item_name:
                matched_dst_name = dst_val
                break
        
        if matched_dst_name:
            try:
                dest_folder_path = os.path.join(base_dest, matched_dst_name)
                os.makedirs(dest_folder_path, exist_ok=True)
                
                final_item_path = os.path.join(dest_folder_path, item_name)
                
                # Handle collision
                if os.path.exists(final_item_path):
                    base, ext = os.path.splitext(item_name)
                    final_item_path = os.path.join(dest_folder_path, f"{base}_{int(time.time())}{ext}")
                
                shutil.move(item_path, final_item_path)
                count += 1
            except Exception as e:
                errors.append(f"Error moviendo {item_name}: {e}")
                
    return {"count": count, "errors": errors}



# --- COMMAND HANDLERS IN WORKER ---

class AgentWorker:
    def __init__(self, username, task_url, result_url):
        self.username = username
        self.task_url = task_url
        self.result_url = result_url
        self.running = False
        self.thread = None
        self.log_callback = None

    def log(self, message):
        print(message)
        if self.log_callback:
            self.log_callback(message)

    def start(self):
        self.running = True
        self.thread = threading.Thread(target=self.run_loop)
        self.thread.daemon = True
        self.thread.start()
        self.log("Worker iniciado.")

    def stop(self):
        self.running = False
        if self.thread:
            self.thread.join(timeout=2)
        self.log("Worker detenido.")

    def run_loop(self):
        while self.running:
            try:
                # Poll for tasks
                resp = requests.get(f"{self.task_url}?username={self.username}", timeout=5)
                if resp.status_code == 200:
                    data = resp.json()
                    if data.get("tasks"):
                        for task in data["tasks"]:
                            self.process_task(task)
            except Exception as e:
                # self.log(f"Error polling tasks: {e}")
                pass
            
            time.sleep(2)

    def process_task(self, task):
        task_id = task.get("id")
        command = task.get("command")
        params = task.get("params", {})
        
        self.log(f"Procesando tarea {task_id}: {command}")
        
        result = {"status": "SUCCESS", "result": None}
        
        try:
            if command == "ping":
                result["result"] = "pong"
                
            elif command == "update_cups":
                folder = params.get("path")
                old_val = params.get("old_val")
                new_val = params.get("new_val")
                
                if folder and old_val and new_val:
                    res = process_update_cups(folder, old_val, new_val)
                    result["result"] = res
                else:
                    result["status"] = "ERROR"
                    result["result"] = {"error": "Faltan parámetros"}

            elif command == "update_key":
                folder = params.get("path")
                key_target = params.get("key")
                new_value = params.get("value")
                
                if folder and key_target and new_value:
                    res = process_update_key(folder, key_target, new_value)
                    result["result"] = res
                else:
                    result["status"] = "ERROR"
                    result["result"] = {"error": "Faltan parámetros"}

            elif command == "update_notes":
                path = params.get("path")
                target = params.get("target")
                note = params.get("note")
                if path:
                    self.log(f"Actualizando Notas en: {path}")
                    result["result"] = process_update_notes(path, target, note)
                else:
                    result["status"] = "ERROR"
                    result["result"] = {"error": "Falta path"}
                    
            elif command == "clean_json":
                path = params.get("path")
                if path:
                    self.log(f"Limpiando JSONs en: {path}")
                    result["result"] = process_clean_json(path)
                else:
                    result["status"] = "ERROR"
                    result["result"] = {"error": "Falta path"}
                    
            elif command == "consolidate_json":
                path = params.get("path")
                if path:
                    self.log(f"Consolidando JSONs en: {path}")
                    result["result"] = process_consolidate_json(path)
                else:
                    result["status"] = "ERROR"
                    result["result"] = {"error": "Falta path"}

            elif command == "desconsolidate_json":
                file_path = params.get("file_path")
                dest_folder = params.get("dest_folder")
                if file_path and dest_folder:
                    self.log(f"Desconsolidando {file_path} a {dest_folder}")
                    result["result"] = process_desconsolidate_json(file_path, dest_folder)
                else:
                    result["status"] = "ERROR"
                    result["result"] = {"error": "Faltan parámetros"}
            
            elif command == "flat_to_excel":
                path = params.get("path")
                if path:
                    self.log(f"Convirtiendo planos en: {path}")
                    result["result"] = process_flat_to_excel(path)
                else:
                    result["status"] = "ERROR"
                    result["result"] = {"error": "Falta path"}
            
            elif command == "bulk_rename":
                path = params.get("path")
                items = params.get("items", [])
                separator = params.get("separator", "_")
                
                if path:
                    res = process_bulk_rename(path, items, separator)
                    result["result"] = res
                else:
                    result["status"] = "ERROR"
                    result["result"] = {"error": "Falta path"}

            elif command == "fill_docx":
                base_path = params.get("base_path")
                tasks = params.get("tasks", [])
                template_b64 = params.get("template_b64")
                
                if not base_path or not os.path.exists(base_path):
                    result["status"] = "ERROR"
                    result["result"] = {"error": "Ruta base inválida o no encontrada"}
                else:
                    self.log(f"Iniciando llenado de DOCX en: {base_path} ({len(tasks)} tareas)")
                    res = process_fill_docx(base_path, tasks, template_b64)
                    result["result"] = res

            elif command == "validate_rips":
                base_path = params.get("base_path")
                api_url = params.get("api_url")
                token = params.get("token")
                verify_ssl = params.get("verify_ssl", True)
                
                if not base_path or not api_url:
                    result["status"] = "ERROR"
                    result["result"] = {"error": "Faltan parámetros (path o api_url)"}
                else:
                    self.log(f"Validando RIPS en: {base_path}")
                    res = process_validate_rips(base_path, api_url, token, verify_ssl)
                    result["result"] = res

            elif command == "copy_files":
                items = params.get("items", [])
                target = params.get("target_folder")
                if items and target:
                    self.log(f"Copiando {len(items)} items a {target}")
                    result["result"] = process_copy_files(items, target)
                else:
                    result["status"] = "ERROR"
                    result["result"] = {"error": "Faltan parámetros"}

            elif command == "move_files":
                items = params.get("items", [])
                target = params.get("target_folder")
                if items and target:
                    self.log(f"Moviendo {len(items)} items a {target}")
                    result["result"] = process_move_files(items, target)
                else:
                    result["status"] = "ERROR"
                    result["result"] = {"error": "Faltan parámetros"}

            elif command == "delete_files":
                items = params.get("items", [])
                force = params.get("force", False)
                if items:
                    self.log(f"Eliminando {len(items)} items (Force={force})")
                    result["result"] = process_delete_files(items, force)
                else:
                    result["status"] = "ERROR"
                    result["result"] = {"error": "Faltan parámetros"}

            elif command == "zip_files":
                items = params.get("items", [])
                target = params.get("target_path")
                if items and target:
                    self.log(f"Zipeando {len(items)} items a {target}")
                    result["result"] = process_zip_files(items, target)
                else:
                    result["status"] = "ERROR"
                    result["result"] = {"error": "Faltan parámetros"}
            
            elif command == "zip_folders_individual":
                items = params.get("items", [])
                target = params.get("target_folder")
                if items:
                    self.log(f"Zipeando individualmente {len(items)} carpetas")
                    result["result"] = process_zip_folders_individually(items, target)
                else:
                    result["status"] = "ERROR"
                    result["result"] = {"error": "Faltan parámetros"}

            elif command == "edit_text":
                items = params.get("items", [])
                find = params.get("find")
                replace = params.get("replace")
                if items and find:
                    self.log(f"Editando texto en {len(items)} archivos")
                    result["result"] = process_edit_text(items, find, replace)
                else:
                    result["status"] = "ERROR"
                    result["result"] = {"error": "Faltan parámetros"}

            elif command == "convert_file":
                file_path = params.get("file_path")
                tipo = params.get("type")
                output_folder = params.get("output_folder")
                sep = params.get("sep", ",")
                
                if file_path and tipo:
                    self.log(f"Convirtiendo archivo: {file_path} ({tipo})")
                    result["result"] = process_convert_file(file_path, tipo, output_folder, sep)
                else:
                    result["status"] = "ERROR"
                    result["result"] = {"error": "Faltan parámetros"}

            elif command == "rename_folders_mapped":
                path = params.get("path")
                mapping = params.get("mapping", {})
                if path:
                    self.log(f"Renombrando carpetas en {path} con {len(mapping)} reglas")
                    result["result"] = process_rename_folders_by_mapping(path, mapping)
                else:
                    result["status"] = "ERROR"
                    result["result"] = {"error": "Faltan parámetros"}
            
            elif command == "organize_files_mapped":
                source = params.get("source_path")
                dest = params.get("dest_path")
                mapping = params.get("mapping", {})
                if source and dest:
                    self.log(f"Organizando archivos de {source} a {dest}")
                    result["result"] = process_organize_files(source, dest, mapping)
                else:
                    result["status"] = "ERROR"
                    result["result"] = {"error": "Faltan parámetros"}

            elif command == "create_folders":
                folders = params.get("folders", [])
                if folders:
                    self.log(f"Creando {len(folders)} carpetas")
                    result["result"] = process_create_folders(folders)
                else:
                    result["status"] = "ERROR"
                    result["result"] = {"error": "Falta folders"}

            elif command == "write_file":
                file_path = params.get("file_path")
                content_b64 = params.get("content_b64")
                if file_path and content_b64:
                    self.log(f"Escribiendo archivo: {file_path}")
                    result["result"] = process_write_file(file_path, content_b64)
                else:
                    result["status"] = "ERROR"
                    result["result"] = {"error": "Faltan parámetros"}

            elif command == "distribute_file":
                paths = params.get("paths", [])
                content_b64 = params.get("content_b64")
                if paths and content_b64:
                    self.log(f"Distribuyendo archivo a {len(paths)} ubicaciones")
                    result["result"] = process_distribute_file(paths, content_b64)
                else:
                    result["status"] = "ERROR"
                    result["result"] = {"error": "Faltan parámetros"}

            elif command == "write_files":
                files_list = params.get("files", [])
                if files_list:
                    self.log(f"Escribiendo {len(files_list)} archivos")
                    result["result"] = process_write_files(files_list)
                else:
                    result["status"] = "ERROR"
                    result["result"] = {"error": "Faltan parámetros"}

            elif command == "convert_bulk":
                folder_path = params.get("folder_path")
                tipo = params.get("type")
                output_folder = params.get("output_folder")
                sep = params.get("sep", ",")
                
                if folder_path and tipo:
                    self.log(f"Conversión masiva en: {folder_path} ({tipo})")
                    result["result"] = process_convert_bulk(folder_path, tipo, output_folder, sep)
                else:
                    result["status"] = "ERROR"
                    result["result"] = {"error": "Faltan parámetros"}

            else:
                result["status"] = "ERROR"
                result["result"] = f"Comando desconocido: {command}"

        except Exception as e:
            result["status"] = "ERROR"
            result["result"] = str(e)
            self.log(f"Error procesando tarea: {e}")

        # Send result
        try:
            requests.post(self.result_url, json={
                "task_id": task_id,
                "status": result["status"],
                "result": result["result"]
            })
            self.log(f"Resultado enviado para {task_id}")
        except Exception as e:
            self.log(f"Error enviando resultado: {e}")

# --- GUI ---
def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r') as f:
                return json.load(f)
        except:
            pass
    return {}

def save_config(config):
    with open(CONFIG_FILE, 'w') as f:
        json.dump(config, f)

class AgentGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Agente Local - Organizador Archivos")
        self.root.geometry("500x400")
        
        self.config = load_config()
        self.worker = None
        
        # UI Elements
        frame = ttk.Frame(root, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text="Usuario:").grid(row=0, column=0, sticky=tk.W)
        self.user_var = tk.StringVar(value=self.config.get("username", ""))
        ttk.Entry(frame, textvariable=self.user_var).grid(row=0, column=1, sticky=tk.EW)
        
        ttk.Label(frame, text="URL Tareas:").grid(row=1, column=0, sticky=tk.W)
        self.url_task_var = tk.StringVar(value=self.config.get("task_url", "http://localhost:8501/api/tasks"))
        ttk.Entry(frame, textvariable=self.url_task_var).grid(row=1, column=1, sticky=tk.EW)
        
        ttk.Label(frame, text="URL Resultados:").grid(row=2, column=0, sticky=tk.W)
        self.url_res_var = tk.StringVar(value=self.config.get("result_url", "http://localhost:8501/api/results"))
        ttk.Entry(frame, textvariable=self.url_res_var).grid(row=2, column=1, sticky=tk.EW)
        
        self.btn_start = ttk.Button(frame, text="Iniciar Agente", command=self.toggle_agent)
        self.btn_start.grid(row=3, column=0, columnspan=2, pady=10)
        
        self.log_area = scrolledtext.ScrolledText(frame, height=15)
        self.log_area.grid(row=4, column=0, columnspan=2, sticky=tk.NSEW)
        
        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(4, weight=1)

    def log(self, msg):
        def _log():
            self.log_area.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")
            self.log_area.see(tk.END)
        self.root.after(0, _log)

    def toggle_agent(self):
        if self.worker and self.worker.running:
            self.worker.stop()
            self.btn_start.config(text="Iniciar Agente")
            self.log("Agente detenido manualmente.")
        else:
            user = self.user_var.get()
            t_url = self.url_task_var.get()
            r_url = self.url_res_var.get()
            
            if not user or not t_url:
                messagebox.showerror("Error", "Configure usuario y URLs")
                return
                
            self.config["username"] = user
            self.config["task_url"] = t_url
            self.config["result_url"] = r_url
            save_config(self.config)
            
            self.worker = AgentWorker(user, t_url, r_url)
            self.worker.log_callback = self.log
            self.worker.start()
            self.btn_start.config(text="Detener Agente")

if __name__ == "__main__":
    root = tk.Tk()
    app = AgentGUI(root)
    root.mainloop()
