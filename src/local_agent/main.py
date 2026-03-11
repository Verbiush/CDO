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
    yield from doc_obj.paragraphs
    for table in doc_obj.tables:
        for row in table.rows:
            for cell in row.cells:
                yield from cell.paragraphs
    
    if doc_obj.element.body is not None:
        for txbx in doc_obj.element.body.iter(qn('w:txbxContent')):
            for p_element in txbx.iter(qn('w:p')):
                yield Paragraph(p_element, doc_obj)

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
            if d.lower() == normalized_key or \
               (d.lower().startswith(normalized_key + "_") and d[len(normalized_key)+1:].isdigit()) or \
               (d.lower().startswith(normalized_key + " (") and d.endswith(")")):
               matching_folders.append(d)
        
        for folder_name in matching_folders:
            folder_path = os.path.join(source_path, folder_name)
            
            # Check if folder name already has suffix
            if not folder_name.endswith(f"{separator}{suffix_val}"):
                new_folder_name = f"{folder_name}{separator}{suffix_val}"
                new_folder_path = os.path.join(source_path, new_folder_name)
                
                try:
                    os.rename(folder_path, new_folder_path)
                    count_renamed += 1
                except Exception as e:
                    errors.append(f"Error renombrando carpeta {folder_name}: {str(e)}")
            
            # Rename internal files
            try:
                if os.path.exists(folder_path): # Might have been renamed
                    target_path = folder_path 
                elif os.path.exists(new_folder_path):
                    target_path = new_folder_path
                else:
                    continue
                    
                for filename in os.listdir(target_path):
                    file_full_path = os.path.join(target_path, filename)
                    if os.path.isfile(file_full_path):
                        base_name, ext = os.path.splitext(filename)
                        if not base_name.endswith(f"{separator}{suffix_val}"):
                            new_name = f"{base_name}{separator}{suffix_val}{ext}"
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
                
                if not base_name.endswith(f"{separator}{suffix_val}"):
                    new_name = f"{base_name}{separator}{suffix_val}{ext}"
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

class AgentWorker:
    def __init__(self, username, task_url, result_url, password=None):
        self.username = username
        self.task_url = task_url
        self.result_url = result_url
        self.password = password
        self.running = False
        self.thread = None
        self.log_callback = None
        self.gui_invoker = None

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
                auth = (self.username, self.password) if self.password else None
                resp = requests.get(f"{self.task_url}?username={self.username}", auth=auth, timeout=5)
                if resp.status_code == 200:
                    data = resp.json()
                    if data.get("tasks"):
                        for task in data["tasks"]:
                            self.process_task(task)
            except Exception as e:
                self.log(f"Error polling tasks: {e}")
                # pass
            
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

            elif command == "list_files":
                path = params.get("path")
                if path:
                    self.log(f"Listando archivos en: {path}")
                    try:
                        if os.path.exists(path) and os.path.isdir(path):
                            items = []
                            # Limit to 500 items to avoid payload issues
                            count = 0
                            with os.scandir(path) as it:
                                for entry in it:
                                    if count > 500: break
                                    items.append({
                                        "name": entry.name,
                                        "is_dir": entry.is_dir(),
                                        "size": entry.stat().st_size if not entry.is_dir() else 0,
                                        "mtime": entry.stat().st_mtime
                                    })
                                    count += 1
                            result["result"] = {"files": items, "count": count}
                        else:
                            result["status"] = "ERROR"
                            result["result"] = {"error": "Ruta no encontrada o no es carpeta"}
                    except Exception as e:
                        result["status"] = "ERROR"
                        result["result"] = {"error": str(e)}
                else:
                    result["status"] = "ERROR"
                    result["result"] = {"error": "Falta path"}

            elif command == "browse_folder":
                title = params.get("title", "Seleccionar Carpeta")
                self.log(f"Abriendo selector de carpeta: {title}")
                
                try:
                    import tkinter.filedialog as fd
                    
                    path = None
                    if self.gui_invoker:
                        # Use the GUI thread invoker if available
                        path = self.gui_invoker(fd.askdirectory, title=title)
                    else:
                        # Fallback (unsafe)
                        path = fd.askdirectory(title=title)
                        
                    if path:
                        result["result"] = {"path": path}
                    else:
                        result["result"] = {"cancelled": True}
                        
                except Exception as e:
                    result["status"] = "ERROR"
                    result["result"] = {"error": str(e)}

            elif command == "browse_file":
                title = params.get("title", "Seleccionar Archivo")
                file_types = params.get("file_types", [])
                self.log(f"Abriendo selector de archivo: {title}")
                
                try:
                    import tkinter.filedialog as fd
                    # Convert file_types from list of lists to list of tuples if needed
                    # Streamlit sends: [[label, pattern], ...]
                    # Tkinter expects: [(label, pattern), ...]
                    ft = []
                    if file_types:
                        for item in file_types:
                            if isinstance(item, list) and len(item) >= 2:
                                ft.append((item[0], item[1]))
                            elif isinstance(item, tuple):
                                ft.append(item)
                    
                    if not ft:
                        ft = [("Todos los archivos", "*.*")]
                        
                    path = None
                    if self.gui_invoker:
                        path = self.gui_invoker(fd.askopenfilename, title=title, filetypes=ft)
                    else:
                        path = fd.askopenfilename(title=title, filetypes=ft)

                    if path:
                        result["result"] = {"path": path}
                    else:
                        result["result"] = {"cancelled": True}
                except Exception as e:
                    result["status"] = "ERROR"
                    result["result"] = {"error": str(e)}

            else:
                result["status"] = "ERROR"
                result["result"] = f"Comando desconocido: {command}"

        except Exception as e:
            result["status"] = "ERROR"
            result["result"] = str(e)
            self.log(f"Error procesando tarea: {e}")

        # Send result
        try:
            # FIX: Use RESTful endpoint for AWS server connection
            # Server expects POST /tasks/{task_id}/result
            base_url = self.result_url.rstrip("/")
            post_url = f"{base_url}/{task_id}/result"
            auth = (self.username, self.password) if self.password else None
            
            requests.post(post_url, json={
                "status": result["status"],
                "result": result["result"]
            }, auth=auth)
            self.log(f"Resultado enviado para {task_id}")
        except Exception as e:
            self.log(f"Error enviando resultado: {e}")

# --- GUI ---
def load_config():
    # 1. Check current directory (priority for portable/dev)
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r') as f:
                return json.load(f)
        except:
            pass
            
    # 2. Check LOCALAPPDATA (installed location)
    try:
        app_data = os.getenv('LOCALAPPDATA', os.path.expanduser("~"))
        config_path = os.path.join(app_data, 'CDO_Organizer', 'agent_config.json')
        if os.path.exists(config_path):
            with open(config_path, 'r') as f:
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
        
        ttk.Label(frame, text="Contraseña:").grid(row=1, column=0, sticky=tk.W)
        self.pass_var = tk.StringVar(value=self.config.get("password", ""))
        ttk.Entry(frame, textvariable=self.pass_var, show="*").grid(row=1, column=1, sticky=tk.EW)

        ttk.Label(frame, text="URL Tareas:").grid(row=2, column=0, sticky=tk.W)
        # Fix: Use correct AWS endpoint (port 8000, /tasks/poll)
        default_task_url = "http://3.142.164.128:8000/tasks/poll"
        # Heuristic: if loaded config has "localhost" or "8501" or "/api/", replace with correct default
        loaded_task_url = self.config.get("task_url", default_task_url)
        if "localhost" in loaded_task_url or "8501" in loaded_task_url or "/api/" in loaded_task_url:
             loaded_task_url = default_task_url
             
        self.url_task_var = tk.StringVar(value=loaded_task_url)
        ttk.Entry(frame, textvariable=self.url_task_var).grid(row=2, column=1, sticky=tk.EW)
        
        ttk.Label(frame, text="URL Resultados:").grid(row=3, column=0, sticky=tk.W)
        default_res_url = "http://3.142.164.128:8000/tasks"
        loaded_res_url = self.config.get("result_url", default_res_url)
        if "localhost" in loaded_res_url or "8501" in loaded_res_url or "/api/" in loaded_res_url:
             loaded_res_url = default_res_url
             
        self.url_res_var = tk.StringVar(value=loaded_res_url)
        ttk.Entry(frame, textvariable=self.url_res_var).grid(row=3, column=1, sticky=tk.EW)
        
        self.btn_start = ttk.Button(frame, text="Iniciar Agente", command=self.toggle_agent)
        self.btn_start.grid(row=4, column=0, columnspan=2, pady=10)
        
        self.log_area = scrolledtext.ScrolledText(frame, height=15)
        self.log_area.grid(row=5, column=0, columnspan=2, sticky=tk.NSEW)
        
        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(5, weight=1)

    def log(self, msg):
        def _log():
            self.log_area.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")
            self.log_area.see(tk.END)
        self.root.after(0, _log)

    def invoke_on_gui(self, func, *args, **kwargs):
        """
        Executes a function on the main GUI thread and returns the result.
        This is thread-safe and blocking for the caller.
        """
        result_container = {}
        event = threading.Event()
        
        def wrapper():
            try:
                result_container["result"] = func(*args, **kwargs)
            except Exception as e:
                result_container["error"] = e
            finally:
                event.set()
        
        self.root.after(0, wrapper)
        event.wait()
        
        if "error" in result_container:
            raise result_container["error"]
        return result_container.get("result")

    def toggle_agent(self):
        if self.worker and self.worker.running:
            self.worker.stop()
            self.btn_start.config(text="Iniciar Agente")
            self.log("Agente detenido manualmente.")
        else:
            user = self.user_var.get()
            password = self.pass_var.get()
            t_url = self.url_task_var.get()
            r_url = self.url_res_var.get()
            
            if not user or not t_url:
                messagebox.showerror("Error", "Configure usuario y URLs")
                return
                
            self.config["username"] = user
            self.config["password"] = password
            self.config["task_url"] = t_url
            self.config["result_url"] = r_url
            save_config(self.config)
            
            self.worker = AgentWorker(user, t_url, r_url, password)
            self.worker.log_callback = self.log
            self.worker.gui_invoker = self.invoke_on_gui
            self.worker.start()
            self.btn_start.config(text="Detener Agente")

if __name__ == "__main__":
    root = tk.Tk()
    app = AgentGUI(root)
    root.mainloop()
