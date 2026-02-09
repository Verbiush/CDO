
import fitz  # PyMuPDF
import pandas as pd
import os
import io
import re
import PyPDF2
import json
import streamlit as st
try:
    import pdfplumber
except ImportError:
    pdfplumber = None

try:
    from PIL import Image
except ImportError:
    Image = None

try:
    import pytesseract
except ImportError as e:
    print(f"Warning: Tesseract OCR missing: {e}")
    pytesseract = None

try:
    import google.generativeai as genai
except ImportError:
    genai = None

from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

# --- Helper: Find Tesseract ---
def find_tesseract():
    # Common paths for Windows
    paths = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        r"D:\Program Files\Tesseract-OCR\tesseract.exe",
        os.path.join(os.getenv('LOCALAPPDATA', ''), 'Tesseract-OCR', 'tesseract.exe')
    ]
    for p in paths:
        if os.path.exists(p):
            return p
    return None

TESSERACT_PATH = find_tesseract()

def extract_text_pdfplumber(pdf_path):
    """Extraction using pdfplumber (robust layout/encoding)"""
    if not pdfplumber: return ""
    try:
        text = ""
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                extracted = page.extract_text(layout=True)
                if extracted:
                    text += extracted + "\n"
        return text
    except Exception as e:
        return ""

def extract_text_pypdf(pdf_path):
    """Fallback extraction using PyPDF2"""
    try:
        with open(pdf_path, 'rb') as f:
            reader = PyPDF2.PdfReader(f)
            text = ""
            for page in reader.pages:
                extracted = page.extract_text()
                if extracted:
                    text += extracted + "\n"
            return text
    except Exception as e:
        return ""

def extract_text_ocr(pdf_path):
    """Fallback extraction using OCR (Tesseract) via PyMuPDF images"""
    # 1. Lazy Import / Re-import attempts
    global pytesseract, Image, TESSERACT_PATH
    
    if pytesseract is None:
        try:
            import pytesseract as pt
            pytesseract = pt
        except ImportError: pass

    if Image is None:
        try:
            from PIL import Image as Img
            Image = Img
        except ImportError: pass
        
    # 2. Re-check Tesseract Path if not found yet
    if not TESSERACT_PATH:
        TESSERACT_PATH = find_tesseract()

    # Force update cmd
    if TESSERACT_PATH and pytesseract:
         pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH

    # 3. Validation
    if not pytesseract or not Image:
        return f"MISSING_LIBRARY_PYTESSERACT_OR_PILLOW (pytesseract={pytesseract is not None}, Image={Image is not None})"
        
    if not TESSERACT_PATH:
        return "TESSERACT_NOT_FOUND (Checked common paths, none found)"
        
    # Ensure cmd is set
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH
    
    try:
        doc = fitz.open(pdf_path)
        full_text = ""
        # Only process first 2 pages to save time, usually enough for SOS
        for i, page in enumerate(doc):
            if i > 1: break
            # Render page to image (dpi=300)
            pix = page.get_pixmap(dpi=300)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            
            # Run OCR (Spanish preferred, fallback to English)
            # Try 'spa' first, if fails (not installed), try 'eng'
            try:
                text = pytesseract.image_to_string(img, lang='spa')
            except:
                text = pytesseract.image_to_string(img, lang='eng')
                
            full_text += text + "\n"
        doc.close()
        return full_text
    except Exception as e:
        return f"OCR_ERROR: {str(e)} (Path tried: {TESSERACT_PATH})"

def extract_sos_data_ai(pdf_path, api_key, model_name):
    """
    Extract data from SOS Authorization PDF using Gemini AI (Vision Mode).
    """
    if not genai:
        return {"Error": "Librería google.generativeai no instalada."}
        
    # Strategy: Vision (Image) -> Text Fallback
    used_strategy = "Gemini_Vision"
    images = []
    
    # 1. Try to convert PDF page to Image
    try:
        # Check if PIL is available
        if Image is None:
            raise ImportError("PIL (Pillow) not available for image processing.")
            
        doc = fitz.open(pdf_path)
        # Process first 2 pages max
        for i in range(min(2, len(doc))):
            page = doc.load_page(i)
            pix = page.get_pixmap(dpi=300)
            img_bytes = pix.tobytes("jpeg")
            images.append(Image.open(io.BytesIO(img_bytes)))
        doc.close()
        
        if not images:
            raise Exception("No images extracted from PDF")
            
    except Exception as e:
        # Fallback to text if image conversion fails
        used_strategy = "Gemini_Text_Fallback"
        # print(f"Vision failed, falling back to text: {e}")

    # 2. Configure Gemini
    try:
        genai.configure(api_key=api_key)
        # Ensure model name has prefix
        if not model_name.startswith("models/"):
            model_name = f"models/{model_name}"
            
        model = genai.GenerativeModel(model_name)
        
        prompt = """
        Actúa como un analista de datos experto en salud.
        Analiza este documento "INFORME AUTORIZADOR EN LINEA" de SOS EPS (puede tener 1 o 2 páginas).
        Extrae los datos VISUALMENTE, tal como los vería un humano.
        
        Devuelve EXCLUSIVAMENTE un objeto JSON válido.
        
        Campos a extraer (usa null si no encuentras el dato):
        - "Fecha Consulta": Fecha en formato DD/MM/YYYY.
        - "Afiliado": Nombre completo del paciente.
        - "Identificación": Número de identificación.
        - "Plan": Nombre del plan (ej. BIENESTAR).
        - "Rango Salarial": (ej. A, B, C).
        - "Derecho": Descripción del derecho.
        - "Ambito": (ej. Ambulatorio).
        - "IPS Primaria": Nombre de la IPS asignada.
        - "IPS Solicitante": Nombre de la IPS que solicita.
        - "Código Prestación": Código CUPS (ej. 890208).
        - "Nombre Prestación": Descripción del servicio (ej. CONSULTA...).
        - "Cantidad": Cantidad autorizada (número, ej. 1).
        - "Respuesta EPS": (ej. Autorizado).
        - "No. Autorización": Número de autorización (generalmente largo, ej. 3781209).
        - "Justificación Resultado": Texto si existe.
        
        IMPORTANTE:
        - Si hay una tabla, extrae los datos de la primera fila de la tabla.
        - El campo "Cantidad" suele ser un número pequeño (1, 2, etc.) en la tabla.
        - El "No. Autorización" está en la última columna de la tabla.
        - Si el documento tiene varias páginas, busca la información en todas ellas.
        """
        
        inputs = [prompt]
        if images:
            inputs.extend(images)
        else:
            # Text Fallback
            text = extract_text_pdfplumber(pdf_path)
            
            # Check for Garbage (CID fonts)
            is_garbage = False
            if text:
                if text.count("(cid:") > 5 or len(text) < 50:
                    is_garbage = True
            
            if not text or is_garbage:
                text = extract_text_ocr(pdf_path)
                
            inputs.append(f"Texto del documento:\n{text[:10000]}")

        response = model.generate_content(inputs)
        raw_content = response.text.strip()
        
        # Robust JSON extraction
        json_start = raw_content.find('{')
        json_end = raw_content.rfind('}') + 1
        
        if json_start != -1 and json_end != -1:
            json_str = raw_content[json_start:json_end]
            data = json.loads(json_str)
        else:
            # If no JSON found, try to parse the text as fallback or return raw error
            raise ValueError(f"No valid JSON found in response. Raw: {raw_content[:200]}...")

        data["_DEBUG_STRATEGY"] = used_strategy
        return data
        
    except Exception as e:
        return {
            "Error": f"Fallo AI: {str(e)}", 
            "_DEBUG_STRATEGY": f"AI_FAIL_{used_strategy}",
            "Raw_Text_Preview": f"Image Mode ({len(images)} pages)" if images else "Text Mode",
            "Raw_Response": raw_content if 'raw_content' in locals() else "No response"
        }

def extract_sos_data_studio(pdf_path):
    """
    Implementation based on 'ai_studio_code (1).py' logic.
    Uses pdfplumber with specific regex and table structure.
    MODIFIED: Captures ALL rows (no break), raw extraction (no type conversion).
    """
    if not pdfplumber: return {}
    
    datos_extraidos = {"valid_extraction": False}
    
    with pdfplumber.open(pdf_path) as pdf:
        pagina = pdf.pages[0]
        texto = pagina.extract_text() or ""
        
        # 1. Regex Extraction (Raw capture)
        # Enhanced patterns to match main strategy
        patrones = {
            "Fecha Consulta": r"Fecha Consulta[:\s]+(\d{2}/\d{2}/\d{4})",
            "Identificación": r"Identificación[:\s]+(\d+)",
            "Afiliado": r"Afiliado[:\s]+(.+?)(?=\s+Identificación|Plan|\n|$)",
            "Plan": r"Plan[:\s]+(.+?)(?=\s+Rango|\n|$)",
            "Derecho": r"Derecho[:\s]+(.+?)(?=\s+Ambito|\s+IPS Primaria|\n|$)",
            "IPS Primaria": r"IPS Primaria[:\s]+(.+?)(?=\s+IPS Solicitante|\n|$)",
            "IPS Solicitante": r"IPS Solicitante[:\s]+(.+?)(?=\n|$)",
            "Ambito": r"Ambito[:\s]+([A-Z\s]+)"
        }

        for clave, patron in patrones.items():
            match = re.search(patron, texto, re.IGNORECASE)
            if match:
                datos_extraidos[clave] = match.group(1).strip()

        # 2. Table Extraction (Iterate ALL rows)
        tabla = pagina.extract_table()
        
        if tabla:
            codigos = []
            nombres = []
            cantidades = []
            respuestas = []
            autorizaciones = []
            
            for fila in tabla:
                # Basic sanity check: row must have content
                if not fila: continue
                
                # Check if it's a header row
                row_str = "".join([str(c) for c in fila if c]).lower()
                if "código" in row_str and "prestación" in row_str:
                    continue
                if "autorizador" in row_str and "linea" in row_str:
                    continue
                # NEW: Exclude Validador Web Headers explicitly
                if "nro." in row_str and "p-autorización" in row_str:
                    continue
                if "justificación" in row_str and "resultado" in row_str:
                    continue

                # GARBAGE CHECK: If the row contains CID font garbage, abort this method
                # to allow fallback to OCR.
                if "(cid:" in row_str:
                    return {"valid_extraction": False}

                # SOS Structure Assumption (Raw Index Access):
                # 0: Codigo, 1: Nombre, 2: Cantidad, 3: Respuesta ... 7: Auth
                
                # Get raw values (replicate, don't convert)
                val_codigo = str(fila[0]).strip() if len(fila) > 0 and fila[0] else ""
                val_nombre = str(fila[1]).strip().replace("\n", " ") if len(fila) > 1 and fila[1] else ""
                val_cant = str(fila[2]).strip() if len(fila) > 2 and fila[2] else ""
                val_resp = str(fila[3]).strip() if len(fila) > 3 and fila[3] else ""
                val_auth = ""

                # --- NEW: Support for 4-column Validador Web format ---
                # [Nro. P-Autorización, Aprobada, Justificación, Nro. Autorización]
                if len(fila) == 4 and (str(fila[0]).isdigit() or str(fila[3]).isdigit()):
                    val_codigo = str(fila[0]).strip() # P-Autorización
                    val_resp = str(fila[1]).strip()   # Aprobada
                    # val_just = str(fila[2]).strip() # Justificación (unused in standard output?)
                    val_auth = str(fila[3]).strip()   # Nro. Autorización
                    val_nombre = f"Ver P-Autorización {val_codigo}" # Placeholder name
                    val_cant = "1"
                else:
                    # Standard 8+ column format
                    if len(fila) > 7 and fila[7]:
                        val_auth = str(fila[7]).strip()
                    elif len(fila) > 0:
                        # Fallback: check last non-empty column
                        non_empty = [c for c in fila if c]
                        if non_empty:
                            last_val = str(non_empty[-1]).strip()
                            # Heuristic: Auth is usually numeric and > 3 chars
                            if len(last_val) > 3 and any(char.isdigit() for char in last_val):
                                # Don't confuse with code if they are the same
                                if last_val != val_codigo:
                                    val_auth = last_val
                
                # Additional Validation: Code should be numeric or empty, not "Nro. P-Au"
                if val_codigo and not val_codigo.replace(".","").isdigit():
                    # It's likely a header that slipped through
                    continue

                # If we have at least a code or a name, treat as valid data row
                if val_codigo or val_nombre:
                    if val_codigo: codigos.append(val_codigo)
                    if val_nombre: nombres.append(val_nombre)
                    if val_cant: cantidades.append(val_cant)
                    if val_resp: respuestas.append(val_resp)
                    if val_auth: autorizaciones.append(val_auth)
                    
                    datos_extraidos["valid_extraction"] = True
            
            # Join multiple items with pipe (or comma if preferred, but pipe is safer for CSV export)
            if codigos: datos_extraidos["Código Prestación"] = " | ".join(codigos)
            if nombres: datos_extraidos["Nombre Prestación"] = " | ".join(nombres)
            if cantidades: datos_extraidos["Cantidad"] = " | ".join(cantidades)
            if respuestas: datos_extraidos["Respuesta EPS"] = " | ".join(respuestas)
            if autorizaciones: datos_extraidos["No. Autorización"] = " | ".join(autorizaciones)
    
    return datos_extraidos

def extract_sos_data(pdf_path):
    """
    Extract data from SOS Authorization PDF.
    Strategy: pdfplumber Tables -> Regex -> PyMuPDF/PyPDF2 -> OCR (Last Resort)
    """
    data = {
        "Fecha Consulta": "",
        "Afiliado": "",
        "Identificación": "",
        "Plan": "",
        "IPS Primaria": "",
        "Código Prestación": "",
        "Nombre Prestación": "",
        "Cantidad": "",
        "Respuesta EPS": "",
        "No. Autorización": "",
        "Ambito": "",
        "Derecho": "",
        "IPS Solicitante": "",
        "_DEBUG_STRATEGY": "None"
    }
    
    # --- Strategy 0: Studio Logic (pdfplumber specific patterns) ---
    # Derived from user provided script "ai_studio_code (1).py"
    if pdfplumber:
        try:
            studio_data = extract_sos_data_studio(pdf_path)
            if studio_data.get("valid_extraction"):
                # Merge studio data into main data
                for k, v in studio_data.items():
                    if k in data and v:
                        data[k] = v
                data["_DEBUG_STRATEGY"] = "Studio_Logic"
                return data
        except Exception as e:
            # print(f"Studio logic failed: {e}")
            pass

    # --- Strategy 1: pdfplumber Table Extraction (Best for multi-line cells) ---
    codigos = []
    nombres = []
    autorizaciones = []
    is_garbage = False
    clean_text = ""
    
    if pdfplumber:
        try:
            with pdfplumber.open(pdf_path) as pdf:
                page = pdf.pages[0] # Assume first page
                text = page.extract_text() or ""
                
                # Check for Garbage (CID fonts)
                is_garbage = "(cid:" in text or len(text.strip()) < 10
                
                if not is_garbage:
                    # Header Extraction via Regex on text
                    # CLEAN TEXT: Remove known header/footer garbage lines to prevent false positives
                    clean_lines = []
                    for line in text.split('\n'):
                        if "INFORME | RESULTADO" in line or "https://" in line or "SOS.COM.CO" in line:
                            continue
                        clean_lines.append(line)
                    clean_text = "\n".join(clean_lines)

                    # Regex Patterns (Enhanced)
                    patterns = {
                        "Fecha Consulta": r"Fecha Consulta[:\s]+(\d{2}/\d{2}/\d{4})",
                        "Afiliado": r"Afiliado[:\s]+(.+?)(?=\s+Identificación|Plan|\n|$)",
                        "Identificación": r"Identificación[:\s]+(\d+)",
                        "Plan": r"Plan[:\s]+(.+?)(?=\s+Rango|\n|$)",
                        "IPS Primaria": r"IPS Primaria[:\s]+(.+?)(?=\s+IPS Solicitante|\n|$)",
                        "Ambito": r"Ambito[:\s]+([A-Z\s]+)",
                        "Derecho": r"Derecho[:\s]+(.+?)(?=\s+Ambito|\s+IPS Primaria|\n|$)",
                        "IPS Solicitante": r"IPS Solicitante[:\s]+(.+?)(?=\n|$)"
                    }
                    
                    for key, pat in patterns.items():
                        m = re.search(pat, clean_text, re.IGNORECASE)
                        if m:
                            val = m.group(1).strip()
                            # Clean common trailing headers if regex overshot
                            val = re.sub(r'\s+(Identificación|Plan|Rango|Ambito|IPS).*$', '', val, flags=re.IGNORECASE)
                            data[key] = val

                    # Table Extraction for Auth Details
                    # Find table below headers
                    table_settings = {
                        "vertical_strategy": "text", 
                        "horizontal_strategy": "text",
                        "intersection_y_tolerance": 5
                    }
                    tables = page.extract_tables(table_settings)
                    
                    codigos = []
                    nombres = []
                    autorizaciones = []
                    respuestas = []
                    cantidades = []
                    
                    for table in tables:
                        for row in table:
                            # Filter None and empty strings
                            row = [str(cell).strip() if cell else "" for cell in row]
                            
                            # Heuristic: Row should contain an Auth Number or Code
                            # SOS Codes are usually 6 digits (e.g. 890208)
                            # Auth Numbers are usually long (e.g. 3781209)
                            
                            # Check if this row looks like a data row (not header)
                            row_upper = [c.upper() for c in row]
                            if "CÓDIGO" in row_upper or "PRESTACIÓN" in row_upper or "P-AUTORIZACIÓN" in row_upper: continue
                            
                            code_val = ""
                            name_val = ""
                            auth_val = ""
                            resp_val = ""
                            qty_val = ""
                            
                            # DETECCIÓN DE FORMATO VALIDADOR WEB (4 columnas)
                            if len(row) == 4 and (row[0].isdigit() or row[3].isdigit()):
                                # [Nro. P-Autorización, Aprobada, Justificación, Nro. Autorización]
                                code_val = row[0]
                                resp_val = row[1]
                                auth_val = row[3]
                                name_val = "Ver P-Autorización " + row[0]
                                qty_val = "1"
                                
                                if code_val: codigos.append(code_val)
                                if name_val: nombres.append(name_val)
                                if resp_val: respuestas.append(resp_val)
                                if qty_val: cantidades.append(qty_val)
                                if auth_val: autorizaciones.append(auth_val)
                                continue

                            # Formato Estándar (8+ columnas) o Genérico
                            # Try to identify columns by content type
                            # 1. Code: 6 digits
                            for cell in row:
                                if re.match(r'^\d{6}$', cell):
                                    code_val = cell
                                    break
                            
                            # 2. Status/Response: "Autorizado", "Negado"
                            for cell in row:
                                if cell.upper() in ["AUTORIZADO", "NEGADO", "PARCIAL"]:
                                    resp_val = cell
                                    break
                            
                            # 2.5 Quantity: Small number (1-100), not Code
                            for cell in row:
                                if re.match(r'^\d{1,3}$', cell) and cell != code_val and cell != auth_val:
                                    # Ensure it's not a part of a date or ID (unlikely in this table structure)
                                    qty_val = cell
                                    # Usually Quantity is between Name and Status
                                    break

                            # 3. Name: Long text, not header
                            # Usually between Code and Status, or first long text
                            for cell in row:
                                if len(cell) > 10 and not re.search(r'\d{5}', cell) and cell != resp_val:
                                    # Exclude common headers if they slipped in
                                    if "CONSULTA" in cell or "PROCEDIMIENTO" in cell or "DERECHO" not in cell:
                                        name_val = cell
                                        break
                                        
                            # 4. Auth Number: Long digits (often > 6)
                            # Or matches specific pattern
                            for cell in row:
                                clean_cell = re.sub(r'[^\d]', '', cell)
                                if len(clean_cell) >= 4 and cell != code_val:
                                    auth_val = cell
                                    break
                                    
                            if code_val or auth_val:
                                if code_val: codigos.append(code_val)
                                if name_val: nombres.append(name_val)
                                if resp_val: respuestas.append(resp_val)
                                if qty_val: cantidades.append(qty_val)
                                
                                # Advanced Auth Extraction (Proximity Fallback)
                                if not auth_val:
                                    # Sometimes Auth is merged or misaligned
                                    # Look for number at end of row
                                    pass

                                # Fallback: Look for number at end of row
                                if not auth_val:
                                    # 1. Priority: Look for exactly 7 digits (Common SOS format)
                                    for cell in reversed(row):
                                        c_clean = re.sub(r'[^\d]', '', str(cell))
                                        # Avoid confusing Code (often 6 digits) with Auth (often 7)
                                        if len(c_clean) == 7 and c_clean != code_val:
                                            auth_val = str(cell)
                                            break
                                            
                                    # 2. Standard check: >= 4 digits
                                    if not auth_val:
                                        for cell in reversed(row):
                                            c_clean = re.sub(r'[^\d]', '', str(cell))
                                            if len(c_clean) >= 4:
                                                # Avoid confusing Code with Auth if they are same value/column
                                                if code_val and str(cell) == code_val:
                                                    continue 
                                                auth_val = str(cell)
                                                break
                                            
                                if auth_val:
                                    autorizaciones.append(auth_val)

                    if nombres:
                        data["Nombre Prestación"] = " | ".join(nombres)
                        data["No. Autorización"] = " | ".join(autorizaciones)
                        
        except Exception as e:
            # print(f"Strategy 1 failed: {e}")
            is_garbage = True # Assume failure means we might need OCR

    # --- Strategy 2: OCR Fallback (If text was garbage or missing) ---
    if is_garbage or not data["Nombre Prestación"]:
        # Try to get text via OCR
        ocr_text = extract_text_ocr(pdf_path)
        
        if ocr_text and "OCR_ERROR" not in ocr_text:
            data["_DEBUG_STRATEGY"] = "OCR_Fallback"
            clean_text = ocr_text
            
            # Apply Regex Patterns on OCR Text
            patterns = {
                "Fecha Consulta": r"Fecha Consulta[:\s]+(\d{2}/\d{2}/\d{4})",
                "Afiliado": r"Afiliado[:\s]+(.+?)(?=\s+Identificación|Plan|$)",
                "Identificación": r"Identificación[:\s]+(\d+)",
                "Plan": r"Plan[:\s]+(.+?)(?=\s+Rango|$)",
                "IPS Primaria": r"IPS Primaria[:\s]+(.+?)(?=\s+IPS Solicitante|$)",
                "Ambito": r"Ambito[:\s]+(.+?)(?=\n|$)",
                "Derecho": r"Derecho[:\s]+(.+?)(?=\s+Ambito|$)",
                "IPS Solicitante": r"IPS Solicitante[:\s]+(.+?)(?=\n|$)"
            }
            
            for key, pat in patterns.items():
                if not data[key]: # Only if not already found
                    m = re.search(pat, clean_text, re.IGNORECASE)
                    if m:
                        data[key] = m.group(1).strip()
            
            # Try to find Auth Number in OCR text
            if not data["No. Autorización"]:
                 auths = re.findall(r'\b\d{7}\b', clean_text)
                 # Filter ID
                 if auths:
                     valid = [a for a in auths if a != data["Identificación"]]
                     if valid:
                         data["No. Autorización"] = " | ".join(list(set(valid)))

            # Try to find Service Name (Harder in unstructured OCR)
            # Look for lines with "AUTORIZADO"
            if not data["Nombre Prestación"]:
                names = []
                for line in clean_text.split('\n'):
                    if "AUTORIZADO" in line.upper():
                        # Assume name is before it
                        # "890208 CONSULTA MEDICA ... AUTORIZADO"
                        parts = line.split("AUTORIZADO")
                        if parts[0]:
                            candidate = parts[0].strip()
                            # Remove trailing digits
                            candidate = re.sub(r'\d+$', '', candidate).strip()
                            if len(candidate) > 5:
                                names.append(candidate)
                if names:
                    data["Nombre Prestación"] = " | ".join(names)

    # --- Strategy 3: PyMuPDF / Regex on clean_text (from Strategy 1 or 2) ---
    if not data["Código Prestación"] and codigos:
        data["Código Prestación"] = " | ".join(codigos)
    if not data["Nombre Prestación"] and nombres:
        data["Nombre Prestación"] = " | ".join(nombres)
    if not data["No. Autorización"] and autorizaciones:
        data["No. Autorización"] = " | ".join(autorizaciones)
    
    # Check if we found critical data via Strategy 1 or 2
    if data["Nombre Prestación"] or data["No. Autorización"]:
         data["_DEBUG_STRATEGY"] = "Strategy_1_or_2"
         return data 
         
    # --- Strategy 1.5: Visual/Spatial Extraction Fallback ---
    # Captura "visual" de columnas basada en la posición de los encabezados "Nombre Prestación" y "No. Autorización"
    if pdfplumber and (not data["Nombre Prestación"] or not data["No. Autorización"]):
        try:
            with pdfplumber.open(pdf_path) as pdf:
                page = pdf.pages[0]
                words = page.extract_words()
                
                # Buscar encabezados clave
                nombre_header_box = None # (x0, top, x1, bottom)
                auth_header_box = None
                
                # Agrupar palabras para encontrar "Nombre Prestación" y "No. Autorización"
                # Buscamos "Prestación" que esté a la derecha de un "Código" o que sea la segunda "Prestación"
                # O simplemente buscamos "Nombre" seguido de "Prestación"
                
                # Simplificación robusta: Buscar palabra "Prestación" y "Autorización"
                # Asumimos que "Nombre Prestación" está aprox en el medio-izquierda y "Autorización" al final derecha
                
                prestacion_candidates = [w for w in words if "prestación" in w['text'].lower() and "código" not in w['text'].lower()]
                autorizacion_candidates = [w for w in words if "autorización" in w['text'].lower() and "no." in w['text'].lower()]
                
                # Fallback candidates if exact match fails
                if not prestacion_candidates:
                    prestacion_candidates = [w for w in words if "descripción" in w['text'].lower()]
                if not autorizacion_candidates:
                    autorizacion_candidates = [w for w in words if "autorización" in w['text'].lower()]

                if prestacion_candidates:
                    # Usar el primero que parezca encabezado (usualmente más arriba)
                    prestacion_candidates.sort(key=lambda x: x['top'])
                    target_prestacion = prestacion_candidates[0]
                    
                    # Definir zona de búsqueda DEBAJO del encabezado
                    x0 = target_prestacion['x0'] - 10
                    top = target_prestacion['bottom'] + 2 # Justo debajo
                    x1 = x0 + 250 # Ancho estimado de la columna Nombre
                    bottom = top + 150 # Altura de búsqueda (varias líneas)
                    
                    if not data["Nombre Prestación"]:
                        # Crop y extraer
                        crop = page.crop((x0, top, x1, bottom))
                        # Extraer texto preservando layout visual para ver si hay varias líneas
                        text_blob = crop.extract_text()
                        if text_blob:
                            # Limpiar: tomar primeras líneas no vacías que no parezcan otro header
                            lines = [l.strip() for l in text_blob.split('\n') if l.strip()]
                            if lines:
                                # Tomar todo lo que parezca texto válido hasta encontrar un número aislado o fin
                                valid_lines = []
                                for l in lines:
                                    if "cantidad" in l.lower() or "respuesta" in l.lower(): break
                                    valid_lines.append(l)
                                data["Nombre Prestación"] = " ".join(valid_lines)
                                data["_DEBUG_STRATEGY"] = "pdfplumber_visual_spatial"

                # Heurística para No. Autorización:
                # Buscar "Autorización" (o No. Autorización)
                # Suele estar a la derecha
                target_auth = None
                if autorizacion_candidates:
                    # Tomar la que esté más a la derecha
                    autorizacion_candidates.sort(key=lambda x: x['x0'], reverse=True)
                    target_auth = autorizacion_candidates[0]
                
                if target_auth:
                    x0 = target_auth['x0'] - 20
                    top = target_auth['bottom'] + 2
                    x1 = page.width # Hasta el final de la página
                    bottom = top + 150
                    
                    if not data["No. Autorización"]:
                        crop = page.crop((x0, top, x1, bottom))
                        text_blob = crop.extract_text()
                        if text_blob:
                            # Buscar números
                            nums = re.findall(r'\d{4,}', text_blob)
                            # Filtrar ID del pie de página si se coló (944...)
                            valid_nums = [n for n in nums if "94433448" not in n]
                            if valid_nums:
                                data["No. Autorización"] = valid_nums[0]
                                if data["_DEBUG_STRATEGY"] == "None":
                                    data["_DEBUG_STRATEGY"] = "pdfplumber_visual_spatial"
                                else:
                                    data["_DEBUG_STRATEGY"] += "+visual_auth"

        except Exception as e:
            pass

    # --- Strategy 2: Text Extraction Fallback (if table failed or garbage) ---
    full_text = ""
    used_strategy = "None"
    
    # Try pdfplumber text
    if pdfplumber:
        plumber_text = extract_text_pdfplumber(pdf_path)
        if "Fecha" in plumber_text and "(cid:" not in plumber_text:
            full_text = plumber_text
            used_strategy = "pdfplumber_text"
    
    # Try PyMuPDF
    if not full_text:
        try:
            doc = fitz.open(pdf_path)
            fitz_text = ""
            for page in doc:
                fitz_text += page.get_text("text") + "\n"
            doc.close()
            if "Fecha" in fitz_text and "(cid:" not in fitz_text:
                full_text = fitz_text
                used_strategy = "PyMuPDF"
        except: pass

    # Try OCR (Last Resort - if text is empty or has CID garbage)
    if not full_text or "(cid:" in full_text:
        ocr_text = extract_text_ocr(pdf_path)
        if "TESSERACT_NOT_FOUND" in ocr_text:
             data["_DEBUG_STRATEGY"] = "MISSING_TESSERACT_OCR"
             data["_DEBUG_RAW_START"] = "Este PDF requiere OCR. Instale Tesseract-OCR."
             return data
        elif "OCR_ERROR" in ocr_text:
             data["_DEBUG_STRATEGY"] = "OCR_FAILED"
             data["_DEBUG_RAW_START"] = ocr_text
             return data
        else:
             full_text = ocr_text
             used_strategy = "OCR_Tesseract"

    data["_DEBUG_STRATEGY"] = used_strategy
    if not full_text: return data
    
    # --- Universal Regex Extraction (Fallback / Supplement) ---
    
    # Pre-clean text: Remove header garbage that often confuses Regex
    # e.g. "AUTORIZACION DE SERVICIOS..." header repeating
    full_text_clean = re.sub(r'AUTORIZACION/AUTORIZADOR', '', full_text, flags=re.IGNORECASE)

    # 1. Header Fields
    if not data["Fecha Consulta"]:
        m = re.search(r"Fecha Consulta[:\s]+(\d{2}/\d{2}/\d{4})", full_text)
        if m: data["Fecha Consulta"] = m.group(1)

    if not data["Afiliado"]:
        # Try to capture name until "Identificación" or "Plan"
        m = re.search(r"Afiliado[:\s]+(.+?)(?=\s+Identificación|\s+Plan|\n)", full_text)
        if m: data["Afiliado"] = m.group(1).strip()
        
    if not data["Identificación"]:
        m = re.search(r"Identificación[:\s]+(\d+)", full_text)
        if m: data["Identificación"] = m.group(1)

    if not data["Plan"]:
        m = re.search(r"Plan[:\s]+(.+?)(?=\s+Rango|\n)", full_text)
        if m: data["Plan"] = m.group(1).strip()
        
    if not data["IPS Primaria"]:
        m = re.search(r"IPS Primaria[:\s]+(.+?)(?=\s+IPS Solicitante|\n|$)", full_text)
        if m: data["IPS Primaria"] = m.group(1).strip()
        
    # 2. Table Data (Row based regex)
    # Look for lines starting with Code (digits)
    
    lines = full_text_clean.split('\n')
    codigos = []
    nombres = []
    autorizaciones = []
    respuestas = []
    
    for line in lines:
        line = line.strip()
        if not line: continue
        
        # Pattern: CODE ... NAME ... QTY ... STATUS ... AUTH
        # This is very hard to regex perfectly due to variable spaces, but we try common patterns
        
        # 1. Identify Status "Autorizado"
        status_match = re.search(r"(Autorizado|Negado|Parcial)", line, re.IGNORECASE)
        
        if status_match:
            # Only process if we haven't found this via table extraction already
            if not data["Código Prestación"]: 
                
                # 2. Extract Code (Look at START of line)
                # "890208 CONSULTA..."
                # Must be at start, digits, > 3 length
                # Relaxed anchor: Allow some garbage/spaces at start, but ensure it's the first number
                code_match = re.search(r"(?:^|\s)(\d{4,})\s+", line)
                if code_match:
                    # Verify it's at the beginning (ignoring non-alphanumeric garbage)
                    pre_match = line[:code_match.start()]
                    if not re.search(r"[a-zA-Z]", pre_match):
                        val = code_match.group(1)
                        if val not in codigos:
                            codigos.append(val)
                    else:
                        code_match = None # Not the code if preceded by letters
                
                # 3. Extract Name (Between Code and Status)
                # "890208 CONSULTA DE PRIMERA VEZ... Autorizado"
                start_idx = 0
                if code_match:
                    start_idx = code_match.end()
                
                end_idx = status_match.start()
                if end_idx > start_idx:
                    raw_name = line[start_idx:end_idx]
                    # Remove trailing digits/spaces (Quantity)
                    # "CONSULTA ... 1 " -> "CONSULTA ..."
                    clean_name = re.sub(r'\s+\d+\s*$', '', raw_name).strip()
                    if clean_name and clean_name not in nombres:
                        nombres.append(clean_name)
                        
                # 4. Extract Auth Number (After Status)
                # "Autorizado ... 3781209"
                after_status = line[status_match.end():]
                # Find long number
                auth_nums = re.findall(r'\d{4,}', after_status)
                if auth_nums:
                    for num in auth_nums:
                        if num not in autorizaciones:
                            autorizaciones.append(num)

    if not data["Código Prestación"] and codigos:
        data["Código Prestación"] = " | ".join(codigos)
    if not data["Nombre Prestación"] and nombres:
        data["Nombre Prestación"] = " | ".join(nombres)
    if not data["No. Autorización"] and autorizaciones:
        data["No. Autorización"] = " | ".join(autorizaciones)
        
    # --- Fallback: Find Auth Number in isolation ---
    if not data["No. Autorización"]:
        # Look for "No. Autorización" followed by digits
        m = re.search(r"No\. Autorización\s*(\d+)", full_text)
        if m: data["No. Autorización"] = m.group(1)
        else:
            # Just look for loose 7-digit numbers at end of text (often the auth number)
            # Dangerous heuristic, but useful for SOS
            candidates = re.findall(r'\b\d{7}\b', full_text)
            # Filter out ID
            id_val = data["Identificación"]
            valid_candidates = [c for c in candidates if c != id_val]
             
            if valid_candidates:
                # Often the authorization is near the "Autorizado" word
                # Check proximity? For now just take first unique
                
                # Special Check: "3781209" style
                # Filter out numbers starting with '890' (often codes) if we aren't sure
                seven_plus = [c for c in valid_candidates if not c.startswith("890")]
                if seven_plus:
                    data["No. Autorización"] = seven_plus[0] # Take first valid long number
                else:
                    data["No. Autorización"] = valid_candidates[0] # Fallback to first valid


    # --- Dynamic Extraction (Catch-all for other Label: Value pairs) ---
    # Captures anything looking like "Label: Value"
    # Filter to reasonable labels (Start with uppercase, no numbers, length < 40)
    
    dynamic_matches = re.finditer(r"(?P<key>[A-ZÁÉÍÓÚÑ][a-zA-Záéíóúñ\s\.]+)\s*[:]\s*(?P<val>.+?)(?=\s\s|\n|$)", full_text)
    for m in dynamic_matches:
        key = m.group("key").strip()
        val = m.group("val").strip()
        
        # Heuristic filters
        if len(key) > 40 or len(key) < 2: continue
        if any(x in key for x in [":", "=", "/"]): continue # Bad keys
        if "www" in key.lower() or "http" in key.lower(): continue
        
        # Normalize key
        key_norm = re.sub(r'\s+', ' ', key).title()
        
        # Normalize specific keys to avoid duplicates
        if "Codigo" in key_norm and "Prestacion" in key_norm: key_norm = "Código Prestación"
        if "Autorizacion" in key_norm: key_norm = "No. Autorización"
        if "Identificacion" in key_norm: key_norm = "Identificación"
        
        # Don't overwrite existing specific extractions
        if key_norm not in data and val:
             data[key_norm] = val

    # --- Strategy 3: Header-Value Proximity (OCR Text Layout) ---
    # Specific fallback for when row-based extraction fails but headers are readable
    if not data["Nombre Prestación"] and "Nombre Prestación" in full_text:
         # Find "Nombre Prestación" index
         idx = full_text.find("Nombre Prestación")
         # Look at text immediately following it, or on next line
         pass 

    return data

def worker_analisis_sos(file_list, use_ai=False, silent_mode=False):
    if not file_list: return None
    pdf_files = [f for f in file_list if f["Ruta completa"].lower().endswith(".pdf")]
    
    extracted_data = []
    
    # Progress bar setup (needs st context, assuming called from app_web)
    progress_bar = None
    if not silent_mode:
        progress_bar = st.progress(0, text="Analizando PDFs...")
    total = len(pdf_files)
    
    # Get API Key once if AI
    api_key = None
    model_name = "models/gemini-1.5-flash-001"
    
    if use_ai:
        api_key = st.session_state.app_config.get("gemini_api_key")
        model_name = st.session_state.app_config.get("gemini_model", "models/gemini-1.5-flash-001")
        if not api_key:
            if not silent_mode:
                st.error("⚠️ Para usar IA, primero configura tu API Key en la barra lateral.")
            return None

    for i, item in enumerate(pdf_files):
        pdf_path = item["Ruta completa"]
        filename = os.path.basename(pdf_path)
        
        if progress_bar:
            progress_bar.progress((i + 1) / total, text=f"Analizando: {filename}")
        
        if use_ai:
            row_data = extract_sos_data_ai(pdf_path, api_key, model_name)
        else:
            row_data = extract_sos_data(pdf_path)
        
        final_row = {'Archivo': filename}
        final_row.update(row_data)
        extracted_data.append(final_row)

    if progress_bar:
        progress_bar.empty()
    if not extracted_data:
        return None

    df = pd.DataFrame(extracted_data)
    
    # Clean illegal characters
    def clean_illegal_chars(val):
        if isinstance(val, str):
            return ILLEGAL_CHARACTERS_RE.sub("", val)
        return val

    df = df.applymap(clean_illegal_chars)
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="Analisis_SOS")
    output.seek(0)
    
    # Generate CSV/TXT output
    csv_output = io.BytesIO()
    df.to_csv(csv_output, index=False, encoding='utf-8-sig', sep=',')
    csv_output.seek(0)
    
    return output, csv_output
