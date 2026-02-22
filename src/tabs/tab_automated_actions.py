import streamlit as st
import os
import time
import shutil
import pandas as pd
import json
import re
import random
import math
from datetime import datetime
import zipfile
import io
import fitz  # PyMuPDF
from PIL import Image, ImageDraw, ImageFont
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import unicodedata
try:
    import openpyxl
except ImportError:
    openpyxl = None
import requests
import base64
import urllib.parse
import xml.etree.ElementTree as ET

# --- CONDITIONAL IMPORTS FOR ANALYSIS WORKERS ---
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
except ImportError:
    pytesseract = None
    
# --- IMPORTS & SETUP ---
try:
    from modules.registraduria_validator import ValidatorRegistraduria
    from modules.adres_validator import ValidatorAdres, ValidatorAdresWeb
except ImportError:
    try:
        from src.modules.registraduria_validator import ValidatorRegistraduria
        from src.modules.adres_validator import ValidatorAdres, ValidatorAdresWeb
    except ImportError:
        pass

try:
    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service
    from webdriver_manager.chrome import ChromeDriverManager
except ImportError:
    pass

try:
    from task_manager import submit_task
    from gui_utils import seleccionar_carpeta_nativa
except ImportError:
    try:
        from src.task_manager import submit_task
        from src.gui_utils import seleccionar_carpeta_nativa
    except ImportError:
        def submit_task(name, func, *args, **kwargs):
            st.warning(f"Task Manager not available. Running {name} synchronously.")
            return func(*args, **kwargs)
        def seleccionar_carpeta_nativa(key):
            return st.text_input(f"Ruta para {key}", key=key)

try:
    from pdf2docx import Converter
    HAS_PDF2DOCX = True
except ImportError:
    HAS_PDF2DOCX = False

try:
    from docx2pdf import convert as convert_docx_to_pdf
    HAS_DOCX2PDF = True
except ImportError:
    HAS_DOCX2PDF = False

try:
    import win32com.client
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False

try:
    import pyperclip
except ImportError:
    pyperclip = None

import google.generativeai as genai

# --- HELPERS ---

def natural_sort_key(s):
    return [int(text) if text.isdigit() else text.lower() for text in re.split('([0-9]+)', s)]

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

def get_val_ci(data_dict, key):
    if not isinstance(data_dict, dict): return None
    for k, v in data_dict.items():
        if k.lower() == key.lower():
            return v
    return None

def recursive_clean_json(data):
    if isinstance(data, dict):
        return {k: recursive_clean_json(v) for k, v in data.items() if v is not None}
    elif isinstance(data, list):
        return [recursive_clean_json(item) for item in data]
    else:
        return data

def recursive_strip(data):
    if isinstance(data, dict):
        return {k: recursive_strip(v) for k, v in data.items()}
    elif isinstance(data, list):
        return [recursive_strip(i) for i in data]
    elif isinstance(data, str):
        return data.strip()
    return data

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

# --- HELPERS: SIGNATURES ---

def _dibujar_trazo_vocal(draw, x_base, y_centro, ascii_val, colores):
    """Dibuja un trazo característico de vocal"""
    color = random.choice(colores)
    grosor = random.randint(2, 3)
    
    # Crear arco o curva
    altura_arco = 20 + (ascii_val % 15)
    puntos = []
    
    for i in range(20):
        angulo = (i / 19.0) * math.pi
        x = x_base + i * 2
        y = y_centro - math.sin(angulo) * altura_arco + random.randint(-2, 2)
        puntos.append((x, y))
    
    # Dibujar curva
    for i in range(len(puntos) - 1):
        draw.line([puntos[i], puntos[i + 1]], fill=color, width=grosor)

def _dibujar_trazo_consonante_dura(draw, x_base, y_centro, ascii_val, colores):
    """Dibuja un trazo característico de consonante dura"""
    color = random.choice(colores)
    grosor = random.randint(3, 4)
    
    # Línea con ángulos y cambios de dirección
    x = x_base
    y = y_centro + random.randint(-10, 10)
    
    # Línea ascendente
    draw.line([(x, y), (x + 15, y - 20)], fill=color, width=grosor)
    # Línea horizontal
    draw.line([(x + 15, y - 20), (x + 30, y - 15)], fill=color, width=grosor)
    # Línea descendente
    draw.line([(x + 30, y - 15), (x + 40, y + 10)], fill=color, width=grosor)

def _dibujar_trazo_generico(draw, x_base, y_centro, ascii_val, colores, width, height):
    """Dibuja un trazo genérico fluido"""
    color = random.choice(colores)
    grosor = random.randint(2, 3)
    
    # Crear línea ondulada
    puntos = []
    for i in range(30):
        x = x_base + i * 1.5
        # Onda basada en el valor ASCII
        onda = math.sin((x - x_base) * 0.2 + ascii_val * 0.1) * 15
        y = y_centro + onda + random.randint(-3, 3)
        puntos.append((x, y))
    
    # Dibujar línea ondulada
    for i in range(len(puntos) - 1):
        draw.line([puntos[i], puntos[i + 1]], fill=color, width=grosor)
    
    # Añadir algunos puntos extra para dar textura
    for _ in range(random.randint(2, 4)):
        punto_x = random.randint(int(x_base), int(x_base + 40))
        punto_y = y_centro + random.randint(-5, 5)
        draw.ellipse([punto_x - 1, punto_y - 1, punto_x + 1, punto_y + 1], fill=color)
    
    # Añadir algunos detalles decorativos adicionales
    # Pequeñas líneas adicionales
    for _ in range(random.randint(1, 3)):
        x_rand = random.randint(20, width - 20)
        y_rand = random.randint(height // 3, 2 * height // 3)
        longitud = random.randint(10, 25)
        angulo = random.uniform(-0.5, 0.5)
        
        x_fin = x_rand + int(longitud * math.cos(angulo))
        y_fin = y_rand + int(longitud * math.sin(angulo))
        
        draw.line([(x_rand, y_rand), (x_fin, y_fin)], fill=random.choice(colores), width=random.randint(1, 2))
    
    # Añadir un toque final con una línea más prominente
    y_linea_principal = height // 2 + random.randint(-10, 10)
    x_inicio_principal = random.randint(25, 45)
    x_fin_principal = width - random.randint(25, 45)
    
    # Crear la línea principal más larga
    puntos_principales = []
    for x in range(x_inicio_principal, x_fin_principal, 3):
        onda = math.sin((x - x_inicio_principal) * 0.08) * 8
        y = y_linea_principal + onda + random.randint(-2, 2)
        puntos_principales.append((x, y))
    
    # Dibujar línea principal
    if len(puntos_principales) > 1:
        for i in range(len(puntos_principales) - 1):
            draw.line([puntos_principales[i], puntos_principales[i + 1]], fill='black', width=3)

def _crear_firma_estilizada(texto):
    """
    Crea una firma digital estilizada sin usar fuentes tipográficas.
    Convierte cada letra del texto en un trazo manuscrito único.
    """
    # Dimensiones de la imagen
    width = max(400, len(texto) * 40)
    height = 150
    
    # Crear imagen base con fondo blanco
    if Image is None: return None
    image = Image.new('RGB', (width, height), color='white')
    draw = ImageDraw.Draw(image)
    
    # Convertir cada letra del texto en un trazo manuscrito
    colores = ['black', 'gray', 'darkgray']
    
    # Procesar cada letra del nombre
    for i, letra in enumerate(texto):
        if letra.isspace():
            continue
            
        # Posición basada en la letra (más espaciado)
        x_base = 30 + (i * (width - 60) // len(texto))
        y_centro = height // 2
        
        # Crear trazo único basado en la letra
        # Usar el código ASCII de la letra para determinar el estilo
        ascii_val = ord(letra.upper()) if letra.isalpha() else ord('A')
        
        # Determinar la forma del trazo basada en la letra
        if letra.upper() in 'AEIOU':
            # Vocales: líneas curvas y abiertas
            _dibujar_trazo_vocal(draw, x_base, y_centro, ascii_val, colores)
        elif letra.upper() in 'BCDFG':
            # Consonantes duras: líneas fuertes y angulares
            _dibujar_trazo_consonante_dura(draw, x_base, y_centro, ascii_val, colores)
        else:
            # Otras letras: líneas fluidas
            _dibujar_trazo_generico(draw, x_base, y_centro, ascii_val, colores, width, height)
    
    return image

# --- WORKERS: ORGANIZATION ---

def worker_mover_por_coincidencia(root_path, silent_mode=False):
    if not silent_mode: st.info(f"Iniciando movimiento por coincidencia en: {root_path}")
    
    items = os.listdir(root_path)
    files = [f for f in items if os.path.isfile(os.path.join(root_path, f))]
    folders = [d for d in items if os.path.isdir(os.path.join(root_path, d))]
    
    count_moved = 0
    progress_bar = None
    if not silent_mode:
        progress_bar = st.progress(0, text="Analizando...")
    total = len(files)
    
    for i, file in enumerate(files):
        if not silent_mode and i % 10 == 0 and total > 0:
             progress_bar.progress(min(i / total, 1.0), text=f"Procesando {i}/{total}")
             
        file_lower = file.lower()
        target = None
        for folder in folders:
            if folder.lower() in file_lower:
                target = folder
                break
        
        if target:
            src = os.path.join(root_path, file)
            dst = os.path.join(root_path, target, file)
            try:
                shutil.move(src, dst)
                count_moved += 1
            except Exception as e:
                if not silent_mode: st.error(f"Error moviendo {file} a {target}: {e}")
                
    msg = f"Proceso completado. {count_moved} archivos organizados."
    if not silent_mode:
        if progress_bar: progress_bar.progress(1.0, text="Finalizado.")
        st.success(msg)
    return msg

def worker_consolidar_subcarpetas(root_path, silent_mode=False):
    if not silent_mode: st.info(f"Consolidando subcarpetas en: {root_path}")
    
    try:
        main_folders = [d for d in os.listdir(root_path) if os.path.isdir(os.path.join(root_path, d))]
    except Exception as e:
        return f"Error leyendo directorio base: {e}"

    if not main_folders: return "No se encontraron carpetas para procesar."
    
    copiados = 0
    conflictos = 0
    errores = 0
    
    progress_bar = None
    if not silent_mode:
        progress_bar = st.progress(0, text="Consolidando...")
    total = len(main_folders)
    
    for i, folder_name in enumerate(main_folders):
        if not silent_mode and total > 0:
             progress_bar.progress(min(i / total, 1.0), text=f"Procesando carpeta {folder_name}")
        
        main_folder_path = os.path.join(root_path, folder_name)
        
        for sub_root, _, files in os.walk(main_folder_path):
            if sub_root == main_folder_path:
                continue
                
            for file_name in files:
                source_path = os.path.join(sub_root, file_name)
                dest_path = os.path.join(main_folder_path, file_name)

                try:
                    if os.path.exists(dest_path):
                        conflictos += 1
                        continue
                    
                    shutil.copy2(source_path, dest_path)
                    copiados += 1
                except Exception as e:
                    errores += 1
                    
    msg = f"Consolidación completada. Copiados: {copiados}, Conflictos: {conflictos}, Errores: {errores}."
    if not silent_mode:
        if progress_bar: progress_bar.progress(1.0, text="Finalizado.")
        st.success(msg)
    return msg

def worker_firmar_docx_con_imagen_masivo(base_path, docx_filename, signature_filename, silent_mode=False):
    try:
        folders_to_process = [d for d in os.listdir(base_path) if os.path.isdir(os.path.join(base_path, d))]
        if not folders_to_process: return "No se encontraron carpetas para procesar."
        
        procesados = 0
        errores = 0
        
        progress_bar = None
        if not silent_mode:
            progress_bar = st.progress(0, text="Modificando documentos...")
            
        for i, folder_name in enumerate(folders_to_process):
            if not silent_mode:
                progress_bar.progress((i + 1) / len(folders_to_process), text=f"Procesando: {folder_name}")
                
            folder_path = os.path.join(base_path, folder_name)
            docx_path = os.path.join(folder_path, docx_filename)
            
            signature_path = None
            possible_sig_paths = [
                os.path.join(folder_path, signature_filename),
                os.path.join(folder_path, "tipografia", signature_filename)
            ]
            for path in possible_sig_paths:
                if os.path.exists(path):
                    signature_path = path
                    break
            
            if not os.path.exists(docx_path) or not signature_path:
                errores += 1
                continue
                
            # Worker logic for single file
            try:
                doc = Document(docx_path)
                anchor_text = "Firma de Aceptacion"
                signature_p_index = -1

                for idx, p in enumerate(doc.paragraphs):
                    if anchor_text.lower() in p.text.lower():
                        target_index = idx + 1
                        if target_index < len(doc.paragraphs):
                            signature_p_index = target_index
                        break
                
                if signature_p_index != -1:
                    signature_p = doc.paragraphs[signature_p_index]
                    p_element = signature_p._p
                    p_element.clear_content()
                    run = signature_p.add_run()
                    run.add_picture(signature_path, width=Inches(1.5))
                    signature_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    doc.save(docx_path)
                    procesados += 1
                else:
                    errores += 1
            except Exception:
                errores += 1
                
        return f"Proceso finalizado. Modificados: {procesados}, Errores/Omitidos: {errores}"
    except Exception as e:
        return f"Error general: {e}"

def worker_txt_a_json_individual(file_list, silent_mode=False):
    count = 0
    errores = 0
    for file_path in file_list:
        try:
            base, _ = os.path.splitext(file_path)
            new_path = base + ".json"
            os.rename(file_path, new_path)
            count += 1
        except:
            errores += 1
    return f"Renombrados {count} archivos. Errores: {errores}"

def worker_organizar_facturas_feov(root_path, target_path, silent_mode=False):
    if not root_path or not target_path:
        return "Error: Rutas de origen o destino no válidas."

    if not silent_mode: st.info("Iniciando organización de facturas FEOV...")
    
    regex = re.compile(r'FEOV(\d+)', re.IGNORECASE)
    destinos_map = {}
    
    # Paso 1: Mapear destinos
    try:
        list_carpetas_destino = [d for d in os.listdir(target_path) if os.path.isdir(os.path.join(target_path, d))]
    except Exception as e:
        return f"Error leyendo destinos: {e}"

    for nombre_carpeta_destino in list_carpetas_destino:
        ruta_carpeta_destino = os.path.join(target_path, nombre_carpeta_destino)
        try:
            for archivo in os.listdir(ruta_carpeta_destino):
                if archivo.lower().endswith('.pdf'):
                    match = regex.search(archivo)
                    if match:
                        numero_factura = match.group(1)
                        destinos_map[numero_factura] = ruta_carpeta_destino
                        break
        except: pass

    if not destinos_map:
        return "No se encontraron facturas FEOV en las carpetas de destino."

    # Paso 2: Mover archivos
    movidos, errores, conflictos = 0, 0, 0
    
    progress_bar = None
    if not silent_mode: progress_bar = st.progress(0, text="Organizando...")
    
    files_to_move = []
    for root, _, files in os.walk(root_path):
        for f in files:
            files_to_move.append((root, f))
            
    total = len(files_to_move)
    
    for i, (root, file_to_move) in enumerate(files_to_move):
        if not silent_mode and i % 10 == 0 and total > 0:
             progress_bar.progress(min(i / total, 1.0), text=f"Procesando {i}/{total}")

        moved = False
        for numero_factura, ruta_destino_final in destinos_map.items():
            if numero_factura in file_to_move:
                try:
                    ruta_origen_archivo = os.path.join(root, file_to_move)
                    ruta_final_archivo = os.path.join(ruta_destino_final, file_to_move)

                    if os.path.exists(ruta_final_archivo):
                        conflictos += 1
                    else:
                        shutil.move(ruta_origen_archivo, ruta_destino_final)
                        movidos += 1
                    moved = True
                    break 
                except Exception:
                    errores += 1
                    break
    
    if not silent_mode: 
        if progress_bar: progress_bar.progress(1.0, text="Finalizado.")
        
    return f"Organización FEOV finalizada. Movidos: {movidos}, Conflictos: {conflictos}, Errores: {errores}"



# --- WORKERS: UNIFICATION & DIVISION (GROUP 1) ---

def worker_unificar_por_carpeta(carpeta_base, nombre_final_base, silent_mode=False):
    if not carpeta_base or not nombre_final_base:
        return "Faltan argumentos (carpeta o nombre)."
    
    subcarpetas = [os.path.join(carpeta_base, d) for d in os.listdir(carpeta_base) if os.path.isdir(os.path.join(carpeta_base, d))]
    if not subcarpetas: return "No se encontraron subcarpetas."

    progress_bar = None
    if not silent_mode: progress_bar = st.progress(0, text="Unificando PDFs...")
    
    procesados = 0
    errores = 0
    
    for i, carpeta in enumerate(subcarpetas):
        if not silent_mode: progress_bar.progress((i + 1) / len(subcarpetas))
        
        archivos_pdf_a_procesar = []
        for num_pdf in range(1, 11):
            nombre_archivo_buscado = f"{num_pdf}.pdf"
            ruta_archivo_buscado = os.path.join(carpeta, nombre_archivo_buscado)
            if os.path.exists(ruta_archivo_buscado):
                archivos_pdf_a_procesar.append(ruta_archivo_buscado)

        if not archivos_pdf_a_procesar:
            continue
        
        try:
            nombre_final = f"{nombre_final_base}.pdf"
            ruta_salida = os.path.join(carpeta, nombre_final)
            
            doc_final = fitz.open()
            for ruta_pdf in archivos_pdf_a_procesar:
                with fitz.open(ruta_pdf) as doc_origen:
                    doc_final.insert_pdf(doc_origen)
            
            if len(doc_final) > 0:
                doc_final.save(ruta_salida, garbage=4, deflate=True)
                procesados += 1
            doc_final.close()
        except Exception:
            errores += 1
            
    return f"Unificación finalizada. Procesados: {procesados}, Errores: {errores}"

def worker_unificar_imagenes_por_carpeta(carpeta_base, nombre_final_base, tipo_imagen, silent_mode=False):
    # tipo_imagen: "JPG" or "PNG"
    ext_map = {
        "JPG": ['.jpg', '.jpeg'],
        "PNG": ['.png']
    }
    
    subcarpetas = [os.path.join(carpeta_base, d) for d in os.listdir(carpeta_base) if os.path.isdir(os.path.join(carpeta_base, d))]
    if not subcarpetas: return "No se encontraron subcarpetas."

    progress_bar = None
    if not silent_mode: progress_bar = st.progress(0, text=f"Unificando {tipo_imagen}...")
    
    procesados = 0
    errores = 0
    
    for i, carpeta in enumerate(subcarpetas):
        if not silent_mode: progress_bar.progress((i + 1) / len(subcarpetas))
        
        image_files = []
        for file_name in os.listdir(carpeta):
            if any(file_name.lower().endswith(ext) for ext in ext_map.get(tipo_imagen, [])):
                try:
                    num = int(os.path.splitext(file_name)[0])
                    image_files.append((num, os.path.join(carpeta, file_name)))
                except: continue
        
        if not image_files: continue
        image_files.sort()
        
        try:
            pdf_path = os.path.join(carpeta, f"{nombre_final_base}.pdf")
            images_to_convert = []
            
            # Open first image
            first_img = Image.open(image_files[0][1])
            if first_img.mode == 'RGBA': first_img = first_img.convert('RGB')
            
            for _, img_path in image_files[1:]:
                img = Image.open(img_path)
                if img.mode == 'RGBA': img = img.convert('RGB')
                images_to_convert.append(img)
            
            first_img.save(pdf_path, "PDF", resolution=100.0, save_all=True, append_images=images_to_convert)
            procesados += 1
        except:
            errores += 1
            
    return f"Unificación {tipo_imagen} finalizada. Procesados: {procesados}, Errores: {errores}"

def worker_unificar_docx_por_carpeta(carpeta_base, nombre_final_base, silent_mode=False):
    if not HAS_DOCX2PDF: return "Librería docx2pdf no instalada."
    
    subcarpetas = [os.path.join(carpeta_base, d) for d in os.listdir(carpeta_base) if os.path.isdir(os.path.join(carpeta_base, d))]
    if not subcarpetas: return "No se encontraron subcarpetas."

    progress_bar = None
    if not silent_mode: progress_bar = st.progress(0, text="Unificando DOCX...")
    
    procesados = 0
    errores = 0
    
    for i, carpeta in enumerate(subcarpetas):
        if not silent_mode: progress_bar.progress((i + 1) / len(subcarpetas))
        
        archivos_docx = []
        for num in range(1, 11):
            f = os.path.join(carpeta, f"{num}.docx")
            if os.path.exists(f): archivos_docx.append(f)
            
        if not archivos_docx: continue
        
        pdfs_temp = []
        try:
            for docx in archivos_docx:
                temp_pdf = os.path.splitext(docx)[0] + "_temp.pdf"
                try:
                    convert_docx_to_pdf(docx, temp_pdf)
                    if os.path.exists(temp_pdf): pdfs_temp.append(temp_pdf)
                except: pass
            
            if pdfs_temp:
                doc_final = fitz.open()
                for p in pdfs_temp:
                    try:
                        with fitz.open(p) as d: doc_final.insert_pdf(d)
                    except: pass
                
                doc_final.save(os.path.join(carpeta, f"{nombre_final_base}.pdf"))
                doc_final.close()
                procesados += 1
                
                # Cleanup
                for p in pdfs_temp:
                    try: os.remove(p)
                    except: pass
        except:
            errores += 1
            
    return f"Unificación DOCX finalizada. Procesados: {procesados}, Errores: {errores}"

def worker_dividir_pdf_en_paginas(archivos_pdf, silent_mode=False):
    if not archivos_pdf: return "No hay archivos seleccionados."
    
    progress_bar = None
    if not silent_mode: progress_bar = st.progress(0, text="Dividiendo PDFs...")
    
    procesados = 0
    errores = 0
    
    for i, pdf_path in enumerate(archivos_pdf):
        if not silent_mode: progress_bar.progress((i + 1) / len(archivos_pdf))
        try:
            base_name = os.path.splitext(os.path.basename(pdf_path))[0]
            out_dir = os.path.join(os.path.dirname(pdf_path), base_name)
            os.makedirs(out_dir, exist_ok=True)
            
            with fitz.open(pdf_path) as doc:
                for page_num in range(len(doc)):
                    new_doc = fitz.open()
                    new_doc.insert_pdf(doc, from_page=page_num, to_page=page_num)
                    new_doc.save(os.path.join(out_dir, f"{page_num + 1}.pdf"))
                    new_doc.close()
            procesados += 1
        except:
            errores += 1
            
    return f"División finalizada. Procesados: {procesados}, Errores: {errores}"

def worker_dividir_pdfs_masivamente(carpeta_base, silent_mode=False):
    pdfs = []
    for root, _, files in os.walk(carpeta_base):
        for f in files:
            if f.lower().endswith('.pdf'): pdfs.append(os.path.join(root, f))
            
    if not pdfs: return "No se encontraron PDFs."
    
    return worker_dividir_pdf_en_paginas(pdfs, silent_mode)

# --- WORKERS: ORGANIZATION (GROUP 2) ---

def worker_copiar_archivos_mapeo(ruta_origen_base, ruta_destino_base, df_mapeo, col_origen, col_destino, silent_mode=False):
    if df_mapeo is None: return "DataFrame inválido."
    
    progress_bar = None
    if not silent_mode: progress_bar = st.progress(0, text="Copiando por mapeo...")
    
    copiados, errores, conflictos, no_enc = 0, 0, 0, 0
    total = len(df_mapeo)
    
    for idx, row in df_mapeo.iterrows():
        if not silent_mode: progress_bar.progress((idx + 1) / total)
        
        try:
            dir_origen = str(row[col_origen]).strip()
            dir_destino = str(row[col_destino]).strip()
            
            full_src = os.path.join(ruta_origen_base, dir_origen)
            full_dst = os.path.join(ruta_destino_base, dir_destino)
            
            if not os.path.isdir(full_src) or not os.path.isdir(full_dst):
                no_enc += 1
                continue
                
            for f in os.listdir(full_src):
                src_file = os.path.join(full_src, f)
                dst_file = os.path.join(full_dst, f)
                
                if os.path.isfile(src_file):
                    if os.path.exists(dst_file):
                        conflictos += 1
                    else:
                        shutil.copy2(src_file, dst_file)
                        copiados += 1
        except:
            errores += 1
            
    return f"Copia finalizada. Copiados: {copiados}, Conflictos: {conflictos}, No encontrados: {no_enc}, Errores: {errores}"

def worker_copiar_archivos_raiz_mapeo(ruta_origen_raiz, ruta_destino_base, df_mapeo, col_id, col_folder, silent_mode=False):
    if df_mapeo is None: return "DataFrame inválido."
    
    files_in_root = [f for f in os.listdir(ruta_origen_raiz) if os.path.isfile(os.path.join(ruta_origen_raiz, f))]
    
    progress_bar = None
    if not silent_mode: progress_bar = st.progress(0, text="Copiando desde raíz...")
    
    copiados, errores, conflictos, no_enc = 0, 0, 0, 0
    total = len(df_mapeo)
    
    for idx, row in df_mapeo.iterrows():
        if not silent_mode: progress_bar.progress((idx + 1) / total)
        
        try:
            file_id = str(row[col_id]).strip()
            folder_name = str(row[col_folder]).strip()
            
            if not file_id or not folder_name: continue
            
            # Find file (case insensitive partial match)
            found_file = None
            for f in files_in_root:
                if file_id.lower() in f.lower():
                    found_file = f
                    break
            
            if not found_file:
                no_enc += 1
                continue
                
            dst_folder = os.path.join(ruta_destino_base, folder_name)
            os.makedirs(dst_folder, exist_ok=True)
            
            src = os.path.join(ruta_origen_raiz, found_file)
            dst = os.path.join(dst_folder, found_file)
            
            if os.path.exists(dst):
                conflictos += 1
            else:
                shutil.copy2(src, dst)
                copiados += 1
        except:
            errores += 1
            
    return f"Copia raíz finalizada. Copiados: {copiados}, Conflictos: {conflictos}, No encontrados: {no_enc}, Errores: {errores}"

def worker_copiar_archivo_a_subcarpetas(archivo_a_copiar, carpeta_destino_base, silent_mode=False):
    if not archivo_a_copiar or not carpeta_destino_base: return "Argumentos inválidos."
    
    subfolders = [os.path.join(carpeta_destino_base, d) for d in os.listdir(carpeta_destino_base) if os.path.isdir(os.path.join(carpeta_destino_base, d))]
    if not subfolders: return "No hay subcarpetas."
    
    progress_bar = None
    if not silent_mode: progress_bar = st.progress(0, text="Distribuyendo archivo...")
    
    copiados, conflictos, errores = 0, 0, 0
    fname = os.path.basename(archivo_a_copiar)
    
    for i, folder in enumerate(subfolders):
        if not silent_mode: progress_bar.progress((i + 1) / len(subfolders))
        try:
            dst = os.path.join(folder, fname)
            if os.path.exists(dst):
                conflictos += 1
            else:
                shutil.copy2(archivo_a_copiar, dst)
                copiados += 1
        except:
            errores += 1
            
    return f"Distribución finalizada. Copiados: {copiados}, Conflictos: {conflictos}, Errores: {errores}"


    resultados_datos = []
    procesados, errores = 0, 0
    
    progress_bar = None
    if not silent_mode: progress_bar = st.progress(0, text="Leyendo Retefuente...")
    
    for i, ruta_pdf in enumerate(pdf_files):
        if not silent_mode: progress_bar.progress((i + 1) / len(pdf_files))
        nombre_archivo = os.path.basename(ruta_pdf)
        try:
            with fitz.open(ruta_pdf) as doc:
                for num_pagina, page in enumerate(doc, start=1):
                    blocks = page.get_text("blocks")
                    blocks.sort(key=lambda b: b[1]) # Sort vertically
                    
                    label_block = None
                    nit_label_block = None
                    nombre_encontrado = "NO ENCONTRADO"
                    nit_encontrado = "NO ENCONTRADO"
                    
                    # 1. Find key labels
                    for b in blocks:
                        text_clean = " ".join(b[4].split()).upper()
                        if "PRACTICO LA RETENCION" in text_clean:
                            label_block = b
                            break
                            
                    if label_block:
                        lx0, ly0, lx1, ly1 = label_block[:4]
                        for b in blocks:
                            bx0, by0 = b[:2]
                            text_clean = " ".join(b[4].split()).upper()
                            if bx0 > lx0 and abs(by0 - ly0) < 30:
                                if "NIT" in text_clean or "C.C." in text_clean:
                                    nit_label_block = b
                                    break
                                    
                    if not nit_label_block:
                        for b in blocks:
                            text_clean = " ".join(b[4].split()).upper()
                            if "NIT." in text_clean and "C.C." in text_clean:
                                nit_label_block = b
                                break
                                
                    # 2. Extract NAME
                    if label_block:
                        lx0, ly0, lx1, ly1 = label_block[:4]
                        candidates = []
                        for b in blocks:
                            if b == label_block: continue
                            bx0, by0 = b[:2]
                            if by0 > ly0 and abs(bx0 - lx0) < 100:
                                candidates.append(b)
                        candidates.sort(key=lambda b: b[1])
                        
                        for cand in candidates:
                            text_cand = cand[4].strip()
                            upper_cand = text_cand.upper()
                            if not text_cand: continue
                            if "DIRECCION" in upper_cand: break
                            if "NIT" in upper_cand or "C.C." in upper_cand: continue
                            nombre_encontrado = " ".join(text_cand.split())
                            break
                            
                    # 3. Extract NIT
                    if nit_label_block:
                        nx0, ny0 = nit_label_block[:2]
                        nit_candidates = []
                        for b in blocks:
                            if b == nit_label_block: continue
                            bx0, by0 = b[:2]
                            if by0 > ny0 and abs(bx0 - nx0) < 80:
                                nit_candidates.append(b)
                        nit_candidates.sort(key=lambda b: b[1])
                        
                        for cand in nit_candidates:
                            text_cand = cand[4].strip()
                            upper_cand = text_cand.upper()
                            if not text_cand: continue
                            if "CIUDAD" in upper_cand: break
                            nit_encontrado = " ".join(text_cand.split())
                            break
                            
                    # Cleanup
                    if nombre_encontrado != "NO ENCONTRADO":
                        match_mix = re.search(r'^(.*?)(\d{6,}[\d\s]*)$', nombre_encontrado)
                        if match_mix:
                            nombre_encontrado = match_mix.group(1).strip()
                            nit_extraido = match_mix.group(2).replace(" ", "").strip()
                            if nit_encontrado == "NO ENCONTRADO" or not any(c.isdigit() for c in nit_encontrado):
                                nit_encontrado = nit_extraido
                                
                    resultados_datos.append({
                        "Archivo": nombre_archivo,
                        "Página": num_pagina,
                        "RAZON SOCIAL / NOMBRE": nombre_encontrado,
                        "NIT / C.C.": nit_encontrado
                    })
            procesados += 1
        except Exception as e:
            errores += 1
            resultados_datos.append({"Archivo": nombre_archivo, "Error": str(e)})
            
    if not resultados_datos: return "No se extrajeron datos."
    
    try:
        df = pd.DataFrame(resultados_datos)
        df.to_excel(output_path, index=False)
        return f"Reporte guardado en {output_path}. Procesados: {procesados}, Errores: {errores}"
    except Exception as e:
        return f"Error guardando Excel: {e}"

# --- WORKERS: EXCEL/FILES ---

def worker_crear_carpetas_desde_excel(excel_path, sheet_name, col_idx, target_folder, visible_only=False, silent_mode=False):
    if not excel_path or not sheet_name or col_idx is None or not target_folder:
        return "Faltan parámetros."
        
    try:
        nombres_carpetas_raw = []
        if visible_only:
            wb = openpyxl.load_workbook(excel_path, data_only=True)
            ws = wb[sheet_name]
            # openpyxl uses 1-based indexing
            col_1based = col_idx + 1 
            for i in range(2, ws.max_row + 1):
                if not ws.row_dimensions[i].hidden:
                    val = ws.cell(row=i, column=col_1based).value
                    if val: nombres_carpetas_raw.append(str(val))
        else:
            df = pd.read_excel(excel_path, sheet_name=sheet_name)
            nombres_carpetas_raw = df.iloc[:, col_idx].dropna().astype(str).tolist()
            
        if not nombres_carpetas_raw: return "No se encontraron nombres."
        
        creadas, errores = 0, 0
        progress_bar = None
        if not silent_mode: progress_bar = st.progress(0, text="Creando carpetas...")
        
        for i, nombre in enumerate(nombres_carpetas_raw):
            if not silent_mode: progress_bar.progress((i + 1) / len(nombres_carpetas_raw))
            nombre_base = "".join(c for c in nombre if c.isalnum() or c in " _-").rstrip()
            if not nombre_base: continue
            
            ruta_final = os.path.join(target_folder, nombre_base)
            if os.path.exists(ruta_final):
                c = 2
                while os.path.exists(os.path.join(target_folder, f"{nombre_base} ({c})")):
                    c += 1
                ruta_final = os.path.join(target_folder, f"{nombre_base} ({c})")
            
            try:
                os.makedirs(ruta_final)
                creadas += 1
            except: errores += 1
            
        return f"Creadas: {creadas}, Errores: {errores}"
    except Exception as e:
        return f"Error: {e}"

# --- WORKERS: PDF ---

def worker_unificar_pdfs_list(file_list, output_path, sort_method="Nombre", silent_mode=False):
    try:
        if not file_list: return "No hay archivos para unificar."
        
        # Sort files
        if sort_method == "Nombre":
            file_list.sort(key=lambda x: x.name if hasattr(x, 'name') else os.path.basename(x))
        
        doc_final = fitz.open()
        for f in file_list:
            try:
                # Handle both file paths and BytesIO/UploadedFile
                if isinstance(f, str):
                    doc = fitz.open(f)
                else:
                    f.seek(0)
                    doc = fitz.open(stream=f.read(), filetype="pdf")
                
                doc_final.insert_pdf(doc)
                doc.close()
            except Exception as e:
                if not silent_mode: st.warning(f"Omitiendo archivo por error: {e}")
        
        doc_final.save(output_path)
        doc_final.close()
        return f"PDF Unificado creado en: {output_path}"
    except Exception as e:
        return f"Error unificando PDFs: {e}"

def worker_dividir_pdf_paginas(input_pdf, output_folder, silent_mode=False):
    try:
        if isinstance(input_pdf, str):
            doc = fitz.open(input_pdf)
            name_base = os.path.splitext(os.path.basename(input_pdf))[0]
        else:
            input_pdf.seek(0)
            doc = fitz.open(stream=input_pdf.read(), filetype="pdf")
            name_base = os.path.splitext(input_pdf.name)[0]
            
        os.makedirs(output_folder, exist_ok=True)
        
        for i in range(len(doc)):
            page = doc.load_page(i)
            pix = page.get_pixmap()
            out_name = f"{name_base}_pag_{i+1}.pdf"
            
            # Create new PDF for single page
            new_doc = fitz.open()
            new_doc.insert_pdf(doc, from_page=i, to_page=i)
            new_doc.save(os.path.join(output_folder, out_name))
            new_doc.close()
            
        return f"PDF dividido en {len(doc)} páginas en {output_folder}"
    except Exception as e:
        return f"Error dividiendo PDF: {e}"

def worker_unificar_imagenes_pdf(folder_path, output_name="Unificado.pdf", silent_mode=False):
    try:
        images = [
            os.path.join(folder_path, f) 
            for f in sorted(os.listdir(folder_path), key=natural_sort_key) 
            if f.lower().endswith(('.jpg', '.jpeg', '.png', '.bmp'))
        ]
        
        if not images:
            return "No se encontraron imágenes en la carpeta."
            
        pdf_path = os.path.join(folder_path, output_name)
        img_list = []
        first_img = None
        
        for img_path in images:
            try:
                img = Image.open(img_path).convert('RGB')
                if first_img is None:
                    first_img = img
                else:
                    img_list.append(img)
            except Exception:
                pass
                
        if first_img:
            first_img.save(pdf_path, save_all=True, append_images=img_list)
            return f"PDF creado exitosamente: {output_name} ({len(images)} imágenes)"
        else:
            return "No se pudieron procesar las imágenes."
    except Exception as e:
        return f"Error: {e}"

# --- WORKERS: PDF EXTENDED ---

def worker_unificar_por_carpeta(carpeta_base, nombre_final_base, silent_mode=False):
    if not carpeta_base or not os.path.isdir(carpeta_base): return "Carpeta base inválida."
    
    subcarpetas = [os.path.join(carpeta_base, d) for d in os.listdir(carpeta_base) if os.path.isdir(os.path.join(carpeta_base, d))]
    if not subcarpetas: return "No se encontraron subcarpetas."

    log = []
    pdfs_creados = 0
    
    for carpeta in subcarpetas:
        nombre_subcarpeta = os.path.basename(carpeta)
        
        archivos_pdf_a_procesar = []
        for num_pdf in range(1, 11):
            nombre_archivo_buscado = f"{num_pdf}.pdf"
            ruta_archivo_buscado = os.path.join(carpeta, nombre_archivo_buscado)
            if os.path.exists(ruta_archivo_buscado):
                archivos_pdf_a_procesar.append(ruta_archivo_buscado)

        if not archivos_pdf_a_procesar:
            continue
        
        nombre_final = f"{nombre_final_base}.pdf"
        ruta_salida = os.path.join(carpeta, nombre_final)
        
        try:
            doc_final = fitz.open()
            for ruta_pdf in archivos_pdf_a_procesar:
                with fitz.open(ruta_pdf) as doc_origen:
                    for page_origen in doc_origen:
                        pix = page_origen.get_pixmap(dpi=300, colorspace=fitz.csGRAY)
                        pagina_nueva = doc_final.new_page(width=pix.width, height=pix.height)
                        pagina_nueva.insert_image(pagina_nueva.rect, pixmap=pix)
            
            if len(doc_final) > 0:
                doc_final.save(ruta_salida, garbage=4, deflate=True)
                pdfs_creados += 1
            doc_final.close()
        except Exception as e:
            log.append(f"Error en {nombre_subcarpeta}: {e}")

    return f"Proceso finalizado. {pdfs_creados} PDFs creados." + (" Errores: " + "; ".join(log) if log else "")

def worker_unificar_imagenes_por_carpeta_rec(carpeta_base, nombre_final_base, tipo_imagen="JPG", silent_mode=False):
    # tipo_imagen: "JPG" or "PNG"
    if not carpeta_base or not os.path.isdir(carpeta_base): return "Carpeta base inválida."
    
    ext_map = {
        "JPG": ['.jpg', '.jpeg'],
        "PNG": ['.png']
    }
    exts = ext_map.get(tipo_imagen, ['.jpg'])

    subcarpetas = [os.path.join(carpeta_base, d) for d in os.listdir(carpeta_base) if os.path.isdir(os.path.join(carpeta_base, d))]
    if not subcarpetas: return "No se encontraron subcarpetas."
    
    pdfs_creados = 0
    log = []

    for carpeta in subcarpetas:
        nombre_subcarpeta = os.path.basename(carpeta)
        
        archivos_img_a_procesar = []
        # Buscar 1.jpg, 2.jpg, ...
        for num_img in range(1, 11):
            ruta_encontrada = None
            for ext in exts:
                nombre_archivo = f"{num_img}{ext}"
                ruta_archivo = os.path.join(carpeta, nombre_archivo)
                if os.path.exists(ruta_archivo):
                    ruta_encontrada = ruta_archivo
                    break
            if ruta_encontrada:
                archivos_img_a_procesar.append(ruta_encontrada)
        
        if not archivos_img_a_procesar:
            continue

        try:
            lista_imagenes_procesadas = []
            for ruta_img in archivos_img_a_procesar:
                img = Image.open(ruta_img)
                if tipo_imagen == "PNG":
                     if img.mode in ('RGBA', 'LA'):
                        background = Image.new('RGB', img.size, (255, 255, 255))
                        background.paste(img, mask=img.split()[-1]) # Use alpha channel as mask
                        img = background
                     elif img.mode != 'RGB':
                        img = img.convert('RGB')
                else:
                    img = img.convert('L') # Grayscale for JPG as per original code

                lista_imagenes_procesadas.append(img)
            
            if lista_imagenes_procesadas:
                nombre_pdf = f"{nombre_final_base}.pdf"
                ruta_salida = os.path.join(carpeta, nombre_pdf)
                lista_imagenes_procesadas[0].save(
                    ruta_salida, 
                    save_all=True, 
                    append_images=lista_imagenes_procesadas[1:], 
                    resolution=300.0
                )
                pdfs_creados += 1
        except Exception as e:
            log.append(f"Error en {nombre_subcarpeta}: {e}")

    return f"Proceso finalizado. {pdfs_creados} PDFs creados." + (" Errores: " + "; ".join(log) if log else "")

def worker_unificar_docx_por_carpeta(carpeta_base, nombre_final_base, silent_mode=False):
    if not HAS_DOCX2PDF: return "docx2pdf no está instalado."
    if not carpeta_base or not os.path.isdir(carpeta_base): return "Carpeta base inválida."
    
    subcarpetas = [os.path.join(carpeta_base, d) for d in os.listdir(carpeta_base) if os.path.isdir(os.path.join(carpeta_base, d))]
    pdfs_creados = 0
    log = []

    for carpeta in subcarpetas:
        nombre_subcarpeta = os.path.basename(carpeta)
        archivos_docx_a_procesar = []
        for num_doc in range(1, 11):
            nombre_archivo = f"{num_doc}.docx"
            ruta_archivo = os.path.join(carpeta, nombre_archivo)
            if os.path.exists(ruta_archivo):
                archivos_docx_a_procesar.append(ruta_archivo)
        
        if not archivos_docx_a_procesar: continue

        pdfs_temporales = []
        try:
            for ruta_docx in archivos_docx_a_procesar:
                nombre_temp_pdf = os.path.splitext(os.path.basename(ruta_docx))[0] + "_temp.pdf"
                ruta_temp_pdf = os.path.join(carpeta, nombre_temp_pdf)
                try:
                    convert_docx_to_pdf(ruta_docx, ruta_temp_pdf)
                    if os.path.exists(ruta_temp_pdf):
                        pdfs_temporales.append(ruta_temp_pdf)
                except: pass
            
            if pdfs_temporales:
                nombre_pdf_final = f"{nombre_final_base}.pdf"
                ruta_salida = os.path.join(carpeta, nombre_pdf_final)
                
                doc_final = fitz.open()
                for pdf_temp in pdfs_temporales:
                    try:
                        with fitz.open(pdf_temp) as doc_temp:
                            doc_final.insert_pdf(doc_temp)
                    except: pass
                
                doc_final.save(ruta_salida)
                doc_final.close()
                pdfs_creados += 1
                
                for pdf_temp in pdfs_temporales:
                    try: os.remove(pdf_temp)
                    except: pass
        except Exception as e:
            log.append(f"Error en {nombre_subcarpeta}: {e}")

    return f"Proceso finalizado. {pdfs_creados} PDFs creados."

def worker_dividir_pdfs_masivamente(carpeta_base, silent_mode=False):
    if not carpeta_base or not os.path.isdir(carpeta_base): return "Carpeta inválida."
    
    pdfs_a_procesar = []
    for root, _, files in os.walk(carpeta_base):
        for file in files:
            if file.lower().endswith('.pdf'):
                pdfs_a_procesar.append(os.path.join(root, file))
    
    if not pdfs_a_procesar: return "No se encontraron PDFs."
    
    count = 0
    for ruta_pdf_original in pdfs_a_procesar:
        try:
            nombre_base_original = os.path.splitext(os.path.basename(ruta_pdf_original))[0]
            directorio_origen = os.path.dirname(ruta_pdf_original)
            ruta_carpeta_salida = os.path.join(directorio_origen, nombre_base_original)
            os.makedirs(ruta_carpeta_salida, exist_ok=True)
            
            with fitz.open(ruta_pdf_original) as doc_origen:
                if len(doc_origen) == 0: continue
                for i in range(len(doc_origen)):
                    doc_nuevo = fitz.open()
                    doc_nuevo.insert_pdf(doc_origen, from_page=i, to_page=i)
                    doc_nuevo.save(os.path.join(ruta_carpeta_salida, f"{i+1}.pdf"))
                    doc_nuevo.close()
            count += 1
        except: pass
        
    return f"Divididos {count} PDFs masivamente."

def worker_pdf_a_escala_grises(files, replace_original=False, silent_mode=False):
    # files: list of file paths
    count = 0
    for entrada in files:
        try:
            ruta_salida = entrada
            ruta_temporal = entrada + "._tmp_grayscale" if replace_original else None
            
            with fitz.open(entrada) as doc_origen:
                if doc_origen.is_encrypted: continue
                doc_final = fitz.open()
                for page in doc_origen:
                    pix = page.get_pixmap(dpi=300, colorspace=fitz.csGRAY)
                    pagina_nueva = doc_final.new_page(width=pix.width, height=pix.height)
                    pagina_nueva.insert_image(pagina_nueva.rect, pixmap=pix)
                
                target = ruta_temporal if replace_original else ruta_salida.replace(".pdf", "_gray.pdf")
                doc_final.save(target, garbage=4, deflate=True)
                doc_final.close()
            
            if replace_original:
                os.replace(ruta_temporal, entrada)
            count += 1
        except Exception:
            if replace_original and ruta_temporal and os.path.exists(ruta_temporal):
                os.remove(ruta_temporal)
                
    return f"Convertidos {count} archivos a escala de grises."

# --- WORKERS: RIPS ---

def worker_json_a_xlsx_ind(file_obj, silent_mode=False):
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
                        item["consecutivoUsuario"] = u_info["consecutivo"]
                        all_services[sheet_name].append(item)

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

def worker_consolidar_json_xlsx(folder_path, silent_mode=False):
    try:
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
        
        for fname in json_files:
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
                            item["archivo_origen"] = fname
                            item["consecutivoUsuario"] = u.get("consecutivo")
                            master_services[s_name].append(item)
                            
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

def worker_xlsx_a_json_ind(file_obj, silent_mode=False):
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

def worker_rips_excel_to_json_original(file_obj, silent_mode=False):
    """
    Worker para convertir Excel a JSON con estructura RIPS (Original).
    Espera hojas: Consultas, Procedimientos, OtrosServicios.
    """
    try:
        if hasattr(file_obj, 'seek'):
            file_obj.seek(0)
        xls = pd.ExcelFile(file_obj)
        
        usuarios_dict = {}
        
        def procesar_hoja(nombre_hoja, clave_servicio):
            if nombre_hoja in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=nombre_hoja)
                df = df.astype(object).where(pd.notnull(df), None)
                
                for _, row in df.iterrows():
                    td = str(row.get("tipo_documento_usuario", ""))
                    doc = str(row.get("documento_usuario", ""))
                    user_key = (td, doc)
                    
                    if user_key not in usuarios_dict:
                        usuarios_dict[user_key] = {
                            "tipoDocumentoIdentificacion": row.get("tipo_documento_usuario"),
                            "numDocumentoIdentificacion": row.get("documento_usuario"),
                            "tipoUsuario": row.get("tipo_usuario"),
                            "fechaNacimiento": row.get("fecha_nacimiento"), 
                            "codSexo": row.get("sexo"),
                            "codPaisResidencia": row.get("pais_residencia"),
                            "municipio_residencia": row.get("municipio_residencia"),
                            "codZonaTerritorialResidencia": row.get("zona_residencia"),
                            "incapacidad": row.get("incapacidad"),
                            "consecutivo": row.get("consecutivo_usuario"),
                            "codPaisOrigen": row.get("pais_origen"),
                            "servicios": {
                                "consultas": [],
                                "procedimientos": [],
                                "otrosServicios": []
                            }
                        }
                    
                    servicio_data = row.to_dict()
                    keys_to_remove = [
                        "tipo_documento_usuario", "documento_usuario", "tipo_usuario", 
                        "fecha_nacimiento", "sexo", "pais_residencia", "municipio_residencia", 
                        "zona_residencia", "incapacidad", "consecutivo_usuario", "pais_origen"
                    ]
                    for k in keys_to_remove:
                        servicio_data.pop(k, None)
                        
                    if any(v is not None for v in servicio_data.values()):
                        usuarios_dict[user_key]["servicios"][clave_servicio].append(servicio_data)

        procesar_hoja("Consultas", "consultas")
        procesar_hoja("Procedimientos", "procedimientos")
        procesar_hoja("OtrosServicios", "otrosServicios")
        
        resultado_final = {
            "usuarios": list(usuarios_dict.values())
        }
        
        return json.dumps(resultado_final, ensure_ascii=False, indent=4), None
        
    except Exception as e:
        return None, str(e)

def worker_rips_json_to_excel_original(file_obj, silent_mode=False):
    """
    Worker para convertir JSON a Excel con estructura RIPS (Original).
    Genera hojas: Consultas, Procedimientos, OtrosServicios.
    """
    try:
        if hasattr(file_obj, 'seek'):
            file_obj.seek(0)
            data = json.load(file_obj)
        else:
            with open(file_obj, "r", encoding="utf-8") as f:
                data = json.load(f)

        consultas = []
        procedimientos = []
        otros_servicios = []

        for usuario in data.get("usuarios", []):
            base_info = {
                "tipo_documento_usuario": usuario.get("tipoDocumentoIdentificacion"),
                "documento_usuario": usuario.get("numDocumentoIdentificacion"),
                "tipo_usuario": usuario.get("tipoUsuario"),
                "fecha_nacimiento": usuario.get("fechaNacimiento"),
                "sexo": usuario.get("codSexo"),
                "pais_residencia": usuario.get("codPaisResidencia"),
                "municipio_residencia": usuario.get("codMunicipioResidencia"),
                "zona_residencia": usuario.get("codZonaTerritorialResidencia"),
                "incapacidad": usuario.get("incapacidad"),
                "consecutivo_usuario": usuario.get("consecutivo"),
                "pais_origen": usuario.get("codPaisOrigen")
            }

            servicios = usuario.get("servicios", {})

            for consulta in servicios.get("consultas", []):
                consultas.append({**base_info, **consulta})

            for procedimiento in servicios.get("procedimientos", []):
                procedimientos.append({**base_info, **procedimiento})

            for otro in servicios.get("otrosServicios", []):
                otros_servicios.append({**base_info, **otro})

        output = io.BytesIO()
        # Use xlsxwriter or openpyxl if installed. Default to None (let pandas choose)
        with pd.ExcelWriter(output) as writer:
            if consultas:
                pd.DataFrame(consultas).to_excel(writer, sheet_name="Consultas", index=False)
            if procedimientos:
                pd.DataFrame(procedimientos).to_excel(writer, sheet_name="Procedimientos", index=False)
            if otros_servicios:
                pd.DataFrame(otros_servicios).to_excel(writer, sheet_name="OtrosServicios", index=False)
            
            if not consultas and not procedimientos and not otros_servicios:
                 pd.DataFrame().to_excel(writer, sheet_name="Vacio", index=False)
                 
        return output.getvalue(), None
    except Exception as e:
        return None, str(e)

def worker_desconsolidar_xlsx_json(file_obj, dest_folder, silent_mode=False):
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
        for fname, header in headers_by_file.items():
            header.pop("archivo_origen", None)
            final = header
            users = []
            for u in users_by_file.get(fname, []):
                u.pop("archivo_origen", None)
                u_cons = u.get("consecutivo")
                u["servicios"] = {k: [] for k in service_map.values()}
                if fname in services_by_file:
                    for s_key, items in services_by_file[fname].items():
                        for item in items:
                            if item.get("consecutivoUsuario") == u_cons:
                                i_clean = item.copy()
                                i_clean.pop("archivo_origen", None)
                                i_clean.pop("consecutivoUsuario", None)
                                u["servicios"][s_key].append(i_clean)
                users.append(u)
            final["usuarios"] = users
            with open(os.path.join(dest_folder, fname), 'w', encoding='utf-8') as f:
                json.dump(final, f, ensure_ascii=False, indent=4)
            count += 1
        return True, f"Desconsolidados {count} archivos."
    except Exception as e:
        return False, str(e)

# --- WORKERS: EXCEL / RENAMING ---

def worker_aplicar_renombrado_excel(excel_path, folder_path, silent_mode=False):
    try:
        df = pd.read_excel(excel_path)
        if "Nombre Actual" not in df.columns or "Nombre Nuevo" not in df.columns:
            return "Excel debe tener columnas 'Nombre Actual' y 'Nombre Nuevo'"
        count = 0
        for _, row in df.iterrows():
            curr = str(row["Nombre Actual"]).strip()
            new = str(row["Nombre Nuevo"]).strip()
            curr_path = os.path.join(folder_path, curr)
            if os.path.exists(curr_path):
                if "." not in new:
                    _, ext = os.path.splitext(curr)
                    new += ext
                try:
                    os.rename(curr_path, os.path.join(folder_path, new))
                    count += 1
                except: pass
        return f"Renombrados {count} archivos."
    except Exception as e:
        return f"Error: {e}"

def worker_anadir_sufijo_excel(excel_path, sheet_name, col_name, col_suffix, folder_path, use_filter=False, silent_mode=False):
    try:
        if isinstance(excel_path, bytes):
            excel_path = io.BytesIO(excel_path)
        excel_path.seek(0)

        data_rows = []
        if use_filter:
            import openpyxl
            wb = openpyxl.load_workbook(excel_path, data_only=True)
            if sheet_name not in wb.sheetnames: return "Hoja no encontrada"
            ws = wb[sheet_name]
            header = [cell.value for cell in ws[1]]
            try:
                # Buscar columnas por nombre exacto
                idx_name = header.index(col_name)
                idx_suffix = header.index(col_suffix)
            except: return f"Columnas '{col_name}' o '{col_suffix}' no encontradas en el encabezado."
            
            for row in ws.iter_rows(min_row=2):
                if not ws.row_dimensions[row[0].row].hidden:
                    # Safe value retrieval
                    v_name = row[idx_name].value
                    v_suffix = row[idx_suffix].value
                    
                    val_name = str(v_name).strip() if v_name is not None else ""
                    val_suffix = str(v_suffix).strip() if v_suffix is not None else ""
                    
                    if val_name and val_suffix:
                        data_rows.append((val_name, val_suffix))
        else:
            df = pd.read_excel(excel_path, sheet_name=sheet_name)
            if col_name not in df.columns or col_suffix not in df.columns:
                return f"Columnas '{col_name}' o '{col_suffix}' no encontradas en Excel."
            for _, row in df.iterrows():
                if pd.notna(row[col_name]) and pd.notna(row[col_suffix]):
                    val_name = str(row[col_name]).strip()
                    val_suffix = str(row[col_suffix]).strip()
                    if val_name and val_suffix:
                        data_rows.append((val_name, val_suffix))

        if not data_rows:
            return "No se encontraron datos válidos (nombres/sufijos no vacíos) para procesar."

        if not os.path.isdir(folder_path):
            return f"La carpeta objetivo no existe: {folder_path}"

        # Listar archivos UNA VEZ
        try:
            existing_files = [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]
        except Exception as e:
            return f"Error leyendo carpeta objetivo: {e}"

        count = 0
        errores = 0
        
        progress_bar = None
        if not silent_mode:
            progress_bar = st.progress(0, text="Añadiendo sufijos...")

        total = len(data_rows)
        processed_files = set() 

        for i, (name, suffix) in enumerate(data_rows):
            if not silent_mode and progress_bar:
                progress_bar.progress((i + 1) / total)
            
            # Buscar coincidencias EXACTAS (ignorando mayúsculas/minúsculas)
            # Match: nombre archivo sin extension == name OR nombre archivo completo == name
            matches = []
            for f in existing_files:
                if f.lower() == name.lower() or os.path.splitext(f)[0].lower() == name.lower():
                    matches.append(f)
            
            for f in matches:
                if f in processed_files: continue
                
                base, ext = os.path.splitext(f)
                
                # Evitar doble sufijo si ya lo tiene
                if f.endswith(f"_{suffix}{ext}"):
                    continue
                    
                new_name = f"{base}_{suffix}{ext}"
                old_path = os.path.join(folder_path, f)
                new_path = os.path.join(folder_path, new_name)
                
                try:
                    if old_path != new_path and not os.path.exists(new_path):
                        os.rename(old_path, new_path)
                        count += 1
                        processed_files.add(f)
                    elif os.path.exists(new_path):
                        errores += 1 
                except Exception:
                    errores += 1

        if not silent_mode and progress_bar: progress_bar.empty()
        return f"Sufijos añadidos a {count} archivos. Errores/Omitidos: {errores}."
    except Exception as e:
        return f"Error crítico: {e}"

# --- WORKERS: DOCX / FIRMAS ---

def worker_unificar_docx_carpeta(folder_path, output_name="Unificado.docx", silent_mode=False):
    try:
        files = sorted([f for f in os.listdir(folder_path) if f.lower().endswith('.docx')], key=natural_sort_key)
        if not files: return "No hay archivos .docx"
        master = Document(os.path.join(folder_path, files[0]))
        master.add_page_break()
        for f in files[1:]:
            doc = Document(os.path.join(folder_path, f))
            for element in doc.element.body:
                master.element.body.append(element)
            master.add_page_break()
        master.save(os.path.join(folder_path, output_name))
        return f"Unificados {len(files)} DOCX."
    except Exception as e:
        return f"Error: {e}"

def worker_crear_firma_nombre(nombre, documento, output_folder, silent_mode=False):
    try:
        img = Image.new('RGB', (400, 100), color='white')
        d = ImageDraw.Draw(img)
        try: font = ImageFont.truetype("arial.ttf", 24)
        except: font = ImageFont.load_default()
        d.text((10, 10), f"Firmado por: {nombre}", fill='black', font=font)
        d.text((10, 50), f"Doc: {documento}", fill='black', font=font)
        out_path = os.path.join(output_folder, f"Firma_{documento}.png")
        img.save(out_path)
        return out_path
    except: return None

def worker_firmar_docx(docx_path, firma_path, output_path, silent_mode=False):
    try:
        doc = Document(docx_path)
        doc.add_picture(firma_path, width=Pt(150))
        doc.save(output_path)
        return True
    except: return False

def worker_modificar_docx_excel(uploaded_file, sheet_name, col_folder, col_val, root_path, mode, silent_mode=False):
    try:
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
        modificados = 0
        docx_pattern = re.compile(r'CRC_.*_FEOV.*\.docx$', re.IGNORECASE)
        col_folder_idx = ord(col_folder.upper()) - ord('A')
        col_val_idx = ord(col_val.upper()) - ord('A')
        for index, row in df.iterrows():
            folder_name = str(row.iloc[col_folder_idx]).strip()
            new_val = str(row.iloc[col_val_idx]).strip()
            if not folder_name or not new_val: continue
            target_dir = os.path.join(root_path, folder_name)
            if not os.path.isdir(target_dir): continue
            target_docx = next((os.path.join(target_dir, f) for f in os.listdir(target_dir) if docx_pattern.match(f)), None)
            if target_docx:
                try:
                    doc = Document(target_docx)
                    modified = False
                    keyword = "REGIMEN:" if mode == "Regimen" else f"{mode}:"
                    for p in doc.paragraphs:
                        if keyword in p.text.upper():
                            p.text = re.sub(rf'({keyword})\s*.*', r'\1 ' + new_val, p.text, flags=re.IGNORECASE)
                            modified = True
                            break
                    if modified:
                        doc.save(target_docx)
                        modificados += 1
                except: pass
        return f"Modificados {modificados} documentos."
    except Exception as e:
        return f"Error: {e}"

def _create_column_map_from_headers(df):
    required_map = {
        'folder': 'Nombre Carpeta', 'date': 'Ciudad y Fecha', 'full_name': 'Nombre Completo',
        'doc_type': 'Tipo Documento', 'doc_num': 'Numero Documento', 'service': 'Servicio',
        'eps': 'EPS', 'tipo_servicio': 'Tipo Servicio', 'regimen': 'Regimen',
        'categoria': 'Categoria', 'cuota': 'Valor Cuota Moderadora', 'auth': 'Numero Autorizacion',
        'fecha_atencion': 'Fecha y Hora Atencion', 'fecha_fin': 'Fecha Finalizacion'
    }
    excel_headers = df.columns
    missing_cols = [header for header in required_map.values() if header not in excel_headers]
    if missing_cols: return None, missing_cols
    return required_map, []

def worker_modificar_docx_completo(uploaded_file, sheet_name, root_path, use_filter=False, silent_mode=False):
    try:
        if isinstance(uploaded_file, bytes): uploaded_file = io.BytesIO(uploaded_file)
        uploaded_file.seek(0)

        df = None
        if use_filter:
            import openpyxl
            wb = openpyxl.load_workbook(uploaded_file, data_only=True)
            if sheet_name not in wb.sheetnames: return "Hoja no encontrada."
            ws = wb[sheet_name]
            
            data = []
            # Read headers from first row
            headers = [cell.value for cell in ws[1]]
            
            # Read visible rows
            for row in ws.iter_rows(min_row=2):
                if not ws.row_dimensions[row[0].row].hidden:
                    data.append([cell.value for cell in row])
            
            if data:
                df = pd.DataFrame(data, columns=headers)
            else:
                return "No hay datos visibles para procesar."
        else:
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name)

        column_map, missing = _create_column_map_from_headers(df)
        if not column_map:
            return f"Error: Faltan columnas en el Excel: {', '.join(missing)}"
            
        modificados = 0
        errores = 0
        
        progress_bar = None
        if not silent_mode:
            progress_bar = st.progress(0, text="Modificando DOCX...")
            
        for index, row in df.iterrows():
            if not silent_mode:
                progress_bar.progress((index + 1) / len(df), text=f"Procesando fila {index+1}")

            try:
                datos = {key: str(row[col_name]).strip() if pd.notna(row[col_name]) else "" for key, col_name in column_map.items()}
                folder_name = datos.get('folder')
                if not folder_name: continue
                
                target_dir = os.path.join(root_path, folder_name)
                if not os.path.isdir(target_dir): 
                    errores += 1
                    continue
                    
                target_docx = next((os.path.join(target_dir, f) for f in os.listdir(target_dir) if f.lower().endswith('.docx') and 'plantilla' in f.lower()), None)
                if not target_docx:
                    errores += 1
                    continue
                
                doc = Document(target_docx)
                for p in doc.paragraphs:
                    if "Santiago de Cali, " in p.text: 
                        p.text = f"Santiago de Cali,  {datos['date']}"
                    
                    if "Yo " in p.text and "identificado con" in p.text:
                        p.text = f"Yo {datos['full_name']} identificado con {datos['doc_type']}, Numero {datos['doc_num']} en calidad de paciente, doy fé y acepto el servicio de {datos['service']} brindado por la IPS OPORTUNIDAD DE VIDA S.A.S"
                    
                    replacements = {
                        "EPS:": datos['eps'], "TIPO SERVICIO:": datos['tipo_servicio'],
                        "REGIMEN:": datos['regimen'], "CATEGORIA:": datos['categoria'],
                        "VALOR CUOTA MODERADORA:": datos['cuota'], "AUTORIZACION:": datos['auth'],
                        "Fecha de Atención:": datos['fecha_atencion'], "Fecha de Finalización:": datos['fecha_fin']
                    }
                    for key, val in replacements.items():
                        if key in p.text:
                            p.text = re.sub(rf'({key})\s*.*', r'\1 ' + val, p.text, count=1)
                
                sig_idx = -1
                for i, p in enumerate(doc.paragraphs):
                    if "FIRMA DE ACEPTACION" in p.text.upper():
                        sig_idx = i
                        break
                if sig_idx != -1 and sig_idx + 2 < len(doc.paragraphs):
                    doc.paragraphs[sig_idx + 2].text = datos['full_name'].upper()
                
                doc.save(target_docx)
                modificados += 1
            except Exception:
                errores += 1
        
        if not silent_mode: progress_bar.empty()
        return f"Proceso finalizado. Modificados: {modificados}, Errores/No encontrados: {errores}"
    except Exception as e:
        return f"Error general: {e}"

def worker_crear_carpetas_excel_avanzado(uploaded_file, sheet_name, col_name, base_path, use_filter=False, silent_mode=False):
    try:
        if not os.path.isdir(base_path):
            return "La ruta base seleccionada no es válida."

        if isinstance(uploaded_file, bytes):
            uploaded_file = io.BytesIO(uploaded_file)
        if hasattr(uploaded_file, 'seek'):
            uploaded_file.seek(0)
            
        if use_filter:
            wb = openpyxl.load_workbook(uploaded_file, data_only=True)
            if sheet_name not in wb.sheetnames:
                return f"La hoja '{sheet_name}' no existe."
            ws = wb[sheet_name]
            
            # Find column index (1-based)
            header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
            col_idx = -1
            for idx, val in enumerate(header_row):
                if str(val).strip() == col_name:
                    col_idx = idx + 1
                    break
            
            if col_idx == -1:
                return f"No se encontró la columna '{col_name}' en la primera fila."
                
            nombres_carpetas = []
            for row in ws.iter_rows(min_row=2):
                if not ws.row_dimensions[row[0].row].hidden:
                    val = row[col_idx-1].value
                    if val: nombres_carpetas.append(str(val))
        else:
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
            if col_name not in df.columns:
                return f"No se encontró la columna '{col_name}' en el Excel."
            nombres_carpetas = df[col_name].dropna().astype(str).tolist()

        creadas = 0
        errores = 0
        
        progress_bar = None
        if not silent_mode:
            progress_bar = st.progress(0, text="Creando carpetas...")
            
        for i, nombre in enumerate(nombres_carpetas):
            if not silent_mode:
                progress_bar.progress((i + 1) / len(nombres_carpetas), text=f"Procesando: {nombre}")
                
            nombre_base = "".join(c for c in nombre if c.isalnum() or c in " _-").rstrip()
            if not nombre_base: continue
            
            ruta_final = os.path.join(base_path, nombre_base)
            
            if os.path.exists(ruta_final):
                contador = 2
                nombre_consecutivo = f"{nombre_base} ({contador})"
                ruta_final = os.path.join(base_path, nombre_consecutivo)
                while os.path.exists(ruta_final):
                    contador += 1
                    nombre_consecutivo = f"{nombre_base} ({contador})"
                    ruta_final = os.path.join(base_path, nombre_consecutivo)
            
            try:
                os.makedirs(ruta_final, exist_ok=True)
                creadas += 1
            except Exception:
                errores += 1
                
        if not silent_mode: progress_bar.empty()
        return f"Proceso finalizado. Carpetas creadas: {creadas}, Errores: {errores}"
        
    except Exception as e:
        return f"Error crítico: {e}"

def worker_mover_archivos_por_coincidencia(base_path, silent_mode=False):
    if not base_path or not os.path.isdir(base_path):
        return "Ruta base inválida."
        
    try:
        elementos = os.listdir(base_path)
        archivos = [os.path.join(base_path, e) for e in elementos if os.path.isfile(os.path.join(base_path, e))]
        carpetas = [os.path.join(base_path, e) for e in elementos if os.path.isdir(os.path.join(base_path, e))]
    except Exception as e:
        return f"Error leyendo directorio: {e}"
        
    if not archivos or not carpetas:
        return "No hay archivos o carpetas suficientes para procesar."
        
    movidos, errores = 0, 0
    progress_bar = None
    if not silent_mode:
        progress_bar = st.progress(0, text="Moviendo archivos...")
        
    for i, ruta_archivo in enumerate(archivos):
        nombre_archivo = os.path.basename(ruta_archivo)
        if not silent_mode:
            progress_bar.progress((i + 1) / len(archivos), text=f"Verificando: {nombre_archivo}")
            
        for ruta_carpeta in carpetas:
            nombre_carpeta = os.path.basename(ruta_carpeta)
            if nombre_carpeta.lower() in nombre_archivo.lower():
                try:
                    shutil.move(ruta_archivo, ruta_carpeta)
                    movidos += 1
                    break
                except Exception:
                    errores += 1
                    break
                    
    if not silent_mode: progress_bar.empty()
    return f"Proceso finalizado. Movidos: {movidos}, Errores: {errores}"

def worker_anadir_sufijo_desde_excel(uploaded_file, sheet_name, col_folder, col_suffix, base_path, use_filter=False, silent_mode=False):
    try:
        data_rows = []
        if isinstance(uploaded_file, bytes):
            uploaded_file = io.BytesIO(uploaded_file)
        if hasattr(uploaded_file, 'seek'):
            uploaded_file.seek(0)

        if use_filter:
            wb = openpyxl.load_workbook(uploaded_file, data_only=True)
            if sheet_name not in wb.sheetnames: return "Hoja no encontrada"
            ws = wb[sheet_name]
            
            header = [cell.value for cell in ws[1]]
            try:
                idx_folder = header.index(col_folder)
                idx_suffix = header.index(col_suffix)
            except ValueError: return "Columnas no encontradas en encabezado"
            
            for row in ws.iter_rows(min_row=2):
                if not ws.row_dimensions[row[0].row].hidden:
                    val_folder = row[idx_folder].value
                    val_suffix = row[idx_suffix].value
                    if val_folder and val_suffix:
                        data_rows.append((str(val_folder).strip(), str(val_suffix).strip()))
        else:
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
            if col_folder not in df.columns or col_suffix not in df.columns:
                return f"Columnas '{col_folder}' o '{col_suffix}' no encontradas."
            for _, row in df.iterrows():
                if pd.notna(row[col_folder]) and pd.notna(row[col_suffix]):
                    data_rows.append((str(row[col_folder]).strip(), str(row[col_suffix]).strip()))

        carpetas_procesadas, archivos_renombrados, errores = 0, 0, 0
        
        progress_bar = None
        if not silent_mode:
            progress_bar = st.progress(0, text="Añadiendo sufijos...")
            
        total = len(data_rows)
        for index, (folder_name, suffix) in enumerate(data_rows):
            if not silent_mode:
                progress_bar.progress((index + 1) / total)
            
            target_dir = os.path.join(base_path, folder_name)
            if os.path.isdir(target_dir):
                carpetas_procesadas += 1
                for fname in os.listdir(target_dir):
                    fpath = os.path.join(target_dir, fname)
                    if os.path.isfile(fpath):
                        name, ext = os.path.splitext(fname)
                        new_name = f"{name}{suffix}{ext}"
                        new_path = os.path.join(target_dir, new_name)
                        
                        if fpath == new_path: continue
                        if os.path.exists(new_path):
                            errores += 1
                            continue
                            
                        try:
                            os.rename(fpath, new_path)
                            archivos_renombrados += 1
                        except Exception:
                            errores += 1
            else:
                errores += 1 # Folder not found
                
        if not silent_mode: progress_bar.empty()
        return f"Finalizado. Carpetas: {carpetas_procesadas}, Renombrados: {archivos_renombrados}, Errores: {errores}"
    except Exception as e:
        return f"Error crítico: {e}"

def worker_autorizacion_docx_desde_excel(uploaded_file, sheet_name, col_folder, col_auth, base_path, use_filter=False, silent_mode=False):
    try:
        data_rows = []
        if isinstance(uploaded_file, bytes):
            uploaded_file = io.BytesIO(uploaded_file)
        if hasattr(uploaded_file, 'seek'):
            uploaded_file.seek(0)

        if use_filter:
            wb = openpyxl.load_workbook(uploaded_file, data_only=True)
            if sheet_name not in wb.sheetnames: return "Hoja no encontrada"
            ws = wb[sheet_name]
            
            header = [cell.value for cell in ws[1]]
            try:
                idx_folder = header.index(col_folder)
                idx_auth = header.index(col_auth)
            except ValueError: return "Columnas no encontradas en encabezado"
            
            for row in ws.iter_rows(min_row=2):
                if not ws.row_dimensions[row[0].row].hidden:
                    val_folder = row[idx_folder].value
                    val_auth = row[idx_auth].value
                    if val_folder and val_auth:
                        data_rows.append((str(val_folder).strip(), str(val_auth).strip()))
        else:
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
            if col_folder not in df.columns or col_auth not in df.columns:
                 return f"Columnas '{col_folder}' o '{col_auth}' no encontradas."
            for _, row in df.iterrows():
                if pd.notna(row[col_folder]) and pd.notna(row[col_auth]):
                    data_rows.append((str(row[col_folder]).strip(), str(row[col_auth]).strip()))

        modificados, errores = 0, 0
        docx_pattern = re.compile(r'CRC_.*_FEOV.*\.docx$', re.IGNORECASE)
        
        progress_bar = None
        if not silent_mode:
            progress_bar = st.progress(0, text="Actualizando autorizaciones...")
            
        total = len(data_rows)
        for index, (folder_name, new_auth) in enumerate(data_rows):
            if not silent_mode:
                progress_bar.progress((index + 1) / total)
            
            if not folder_name or not new_auth: continue
            
            target_dir = os.path.join(base_path, folder_name)
            if not os.path.isdir(target_dir):
                errores += 1
                continue
                
            target_docx = next((os.path.join(target_dir, f) for f in os.listdir(target_dir) if docx_pattern.match(f)), None)
            if not target_docx:
                errores += 1
                continue
                
            try:
                doc = Document(target_docx)
                changed = False
                for p in doc.paragraphs:
                    if "AUTORIZACION:" in p.text.upper():
                        p.text = re.sub(r'(AUTORIZACION:)\s*.*', r'\1 ' + new_auth, p.text, flags=re.IGNORECASE)
                        changed = True
                
                if changed:
                    doc.save(target_docx)
                    modificados += 1
            except Exception:
                errores += 1
                
        if not silent_mode: progress_bar.empty()
        return f"Finalizado. Modificados: {modificados}, Errores/No encontrados: {errores}"
    except Exception as e:
        return f"Error crítico: {e}"



    try:
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
        count = 0
        os.makedirs(root_path, exist_ok=True)
        try: font = ImageFont.truetype(ttf_path, size) if ttf_path else ImageFont.load_default()
        except: font = ImageFont.load_default()
        for _, row in df.iterrows():
            name = str(row[col_full_name]).strip()
            if name:
                img = Image.new('RGB', (600, 150), color='white')
                d = ImageDraw.Draw(img)
                d.text((20, 50), name, fill='black', font=font)
                safe_name = "".join(c for c in name if c.isalnum() or c in " _-")
                out_path = os.path.join(root_path, f"Firma_{safe_name}.png")
                img.save(out_path)
                count += 1
        return f"Generadas {count} firmas."
    except Exception as e:
        return f"Error: {e}"



# --- WORKERS: MISSING FILE OPS ---

def worker_crear_carpetas_desde_excel(excel_path, sheet_name, col_name, output_folder, filter_hidden=False, silent_mode=False):
    try:
        if filter_hidden:
            wb = openpyxl.load_workbook(excel_path, data_only=True)
            ws = wb[sheet_name]
            df_col_idx = -1
            for idx, cell in enumerate(ws[1]):
                if cell.value == col_name:
                    df_col_idx = idx
                    break
            if df_col_idx == -1: return "Columna no encontrada."
            names = []
            for row in ws.iter_rows(min_row=2):
                if not ws.row_dimensions[row[0].row].hidden:
                    val = row[df_col_idx].value
                    if val: names.append(str(val))
        else:
            df = pd.read_excel(excel_path, sheet_name=sheet_name)
            if col_name not in df.columns: return "Columna no encontrada."
            names = df[col_name].dropna().astype(str).tolist()
            
        count = 0
        errores = 0
        for name in names:
            safe_name = "".join(c for c in name if c.isalnum() or c in " _-").strip()
            if not safe_name: continue
            target = os.path.join(output_folder, safe_name)
            if os.path.exists(target):
                i = 2
                while os.path.exists(f"{target} ({i})"): i += 1
                target = f"{target} ({i})"
            try:
                os.makedirs(target)
                count += 1
            except: errores += 1
        return f"Creadas {count} carpetas. Errores: {errores}"
    except Exception as e:
        return f"Error: {e}"

def worker_copiar_archivos_desde_mapeo(excel_path, sheet_name, col_src, col_dst, base_src, base_dst, silent_mode=False):
    try:
        df = pd.read_excel(excel_path, sheet_name=sheet_name)
        if col_src not in df.columns or col_dst not in df.columns: return "Columnas no encontradas."
        
        copiados = 0
        conflictos = 0
        errores = 0
        
        for _, row in df.iterrows():
            src_folder = str(row[col_src]).strip()
            dst_folder = str(row[col_dst]).strip()
            if not src_folder or not dst_folder: continue
            
            full_src = os.path.join(base_src, src_folder)
            full_dst = os.path.join(base_dst, dst_folder)
            
            if not os.path.isdir(full_src) or not os.path.isdir(full_dst): continue
            
            for f in os.listdir(full_src):
                f_src = os.path.join(full_src, f)
                if os.path.isfile(f_src):
                    f_dst = os.path.join(full_dst, f)
                    if os.path.exists(f_dst):
                        conflictos += 1
                        continue
                    try:
                        shutil.copy2(f_src, f_dst)
                        copiados += 1
                    except: errores += 1
        return f"Copiados {copiados}. Conflictos: {conflictos}. Errores: {errores}."
    except Exception as e:
        return f"Error: {e}"





def worker_txt_a_json_masivo(folder_path, silent_mode=False):
    try:
        count = 0
        for f in os.listdir(folder_path):
            if f.lower().endswith('.txt'):
                try:
                    base = os.path.splitext(f)[0]
                    src = os.path.join(folder_path, f)
                    dst = os.path.join(folder_path, base + ".json")
                    if not os.path.exists(dst):
                        os.rename(src, dst)
                        count += 1
                except: pass
        return f"Renombrados {count} TXT a JSON."
    except Exception as e:
        return f"Error: {e}"


def worker_analisis_carpetas(root_path, silent_mode=False):
    if not silent_mode: st.info(f"Analizando: {root_path}")
    data = []
    for root, dirs, files in os.walk(root_path):
        folder_name = os.path.basename(root)
        count = len(files)
        size = sum(os.path.getsize(os.path.join(root, f)) for f in files)
        data.append({"Carpeta": folder_name, "Archivos": count, "Peso (KB)": round(size/1024, 2), "Ruta": root})
    
    if not data:
        if not silent_mode: st.warning("Carpeta vacía o sin acceso.")
        return None

    df = pd.DataFrame(data)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    
    return {
        "files": [{
            "name": f"Reporte_Carpetas_{int(time.time())}.xlsx",
            "data": output.getvalue(),
            "label": "Descargar Reporte"
        }],
        "message": f"Reporte generado con {len(data)} carpetas analizadas."
    }

# --- WORKERS: CONVERSION & AI ---

def _pdf_a_docx(input_path, output_path):
    try:
        cv = Converter(input_path)
        cv.convert(output_path)
        cv.close()
    except Exception as e:
        print(f"Error pdf2docx: {e}")

def _jpg_a_pdf(input_path, output_path):
    img = Image.open(input_path)
    if img.mode == 'RGBA':
        img = img.convert('RGB')
    res = st.session_state.app_config.get("image_resolution", 100.0) if 'app_config' in st.session_state else 100.0
    img.save(output_path, "PDF", resolution=res)

def _docx_a_pdf(input_path, output_path):
    """
    Convierte DOCX a PDF usando automatización COM directa (win32com) si está disponible.
    Fallback a docx2pdf.
    """
    success = False
    if HAS_WIN32COM:
        try:
            import pythoncom
            pythoncom.CoInitialize()
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            
            in_file_abs = os.path.abspath(input_path)
            out_file_abs = os.path.abspath(output_path)
            
            doc = word.Documents.Open(in_file_abs)
            doc.SaveAs(out_file_abs, FileFormat=17) # wdFormatPDF = 17
            doc.Close(False)
            word.Quit()
            success = True
        except Exception as e:
            print(f"Error win32com DOCX->PDF: {e}. Intentando fallback...")
    
    if not success and HAS_DOCX2PDF:
        try:
            # docx2pdf sometimes needs pythoncom init too if running in thread
            try:
                import pythoncom
                pythoncom.CoInitialize() 
            except: pass
            convert_docx_to_pdf(input_path, output_path)
        except Exception as e:
            print(f"Error docx2pdf: {e}")

def _pdf_a_jpg(input_path, output_base):
    doc = fitz.open(input_path)
    for i, page in enumerate(doc):
        pix = page.get_pixmap()
        out = f"{output_base}_p{i+1}.jpg" if len(doc) > 1 else f"{output_base}.jpg"
        pix.save(out)
    doc.close()

def _png_a_jpg(input_path, output_path):
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

def _pdf_escala_grises(input_path, output_path):
    doc = fitz.open(input_path)
    doc_final = fitz.open()
    dpi = st.session_state.app_config.get("pdf_dpi", 600) if 'app_config' in st.session_state else 600
    matrix_scale = dpi / 72.0
    mat = fitz.Matrix(matrix_scale, matrix_scale)
    for page in doc:
        pix = page.get_pixmap(matrix=mat, colorspace=fitz.csGRAY)
        new_page = doc_final.new_page(width=page.rect.width, height=page.rect.height)
        new_page.insert_image(new_page.rect, pixmap=pix)
    doc.close()
    compression = st.session_state.app_config.get("pdf_compression", 4) if 'app_config' in st.session_state else 4
    doc_final.save(output_path, garbage=compression, deflate=True)
    doc_final.close()

def worker_convertir_archivo(file_path, tipo, output_folder=None, silent_mode=False):
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
            out_base = os.path.join(folder, name_no_ext)
            _pdf_a_jpg(file_path, out_base)
        elif tipo == "PNG2JPG":
            out = os.path.join(folder, f"{name_no_ext}.jpg")
            _png_a_jpg(file_path, out)
        elif tipo == "TXT2JSON":
            out = os.path.join(folder, f"{name_no_ext}.json")
            _txt_a_json(file_path, out)
        elif tipo == "PDF_GRAY":
            temp_out = os.path.join(folder, f"{name_no_ext}_temp_gray.pdf")
            _pdf_escala_grises(file_path, temp_out)
            if os.path.exists(temp_out):
                try:
                    os.replace(temp_out, file_path)
                except OSError:
                    time.sleep(0.5)
                    os.remove(file_path)
                    os.rename(temp_out, file_path)
        return True, "Conversión exitosa"
    except Exception as e:
        return False, str(e)

def worker_convertir_masivo(folder_path, tipo, silent_mode=False):
    if not folder_path or not os.path.exists(folder_path):
        return 0, "Carpeta no encontrada"
    count = 0
    files_to_process = []
    for r, d, f in os.walk(folder_path):
        for file in f:
            files_to_process.append(os.path.join(r, file))
    total = len(files_to_process)
    if total == 0:
        return 0, "Carpeta vacía"
    progress_bar = None
    if not silent_mode:
        progress_bar = st.progress(0, text="Convirtiendo...")
    for i, full_path in enumerate(files_to_process):
        if not silent_mode and i % 5 == 0: 
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
        elif tipo == "PDF_GRAY" and f_lower.endswith(".pdf"): process = True
        if process:
            ok, msg = worker_convertir_archivo(full_path, tipo, silent_mode=True)
            if ok: count += 1
            else: 
                if not silent_mode: print(f"Error convirtiendo {f}: {msg}")
    if not silent_mode:
        progress_bar.progress(1.0, text="Finalizado.")
    return count, f"Procesados {count} archivos."

def worker_consultar_gemini(prompt, file_context=None, silent_mode=False):
    api_key = st.session_state.app_config.get("gemini_api_key") if 'app_config' in st.session_state else None
    if not api_key: return "⚠️ Configura tu API Key de Google Gemini."
    try:
        genai.configure(api_key=api_key.strip())
        model_name = st.session_state.app_config.get("gemini_model", "gemini-1.5-flash") if 'app_config' in st.session_state else "gemini-1.5-flash"
        model = genai.GenerativeModel(model_name)
        full_prompt = f"Contexto:\n{file_context}\n\n{prompt}" if file_context else prompt
        response = model.generate_content(full_prompt)
        return response.text
    except Exception as e:
        return f"Error Gemini: {e}"

# --- WORKERS: MAPEO / OTROS ---

def worker_copiar_raiz_mapeo(uploaded_file, sheet_name, col_id, col_dst, path_src_base, path_dst_base, silent_mode=False):
    try:
        if isinstance(uploaded_file, bytes):
            uploaded_file = io.BytesIO(uploaded_file)
        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
        count_files = 0
        files_in_root = {f.lower(): f for f in os.listdir(path_src_base) if os.path.isfile(os.path.join(path_src_base, f))}
        progress_bar = None
        if not silent_mode:
            progress_bar = st.progress(0, text="Copiando...")
        total_rows = len(df)
        for idx, row in df.iterrows():
            if not silent_mode and idx % 10 == 0 and total_rows > 0:
                progress_bar.progress(min(idx / total_rows, 1.0), text=f"Procesando {idx}/{total_rows}")
            id_val = str(row[col_id]).strip().lower()
            dst_folder_name = str(row[col_dst]).strip()
            if not id_val or not dst_folder_name: continue
            for f_lower, f_real in files_in_root.items():
                if id_val in f_lower:
                    src = os.path.join(path_src_base, f_real)
                    dst_folder = os.path.join(path_dst_base, dst_folder_name)
                    if not os.path.exists(dst_folder):
                        try: os.makedirs(dst_folder)
                        except: pass
                    try:
                        shutil.copy2(src, os.path.join(dst_folder, f_real))
                        count_files += 1
                    except Exception: pass
        msg = f"Copia completada. {count_files} archivos copiados."
        if not silent_mode:
            if progress_bar: progress_bar.progress(1.0, text="Finalizado.")
            st.success(msg)
        return msg
    except Exception as e:
        return f"Error: {e}"

def worker_exportar_renombrado(search_results, silent_mode=False):
    if not search_results: return None
    data = []
    for item in search_results:
        path = item.get("Ruta completa", "")
        if path:
            data.append({"Ruta actual": path, "Nuevo nombre": os.path.basename(path)})
    return pd.DataFrame(data)

def worker_renombrar_mapeo_excel(uploaded_file, sheet_name, col_src, col_dst, use_filter, root_path=None, silent_mode=False):
    try:
        if isinstance(uploaded_file, bytes):
            uploaded_file = io.BytesIO(uploaded_file)
        if hasattr(uploaded_file, 'seek'):
            uploaded_file.seek(0)
        data_rows = []
        if use_filter:
            import openpyxl
            wb = openpyxl.load_workbook(uploaded_file, data_only=True)
            if sheet_name not in wb.sheetnames: return "Hoja no encontrada."
            ws = wb[sheet_name]
            header = [cell.value for cell in ws[1]]
            try:
                idx_src = header.index(col_src)
                idx_dst = header.index(col_dst)
            except: return "Columnas no encontradas."
            for row in ws.iter_rows(min_row=2):
                if not ws.row_dimensions[row[0].row].hidden:
                    val_src = row[idx_src].value
                    val_dst = row[idx_dst].value
                    if val_src and val_dst:
                        data_rows.append((str(val_src).strip(), str(val_dst).strip()))
        else:
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
            if col_src not in df.columns or col_dst not in df.columns: return "Columnas no encontradas."
            for _, row in df.iterrows():
                if pd.notna(row[col_src]) and pd.notna(row[col_dst]):
                    data_rows.append((str(row[col_src]).strip(), str(row[col_dst]).strip()))
        count = 0
        progress_bar = None
        if not silent_mode: progress_bar = st.progress(0, text="Renombrando...")
        for i, (src_name, dst_name) in enumerate(data_rows):
            if not silent_mode: progress_bar.progress(min((i+1)/len(data_rows), 1.0))
            src_path = os.path.join(root_path, src_name)
            dst_path = os.path.join(root_path, dst_name)
            if os.path.exists(src_path) and src_path != dst_path:
                try:
                    os.rename(src_path, dst_path)
                    count += 1
                except: pass
        if not silent_mode: st.success(f"Renombrados {count} archivos.")
        return f"Renombrados {count} archivos."
    except Exception as e: return f"Error: {e}"



def worker_copiar_mapeo_subcarpetas(uploaded_file, sheet_name, col_src, col_dst, path_src_base, path_dst_base, use_filter=False, silent_mode=False):
    try:
        if isinstance(uploaded_file, bytes): uploaded_file = io.BytesIO(uploaded_file)
        uploaded_file.seek(0)
        
        df = None
        if use_filter:
            import openpyxl
            wb = openpyxl.load_workbook(uploaded_file, data_only=True)
            if sheet_name not in wb.sheetnames: return "Hoja no encontrada."
            ws = wb[sheet_name]
            
            data = []
            headers = [cell.value for cell in ws[1]]
            
            for row in ws.iter_rows(min_row=2):
                if not ws.row_dimensions[row[0].row].hidden:
                    data.append([cell.value for cell in row])
            
            if data:
                df = pd.DataFrame(data, columns=headers)
            else:
                return "No hay datos visibles."
        else:
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name)

        if col_src not in df.columns or col_dst not in df.columns:
            return f"Columnas no encontradas: {col_src}, {col_dst}"

        copiados = 0
        for _, row in df.iterrows():
            src_n = str(row[col_src]).strip()
            dst_n = str(row[col_dst]).strip()
            if not src_n or not dst_n: continue
            src_full = os.path.join(path_src_base, src_n)
            dst_full = os.path.join(path_dst_base, dst_n)
            
            if os.path.isdir(src_full):
                os.makedirs(dst_full, exist_ok=True)
                for f in os.listdir(src_full):
                    src_f = os.path.join(src_full, f)
                    dst_f = os.path.join(dst_full, f)
                    if os.path.isfile(src_f) and not os.path.exists(dst_f):
                        try:
                            shutil.copy2(src_f, dst_f)
                            copiados += 1
                        except: pass
        return f"Copiados: {copiados}"
    except Exception as e: return f"Error: {e}"

# --- WORKERS: FIRMA DIGITAL & CONSOLIDACION ---

def _crear_firma_estilizada(texto):
    """
    Crea una firma digital estilizada sin usar fuentes tipográficas.
    Convierte cada letra del texto en un trazo manuscrito único.
    """
    width = max(400, len(texto) * 40)
    height = 150
    image = Image.new('RGB', (width, height), color='white')
    draw = ImageDraw.Draw(image)
    colores = ['black', 'gray', 'darkgray']

    def _dibujar_trazo_vocal(draw, x_base, y_centro, ascii_val, colores):
        color = random.choice(colores)
        grosor = random.randint(2, 3)
        altura_arco = 20 + (ascii_val % 15)
        puntos = []
        for i in range(20):
            angulo = (i / 19.0) * math.pi
            x = x_base + i * 2
            y = y_centro - math.sin(angulo) * altura_arco + random.randint(-2, 2)
            puntos.append((x, y))
        for i in range(len(puntos) - 1):
            draw.line([puntos[i], puntos[i + 1]], fill=color, width=grosor)

    def _dibujar_trazo_consonante_dura(draw, x_base, y_centro, ascii_val, colores):
        color = random.choice(colores)
        grosor = random.randint(3, 4)
        x = x_base
        y = y_centro + random.randint(-10, 10)
        draw.line([(x, y), (x + 15, y - 20)], fill=color, width=grosor)
        draw.line([(x + 15, y - 20), (x + 30, y - 15)], fill=color, width=grosor)
        draw.line([(x + 30, y - 15), (x + 40, y + 10)], fill=color, width=grosor)

    def _dibujar_trazo_generico(draw, x_base, y_centro, ascii_val, colores):
        color = random.choice(colores)
        grosor = random.randint(2, 3)
        puntos = []
        for i in range(30):
            x = x_base + i * 1.5
            onda = math.sin((x - x_base) * 0.2 + ascii_val * 0.1) * 15
            y = y_centro + onda + random.randint(-3, 3)
            puntos.append((x, y))
        for i in range(len(puntos) - 1):
            draw.line([puntos[i], puntos[i + 1]], fill=color, width=grosor)
        for _ in range(random.randint(2, 4)):
            punto_x = random.randint(int(x_base), int(x_base + 40))
            punto_y = int(y_centro + random.randint(-5, 5))
            draw.ellipse([punto_x - 1, punto_y - 1, punto_x + 1, punto_y + 1], fill=color)

    for i, letra in enumerate(texto):
        if letra.isspace(): continue
        x_base = 30 + (i * (width - 60) // len(texto))
        y_centro = height // 2
        ascii_val = ord(letra.upper()) if letra.isalpha() else ord('A')
        if letra.upper() in 'AEIOU':
            _dibujar_trazo_vocal(draw, x_base, y_centro, ascii_val, colores)
        elif letra.upper() in 'BCDFG':
            _dibujar_trazo_consonante_dura(draw, x_base, y_centro, ascii_val, colores)
        else:
            _dibujar_trazo_generico(draw, x_base, y_centro, ascii_val, colores)
            
    return image

def worker_crear_firma_digital(base_path, font_path, font_size, silent_mode=False):
    if not base_path or not font_path: return "Error: Rutas inválidas."
    try:
        font = ImageFont.truetype(font_path, font_size)
    except Exception as e:
        return f"Error cargando fuente: {e}"
        
    try:
        folders = [d for d in os.listdir(base_path) if os.path.isdir(os.path.join(base_path, d))]
    except Exception as e:
        return f"Error leyendo carpetas: {e}"
        
    if not folders: return "No hay carpetas."
    
    creadas, errores = 0, 0
    progress_bar = None
    if not silent_mode: progress_bar = st.progress(0, text="Creando firmas...")
    
    for i, folder_name in enumerate(folders):
        if not silent_mode: progress_bar.progress((i + 1) / len(folders))
        try:
            text_to_draw = folder_name
            temp_img = Image.new('RGB', (1, 1))
            draw_temp = ImageDraw.Draw(temp_img)
            bbox = draw_temp.textbbox((0, 0), text_to_draw, font=font)
            text_width = bbox[2] - bbox[0]
            text_height = bbox[3] - bbox[1]
            padding = 20
            img_width = text_width + (2 * padding)
            img_height = text_height + (2 * padding)
            
            image = Image.new('RGB', (img_width, img_height), color='white')
            draw = ImageDraw.Draw(image)
            draw.text((padding, padding), text_to_draw, font=font, fill='black')
            
            tipografia_folder = os.path.join(base_path, folder_name, "tipografia")
            os.makedirs(tipografia_folder, exist_ok=True)
            image.save(os.path.join(tipografia_folder, "firma.jpg"), 'JPEG', quality=95)
            creadas += 1
        except:
            errores += 1
            
    if not silent_mode: progress_bar.empty()
    return f"Firmas creadas: {creadas}. Errores: {errores}."

def worker_consolidar_archivos_subcarpetas(base_path, silent_mode=False):
    if not base_path or not os.path.isdir(base_path): return "Ruta inválida."
    
    try:
        main_folders = [d for d in os.listdir(base_path) if os.path.isdir(os.path.join(base_path, d))]
    except Exception as e: return f"Error leyendo base: {e}"
    
    if not main_folders: return "No hay carpetas."
    
    copiados, conflictos, errores = 0, 0, 0
    progress_bar = None
    if not silent_mode: progress_bar = st.progress(0, text="Consolidando...")
    
    for i, folder_name in enumerate(main_folders):
        if not silent_mode: progress_bar.progress((i + 1) / len(main_folders))
        main_folder_path = os.path.join(base_path, folder_name)
        
        for sub_root, _, files in os.walk(main_folder_path):
            if sub_root == main_folder_path: continue
            for file_name in files:
                src = os.path.join(sub_root, file_name)
                dst = os.path.join(main_folder_path, file_name)
                try:
                    if os.path.exists(dst):
                        conflictos += 1
                        continue
                    shutil.copy2(src, dst)
                    copiados += 1
                except: errores += 1
                
    if not silent_mode: progress_bar.empty()
    return f"Copiados: {copiados}. Conflictos: {conflictos}. Errores: {errores}."

def worker_copiar_archivos_desde_raiz_mapeo(uploaded_file, sheet_name, col_id, col_folder, root_src, root_dst, use_filter=False, silent_mode=False):
    try:
        if isinstance(uploaded_file, bytes): uploaded_file = io.BytesIO(uploaded_file)
        uploaded_file.seek(0)
        
        df = None
        if use_filter:
            import openpyxl
            wb = openpyxl.load_workbook(uploaded_file, data_only=True)
            if sheet_name not in wb.sheetnames: return "Hoja no encontrada."
            ws = wb[sheet_name]
            
            data = []
            headers = [cell.value for cell in ws[1]]
            
            for row in ws.iter_rows(min_row=2):
                if not ws.row_dimensions[row[0].row].hidden:
                    data.append([cell.value for cell in row])
            
            if data:
                df = pd.DataFrame(data, columns=headers)
            else:
                return "No hay datos visibles."
        else:
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
        
        archivos_origen = [f for f in os.listdir(root_src) if os.path.isfile(os.path.join(root_src, f))]
        copiados, no_encontrados, conflictos, errores = 0, 0, 0, 0
        
        progress_bar = None
        if not silent_mode: progress_bar = st.progress(0, text="Copiando...")
        
        for i, row in df.iterrows():
            if not silent_mode: progress_bar.progress((i + 1) / len(df))
            
            id_val = str(row[col_id]).strip()
            folder_val = str(row[col_folder]).strip()
            if not id_val or not folder_val: continue
            
            found_file = None
            for f in archivos_origen:
                if id_val.lower() in f.lower():
                    found_file = f
                    break
            
            if not found_file:
                no_encontrados += 1
                continue
                
            dst_dir = os.path.join(root_dst, folder_val)
            os.makedirs(dst_dir, exist_ok=True)
            dst_file = os.path.join(dst_dir, found_file)
            
            if os.path.exists(dst_file):
                conflictos += 1
                continue
                
            try:
                shutil.copy2(os.path.join(root_src, found_file), dst_file)
                copiados += 1
            except: errores += 1
            
        if not silent_mode: progress_bar.empty()
        return f"Copiados: {copiados}. No encontrados: {no_encontrados}. Conflictos: {conflictos}. Errores: {errores}."
    except Exception as e: return f"Error: {e}"

def worker_copiar_archivo_a_subcarpetas(file_path, dest_base_path, silent_mode=False):
    if not file_path or not dest_base_path: return "Rutas inválidas."
    
    try:
        subcarpetas = [os.path.join(dest_base_path, d) for d in os.listdir(dest_base_path) if os.path.isdir(os.path.join(dest_base_path, d))]
    except Exception as e: return f"Error leyendo destinos: {e}"
    
    if not subcarpetas: return "No hay subcarpetas."
    
    copiados, conflictos, errores = 0, 0, 0
    fname = os.path.basename(file_path)
    
    progress_bar = None
    if not silent_mode: progress_bar = st.progress(0, text=f"Copiando {fname}...")
    
    for i, sub in enumerate(subcarpetas):
        if not silent_mode: progress_bar.progress((i + 1) / len(subcarpetas))
        dst = os.path.join(sub, fname)
        if os.path.exists(dst):
            conflictos += 1
            continue
        try:
            shutil.copy2(file_path, dst)
            copiados += 1
        except: errores += 1
        
    if not silent_mode: progress_bar.empty()
    return f"Copiados: {copiados}. Conflictos: {conflictos}. Errores: {errores}."

def worker_descargar_historias_hospitalizacion_ovida(uploaded_file, sheet_name, col_map, save_path, silent_mode=False):
    # Requires Selenium
    try:
        from selenium import webdriver
        from selenium.webdriver.chrome.service import Service
        from webdriver_manager.chrome import ChromeDriverManager
        import base64
    except ImportError: return "Falta Selenium/WebDriverManager."

    try:
        if isinstance(uploaded_file, bytes): uploaded_file = io.BytesIO(uploaded_file)
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
    except Exception as e: return f"Error Excel: {e}"
    
    driver = None
    try:
        options = webdriver.ChromeOptions()
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        driver.get("https://ovidazs.siesacloud.com/ZeusSalud/ips/iniciando.php")
        
        # Simple wait loop for login (simplified from original for modularity)
        # In a real scenario, we might want a way to signal user readiness via UI, but here we'll wait or rely on user interaction before submitting if possible.
        # However, since this is a background worker, we can't easily ask user "OK" in the middle unless we break it up.
        # For now, we'll assume the user logs in within a generous timeout or we just wait for the URL change.
        
        timeout = 300
        start = time.time()
        logged_in = False
        while time.time() - start < timeout:
            if "App/Vistas" in driver.current_url:
                logged_in = True
                break
            time.sleep(2)
            
        if not logged_in:
            driver.quit()
            return "No se detectó inicio de sesión."
            
        descargados, errores, conflictos = 0, 0, 0
        progress_bar = None
        if not silent_mode: progress_bar = st.progress(0, text="Descargando hospitalización...")
        
        for i, row in df.iterrows():
            if not silent_mode: progress_bar.progress((i + 1) / len(df))
            try:
                estudio = str(int(row[col_map['estudio']])).strip()
                ingreso = pd.to_datetime(row[col_map['ingreso']]).strftime('%Y/%m/%d')
                egreso = pd.to_datetime(row[col_map['egreso']]).strftime('%Y/%m/%d')
                carpeta = str(row[col_map['carpeta']]).strip()
                
                base_url = "https://ovidazs.siesacloud.com/ZeusSalud/Reportes/Cliente//html/reporte_historia_general.php"
                params = {
                    'estudio': estudio, 'fecha_inicio': ingreso, 'fecha_fin': egreso,
                    'verHC': 1, 'verEvo': 1, 'verPar': 1, 'ImprimirOrdenamiento': 1,
                    'ImprimirSolOrdenesExt': 1, 'ImprimirGraficasHC': 1,
                    'ImprimirFormatos': 1, 'ImprimirRegistroAdmon': 1,
                    'ImprimirNotasEnfermeria': 1
                }
                full_url = f"{base_url}?{urllib.parse.urlencode(params)}"
                
                dest_dir = os.path.join(save_path, carpeta)
                os.makedirs(dest_dir, exist_ok=True)
                dest_file = os.path.join(dest_dir, f"HC_{estudio}.pdf")
                
                if os.path.exists(dest_file):
                    conflictos += 1
                    continue
                    
                driver.get(full_url)
                time.sleep(2)
                pdf_b64 = driver.execute_cdp_cmd("Page.printToPDF", {"landscape": False, "printBackground": True})
                with open(dest_file, 'wb') as f:
                    f.write(base64.b64decode(pdf_b64['data']))
                descargados += 1
            except: errores += 1
            
        if not silent_mode: progress_bar.empty()
        return f"Descargados: {descargados}. Errores: {errores}. Conflictos: {conflictos}."
    except Exception as e:
        return f"Error crítico: {e}"
    finally:
        if driver: driver.quit()

# --- WORKERS: CDO VALIDATORS ---

def worker_registraduria_masiva(df, col_cedula, headless=True, update_progress=None, silent_mode=False):
    try:
        validator = ValidatorRegistraduria(headless=headless)
    except Exception as e:
        if not silent_mode: st.error(f"Error initializing ValidatorRegistraduria: {e}")
        return None, f"Error: {e}"
    cb = update_progress if update_progress else lambda c, t, **kwargs: None
    try:
        df_results = validator.process_massive(df, col_cedula, progress_callback=cb)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_results.to_excel(writer, index=False)
        return output.getvalue(), f"Procesados {len(df_results)} registros."
    except Exception as e:
        return None, f"Error processing massive Registraduria: {e}"

def worker_adres_api_masiva(df, col_cedula, col_tipo_doc=None, default_tipo_doc="CC", update_progress=None, silent_mode=False):
    try:
        validator = ValidatorAdres()
    except Exception as e:
        return None, f"Error initializing ValidatorAdres: {e}"
    cb = update_progress if update_progress else lambda c, t, **kwargs: None
    try:
        df_results = validator.process_massive(df, col_cedula, tipo_doc_col=col_tipo_doc, default_tipo_doc=default_tipo_doc, progress_callback=cb)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_results.to_excel(writer, index=False)
        return output.getvalue(), f"Procesados {len(df_results)} registros."
    except Exception as e:
        return None, f"Error processing massive ADRES API: {e}"

def worker_adres_web_massive(df, col_cedula, col_tipo_doc=None, default_tipo_doc="CC", update_progress=None, silent_mode=False):
    try:
        validator = ValidatorAdresWeb(headless=False)
    except Exception as e:
        return None, f"Error initializing ValidatorAdresWeb: {e}"
    cb = update_progress if update_progress else lambda c, t, **kwargs: None
    try:
        df_results = validator.process_massive(df, col_cedula, tipo_doc_col=col_tipo_doc, default_tipo_doc=default_tipo_doc, progress_callback=cb)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_results.to_excel(writer, index=False)
        return output.getvalue(), f"Procesados {len(df_results)} registros."
    except Exception as e:
        return None, f"Error processing massive ADRES Web: {e}"

# --- DIALOGS (Proxies to Workers with UI) ---

@st.dialog("Importar Excel para Renombrado")
def dialog_importar_excel():
    st.write("### Renombrar archivos usando Excel")
    uploaded = st.file_uploader("Subir Excel", type=["xlsx", "xls"])
    if uploaded:
        folder = st.text_input("Carpeta donde aplicar cambios", value=st.session_state.get('current_path', ''))
        if st.button("Aplicar Renombrado"):
            submit_task("Renombrar Excel", worker_aplicar_renombrado_excel, uploaded, folder)

@st.dialog("Añadir Sufijo desde Excel")
def dialog_sufijo():
    st.write("### Añadir Sufijo a Archivos")
    uploaded = st.file_uploader("Subir Excel", type=["xlsx", "xls"])
    if uploaded:
        try:
            xl = pd.ExcelFile(uploaded)
            sheet = st.selectbox("Seleccione la Hoja", xl.sheet_names, key="suf_sheet")
            df_preview = pd.read_excel(uploaded, sheet_name=sheet, nrows=1)
            cols = df_preview.columns.tolist()
            
            col_name = st.selectbox("Columna Nombre Archivo (Inicio)", cols, index=0, key="suf_col_name")
            col_suffix = st.selectbox("Columna Sufijo", cols, index=min(1, len(cols)-1), key="suf_col_suf")
            
            use_filter = st.checkbox("Usar filtro de Excel (Filas visibles)", value=False, key="suf_filter")
            
            folder = st.text_input("Carpeta Objetivo", value=st.session_state.get('current_path', ''))
            
            if st.button("Aplicar Sufijos"):
                # Reset pointer for worker
                uploaded.seek(0)
                submit_task("Sufijos", worker_anadir_sufijo_excel, uploaded, sheet, col_name, col_suffix, folder, use_filter)
        except Exception as e:
            st.error(f"Error leyendo Excel: {e}")

@st.dialog("Renombrar por Mapeo Excel")
def dialog_renombrar_mapeo_excel():
    st.write("### Renombrar Archivos (Mapeo Columna A -> Columna B)")
    uploaded = st.file_uploader("Subir Excel", type=["xlsx", "xls"])
    
    sheet = None
    col_src = None
    col_dst = None
    use_filter = True
    
    if uploaded:
        try:
            if hasattr(uploaded, 'seek'):
                uploaded.seek(0)
            xls = pd.ExcelFile(uploaded)
            sheet = st.selectbox("Nombre Hoja", xls.sheet_names, key="ren_map_sheet")
            if sheet:
                df_preview = pd.read_excel(uploaded, sheet_name=sheet, nrows=5)
                c1, c2 = st.columns(2)
                col_src = c1.selectbox("Columna Nombre Actual", df_preview.columns, key="ren_map_src")
                col_dst = c2.selectbox("Columna Nombre Nuevo", df_preview.columns, key="ren_map_dst")
                use_filter = st.checkbox("Usar filtro de Excel (solo visibles)", value=True, key="ren_map_filter")
        except Exception as e:
            st.error(f"Error: {e}")

    folder = seleccionar_carpeta_nativa("ren_map_folder", initial_dir=st.session_state.get('current_path', ''))
    
    if st.button("Renombrar"):
        if uploaded and sheet and col_src and col_dst and folder:
            if hasattr(uploaded, 'seek'):
                uploaded.seek(0)
            submit_task("Renombrar Mapeo", worker_renombrar_mapeo_excel, uploaded, sheet, col_src, col_dst, use_filter, folder)
            st.rerun()

@st.dialog("Modificar DOCX Completo")
def dialog_modif_docx_completo():
    st.write("### Modificación Masiva de DOCX (Plantillas)")
    uploaded = st.file_uploader("Subir Excel de Datos", type=["xlsx"])
    
    sheet = None
    use_filter = False
    
    if uploaded:
        try:
            uploaded.seek(0)
            xls = pd.ExcelFile(uploaded)
            sheet = st.selectbox("Nombre Hoja", xls.sheet_names, key="mod_full_sheet")
            use_filter = st.checkbox("Usar filtros de Excel (solo filas visibles)", value=False, key="mod_full_filter")
        except Exception as e:
            st.error(f"Error: {e}")

    folder = seleccionar_carpeta_nativa("mod_full_folder", initial_dir=st.session_state.get('current_path', ''))
    
    if st.button("Ejecutar Modificación"):
        if uploaded and sheet and folder:
            uploaded.seek(0)
            submit_task("Modif DOCX Completo", worker_modificar_docx_completo, uploaded, sheet, folder, use_filter)
            st.rerun()

@st.dialog("Insertar Firma en DOCX (Masivo)")
def dialog_insertar_firma_docx():
    st.write("### Insertar Firma (Imagen) en DOCX")
    st.write("Inserta una imagen de firma en documentos DOCX dentro de subcarpetas.")
    st.info("Busca la imagen de firma y la inserta en el DOCX donde diga 'Firma de Aceptacion'.")
    
    # Validación de Modo
    if not st.session_state.get("force_native_mode", True):
        st.warning("⚠️ Modo Web: La selección de carpetas nativa no está disponible.")

    base_path = seleccionar_carpeta_nativa("firma_docx_base", initial_dir=st.session_state.get("current_path", os.path.expanduser("~")))
    docx_name = st.text_input("Nombre del DOCX", value="Consentimiento.docx")
    sig_name = st.text_input("Nombre de la Firma (Imagen)", value="firma.jpg")
    
    if st.button("Iniciar Inserción de Firmas"):
        if base_path and docx_name and sig_name:
            submit_task("Insertar Firmas", worker_firmar_docx_con_imagen_masivo, base_path, docx_name, sig_name)
            st.rerun()
        else:
            st.error("Complete todos los campos.")

@st.dialog("Generar CUV (FEVRIPS)")
def dialog_generar_cuv():
    st.write("### Generar CUV masivo")
    st.info("Funcionalidad pendiente de integración completa.")

@st.dialog("RIPS: Limpieza JSON")
def dialog_rips_limpieza_json():
    st.write("### Limpieza de espacios en JSON")
    st.info("Esta herramienta elimina espacios extra en claves y valores de archivos JSON.")
    # Implementation placeholder

@st.dialog("RIPS: Actualizar Clave")
def dialog_rips_update_key():
    st.write("### Actualizar Clave en JSON")
    uploaded_files = st.file_uploader("Seleccionar archivos JSON", type=["json"], accept_multiple_files=True)
    key_to_update = st.text_input("Clave a buscar")
    new_value = st.text_input("Nuevo valor")
    if uploaded_files and key_to_update and st.button("Actualizar Clave"):
        # Placeholder
        pass

# --- WORKERS: ANALYSIS & EXTRACTION ---

def worker_analisis_historia_clinica(file_list, silent_mode=False):
    """
    Analiza masivamente los archivos PDF para extraer datos de historias clínicas.
    Retorna bytes de Excel.
    """
    if not file_list:
        if not silent_mode: st.error("No hay archivos para analizar.")
        return None

    # Filter PDFs
    archivos_pdf = [f for f in file_list if f.lower().endswith('.pdf')]
    if not archivos_pdf:
        if not silent_mode: st.warning("No se encontraron archivos PDF.")
        return None

    patterns = {
        'Paciente': re.compile(r"Paciente:\s*(.*?)(?:\n|$)", re.IGNORECASE),
        'Estrato': re.compile(r"Estrato:\s*(.*?)\s*Municipio:", re.IGNORECASE | re.DOTALL),
        'Contrato': re.compile(r"Contrato:\s*(.*?)(?:\n|$)", re.IGNORECASE),
        'HISTORIA CLÍNICA': re.compile(r"DATOS HISTORIA CL[IÍ]NICA\s*(.*?)\s*¿ES V[IÍ]CTIMA DE VIOLENCIA\?", re.IGNORECASE | re.DOTALL)
    }

    extracted_data = []
    
    progress_bar = None
    if not silent_mode:
        progress_bar = st.progress(0)
        status_text = st.empty()

    for i, pdf_path in enumerate(archivos_pdf):
        if not silent_mode and progress_bar:
            progress_bar.progress((i + 1) / len(archivos_pdf))
            status_text.text(f"Procesando: {os.path.basename(pdf_path)}")

        try:
            full_text = ""
            with fitz.open(pdf_path) as doc:
                for page in doc:
                    full_text += page.get_text("text") + "\n"
            
            record = {'Archivo': os.path.basename(pdf_path)}
            for key, pattern in patterns.items():
                match = pattern.search(full_text)
                if match:
                    val = match.group(1).strip()
                    record[key] = val
                else:
                    record[key] = "No encontrado"

            extracted_data.append(record)

        except Exception as e:
            if not silent_mode: st.error(f"Error procesando {os.path.basename(pdf_path)}: {e}")

    if not extracted_data:
        if not silent_mode: st.warning("No se extrajeron datos.")
        return None

    try:
        column_order = ['Archivo', 'Paciente', 'Estrato', 'Contrato', 'HISTORIA CLÍNICA']
        df = pd.DataFrame(extracted_data)
        # Reorder if columns exist
        cols = [c for c in column_order if c in df.columns] + [c for c in df.columns if c not in column_order]
        df = df[cols]
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
        return {
            "files": [{
                "name": "Analisis_Historia_Clinica.xlsx",
                "data": output.getvalue(),
                "label": "Descargar Análisis HC"
            }],
            "message": f"Procesados: {len(extracted_data)} registros."
        }
    except Exception as e:
        if not silent_mode: st.error(f"Error generando Excel: {e}")
        return None

def worker_leer_pdf_retefuente(file_list, silent_mode=False):
    """
    Lee archivos PDF de Retefuente y extrae RAZON SOCIAL y NIT.
    Retorna bytes de Excel.
    """
    if not file_list: return None
    
    archivos_pdf = [f for f in file_list if f.lower().endswith('.pdf')]
    if not archivos_pdf: return None

    resultados_datos = []
    
    progress_bar = None
    if not silent_mode:
        progress_bar = st.progress(0)
        status_text = st.empty()

    for i, ruta_pdf in enumerate(archivos_pdf):
        if not silent_mode and progress_bar:
            progress_bar.progress((i + 1) / len(archivos_pdf))
            status_text.text(f"Procesando: {os.path.basename(ruta_pdf)}")
            
        try:
            with fitz.open(ruta_pdf) as doc:
                for num_pagina, page in enumerate(doc, start=1):
                    blocks = page.get_text("blocks")
                    blocks.sort(key=lambda b: b[1]) # Sort vertically
                    
                    label_block = None
                    nit_label_block = None
                    nombre_encontrado = "NO ENCONTRADO"
                    nit_encontrado = "NO ENCONTRADO"
                    
                    # 1. Find key labels
                    for b in blocks:
                        text_clean = " ".join(b[4].split()).upper()
                        if "PRACTICO LA RETENCION" in text_clean:
                            label_block = b
                            break
                    
                    # Find NIT label
                    if label_block:
                        lx0, ly0, lx1, ly1 = label_block[:4]
                        for b in blocks:
                            bx0, by0 = b[:2]
                            text_clean = " ".join(b[4].split()).upper()
                            if bx0 > lx0 and abs(by0 - ly0) < 30:
                                if "NIT" in text_clean or "C.C." in text_clean:
                                    nit_label_block = b
                                    break
                    
                    if not nit_label_block:
                        for b in blocks:
                            text_clean = " ".join(b[4].split()).upper()
                            if "NIT." in text_clean and "C.C." in text_clean:
                                nit_label_block = b
                                break

                    # 2. Extract NAME
                    if label_block:
                        lx0, ly0, lx1, ly1 = label_block[:4]
                        candidates = []
                        for b in blocks:
                            if b == label_block: continue
                            bx0, by0 = b[:2]
                            if by0 > ly0 and abs(bx0 - lx0) < 100: 
                                candidates.append(b)
                        candidates.sort(key=lambda b: b[1])
                        
                        for cand in candidates:
                            text_cand = cand[4].strip()
                            upper_cand = text_cand.upper()
                            if not text_cand: continue
                            if "DIRECCION" in upper_cand: break
                            if "NIT" in upper_cand or "C.C." in upper_cand: continue
                            nombre_encontrado = " ".join(text_cand.split())
                            break
                            
                    # 3. Extract NIT
                    if nit_label_block:
                        nx0, ny0 = nit_label_block[:2]
                        nit_candidates = []
                        for b in blocks:
                            if b == nit_label_block: continue
                            bx0, by0 = b[:2]
                            if by0 > ny0 and abs(bx0 - nx0) < 80:
                                nit_candidates.append(b)
                        nit_candidates.sort(key=lambda b: b[1])
                        
                        for cand in nit_candidates:
                            text_cand = cand[4].strip()
                            upper_cand = text_cand.upper()
                            if not text_cand: continue
                            if "CIUDAD" in upper_cand: break
                            nit_encontrado = " ".join(text_cand.split())
                            break
                    
                    # Fallback for Name
                    if nombre_encontrado == "NO ENCONTRADO":
                        full_text = page.get_text("text")
                        lines = [l.strip() for l in full_text.split('\n') if l.strip()]
                        for idx, line in enumerate(lines):
                            if "PRACTICO LA RETENCION" in line.upper():
                                if idx + 1 < len(lines):
                                    potential = lines[idx+1]
                                    if "NIT" not in potential.upper() and "DIRECCION" not in potential.upper():
                                         nombre_encontrado = potential
                                break
                    
                    # Cleanup and separation
                    if nombre_encontrado != "NO ENCONTRADO":
                        match_mix = re.search(r'^(.*?)(\d{6,}[\d\s]*)$', nombre_encontrado)
                        if match_mix:
                            nombre_limpio = match_mix.group(1).strip()
                            nit_extraido = match_mix.group(2).replace(" ", "").strip()
                            nombre_encontrado = nombre_limpio
                            if nit_encontrado == "NO ENCONTRADO" or not any(c.isdigit() for c in nit_encontrado):
                                nit_encontrado = nit_extraido
                            elif any(c.isalpha() for c in nit_encontrado):
                                 nit_encontrado = nit_extraido

                    if nombre_encontrado != "NO ENCONTRADO" and not nombre_encontrado.lower().endswith('.pdf'):
                        nombre_encontrado += ".pdf"

                    resultados_datos.append({
                        "Archivo": os.path.basename(ruta_pdf),
                        "Página": num_pagina,
                        "RAZON SOCIAL / NOMBRE": nombre_encontrado,
                        "NIT / C.C.": nit_encontrado
                    })
                    
        except Exception as e:
            resultados_datos.append({
                "Archivo": os.path.basename(ruta_pdf),
                "Página": "Error",
                "RAZON SOCIAL / NOMBRE": f"ERROR: {str(e)}"
            })

    if resultados_datos:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            pd.DataFrame(resultados_datos).to_excel(writer, index=False)
        return {
            "files": [{
                "name": "Analisis_Retefuente.xlsx",
                "data": output.getvalue(),
                "label": "Descargar Retefuente"
            }],
            "message": f"Procesados: {len(resultados_datos)} registros."
        }
    return None

def worker_analisis_sos(file_list, silent_mode=False, use_ai=False, api_key=None):
    """
    Analiza archivos PDF de SOS (Autorizaciones).
    Soporta modo 'Studio' (Reglas/PDFPlumber) y opcionalmente IA.
    """
    if not file_list: return None
    
    archivos_pdf = [f for f in file_list if f.lower().endswith('.pdf')]
    if not archivos_pdf: return None
    
    extracted_data = []
    
    # Helper for SOS extraction (Studio Logic)
    def extract_sos_studio(pdf_path):
        if not pdfplumber: return {}
        data_res = {"valid_extraction": False}
        try:
            with pdfplumber.open(pdf_path) as pdf:
                if not pdf.pages: return {}
                pagina = pdf.pages[0]
                texto = pagina.extract_text() or ""
                
                # Regex Extraction
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
                for k, p in patrones.items():
                    m = re.search(p, texto, re.IGNORECASE)
                    if m: data_res[k] = m.group(1).strip()

                # Table Extraction
                tabla = pagina.extract_table()
                if tabla:
                    codigos, nombres, cantidades, respuestas, autorizaciones = [], [], [], [], []
                    for fila in tabla:
                        if not fila: continue
                        row_str = "".join([str(c) for c in fila if c]).lower()
                        if "código" in row_str and "prestación" in row_str: continue
                        if "autorizador" in row_str and "linea" in row_str: continue
                        if "(cid:" in row_str: return {"valid_extraction": False} # Garbage check

                        val_codigo, val_nombre, val_cant, val_resp, val_auth = "", "", "", "", ""

                        # 4-col format check
                        if len(fila) == 4 and (str(fila[0]).isdigit() or str(fila[3]).isdigit()):
                            val_codigo = str(fila[0]).strip()
                            val_resp = str(fila[1]).strip()
                            val_auth = str(fila[3]).strip()
                            val_nombre = f"Ver P-Autorización {val_codigo}"
                            val_cant = "1"
                        else:
                            val_codigo = str(fila[0]).strip() if len(fila) > 0 and fila[0] else ""
                            val_nombre = str(fila[1]).strip().replace("\n", " ") if len(fila) > 1 and fila[1] else ""
                            val_cant = str(fila[2]).strip() if len(fila) > 2 and fila[2] else ""
                            val_resp = str(fila[3]).strip() if len(fila) > 3 and fila[3] else ""
                            if len(fila) > 7 and fila[7]: val_auth = str(fila[7]).strip()
                        
                        if val_codigo and not val_codigo.replace(".","").isdigit(): continue

                        if val_codigo or val_nombre:
                            if val_codigo: codigos.append(val_codigo)
                            if val_nombre: nombres.append(val_nombre)
                            if val_cant: cantidades.append(val_cant)
                            if val_resp: respuestas.append(val_resp)
                            if val_auth: autorizaciones.append(val_auth)
                            data_res["valid_extraction"] = True
                    
                    if codigos: data_res["Código Prestación"] = " | ".join(codigos)
                    if nombres: data_res["Nombre Prestación"] = " | ".join(nombres)
                    if cantidades: data_res["Cantidad"] = " | ".join(cantidades)
                    if respuestas: data_res["Respuesta EPS"] = " | ".join(respuestas)
                    if autorizaciones: data_res["No. Autorización"] = " | ".join(autorizaciones)
            
            return data_res
        except Exception:
            return {}

    progress_bar = None
    if not silent_mode:
        progress_bar = st.progress(0)
        status_text = st.empty()

    for i, pdf_path in enumerate(archivos_pdf):
        if not silent_mode and progress_bar:
            progress_bar.progress((i + 1) / len(archivos_pdf))
            status_text.text(f"Procesando: {os.path.basename(pdf_path)}")
            
        record = {
            "Archivo": os.path.basename(pdf_path),
            "Fecha Consulta": "", "Afiliado": "", "Identificación": "", "Plan": "",
            "IPS Primaria": "", "Código Prestación": "", "Nombre Prestación": "",
            "Cantidad": "", "Respuesta EPS": "", "No. Autorización": "",
            "Ambito": "", "Derecho": "", "IPS Solicitante": ""
        }
        
        # Try Studio Extraction first
        studio_data = extract_sos_studio(pdf_path)
        if studio_data.get("valid_extraction"):
            record.update(studio_data)
            record["_DEBUG_STRATEGY"] = "Studio_Rules"
        elif use_ai and api_key and genai:
             # Placeholder for AI logic if needed, but keeping it simple for now
             # Since the module had complex AI logic, we can defer or implement later if requested.
             # For now, we focus on the rule-based part which is robust.
             record["_DEBUG_STRATEGY"] = "AI_Not_Implemented_Yet"
        else:
            record["_DEBUG_STRATEGY"] = "Failed"

        extracted_data.append(record)

    if extracted_data:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            pd.DataFrame(extracted_data).to_excel(writer, index=False)
        return {
            "files": [{
                "name": "Analisis_SOS.xlsx",
                "data": output.getvalue(),
                "label": "Descargar Análisis SOS"
            }],
            "message": f"Procesados: {len(extracted_data)} registros."
        }
    return None

def worker_analisis_autorizacion_nueva_eps(file_list, silent_mode=False):
    """
    Analiza archivos PDF de Autorizaciones Nueva EPS usando PyMuPDF (fitz).
    Retorna bytes de Excel.
    """
    if not fitz:
        if not silent_mode: st.error("Librería 'fitz' (PyMuPDF) no instalada.")
        return None

    data_res = []
    
    # Regex patterns (from OrganizadorArchivos_v1.py)
    patterns = {
        'Afiliado': re.compile(r"Afiliado:\s*(.*?)(?:\n|$)", re.IGNORECASE),
        'N° Autorización': re.compile(r"N° Autorización:\s*(.*?)(?:\n|$)", re.IGNORECASE),
        'Autorizada el': re.compile(r"Autorizada el:\s*(.*?)(?:\n|$)", re.IGNORECASE),
        'Descripción Servicio': re.compile(r"Descripción Servicio\s*\n\s*\d+\s+\d+\s+(.*?)(?:\n|$)", re.IGNORECASE | re.DOTALL),
        'Info de Pago': re.compile(r"(Afiliado (?:No )?Cancela.*?)(?:\n|$)", re.IGNORECASE)
    }

    progress_bar = None
    if not silent_mode:
        progress_bar = st.progress(0, text="Analizando Autorizaciones...")

    for i, file_path in enumerate(file_list):
        if not silent_mode and progress_bar:
            progress_bar.progress((i + 1) / len(file_list), text=f"Procesando: {os.path.basename(file_path)}")

        try:
            full_text = ""
            with fitz.open(file_path) as doc:
                for page in doc:
                    full_text += page.get_text("text") + "\n"
            
            row = {'Archivo': os.path.basename(file_path)}
            for key, pattern in patterns.items():
                match = pattern.search(full_text)
                if match:
                    val = match.group(1).strip()
                    # Clean up 'Descripción Servicio' which might capture too much
                    if key == 'Descripción Servicio':
                        val = val.split('\n')[0].strip()
                    row[key] = val
                else:
                    row[key] = ""
            data_res.append(row)
        except Exception as e:
            if not silent_mode: st.warning(f"Error en {os.path.basename(file_path)}: {e}")
            data_res.append({'Archivo': os.path.basename(file_path), 'Error': str(e)})

    if data_res:
        # Define column order as in original
        column_order = ['Archivo', 'Afiliado', 'N° Autorización', 'Autorizada el', 'Descripción Servicio', 'Info de Pago']
        # Ensure all columns exist
        for col in column_order:
            if col not in data_res[0]: # Check first row structure mostly
                 pass # Pandas handles missing cols but better to be safe? 
                      # Actually constructing DataFrame from dict list handles it fine.
        
        df = pd.DataFrame(data_res)
        # Reorder if columns present
        existing_cols = [c for c in column_order if c in df.columns]
        other_cols = [c for c in df.columns if c not in column_order]
        df = df[existing_cols + other_cols]

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
        return {
            "files": [{
                "name": "Analisis_Autorizaciones_NuevaEPS.xlsx",
                "data": output.getvalue(),
                "label": "Descargar Autorizaciones"
            }],
            "message": f"Procesados: {len(data_res)} registros."
        }
    return None

def worker_analisis_cargue_sanitas(file_list, silent_mode=False):
    """
    Analiza archivos PDF de Cargue Sanitas (FEOV) usando PyMuPDF (fitz).
    Retorna bytes de Excel.
    """
    if not fitz:
        if not silent_mode: st.error("Librería 'fitz' (PyMuPDF) no instalada.")
        return None

    data_res = []
    
    patterns = {
        'Factura (FEOV)': re.compile(r"FEOV(\d+)", re.IGNORECASE),
        'Fecha y hora de cargue': re.compile(r"(\d{1,2}\s+\w+\s+\d{4}\s*-\s*\d{1,2}:\d{2})", re.IGNORECASE)
    }

    progress_bar = None
    if not silent_mode:
        progress_bar = st.progress(0, text="Analizando Cargue Sanitas...")

    for i, file_path in enumerate(file_list):
        if not silent_mode and progress_bar:
            progress_bar.progress((i + 1) / len(file_list), text=f"Procesando: {os.path.basename(file_path)}")

        try:
            full_text = ""
            with fitz.open(file_path) as doc:
                for page in doc:
                    full_text += page.get_text("text") + "\n"
            
            row = {'Archivo': os.path.basename(file_path)}
            for key, pattern in patterns.items():
                match = pattern.search(full_text)
                row[key] = match.group(1).strip() if match else ""
            
            # Extract numeric FEOV
            feov_match = patterns['Factura (FEOV)'].search(full_text)
            if feov_match:
                row['Factura (FEOV)'] = feov_match.group(1)

            data_res.append(row)
        except Exception as e:
            if not silent_mode: st.warning(f"Error en {os.path.basename(file_path)}: {e}")
            data_res.append({'Archivo': os.path.basename(file_path), 'Error': str(e)})

    if data_res:
        column_order = ['Archivo', 'Factura (FEOV)', 'Fecha y hora de cargue']
        df = pd.DataFrame(data_res)
        
        # Reorder
        existing_cols = [c for c in column_order if c in df.columns]
        other_cols = [c for c in df.columns if c not in column_order]
        df = df[existing_cols + other_cols]

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
        return {
            "files": [{
                "name": "Analisis_Cargue_Sanitas.xlsx",
                "data": output.getvalue(),
                "label": "Descargar Sanitas"
            }],
            "message": f"Procesados: {len(data_res)} registros."
        }
    return None

# --- WORKERS: WEB SCRAPING & DOWNLOADS ---

def worker_descargar_firmas(uploaded_file, sheet_name, col_id, col_folder, root_path, silent_mode=False):
    """
    Descarga firmas desde una URL base usando un Excel para mapear IDs a carpetas.
    """
    if not requests or not Image:
        if not silent_mode: st.error("Faltan librerías: requests o Pillow.")
        return "Error: Librerías faltantes."

    try:
        if isinstance(uploaded_file, bytes):
            uploaded_file = io.BytesIO(uploaded_file)
        
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
        base_url = "https://oportunidaddevida.com/opvcitas/admisionescall/firmas/"
        
        descargados = 0
        errores = 0
        
        progress_bar = None
        if not silent_mode:
            progress_bar = st.progress(0)
            status_text = st.empty()
            
        total = len(df)
        
        for i, row in df.iterrows():
            if not silent_mode and progress_bar:
                progress_bar.progress((i + 1) / total)
                
            id_firma = str(row[col_id]).strip()
            nombre_carpeta = str(row[col_folder]).strip()
            
            if not id_firma or not nombre_carpeta or pd.isna(row[col_id]) or pd.isna(row[col_folder]):
                continue
                
            if not silent_mode: status_text.text(f"Procesando: {id_firma}")
            
            url_completa = f"{base_url}{id_firma}.png"
            target_dir = os.path.join(root_path, nombre_carpeta)
            os.makedirs(target_dir, exist_ok=True)
            
            try:
                response = requests.get(url_completa, stream=True, timeout=10)
                if response.status_code == 200:
                    if not response.content:
                        raise ValueError("Contenido vacío")
                    img = Image.open(io.BytesIO(response.content)).convert('RGB')
                    img.save(os.path.join(target_dir, "firma.jpg"), "JPEG")
                    descargados += 1
                else:
                    with open(os.path.join(target_dir, "no tiene firma.txt"), 'w') as f:
                        f.write(f"No firma: {url_completa} - {response.status_code}")
                    errores += 1
            except Exception:
                errores += 1
                
        return f"Proceso finalizado. Descargados: {descargados}. Errores/No encontrados: {errores}."
        
    except Exception as e:
        return f"Error crítico: {e}"

def worker_descargar_historias_ovida(uploaded_file, sheet_name, col_estudio, col_ingreso, col_egreso, col_carpeta, download_path, silent_mode=False):
    """
    Descarga historias clínicas de OVIDA usando Selenium (Chrome Headless/GUI).
    """
    try:
        from selenium import webdriver
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.webdriver.chrome.service import Service
        from webdriver_manager.chrome import ChromeDriverManager
    except ImportError:
        return "Error: Selenium/WebDriverManager no instalado."

    if not os.path.isdir(download_path):
        return "Error: Carpeta de descarga inválida."

    try:
        if isinstance(uploaded_file, bytes):
            uploaded_file = io.BytesIO(uploaded_file)
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
    except Exception as e:
        return f"Error leyendo Excel: {e}"

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
        
        if not silent_mode:
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
        start_time = time.time()
        logged_in = False
        
        while time.time() - start_time < timeout:
            try:
                # Check for an element present only in the logged-in area
                if "App/Vistas" in driver.current_url:
                    logged_in = True
                    break
            except: pass
            time.sleep(2)
            
        if not logged_in:
            return "Error: No se detectó inicio de sesión en 5 minutos."

        if not silent_mode: st.info("Inicio de sesión detectado. Comenzando descargas...")

        descargados = 0
        errores = 0
        conflictos = 0
        
        progress_bar = None
        if not silent_mode:
            progress_bar = st.progress(0)
            status_text = st.empty()
            
        total = len(df)
        
        for i, row in df.iterrows():
            if not silent_mode and progress_bar:
                progress_bar.progress((i + 1) / total)
            
            try:
                estudio = str(int(row[col_estudio])).strip()
                f_ing = pd.to_datetime(row[col_ingreso]).strftime('%Y/%m/%d')
                f_egr = pd.to_datetime(row[col_egreso]).strftime('%Y/%m/%d')
                carpeta = str(row[col_carpeta]).strip()
                
                if not all([estudio, f_ing, f_egr, carpeta]):
                    errores += 1
                    continue

                dest_dir = os.path.join(download_path, carpeta)
                os.makedirs(dest_dir, exist_ok=True)
                final_path = os.path.join(dest_dir, f"HC_{estudio}.pdf")
                
                if os.path.exists(final_path):
                    conflictos += 1
                    continue
                    
                if not silent_mode: status_text.text(f"Descargando Estudio: {estudio}")

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
                if not silent_mode: st.warning(f"Error en estudio {estudio}: {e}")
                
        return f"Finalizado. Descargados: {descargados}, Errores: {errores}, Conflictos: {conflictos}."

    except Exception as e:
        return f"Error crítico: {e}"
    finally:
        if driver: driver.quit()

# --- DIALOGS FOR DOWNLOADS ---

@st.dialog("Descargar Firmas (Excel)")
def dialog_descargar_firmas():
    st.write("Cargue un Excel con IDs de firma y nombres de carpeta.")
    
    uploaded = st.file_uploader("Archivo Excel", type=["xlsx", "xls"], key="firmas_up")
    
    sheet_name = "Hoja1"
    cols = []
    
    if uploaded:
        try:
            xl = pd.ExcelFile(uploaded)
            sheet_name = st.selectbox("Seleccione la Hoja", xl.sheet_names, key="firmas_sheet_sel")
            df_preview = pd.read_excel(uploaded, sheet_name=sheet_name, nrows=1)
            cols = df_preview.columns.tolist()
        except Exception as e:
            st.error(f"Error leyendo Excel: {e}")

    if cols:
        col_id = st.selectbox("Columna ID Firma", cols, index=cols.index("id_firma") if "id_firma" in cols else 0, key="firmas_col_id_sel")
        col_folder = st.selectbox("Columna Nombre Carpeta", cols, index=cols.index("nombre_carpeta") if "nombre_carpeta" in cols else 0, key="firmas_col_folder_sel")
    else:
        col_id = st.text_input("Columna ID Firma", value="id_firma", key="firmas_col_id")
        col_folder = st.text_input("Columna Nombre Carpeta", value="nombre_carpeta", key="firmas_col_folder")
    
    root_path = st.text_input("Ruta Raíz Descarga", value=st.session_state.get('current_path', 'C:/Firmas'), key="firmas_path")
    
    if st.button("Iniciar Descarga"):
        if uploaded and sheet_name and col_id and col_folder and root_path:
            uploaded.seek(0)
            submit_task("Descargar Firmas", worker_descargar_firmas, uploaded, sheet_name, col_id, col_folder, root_path)
            st.rerun()
        else:
            st.error("Complete todos los campos.")

@st.dialog("Descargar Historias OVIDA")
def dialog_descargar_historias_ovida():
    st.write("Automatización de descargas desde OVIDA (Requiere Credenciales).")
    st.warning("Se abrirá un navegador Chrome. Debe iniciar sesión manualmente cuando se indique.")
    
    uploaded = st.file_uploader("Archivo Excel (Pacientes)", type=["xlsx", "xls"], key="ovida_up")
    
    sheet_name = "Hoja1"
    cols = []
    
    if uploaded:
        try:
            xl = pd.ExcelFile(uploaded)
            sheet_name = st.selectbox("Seleccione la Hoja", xl.sheet_names, key="ovida_sheet_sel")
            df_preview = pd.read_excel(uploaded, sheet_name=sheet_name, nrows=1)
            cols = df_preview.columns.tolist()
        except Exception as e:
            st.error(f"Error leyendo Excel: {e}")

    c1, c2 = st.columns(2)
    with c1:
        if cols:
            col_estudio = st.selectbox("Columna Estudio", cols, index=cols.index("estudio") if "estudio" in cols else 0, key="ovida_est_sel")
            col_ingreso = st.selectbox("Columna Fecha Ingreso", cols, index=cols.index("f_ingreso") if "f_ingreso" in cols else 0, key="ovida_ing_sel")
        else:
            col_estudio = st.text_input("Columna Estudio", value="estudio", key="ovida_est")
            col_ingreso = st.text_input("Columna Fecha Ingreso", value="f_ingreso", key="ovida_ing")
            
    with c2:
        if cols:
            col_egreso = st.selectbox("Columna Fecha Egreso", cols, index=cols.index("f_egreso") if "f_egreso" in cols else 0, key="ovida_egr_sel")
            col_carpeta = st.selectbox("Columna Carpeta Destino", cols, index=cols.index("carpeta") if "carpeta" in cols else 0, key="ovida_carp_sel")
        else:
            col_egreso = st.text_input("Columna Fecha Egreso", value="f_egreso", key="ovida_egr")
            col_carpeta = st.text_input("Columna Carpeta Destino", value="carpeta", key="ovida_carp")
        
    download_path = st.text_input("Ruta Descarga Base", value=st.session_state.get('current_path', 'C:/Historias'), key="ovida_path")
    
    if st.button("Iniciar Descarga Masiva"):
        if uploaded and sheet_name and col_estudio and col_ingreso and col_egreso and col_carpeta and download_path:
            # Re-read file to ensure pointer is at start or handled by worker
            uploaded.seek(0)
            submit_task("Descargar OVIDA", worker_descargar_historias_ovida, uploaded, sheet_name, col_estudio, col_ingreso, col_egreso, col_carpeta, download_path)
            st.rerun()
        else:
            st.error("Complete todos los campos.")


def worker_organizar_facturas_por_pdf_avanzado(carpeta_destinos, carpeta_origen, silent_mode=False):
    try:
        regex = re.compile(r'FEOV(\d+)', re.IGNORECASE)
        destinos_map = {}
        
        # 1. Map destinations
        list_carpetas_destino = [d for d in os.listdir(carpeta_destinos) if os.path.isdir(os.path.join(carpeta_destinos, d))]
        
        for nombre_carpeta_destino in list_carpetas_destino:
            ruta_carpeta_destino = os.path.join(carpeta_destinos, nombre_carpeta_destino)
            for archivo in os.listdir(ruta_carpeta_destino):
                if archivo.lower().endswith('.pdf'):
                    match = regex.search(archivo)
                    if match:
                        numero_factura = match.group(1)
                        destinos_map[numero_factura] = ruta_carpeta_destino
                        break
        
        if not destinos_map:
            return "No se encontraron PDFs con patrón FEOV en las carpetas de destino."
            
        # 2. Move files
        movidos, conflictos, errores = 0, 0, 0
        
        files_to_move = []
        for root, _, files in os.walk(carpeta_origen):
            for f in files:
                files_to_move.append((root, f))
                
        if not files_to_move:
            return "No hay archivos en la carpeta de origen."
            
        if not silent_mode:
            progress_bar = st.progress(0, text="Organizando...")
            
        for i, (root, file_to_move) in enumerate(files_to_move):
            if not silent_mode and i % 10 == 0:
                progress_bar.progress((i + 1) / len(files_to_move), text=f"Procesando: {file_to_move}")
                
            moved_this = False
            for numero_factura, ruta_destino_final in destinos_map.items():
                if numero_factura in file_to_move:
                    try:
                        ruta_origen_archivo = os.path.join(root, file_to_move)
                        ruta_final_archivo = os.path.join(ruta_destino_final, file_to_move)

                        if os.path.exists(ruta_final_archivo):
                            conflictos += 1
                        else:
                            shutil.move(ruta_origen_archivo, ruta_destino_final)
                            movidos += 1
                        moved_this = True
                        break
                    except Exception:
                        errores += 1
                        break
            
        if not silent_mode:
            progress_bar.empty()
            
        return f"Movidos: {movidos}, Conflictos: {conflictos}, Errores: {errores}"
    except Exception as e:
        return f"Error: {e}"

def worker_json_evento_a_xlsx_masivo(carpeta_origen, archivo_salida, silent_mode=False):
    if not carpeta_origen or not archivo_salida: return "Rutas inválidas."
    
    archivos_json = []
    for root, dirs, files in os.walk(carpeta_origen):
        for file in files:
            if file.lower().endswith(".json"):
                archivos_json.append(os.path.join(root, file))
    
    if not archivos_json:
        return "No se encontraron archivos JSON en la carpeta seleccionada."
        
    progress_bar = None
    if not silent_mode:
        progress_bar = st.progress(0, text="Consolidando JSONs...")
    
    try:
        todas_consultas = []
        todos_procedimientos = []
        todos_otros_servicios = []
        errores = 0
        
        for i, ruta_json in enumerate(archivos_json):
            if not silent_mode:
                progress_bar.progress((i + 1) / len(archivos_json), text=f"Procesando: {os.path.basename(ruta_json)}")
            
            try:
                with open(ruta_json, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                
                usuarios_lista = data.get("usuarios", []) if isinstance(data, dict) else []
                
                for usuario in usuarios_lista:
                    base_info = {
                        "archivo_origen": os.path.basename(ruta_json),
                        "tipo_documento_usuario": usuario.get("tipoDocumentoIdentificacion"),
                        "documento_usuario": usuario.get("numDocumentoIdentificacion"),
                        "tipo_usuario": usuario.get("tipoUsuario"),
                        "fecha_nacimiento": usuario.get("fechaNacimiento"),
                        "sexo": usuario.get("codSexo"),
                        "pais_residencia": usuario.get("codPaisResidencia"),
                        "municipio_residencia": usuario.get("codMunicipioResidencia"),
                        "zona_residencia": usuario.get("codZonaTerritorialResidencia"),
                        "incapacidad": usuario.get("incapacidad"),
                        "consecutivo_usuario": usuario.get("consecutivo"),
                        "pais_origen": usuario.get("codPaisOrigen")
                    }

                    servicios = usuario.get("servicios", {})

                    for consulta in servicios.get("consultas", []):
                        todas_consultas.append({**base_info, **consulta})

                    for procedimiento in servicios.get("procedimientos", []):
                        todos_procedimientos.append({**base_info, **procedimiento})

                    for otro in servicios.get("otrosServicios", []):
                        todos_otros_servicios.append({**base_info, **otro})
                    
            except Exception:
                errores += 1
        
        if todas_consultas or todos_procedimientos or todos_otros_servicios:
            with pd.ExcelWriter(archivo_salida, engine="openpyxl") as writer:
                if todas_consultas:
                    pd.DataFrame(todas_consultas).to_excel(writer, sheet_name="Consultas", index=False)
                if todos_procedimientos:
                    pd.DataFrame(todos_procedimientos).to_excel(writer, sheet_name="Procedimientos", index=False)
                if todos_otros_servicios:
                    pd.DataFrame(todos_otros_servicios).to_excel(writer, sheet_name="OtrosServicios", index=False)
                
                if not (todas_consultas or todos_procedimientos or todos_otros_servicios):
                     pd.DataFrame().to_excel(writer, sheet_name="Vacio", index=False)

            if not silent_mode: progress_bar.empty()
            total_reg = len(todas_consultas) + len(todos_procedimientos) + len(todos_otros_servicios)
            return f"Consolidación completada. Registros: {total_reg}, Errores: {errores}"
        else:
            if not silent_mode: progress_bar.empty()
            return "No se encontraron datos RIPS válidos para exportar."
            
    except Exception as e:
        return f"Error general: {e}"

def worker_xlsx_evento_a_json_masivo(archivo_excel, carpeta_destino, silent_mode=False):
    if not archivo_excel or not carpeta_destino: return "Rutas inválidas."
    
    try:
        xls = pd.ExcelFile(archivo_excel)
        df_consultas = pd.DataFrame()
        df_procedimientos = pd.DataFrame()
        df_otros = pd.DataFrame()
        
        if "Consultas" in xls.sheet_names:
            df_consultas = pd.read_excel(xls, sheet_name="Consultas")
        if "Procedimientos" in xls.sheet_names:
            df_procedimientos = pd.read_excel(xls, sheet_name="Procedimientos")
        if "OtrosServicios" in xls.sheet_names:
            df_otros = pd.read_excel(xls, sheet_name="OtrosServicios")
            
        df_consultas = df_consultas.astype(object).where(pd.notnull(df_consultas), None)
        df_procedimientos = df_procedimientos.astype(object).where(pd.notnull(df_procedimientos), None)
        df_otros = df_otros.astype(object).where(pd.notnull(df_otros), None)

        archivos_unicos = set()
        if "archivo_origen" in df_consultas.columns:
            archivos_unicos.update(df_consultas["archivo_origen"].dropna().unique())
        if "archivo_origen" in df_procedimientos.columns:
            archivos_unicos.update(df_procedimientos["archivo_origen"].dropna().unique())
        if "archivo_origen" in df_otros.columns:
            archivos_unicos.update(df_otros["archivo_origen"].dropna().unique())
        
        if not archivos_unicos:
            return "No se encontró la columna 'archivo_origen' o está vacía."

        progress_bar = None
        if not silent_mode:
            progress_bar = st.progress(0, text="Generando JSONs...")
        
        errores = 0
        generados = 0
        
        for i, nombre_archivo in enumerate(archivos_unicos):
            if not silent_mode:
                progress_bar.progress((i + 1) / len(archivos_unicos), text=f"Generando: {nombre_archivo}")
            
            try:
                usuarios_dict = {}
                
                def procesar_df(df_origen, clave_servicio):
                    if df_origen.empty or "archivo_origen" not in df_origen.columns: return
                    df_filtrado = df_origen[df_origen["archivo_origen"] == nombre_archivo]
                    
                    for _, row in df_filtrado.iterrows():
                        td = str(row.get("tipo_documento_usuario", ""))
                        doc = str(row.get("documento_usuario", ""))
                        user_key = (td, doc)
                        
                        if user_key not in usuarios_dict:
                            usuarios_dict[user_key] = {
                                "tipoDocumentoIdentificacion": row.get("tipo_documento_usuario"),
                                "numDocumentoIdentificacion": row.get("documento_usuario"),
                                "tipoUsuario": row.get("tipo_usuario"),
                                "fechaNacimiento": row.get("fecha_nacimiento"), 
                                "codSexo": row.get("sexo"),
                                "codPaisResidencia": row.get("pais_residencia"),
                                "codMunicipioResidencia": row.get("municipio_residencia"),
                                "codZonaTerritorialResidencia": row.get("zona_residencia"),
                                "incapacidad": row.get("incapacidad"),
                                "consecutivo": row.get("consecutivo_usuario"),
                                "codPaisOrigen": row.get("pais_origen"),
                                "servicios": { "consultas": [], "procedimientos": [], "otrosServicios": [] }
                            }
                        
                        servicio_data = row.to_dict()
                        keys_to_remove = [
                            "tipo_documento_usuario", "documento_usuario", "tipo_usuario", 
                            "fecha_nacimiento", "sexo", "pais_residencia", "municipio_residencia", 
                            "zona_residencia", "incapacidad", "consecutivo_usuario", "pais_origen",
                            "archivo_origen"
                        ]
                        for k in keys_to_remove: servicio_data.pop(k, None)
                        if any(v is not None for v in servicio_data.values()):
                            usuarios_dict[user_key]["servicios"][clave_servicio].append(servicio_data)

                procesar_df(df_consultas, "consultas")
                procesar_df(df_procedimientos, "procedimientos")
                procesar_df(df_otros, "otrosServicios")
                
                resultado_final = { "usuarios": list(usuarios_dict.values()) }
                ruta_salida = os.path.join(carpeta_destino, nombre_archivo)
                if not ruta_salida.lower().endswith(".json"): ruta_salida += ".json"

                with open(ruta_salida, 'w', encoding='utf-8') as f:
                    json.dump(resultado_final, f, ensure_ascii=False, indent=4)
                generados += 1
                
            except Exception:
                errores += 1
        
        if not silent_mode: progress_bar.empty()
        return f"Se generaron {generados} archivos JSON. Errores: {errores}"
        
    except Exception as e:
        return f"Error leyendo Excel: {e}"

def worker_autorizacion_docx_desde_excel(carpeta_origen, archivo_excel, sheet_name, col_carpeta, col_auth, use_filter=False, silent_mode=False):
    if not carpeta_origen or not archivo_excel: return "Rutas inválidas."
    
    try:
        if isinstance(archivo_excel, bytes): archivo_excel = io.BytesIO(archivo_excel)
        archivo_excel.seek(0)
        
        df = None
        if use_filter:
            import openpyxl
            wb = openpyxl.load_workbook(archivo_excel, data_only=True)
            if sheet_name not in wb.sheetnames: return "Hoja no encontrada."
            ws = wb[sheet_name]
            
            data = []
            headers = [cell.value for cell in ws[1]]
            
            for row in ws.iter_rows(min_row=2):
                if not ws.row_dimensions[row[0].row].hidden:
                    data.append([cell.value for cell in row])
            
            if data:
                df = pd.DataFrame(data, columns=headers)
            else:
                return "No hay datos visibles."
        else:
            df = pd.read_excel(archivo_excel, sheet_name=sheet_name)
        
        if col_carpeta not in df.columns or col_auth not in df.columns:
            return f"Columnas no encontradas: {col_carpeta}, {col_auth}"

        modificados, errores_carpeta, errores_docx, errores_proceso = 0, 0, 0, 0
        docx_pattern = re.compile(r'CRC_.*_FEOV.*\.docx$', re.IGNORECASE)
        
        progress_bar = None
        if not silent_mode:
            progress_bar = st.progress(0, text="Modificando DOCX...")
            
        for index, fila in df.iterrows():
            if not silent_mode:
                progress_bar.progress((index + 1) / len(df), text=f"Procesando fila {index+1}")
                
            nombre_carpeta = fila[col_carpeta]
            nueva_autorizacion = fila[col_auth]
            
            if pd.isna(nombre_carpeta) or pd.isna(nueva_autorizacion): continue

            nombre_carpeta = str(nombre_carpeta).strip()
            nueva_autorizacion = str(int(nueva_autorizacion)) if isinstance(nueva_autorizacion, (float, int)) else str(nueva_autorizacion).strip()
            
            if not nombre_carpeta or not nueva_autorizacion: continue
            
            ruta_carpeta_especifica = os.path.join(carpeta_origen, nombre_carpeta)
            if not os.path.isdir(ruta_carpeta_especifica):
                errores_carpeta += 1
                continue
            
            ruta_docx_encontrada = next((os.path.join(ruta_carpeta_especifica, f) for f in os.listdir(ruta_carpeta_especifica) if docx_pattern.match(f)), None)

            if not ruta_docx_encontrada:
                errores_docx += 1
                continue

            try:
                doc = Document(ruta_docx_encontrada)
                fue_modificado = False
                for p in doc.paragraphs:
                    if "AUTORIZACION:" in p.text.upper():
                        p.text = re.sub(r'(AUTORIZACION:)\s*.*', r'\1 ' + str(nueva_autorizacion), p.text, flags=re.IGNORECASE)
                        fue_modificado = True
                        break
                
                if fue_modificado:
                    doc.save(ruta_docx_encontrada)
                    modificados += 1
                else:
                    errores_proceso += 1
            except Exception:
                errores_proceso += 1
        
        if not silent_mode: progress_bar.empty()
        return f"Modificados: {modificados}, Carpetas no encontradas: {errores_carpeta}, DOCX no encontrados: {errores_docx}, Errores proceso: {errores_proceso}"
    except Exception as e:
        return f"Error general: {e}"

def worker_regimen_docx_desde_excel(carpeta_origen, archivo_excel, sheet_name, col_carpeta, col_regimen, use_filter=False, silent_mode=False):
    if not carpeta_origen or not archivo_excel: return "Rutas inválidas."
    
    try:
        if isinstance(archivo_excel, bytes): archivo_excel = io.BytesIO(archivo_excel)
        archivo_excel.seek(0)
        
        df = None
        if use_filter:
            import openpyxl
            wb = openpyxl.load_workbook(archivo_excel, data_only=True)
            if sheet_name not in wb.sheetnames: return "Hoja no encontrada."
            ws = wb[sheet_name]
            
            data = []
            headers = [cell.value for cell in ws[1]]
            
            for row in ws.iter_rows(min_row=2):
                if not ws.row_dimensions[row[0].row].hidden:
                    data.append([cell.value for cell in row])
            
            if data:
                df = pd.DataFrame(data, columns=headers)
            else:
                return "No hay datos visibles."
        else:
            df = pd.read_excel(archivo_excel, sheet_name=sheet_name)
        
        if col_carpeta not in df.columns or col_regimen not in df.columns:
            return f"Columnas no encontradas: {col_carpeta}, {col_regimen}"

        modificados, errores_carpeta, errores_docx, errores_proceso = 0, 0, 0, 0
        docx_pattern = re.compile(r'CRC_.*_FEOV.*\.docx$', re.IGNORECASE)
        
        progress_bar = None
        if not silent_mode:
            progress_bar = st.progress(0, text="Modificando Régimen...")
            
        for index, fila in df.iterrows():
            if not silent_mode:
                progress_bar.progress((index + 1) / len(df), text=f"Procesando fila {index+1}")
                
            nombre_carpeta = fila[col_carpeta]
            nuevo_regimen = fila[col_regimen]
            
            if pd.isna(nombre_carpeta) or pd.isna(nuevo_regimen): continue

            nombre_carpeta = str(nombre_carpeta).strip()
            nuevo_regimen = str(nuevo_regimen).strip()
            
            if not nombre_carpeta or not nuevo_regimen: continue
            
            ruta_carpeta_especifica = os.path.join(carpeta_origen, nombre_carpeta)
            if not os.path.isdir(ruta_carpeta_especifica):
                errores_carpeta += 1
                continue
            
            ruta_docx_encontrada = next((os.path.join(ruta_carpeta_especifica, f) for f in os.listdir(ruta_carpeta_especifica) if docx_pattern.match(f)), None)

            if not ruta_docx_encontrada:
                errores_docx += 1
                continue

            try:
                doc = Document(ruta_docx_encontrada)
                fue_modificado = False
                for p in doc.paragraphs:
                    if "REGIMEN:" in p.text.upper():
                        p.text = re.sub(r'(REGIMEN:)\s*.*', r'\1 ' + str(nuevo_regimen), p.text, flags=re.IGNORECASE)
                        fue_modificado = True
                        break
                
                if fue_modificado:
                    doc.save(ruta_docx_encontrada)
                    modificados += 1
                else:
                    errores_proceso += 1
            except Exception:
                errores_proceso += 1
        
        if not silent_mode: progress_bar.empty()
        return f"Modificados: {modificados}, Carpetas no encontradas: {errores_carpeta}, DOCX no encontrados: {errores_docx}, Errores proceso: {errores_proceso}"
    except Exception as e:
        return f"Error general: {e}"

# --- DIALOGS FOR NEW WORKERS ---

# --- WORKERS: FIRMA DIGITAL ---

def worker_crear_firma_nombre(root_path, ttf_path, size, humanize=False, silent_mode=False):
    try:
        from PIL import ImageDraw, ImageFont
        font = ImageFont.truetype(ttf_path, size)
    except Exception as e:
        msg = f"Error cargando fuente: {e}"
        if not silent_mode: st.error(msg)
        return msg

    count = 0
    if not silent_mode:
        progress_bar = st.progress(0, text="Generando firmas...")
    
    try:
        subfolders = [d for d in os.listdir(root_path) if os.path.isdir(os.path.join(root_path, d))]
    except Exception as e:
        return f"Error leyendo carpetas: {e}"
        
    total = len(subfolders)
    
    for i, sub in enumerate(subfolders):
        if not silent_mode and i % 10 == 0: progress_bar.progress(min(i/total, 1.0))
        
        text = sub # Nombre de carpeta es el texto
        
        try:
            # Dummy draw para calcular tamaño base
            dummy_img = Image.new('RGB', (1, 1))
            dummy_draw = ImageDraw.Draw(dummy_img)
            bbox = dummy_draw.textbbox((0, 0), text, font=font)
            text_width = bbox[2] - bbox[0]
            text_height = bbox[3] - bbox[1]
            
            # Crear imagen SOLO del texto primero
            img_text = Image.new('RGB', (text_width + 20, text_height + 20), (255, 255, 255))
            draw_text = ImageDraw.Draw(img_text)
            draw_text.text((10, 10), text, font=font, fill=(0, 0, 0))
            
            final_img = img_text
            
            if humanize:
                import random
                angle = random.uniform(-8, 8) # Rotación aleatoria
                final_img = img_text.rotate(angle, expand=True, fillcolor=(255, 255, 255))
                
            # Añadir padding final consistente
            fw, fh = final_img.size
            bg_w, bg_h = fw + 60, fh + 40
            bg = Image.new('RGB', (bg_w, bg_h), (255, 255, 255))
            
            # Centrar
            offset_x = (bg_w - fw) // 2
            offset_y = (bg_h - fh) // 2
            bg.paste(final_img, (offset_x, offset_y))
            
            target_dir = os.path.join(root_path, sub, "tipografia")
            if not os.path.exists(target_dir):
                os.makedirs(target_dir)
                
            bg.save(os.path.join(target_dir, "firma.jpg"))
            count += 1
            
        except Exception as e:
            if not silent_mode: st.warning(f"Error en firma carpeta {sub}: {e}")
            
    msg = f"Generadas {count} firmas."
    if not silent_mode:
        progress_bar.progress(1.0, text="Finalizado.")
        st.success(msg)
    return msg

def worker_crear_firma_excel(root_path, ttf_path, size, excel_file, sheet_name, col_folder, col_full_name, humanize=False, silent_mode=False):
    try:
        from PIL import ImageDraw, ImageFont
        font = ImageFont.truetype(ttf_path, int(size))
    except Exception as e:
        msg = f"Error cargando fuente: {e}"
        if not silent_mode: st.error(msg)
        return msg

    try:
        if isinstance(excel_file, bytes):
            excel_file = io.BytesIO(excel_file)
        if hasattr(excel_file, 'seek'):
            excel_file.seek(0)
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
    except Exception as e:
        msg = f"Error leyendo Excel: {e}"
        if not silent_mode: st.error(msg)
        return msg

    count = 0
    if not silent_mode:
        progress_bar = st.progress(0, text="Generando firmas desde Excel...")
    total = len(df)
    
    for idx, row in df.iterrows():
        if not silent_mode and idx % 5 == 0: progress_bar.progress(min(idx/total, 1.0))
        
        folder_name = str(row[col_folder]).strip()
        if not folder_name or str(folder_name).lower() == 'nan': continue
        
        # Construir ruta objetivo
        target_dir = os.path.join(root_path, folder_name)
        if not os.path.exists(target_dir):
            continue
            
        # Extraer nombre completo
        full_name = str(row[col_full_name]).strip()
        if not full_name or full_name.lower() == 'nan': full_name = ""
        
        # Lógica inteligente para Primer Nombre + Primer Apellido
        parts = full_name.split()
        name_part = ""
        surname_part = ""
        
        if len(parts) >= 1:
            name_part = parts[0].capitalize() 
        
        if len(parts) >= 4:
            surname_part = parts[2].capitalize()
        elif len(parts) == 3:
            surname_part = parts[1].capitalize()
        elif len(parts) == 2:
            surname_part = parts[1].capitalize()
            
        # Construir texto final
        final_text = f"{name_part} {surname_part}".strip()
        if not final_text: 
            final_text = folder_name 
            
        # Generar Imagen
        try:
            # Dummy draw 
            dummy_img = Image.new('RGB', (1, 1))
            dummy_draw = ImageDraw.Draw(dummy_img)
            bbox = dummy_draw.textbbox((0, 0), final_text, font=font)
            text_width = bbox[2] - bbox[0]
            text_height = bbox[3] - bbox[1]
            
            # Crear imagen base texto
            img_text = Image.new('RGB', (text_width + 20, text_height + 20), (255, 255, 255))
            draw_text = ImageDraw.Draw(img_text)
            draw_text.text((10, 10), final_text, font=font, fill=(0, 0, 0))
            
            final_img = img_text
            
            if humanize:
                import random
                angle = random.uniform(-8, 8)
                final_img = img_text.rotate(angle, expand=True, fillcolor=(255, 255, 255))
            
            # Composition final
            fw, fh = final_img.size
            bg_w, bg_h = fw + 60, fh + 40
            bg = Image.new('RGB', (bg_w, bg_h), (255, 255, 255))
            
            offset_x = (bg_w - fw) // 2
            offset_y = (bg_h - fh) // 2
            bg.paste(final_img, (offset_x, offset_y))
            
            # Guardar
            tipografia_dir = os.path.join(target_dir, "tipografia")
            if not os.path.exists(tipografia_dir):
                os.makedirs(tipografia_dir)
                
            bg.save(os.path.join(tipografia_dir, "firma.jpg"))
            count += 1
        except Exception as e:
            if not silent_mode: st.warning(f"Error generando firma para {folder_name}: {e}")

    msg = f"Generadas {count} firmas desde Excel."
    if not silent_mode:
        progress_bar.progress(1.0, text="Finalizado.")
        st.success(msg)
    return msg

# --- DIALOGS FOR NEW WORKERS ---

@st.dialog("Crear Firma Digital desde Nombre")
def dialog_crear_firma():
    st.write("Genera una imagen JPG con firma manuscrita.")
    
    # Resolver rutas de assets
    base_dir = os.path.dirname(os.path.abspath(__file__))
    # Ajustar ruta para subir dos niveles desde src/tabs hasta la raíz, y luego a assets
    assets_fonts = os.path.join(base_dir, "..", "..", "assets", "fonts")
    
    # Selección de Fuente (común para ambos modos)
    option = st.radio("Fuente:", ["Subir fuente", "Pacifico (Predeterminada)", "MyUglyHandwriting"], index=1, horizontal=True)
    
    font_path = None
    if option == "Subir fuente":
        uploaded_font = st.file_uploader("Fuente TTF:", type=["ttf", "otf"])
        if uploaded_font:
            with open("temp_font.ttf", "wb") as f:
                f.write(uploaded_font.getbuffer())
            font_path = "temp_font.ttf"
    elif option == "Pacifico (Predeterminada)":
        font_path = os.path.join(assets_fonts, "Pacifico.ttf")
    elif option == "MyUglyHandwriting":
        font_path = os.path.join(assets_fonts, "MyUglyHandwriting-Regular.otf")
        
    c1_opt, c2_opt = st.columns(2)
    with c1_opt:
        size = st.number_input("Tamaño Fuente:", value=70)
    with c2_opt:
        humanize = st.checkbox("🎨 Estilo Natural", value=True, help="Aplica rotación aleatoria e imperfecciones.")

    # Tabs para los modos
    tab1, tab2 = st.tabs(["📁 Usar Nombre Carpeta", "📊 Usar Excel"])
    
    current_path = st.session_state.get("current_path", os.getcwd())
    
    with tab1:
        st.write("Usa el nombre de la subcarpeta como texto de la firma.")
        st.info(f"Ruta actual: {current_path}")
        if st.button("🚀 Generar (Carpeta)"):
            if font_path and os.path.exists(font_path):
                submit_task("Crear Firmas (Carpeta)", worker_crear_firma_nombre, current_path, font_path, size, humanize)
                st.rerun()
            else:
                st.error(f"No se encontró la fuente en: {font_path}")

    with tab2:
        st.write("Usa nombres extraídos de una COLUMNA ÚNICA (detecta 1er Nombre + 1er Apellido).")
        uploaded = st.file_uploader("Excel:", type=["xlsx", "xls"], key="excel_firma")
        
        if uploaded:
            try:
                if hasattr(uploaded, 'seek'): uploaded.seek(0)
                xl = pd.ExcelFile(uploaded)
                sheet = st.selectbox("Hoja:", xl.sheet_names, key="sheet_firma")
                df_prev = xl.parse(sheet_name=sheet, nrows=1)
                cols = df_prev.columns.tolist()
                
                c1, c2 = st.columns(2)
                with c1: col_folder = st.selectbox("Col. Carpeta (Match):", cols, key="col_match_firma")
                with c2: col_full_name = st.selectbox("Col. Nombre Completo:", cols, key="col_full_name_firma")
                
                st.info(f"Ruta actual: {current_path}")
                
                if st.button("🚀 Generar (Excel)"):
                     uploaded.seek(0)
                     file_bytes = uploaded.getvalue()
                     if font_path and os.path.exists(font_path):
                        submit_task("Crear Firmas (Excel)", worker_crear_firma_excel, current_path, font_path, size, file_bytes, sheet, col_folder, col_full_name, humanize)
                        st.rerun()
                     else:
                        st.error(f"No se encontró la fuente en: {font_path}")
            except Exception as e:
                st.error(f"Error leyendo Excel: {e}")

@st.dialog("Organización FEOV Avanzada")
def dialog_organizar_feov_avanzado():
    st.write("Mueve archivos de 'Origen' a subcarpetas en 'Destino' basándose en el número de factura FEOV del PDF destino.")
    
    # Validación de Modo
    if not st.session_state.get("force_native_mode", True):
        st.warning("⚠️ Modo Web: La selección de carpetas nativa no está disponible.")

    st.write("1. Carpeta Destino (contiene subcarpetas con PDFs FEOV...)")
    path_dest = seleccionar_carpeta_nativa("feov_adv_dest", initial_dir=st.session_state.get("current_path", os.path.expanduser("~")))
    
    st.write("2. Carpeta Origen (archivos a mover)")
    path_orig = seleccionar_carpeta_nativa("feov_adv_orig", initial_dir=st.session_state.get("current_path", os.path.expanduser("~")))
    
    if st.button("Iniciar Organización Avanzada"):
        if path_dest and path_orig:
            submit_task("Org. FEOV Avanzado", worker_organizar_facturas_por_pdf_avanzado, path_dest, path_orig)
            st.rerun()
        else:
            st.error("Seleccione ambas carpetas.")

@st.dialog("Autorización DOCX desde Excel")
def dialog_autorizacion_docx():
    st.write("Modifica el campo AUTORIZACION en DOCX masivamente.")
    
    # Validación de Modo
    if not st.session_state.get("force_native_mode", True):
        st.warning("⚠️ Modo Web: La selección de carpetas nativa no está disponible.")

    uploaded = st.file_uploader("Excel", type=["xlsx"], key="auth_up")
    sheet = None
    col_folder = None
    col_auth = None
    use_filter = False
    
    if uploaded:
        try:
            uploaded.seek(0)
            xls = pd.ExcelFile(uploaded)
            sheet = st.selectbox("Hoja", xls.sheet_names, key="auth_sheet")
            if sheet:
                df_preview = pd.read_excel(uploaded, sheet_name=sheet, nrows=5)
                c1, c2 = st.columns(2)
                col_folder = c1.selectbox("Columna Carpeta", df_preview.columns, key="auth_col_folder")
                col_auth = c2.selectbox("Columna Autorización", df_preview.columns, key="auth_col_val")
                use_filter = st.checkbox("Usar filtros de Excel (solo filas visibles)", value=False, key="auth_filter")
        except Exception as e:
            st.error(f"Error: {e}")

    base_path = seleccionar_carpeta_nativa("auth_base", initial_dir=st.session_state.get("current_path", os.path.expanduser("~")))
    
    if st.button("Iniciar Modificación"):
        if uploaded and base_path and col_folder and col_auth:
            uploaded.seek(0)
            submit_task("Autorización DOCX", worker_autorizacion_docx_desde_excel, uploaded, sheet, col_folder, col_auth, base_path, use_filter)
            st.rerun()

@st.dialog("Régimen DOCX desde Excel")
def dialog_regimen_docx():
    st.write("Modifica el campo REGIMEN en DOCX masivamente.")
    
    # Validación de Modo
    if not st.session_state.get("force_native_mode", True):
        st.warning("⚠️ Modo Web: La selección de carpetas nativa no está disponible.")

    uploaded = st.file_uploader("Excel", type=["xlsx"], key="reg_up")
    sheet = None
    col_folder = None
    col_reg = None
    use_filter = False
    
    if uploaded:
        try:
            uploaded.seek(0)
            xls = pd.ExcelFile(uploaded)
            sheet = st.selectbox("Hoja", xls.sheet_names, key="reg_sheet")
            if sheet:
                df_preview = pd.read_excel(uploaded, sheet_name=sheet, nrows=5)
                c1, c2 = st.columns(2)
                col_folder = c1.selectbox("Columna Carpeta", df_preview.columns, key="reg_col_folder")
                col_reg = c2.selectbox("Columna Régimen", df_preview.columns, key="reg_col_val")
                use_filter = st.checkbox("Usar filtros de Excel (solo filas visibles)", value=False, key="reg_filter")
        except Exception as e:
            st.error(f"Error: {e}")

    base_path = seleccionar_carpeta_nativa("reg_base", initial_dir=st.session_state.get("current_path", os.path.expanduser("~")))
    
    if st.button("Iniciar Modificación"):
        if uploaded and base_path and col_folder and col_reg:
            uploaded.seek(0)
            submit_task("Régimen DOCX", worker_regimen_docx_desde_excel, base_path, uploaded, sheet, col_folder, col_reg, use_filter)
            st.rerun()

@st.dialog("Crear Carpetas desde Excel")
def dialog_crear_carpetas_excel():
    st.write("Crea estructura de carpetas basada en una columna de Excel.")
    
    # Validación de Modo
    if not st.session_state.get("force_native_mode", True):
        st.warning("⚠️ Modo Web: La selección de carpetas nativa no está disponible.")

    uploaded = st.file_uploader("Excel", type=["xlsx", "xls"], key="create_fold_up")
    
    sheet = None
    col_name = None
    use_filter = False
    
    if uploaded:
        try:
            uploaded.seek(0)
            xls = pd.ExcelFile(uploaded)
            sheet = st.selectbox("Hoja", xls.sheet_names, key="create_fold_sheet")
            if sheet:
                df_preview = pd.read_excel(uploaded, sheet_name=sheet, nrows=5)
                col_name = st.selectbox("Nombre Columna Carpetas", df_preview.columns, key="create_fold_col")
                use_filter = st.checkbox("Usar filtros de Excel (solo filas visibles)", value=False, key="create_fold_filter")
        except Exception as e:
            st.error(f"Error leyendo Excel: {e}")
            
    base_path = seleccionar_carpeta_nativa("create_fold_base", initial_dir=st.session_state.get("current_path", os.path.expanduser("~")))
    
    if st.button("Crear Carpetas"):
        if uploaded and base_path and col_name:
            uploaded.seek(0)
            submit_task("Crear Carpetas", worker_crear_carpetas_excel_avanzado, uploaded, sheet, col_name, base_path, use_filter)
            st.rerun()

@st.dialog("Copiar Mapeo Subcarpetas")
def dialog_copiar_mapeo():
    st.write("Copia archivos entre carpetas basándose en un mapeo Excel.")
    uploaded = st.file_uploader("Excel", type=["xlsx", "xls"], key="copy_map_up")
    
    sheet = None
    col_src = None
    col_dst = None
    use_filter = False
    
    if uploaded:
        try:
            uploaded.seek(0)
            xls = pd.ExcelFile(uploaded)
            sheet = st.selectbox("Hoja", xls.sheet_names, key="copy_map_sheet")
            if sheet:
                df_preview = pd.read_excel(uploaded, sheet_name=sheet, nrows=5)
                c1, c2 = st.columns(2)
                col_src = c1.selectbox("Columna Origen", df_preview.columns, key="copy_map_src")
                col_dst = c2.selectbox("Columna Destino", df_preview.columns, key="copy_map_dst")
                use_filter = st.checkbox("Usar filtros de Excel (solo filas visibles)", value=False, key="copy_map_filter")
        except Exception as e:
            st.error(f"Error leyendo Excel: {e}")
    
    st.write("Rutas Base:")
    src_base = seleccionar_carpeta_nativa("copy_map_src_base", initial_dir=st.session_state.get("current_path", os.path.expanduser("~")))
    dst_base = seleccionar_carpeta_nativa("copy_map_dst_base", initial_dir=st.session_state.get("current_path", os.path.expanduser("~")))
    
    if st.button("Iniciar Copia"):
        if uploaded and src_base and dst_base and col_src and col_dst:
            uploaded.seek(0)
            submit_task("Copiar Mapeo", worker_copiar_mapeo_subcarpetas, uploaded, sheet, col_src, col_dst, src_base, dst_base, use_filter)
            st.rerun()

@st.dialog("Copiar desde Raíz (Mapeo)")
def dialog_copiar_raiz():
    st.write("Copia archivos desde una raíz única a carpetas destino según Excel.")
    
    # Validación de Modo
    if not st.session_state.get("force_native_mode", True):
        st.warning("⚠️ Modo Web: La selección de carpetas nativa no está disponible.")

    uploaded = st.file_uploader("Excel", type=["xlsx", "xls"], key="copy_root_up")
    
    sheet = None
    col_id = None
    col_folder = None
    use_filter = False
    
    if uploaded:
        try:
            uploaded.seek(0)
            xls = pd.ExcelFile(uploaded)
            sheet = st.selectbox("Hoja", xls.sheet_names, key="copy_root_sheet")
            if sheet:
                df_preview = pd.read_excel(uploaded, sheet_name=sheet, nrows=5)
                c1, c2 = st.columns(2)
                col_id = c1.selectbox("Columna ID/Nombre Archivo", df_preview.columns, key="copy_root_id")
                col_folder = c2.selectbox("Columna Carpeta Destino", df_preview.columns, key="copy_root_folder")
                use_filter = st.checkbox("Usar filtros de Excel (solo filas visibles)", value=False, key="copy_root_filter")
        except Exception as e:
            st.error(f"Error leyendo Excel: {e}")

    st.write("Rutas:")
    root_src = seleccionar_carpeta_nativa("copy_root_src_base", initial_dir=st.session_state.get("current_path", os.path.expanduser("~")))
    root_dst = seleccionar_carpeta_nativa("copy_root_dst_base", initial_dir=st.session_state.get("current_path", os.path.expanduser("~")))
    
    if st.button("Iniciar Copia"):
        if uploaded and root_src and root_dst and col_id and col_folder:
            uploaded.seek(0)
            submit_task("Copiar Raíz", worker_copiar_archivos_desde_raiz_mapeo, uploaded, sheet, col_id, col_folder, root_src, root_dst, use_filter)
            st.rerun()

@st.dialog("RIPS Eventos Masivos")
def dialog_rips_masivos():
    st.write("Conversión masiva entre JSON (Eventos) y Excel.")
    
    # Validación de Modo
    if not st.session_state.get("force_native_mode", True):
        st.warning("⚠️ Modo Web: La selección de carpetas nativa no está disponible.")

    mode = st.radio("Modo", ["JSON -> Excel", "Excel -> JSON"])
    
    if mode == "JSON -> Excel":
        folder_src = seleccionar_carpeta_nativa("rips_json_src", initial_dir=st.session_state.get("current_path", os.path.expanduser("~")))
        file_dst = st.text_input("Nombre Archivo Salida (.xlsx)", "Consolidado.xlsx")
        if st.button("Convertir JSONs a Excel"):
            if folder_src:
                submit_task("JSON Evento -> Excel", worker_json_evento_a_xlsx_masivo, folder_src, os.path.join(folder_src, file_dst))
                st.rerun()
    else:
        file_src = st.file_uploader("Excel Eventos", type=["xlsx"])
        folder_dst = seleccionar_carpeta_nativa("rips_excel_dst", initial_dir=st.session_state.get("current_path", os.path.expanduser("~")))
        if st.button("Convertir Excel a JSONs"):
            if file_src and folder_dst:
                # Need to save temp excel first? worker takes path
                t_path = os.path.join(folder_dst, "temp_eventos.xlsx")
                with open(t_path, "wb") as f: f.write(file_src.getbuffer())
                submit_task("Excel Evento -> JSON", worker_xlsx_evento_a_json_masivo, t_path, folder_dst)
                st.rerun()

# --- DIALOGS: RENAMING ---

@st.dialog("Exportar para Renombrar")
def dialog_exportar_renombrado():
    st.write("Genera un Excel con los archivos de una carpeta para renombrarlos masivamente.")
    folder = seleccionar_carpeta_nativa("renombrar_export_src", initial_dir=st.session_state.get("current_path", os.path.expanduser("~")))
    if st.button("Generar Excel"):
        if folder:
            # Create a simple Excel with OldName, NewName
            data = []
            for f in os.listdir(folder):
                if os.path.isfile(os.path.join(folder, f)):
                    data.append({"NombreActual": f, "NuevoNombre": f})
            
            if data:
                df = pd.DataFrame(data)
                out_path = os.path.join(folder, "Renombrar_Archivos.xlsx")
                df.to_excel(out_path, index=False)
                st.success(f"Excel generado en: {out_path}")
            else:
                st.warning("No se encontraron archivos en la carpeta.")

@st.dialog("Aplicar Renombrado (Excel)")
def dialog_aplicar_renombrado():
    st.write("Renombra archivos basándose en un Excel (NombreActual -> NuevoNombre).")
    excel_file = st.file_uploader("Archivo Excel", type=["xlsx"])
    folder = seleccionar_carpeta_nativa("renombrar_apply_src", initial_dir=st.session_state.get("current_path", os.path.expanduser("~")))
    
    if st.button("Aplicar Cambios"):
        if excel_file and folder:
            # Save temp excel
            t_path = os.path.join(folder, "temp_renombrar.xlsx")
            with open(t_path, "wb") as f: f.write(excel_file.getbuffer())
            submit_task("Renombrado Masivo", worker_aplicar_renombrado_excel, t_path, folder)
            st.rerun()

@st.dialog("Copiar Archivo a Subcarpetas")
def dialog_copiar_archivo_a_subcarpetas():
    st.write("Copia un archivo seleccionado a todas las subcarpetas del destino.")
    
    # Validación de Modo
    if not st.session_state.get("force_native_mode", True):
        st.warning("⚠️ Modo Web: La selección de carpetas nativa no está disponible.")

    file_to_copy = st.file_uploader("Archivo a Copiar", key="copy_sub_file")
    dest_base_path = seleccionar_carpeta_nativa("copy_sub_dest", initial_dir=st.session_state.get("current_path", os.path.expanduser("~")))
    
    if st.button("Iniciar Copia a Subcarpetas"):
        if file_to_copy and dest_base_path:
            # Save temp file
            t_path = os.path.join(dest_base_path, file_to_copy.name)
            with open(t_path, "wb") as f:
                f.write(file_to_copy.getbuffer())
            
            submit_task("Copiar a Subcarpetas", worker_copiar_archivo_a_subcarpetas, t_path, dest_base_path)
            st.rerun()
        else:
            st.error("Seleccione archivo y carpeta destino.")

@st.dialog("Organizar Facturas (FEOV)")
def dialog_organizar_feov():
    st.write("Organiza facturas PDF moviéndolas a subcarpetas según su número FEOV.")
    st.info("1. Selecciona la carpeta DESTINO (donde están las carpetas numeradas).\n2. Selecciona la carpeta ORIGEN (donde están los archivos desordenados).")
    
    # Validación de Modo
    if not st.session_state.get("force_native_mode", True):
        st.warning("⚠️ Modo Web: La selección de carpetas nativa no está disponible. Use las rutas manuales.")

    target_path = seleccionar_carpeta_nativa("feov_target", "Carpeta DESTINO (Subcarpetas)", initial_dir=st.session_state.get("current_path", os.path.expanduser("~")))
    source_path = seleccionar_carpeta_nativa("feov_source", "Carpeta ORIGEN (Archivos)", initial_dir=st.session_state.get("current_path", os.path.expanduser("~")))
    
    if st.button("Organizar Facturas"):
        if target_path and source_path:
            submit_task("Organizar FEOV", worker_organizar_facturas_feov, source_path, target_path)
            st.rerun()
        else:
            st.warning("Seleccione ambas carpetas.")

@st.dialog("Convertir PDF a Escala de Grises")
def dialog_escala_grises():
    st.write("Convierte PDFs a escala de grises para reducir tamaño.")
    
    # Validación de Modo
    if not st.session_state.get("force_native_mode", True):
        st.warning("⚠️ Modo Web: La selección de carpetas nativa no está disponible.")

    tab1, tab2 = st.tabs(["Individual/Manual", "Resultados de Búsqueda"])
    
    with tab1:
        files = st.file_uploader("Seleccionar PDFs", type=["pdf"], accept_multiple_files=True, key="gray_manual_files")
        replace = st.checkbox("Reemplazar originales", value=True, key="gray_manual_replace")
        
        if st.button("Convertir Seleccionados"):
            if files:
                # Save temp files if uploaded? 
                # worker expects paths. 
                # If we use file_uploader, we have BytesIO.
                # We need to save them to temp or use a version of worker that accepts file objects (but fitz needs path or bytes).
                # The original app used filedialog, so it had paths.
                # Here we can use paths if we ask for a folder, OR handle uploaded files.
                # Since user wants "Native" feel, maybe we should ask for a folder to process?
                # Or just handle the uploaded files by saving them to a temp dir.
                
                # However, to support "Reemplazar originales", we really need the original paths.
                # So file_uploader is not ideal for "Reemplazar".
                # Better to ask for a folder or file paths via dialog? 
                # Streamlit doesn't support picking files with paths natively securely.
                # But we have `seleccionar_carpeta_nativa`.
                pass
                
        st.write("O procesar carpeta completa:")
        folder = seleccionar_carpeta_nativa("gray_folder", initial_dir=st.session_state.get("current_path", os.path.expanduser("~")))
        if st.button("Convertir Carpeta"):
            if folder:
                pdfs = [os.path.join(folder, f) for f in os.listdir(folder) if f.lower().endswith('.pdf')]
                if pdfs:
                    submit_task("Grayscale Folder", worker_pdf_a_escala_grises, pdfs, replace)
                    st.rerun()
                else:
                    st.warning("No hay PDFs en la carpeta.")

    with tab2:
        results = st.session_state.get('search_results', [])
        pdfs_res = [f for f in results if f.lower().endswith('.pdf')] if results else []
        
        st.write(f"PDFs encontrados en última búsqueda: {len(pdfs_res)}")
        replace_res = st.checkbox("Reemplazar originales", value=True, key="gray_res_replace")
        
        if st.button("Convertir Resultados"):
            if pdfs_res:
                submit_task("Grayscale Results", worker_pdf_a_escala_grises, pdfs_res, replace_res)
                st.rerun()
            else:
                st.warning("No hay resultados PDF disponibles.")

# --- RENDER MAIN TAB ---

def render():
    st.markdown("## ⚙️ Acciones Automatizadas")
    
    # Obtener ruta global por defecto (la buscada al inicio)
    default_path = st.session_state.get("current_path", os.path.expanduser("~"))

    # Crear pestañas principales
    tab_unif, tab_org, tab_modif, tab_an, tab_create = st.tabs([
        "Unificación y División", 
        "Organización", 
        "Modificación y Renombrado", 
        "Análisis", 
        "Creación y Otros"
    ])

    # --- TAB 1: Unificación y División ---
    with tab_unif:
        st.caption("Operaciones de unión y división de archivos PDF, imágenes y DOCX.")
        
        col_u1, col_u2 = st.columns(2)
        
        with col_u1:
            st.subheader("Operaciones por Carpeta")
            path_unif = seleccionar_carpeta_nativa("Carpeta de Trabajo (Unificación)", initial_dir=default_path, key="tab_unif_folder")
            
            if st.button("🗂️ Unificar PDF por Carpeta", key="btn_unif_pdf"):
                submit_task("Unif. PDF", worker_unificar_por_carpeta, path_unif, "Unificado")
            
            if st.button("🖼️ Unificar JPG por Carpeta", key="btn_unif_jpg"):
                submit_task("Unif. JPG", worker_unificar_imagenes_por_carpeta_rec, path_unif, "Unificado.pdf", "JPG")
                
            if st.button("🖼️ Unificar PNG por Carpeta", key="btn_unif_png"):
                submit_task("Unif. PNG", worker_unificar_imagenes_por_carpeta_rec, path_unif, "Unificado.pdf", "PNG")
                
            if st.button("📄 Unificar DOCX por Carpeta", key="btn_unif_docx"):
                submit_task("Unif. DOCX", worker_unificar_docx_por_carpeta, path_unif, "Unificado.docx")

            st.divider()
            if st.button("✂️ Dividir PDFs Masivamente", key="btn_split_mass"):
                submit_task("Dividir Masivo", worker_dividir_pdfs_masivamente, path_unif)

        with col_u2:
            st.subheader("Operaciones Manuales")
            
            # Manual PDF Unify
            uploaded_pdfs = st.file_uploader("Unificar PDFs (Manual)", type=['pdf'], accept_multiple_files=True, key="col1_pdf_man")
            if st.button("🧷 Unificar Seleccionados", key="btn_unif_sel"):
                 if uploaded_pdfs:
                     out_path = os.path.join(st.session_state.get('current_path', '.'), "Unificado_Manual.pdf")
                     submit_task("Unificar Manual", worker_unificar_pdfs_list, uploaded_pdfs, out_path)

            st.divider()

            # Manual Split
            uploaded_split = st.file_uploader("Dividir PDF (Manual)", type=['pdf'], key="col1_split_man")
            if st.button("✂️ Dividir en Páginas", key="btn_split_man"):
                if uploaded_split:
                    out_folder = os.path.join(st.session_state.get('current_path', '.'), "Dividido")
                    submit_task("Dividir PDF", worker_dividir_pdf_paginas, uploaded_split, out_folder)

    # --- TAB 2: Organización ---
    with tab_org:
        st.caption("Organización de facturas, movimiento por coincidencia y consolidación.")
        
        path_org = seleccionar_carpeta_nativa("Seleccionar Carpeta de Trabajo", initial_dir=default_path, key="tab_org_folder")
        
        col_o1, col_o2 = st.columns(2)
        with col_o1:
            if st.button("📥 Organizar Facturas (FEOV)", key="btn_org_feov"):
                dialog_organizar_feov()
                
            if st.button("📂➡️📁 Mover por Coincidencia", key="btn_org_move"):
                submit_task("Mover Coincidencia", worker_mover_por_coincidencia, path_org)
                
            if st.button("⚫⚪ PDF a Escala de Grises", key="btn_org_gray"):
                dialog_escala_grises()

        with col_o2:
            if st.button("🗺️ Copiar Archivos (Mapeo Sub)", key="btn_org_map_sub"):
                dialog_copiar_mapeo()
                
            if st.button("📜 Copiar Archivos Raíz (Mapeo)", key="btn_org_map_root"):
                dialog_copiar_raiz()
                
            if st.button("📤 Consolidar Subcarpetas", key="btn_org_consol"):
                submit_task("Consolidar", worker_consolidar_subcarpetas, path_org)

    # --- TAB 3: Modificación ---
    with tab_modif:
        st.caption("Renombrado masivo con Excel y modificación de documentos DOCX.")
        
        col_m1, col_m2 = st.columns(2)
        with col_m1:
            if st.button("📤 Exportar para renombrar", key="btn_mod_exp"):
                dialog_exportar_renombrado()
                
            if st.button("📥 Aplicar renombrado Excel", key="btn_mod_app"):
                dialog_aplicar_renombrado()
                
            if st.button("🏷️ Añadir Sufijo desde Excel", key="btn_mod_suf"):
                dialog_sufijo()

            if st.button("📝 Renombrar Masivo por Mapeo Excel", key="btn_mod_map"):
                dialog_renombrar_mapeo_excel()
                
        with col_m2:
            if st.button("✍️ Modif. Autorización DOCX", key="btn_mod_auth"):
                dialog_autorizacion_docx()
                
            if st.button("✍️ Modif. Régimen DOCX", key="btn_mod_reg"):
                dialog_regimen_docx()
                
            if st.button("✍️ Modif. DOCX Completo", key="btn_mod_full"):
                dialog_modif_docx_completo()
                
            if st.button("🖋️ Firmar DOCX con Imagen", key="btn_mod_sign"):
                dialog_insertar_firma_docx()

    # --- TAB 4: Análisis ---
    with tab_an:
        st.caption("Análisis y extracción de datos de historias clínicas y otros documentos.")
        
        path_an = seleccionar_carpeta_nativa("Carpeta de Análisis", initial_dir=default_path, key="tab_an_folder")
        
        # Obtener lista de PDFs para análisis
        files_pdf = []
        if path_an and os.path.exists(path_an):
             files_pdf = [os.path.join(path_an, f) for f in os.listdir(path_an) if f.lower().endswith('.pdf')]
        
        col_a1, col_a2 = st.columns(2)

        def run_analysis_sync(func, args, key_prefix):
            try:
                with st.spinner("Procesando..."):
                    result = func(*args)
                
                if result and isinstance(result, dict) and "files" in result:
                    st.success(result.get("message", "Análisis completado."))
                    for i, f in enumerate(result["files"]):
                        data = f["data"]
                        if hasattr(data, "getvalue"): data = data.getvalue()
                        st.download_button(
                            label=f"📥 {f.get('label', 'Descargar')}",
                            data=data,
                            file_name=f["name"],
                            mime=f.get("mime", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
                            key=f"{key_prefix}_dl_{i}"
                        )
                elif result:
                     st.warning("El resultado no tiene el formato esperado para descarga directa.")
                else:
                    st.warning("No se generaron resultados.")
            except Exception as e:
                st.error(f"Error: {e}")

        with col_a1:
            if st.button("📊 Análisis Carpetas (Excel)", key="btn_an_folders"):
                if path_an and os.path.exists(path_an):
                     run_analysis_sync(worker_analisis_carpetas, [path_an], "an_folders")
                else:
                    st.warning("Seleccione una carpeta válida.")
            
            if st.button("📊 Análisis SOS", key="btn_an_sos"):
                 if files_pdf: 
                     run_analysis_sync(worker_analisis_sos, [files_pdf], "an_sos")
                 else: 
                     st.warning("No se encontraron PDFs.")

        with col_a2:
            if st.button("📊 Análisis Historia Clínica", key="btn_an_hc"):
                if files_pdf:
                    run_analysis_sync(worker_analisis_historia_clinica, [files_pdf], "an_hc")
                else:
                    st.warning("No se encontraron PDFs.")
            
            if st.button("📊 Análisis Autoriz. Nueva EPS", key="btn_an_neps"):
                if files_pdf:
                    run_analysis_sync(worker_analisis_autorizacion_nueva_eps, [files_pdf], "an_neps")
                else:
                    st.warning("No se encontraron PDFs.")

            if st.button("📊 Análisis Cargue Sanitas", key="btn_an_sanitas"):
                 if files_pdf:
                    run_analysis_sync(worker_analisis_cargue_sanitas, [files_pdf], "an_sanitas")
                 else:
                    st.warning("No se encontraron PDFs.")

            if st.button("📊 Análisis Retefuente/ICA", key="btn_an_rete"):
                 if files_pdf:
                    run_analysis_sync(worker_leer_pdf_retefuente, [files_pdf], "an_rete")
                 else:
                    st.warning("No se encontraron PDFs.")

    # --- TAB 5: Creación y Otros ---
    with tab_create:
        st.caption("Creación de carpetas, firmas digitales y distribución de archivos.")
        
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            st.subheader("Creación")
            if st.button("📂 Crear Carpetas (Excel)", key="btn_cr_folders"):
                dialog_crear_carpetas_excel()
                
            if st.button("⬇️ Descargar Firmas", key="btn_cr_sigs"):
                dialog_descargar_firmas()
                
            if st.button("⬇️ Descargar Hist. OVIDA", key="btn_cr_ovida"):
                dialog_descargar_historias_ovida()
                
            if st.button("✒️ Crear Firma Digital", key="btn_cr_dig_sig"):
                dialog_crear_firma()

        with col_c2:
            # st.subheader("Distribución / Otros")
            
            # if st.button("📤 Copiar a Subcarpetas (Dialogo)", key="btn_dlg_dist_sub"):
            #    dialog_copiar_archivo_a_subcarpetas()
            pass


