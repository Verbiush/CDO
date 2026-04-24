import os
import time
import tempfile
import shutil
import zipfile
import base64
import traceback
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager

def worker_fomag_massive(df, col_cedula, silent_mode=False):
    temp_dir = tempfile.mkdtemp()
    downloads_dir = os.path.join(temp_dir, "downloads")
    results_dir = os.path.join(temp_dir, "results")
    os.makedirs(downloads_dir, exist_ok=True)
    os.makedirs(results_dir, exist_ok=True)

    options = Options()
    # Modo visible (no headless) porque el usuario necesita iniciar sesión manualmente
    prefs = {
        "download.default_directory": downloads_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True,
        "profile.default_content_settings.popups": 0,
        "profile.default_content_setting_values.automatic_downloads": 1
    }
    options.add_experimental_option("prefs", prefs)

    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
    except Exception as e:
        return {"error": f"Error inicializando ChromeDriver: {e}"}
    
    try:
        # Abrir página inicial para logueo
        driver.get("https://horus2.horus-health.com/")
        if not silent_mode:
            print("Esperando inicio de sesión del usuario en FOMAG...")
        
        # Esperar hasta que el usuario inicie sesión y llegue a la sección de verificación
        # Damos un tiempo largo (10 minutos) para que el usuario pueda ingresar
        try:
            WebDriverWait(driver, 600).until(
                EC.url_contains("aseguramiento/verificacion")
            )
        except:
            return {"error": "Tiempo de espera agotado (10 min) para el inicio de sesión en FOMAG."}

        # Extraer lista de cédulas
        cedulas = df[col_cedula].dropna().astype(str).tolist()
        
        for cedula in cedulas:
            cedula = cedula.strip()
            if not cedula or cedula.lower() == 'nan':
                continue
            
            # Limpiar carpeta de descargas temporal antes de procesar
            for f in os.listdir(downloads_dir):
                os.remove(os.path.join(downloads_dir, f))
            
            # Asegurarnos de estar en la página de verificación
            if "aseguramiento/verificacion" not in driver.current_url:
                driver.get("https://horus2.horus-health.com/aseguramiento/verificacion")
                time.sleep(3)
            
            try:
                # 1. Encontrar el input de número de documento
                search_input = None
                type_input = None
                try:
                    # Intento 1: Buscar estrictamente por formcontrolname o name
                    search_input = driver.find_element(By.XPATH, "//input[@formcontrolname='numeroDocumento' or @name='numeroDocumento']")
                except:
                    pass
                
                if not search_input:
                    try:
                        # Intento 2: Buscar todos los inputs visibles
                        inputs = driver.find_elements(By.XPATH, "//input[not(@type='hidden') and not(@type='checkbox') and not(@type='radio')]")
                        visible_inputs = [inp for inp in inputs if inp.is_displayed() and inp.is_enabled()]
                        
                        # En la estructura de FOMAG, la primera casilla visible SIEMPRE es el dropdown "Tipo de documento" (falso input)
                        # La SEGUNDA casilla es el "Número documento" real.
                        if len(visible_inputs) >= 2:
                            type_input = visible_inputs[0]
                            search_input = visible_inputs[1] # Forzamos a elegir la segunda casilla
                        elif len(visible_inputs) == 1:
                            search_input = visible_inputs[0]
                    except:
                        pass
                
                if not search_input:
                    if not silent_mode: print(f"No se encontró input para la cédula {cedula}")
                    # Crear archivo txt
                    with open(os.path.join(results_dir, f"{cedula}.txt"), "w") as f:
                        f.write(f"Error: No se encontró la casilla para escribir el documento {cedula}")
                    continue

                # Opcional: Seleccionar 'Cedula ciudadania' en el primer campo
                if type_input:
                    try:
                        type_input.click()
                        time.sleep(0.2)
                        type_input.send_keys(Keys.CONTROL, 'a')
                        type_input.send_keys(Keys.DELETE)
                        type_input.send_keys("Cedula ciudadania")
                        time.sleep(0.2)
                        type_input.send_keys(Keys.ENTER)
                        time.sleep(0.2)
                    except:
                        pass

                # Escribir la cédula limpiando el campo completamente (por problemas de Angular)
                try:
                    search_input.click()
                    time.sleep(0.1)
                except:
                    pass
                search_input.send_keys(Keys.CONTROL, 'a')
                search_input.send_keys(Keys.DELETE)
                search_input.clear()
                time.sleep(0.2)
                search_input.send_keys(cedula)
                time.sleep(0.5)

                # 2. Encontrar y hacer clic en el botón BUSCAR
                try:
                    buscar_btn = driver.find_element(By.XPATH, "//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'buscar')] | //a[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'buscar')]")
                    buscar_btn.click()
                except Exception as e:
                    if not silent_mode: print(f"No se pudo hacer clic en BUSCAR para {cedula}: {e}")
                    continue

                # 3. Esperar a que carguen los resultados y aparezca el botón del certificado
                time.sleep(2) # Pausa breve para que procese la búsqueda
                try:
                    cert_btn = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'certificado de afiliaci')]"))
                    )
                    cert_btn.click()
                    
                    # 4. Esperar a que se descargue el archivo PDF
                    downloaded_file = None
                    for _ in range(30): # Máximo 15 segundos de espera para la descarga
                        files = os.listdir(downloads_dir)
                        if files:
                            # Ignorar archivos en proceso de descarga (.crdownload)
                            crdownloads = [f for f in files if f.endswith('.crdownload')]
                            if not crdownloads:
                                pdfs = [f for f in files if f.endswith('.pdf') or f.endswith('.PDF')]
                                if pdfs:
                                    downloaded_file = pdfs[0]
                                    break
                        time.sleep(0.5)
                    
                    # 5. Renombrar y guardar el archivo descargado
                    if downloaded_file:
                        src_path = os.path.join(downloads_dir, downloaded_file)
                        dst_path = os.path.join(results_dir, f"{cedula}.pdf")
                        shutil.move(src_path, dst_path)
                        if not silent_mode: print(f"Certificado guardado: {cedula}.pdf")
                    else:
                        if not silent_mode: print(f"No se descargó ningún PDF para {cedula}")
                        # Crear archivo txt indicando que no se descargó o encontró
                        with open(os.path.join(results_dir, f"{cedula}.txt"), "w") as f:
                            f.write(f"No se encontró certificado o no se pudo descargar para el documento: {cedula}")

                except Exception as e:
                    if not silent_mode: print(f"No se encontró botón de certificado para {cedula} (puede que no exista): {e}")
                    # Crear archivo txt indicando que no se encontró
                    with open(os.path.join(results_dir, f"{cedula}.txt"), "w") as f:
                        f.write(f"No se encontró certificado o no tiene afiliación en FOMAG para el documento: {cedula}")
                    pass # Pasamos al siguiente registro

            except Exception as e:
                if not silent_mode: print(f"Error procesando documento {cedula}: {e}")
                # Crear archivo txt de error general para esa cédula
                with open(os.path.join(results_dir, f"{cedula}.txt"), "w") as f:
                    f.write(f"Error procesando el documento {cedula}: {str(e)}")
                continue

        # Al terminar todas las cédulas, empaquetar resultados en un ZIP
        zip_path = os.path.join(temp_dir, "Certificados_FOMAG.zip")
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, _, files in os.walk(results_dir):
                for file in files:
                    zipf.write(os.path.join(root, file), file)

        # Leer el ZIP para enviarlo por la red
        if os.path.exists(zip_path) and os.path.getsize(zip_path) > 22: # > 22 bytes means not empty zip
            with open(zip_path, 'rb') as f:
                zip_data = f.read()

            return {
                "files": [{
                    "name": "Certificados_FOMAG.zip",
                    "data": zip_data,
                    "label": "Descargar ZIP"
                }],
                "message": "Descarga de certificados completada exitosamente."
            }
        else:
            return {"error": "El proceso finalizó pero no se pudo descargar ningún certificado. Verifique que los usuarios existan en FOMAG."}

    except Exception as e:
        return {"error": f"Error general en proceso FOMAG: {str(e)}\n{traceback.format_exc()}"}
    finally:
        try:
            driver.quit()
        except:
            pass
        try:
            shutil.rmtree(temp_dir, ignore_errors=True)
        except:
            pass
