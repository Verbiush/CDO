import time
import pandas as pd
import re
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.common.exceptions import TimeoutException, NoSuchElementException

class ValidatorRegistraduria:
    def __init__(self, headless=False):
        self.headless = headless
        self.driver = None

    def start_driver(self):
        try:
            service = Service(ChromeDriverManager().install())
            options = webdriver.ChromeOptions()
            if self.headless:
                options.add_argument("--headless")
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")
            options.add_argument("--start-maximized")
            self.driver = webdriver.Chrome(service=service, options=options)
        except Exception as e:
            print(f"Error starting Chrome: {e}. Trying Edge...")
            try:
                service = Service(EdgeChromiumDriverManager().install())
                options = webdriver.EdgeOptions()
                if self.headless:
                    options.add_argument("--headless")
                self.driver = webdriver.Edge(service=service, options=options)
            except Exception as e2:
                raise Exception(f"Could not start browser: {e2}")

    def close_driver(self):
        if self.driver:
            self.driver.quit()
            self.driver = None

    def validate_cedula(self, cedula):
        if not self.driver:
            self.start_driver()
        
        try:
            self.driver.get("https://defunciones.registraduria.gov.co/")
            
            # Wait for input
            wait = WebDriverWait(self.driver, 10)
            input_el = wait.until(EC.presence_of_element_located((By.ID, "nuip")))
            
            input_el.clear()
            input_el.send_keys(str(cedula))
            
            # Click search
            search_btn = self.driver.find_element(By.CSS_SELECTOR, "button.btn-primary")
            search_btn.click()
            
            # Wait for result
            time.sleep(2) 
            
            # Extract info
            body_text = self.driver.find_element(By.TAG_NAME, "body").text
            
            status = "Desconocido"
            fecha_defuncion = ""
            detalle = ""
            fecha_consulta = time.strftime("%Y-%m-%d %H:%M:%S")
            
            # Clean text for regex
            clean_text = body_text.replace("\n", " ").strip()
            
            # --- NEW EXACT EXTRACTION STRATEGY ---
            # Match user request: "El número de documento [CEDULA] se encuentra en el archivo nacional de identificación con estado [ESTADO]"
            # Regex to capture the exact phrase shown in screenshots
            # We use a broad catch for the status until the end of the sentence/line
            match_exact = re.search(r"El número de documento\s+(\d+)\s+se encuentra en el archivo nacional de identificación con estado\s+([^.]+)", clean_text, re.IGNORECASE)
            
            if match_exact:
                # We found the exact sentence pattern
                cedula_found = match_exact.group(1)
                status_found = match_exact.group(2).strip()
                
                # Check if it matches the queried cedula (just in case)
                if str(cedula) in str(cedula_found):
                    status = status_found
                    detalle = match_exact.group(0) # The full sentence
                    
                    # If deceased, try to find date
                    if "FALLECIDO" in status.upper() or "MUERTE" in status.upper():
                         match_date = re.search(r"Fecha de defunción[:\s]+(\d{2}/\d{2}/\d{4})", clean_text, re.IGNORECASE)
                         if match_date:
                             fecha_defuncion = match_date.group(1)
                else:
                    # Mismatch in cedula found vs queried? Fallback to heuristic
                    pass
            
            # --- FALLBACK HEURISTIC STRATEGY (If regex fails) ---
            if status == "Desconocido":
                if "no aparece en la base de datos" in body_text.lower():
                    status = "NO APARECE EN BASE DE DEFUNCIONES"
                    detalle = "El número de documento no se encuentra registrado como fallecido."
                elif "FALLECIDO" in body_text.upper():
                    status = "FALLECIDO"
                    # Try to extract Date of Death
                    match_date = re.search(r"Fecha de defunción[:\s]+(\d{2}/\d{2}/\d{4})", clean_text, re.IGNORECASE)
                    if match_date:
                        fecha_defuncion = match_date.group(1)
                    detalle = "Registrado como fallecido."
                else:
                    status = "CONSULTADO (Verificar)"
                    detalle = "Respuesta no estandarizada."

            # Try to grab specific result text from a likely container
            try:
                # Common containers for results
                result_containers = self.driver.find_elements(By.CSS_SELECTOR, ".card-body, .alert, .jumbotron, .result-box")
                for container in result_containers:
                    if str(cedula) in container.text:
                        detalle = container.text.replace("\n", " | ")
                        break
            except:
                pass

            return {
                "Cedula": cedula,
                "Estado": status,
                "Fecha Defuncion": fecha_defuncion,
                "Detalle": detalle,
                "Fecha Consulta": fecha_consulta
            }

        except Exception as e:
            return {
                "Cedula": cedula,
                "Estado": "ERROR",
                "Fecha Defuncion": "",
                "Detalle": f"Error técnico: {str(e)}",
                "Fecha Consulta": time.strftime("%Y-%m-%d %H:%M:%S")
            }

    def process_massive(self, df, cedula_col, progress_callback=None):
        results = []
        total = len(df)
        
        self.start_driver() # Start once for the batch
        
        try:
            for i, row in df.iterrows():
                cedula = row[cedula_col]
                res = self.validate_cedula(cedula)
                results.append(res)
                
                if progress_callback:
                    progress_callback(i + 1, total, message=f"Procesando Cédula: {cedula}...")
                    
        finally:
            self.close_driver()
            
        return pd.DataFrame(results)

