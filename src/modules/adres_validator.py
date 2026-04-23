import time
import pandas as pd
import requests
import json
import urllib3
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# Suppress SSL warnings as we might need verify=False for some gov sites
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

class ValidatorAdres:
    def __init__(self):
        # API Base URL
        # Format: .../api/adres/{tipo_doc}/{numero_doc}
        self.base_url = "https://pqrdsuperargo.supersalud.gov.co/api/api/adres/"
        self.headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
        }
        self.tipo_doc_map = {
            "CC": "0", # Based on previous code assumption
            "TI": "1",
            "CE": "2",
            "RC": "3",
            "PA": "4",
            "PE": "5" # Permiso Especial?
        }

    def start_driver(self):
        # No driver needed for API
        pass

    def close_driver(self):
        # No driver needed for API
        pass

    def validate_cedula(self, cedula, tipo_doc="CC", timeout=None):
        """
        Validates a single cedula using the Supersalud API.
        """
        try:
            # Map text tipo_doc to API code
            tipo_code = self.tipo_doc_map.get(str(tipo_doc).upper(), "0")
            
            url = f"{self.base_url}{tipo_code}/{cedula}"
            response = requests.get(url, headers=self.headers, verify=False, timeout=20)
            
            if response.status_code == 200:
                data = response.json()
                
                # Extract fields from JSON
                nombres = f"{data.get('nombre', '')} {data.get('s_nombre', '')}".strip()
                apellidos = f"{data.get('apellido', '')} {data.get('s_apellido', '')}".strip()
                
                return {
                    "Tipo Doc": tipo_doc,
                    "Cedula": str(cedula),
                    "Nombres": nombres,
                    "Apellidos": apellidos,
                    "Fecha Nacimiento": data.get("fecha_nacimiento", ""),
                    "Departamento": data.get("departamento_id", ""),
                    "Municipio": data.get("municipio_id", ""),
                    "Estado": data.get("estado_afiliacion", ""),
                    "Entidad": data.get("eps", ""),
                    "Regimen": str(data.get("eps_tipo", "")),
                    "Tipo Afiliado": str(data.get("tipo_de_afiliado", ""))
                }
            else:
                return {
                    "Tipo Doc": tipo_doc,
                    "Cedula": cedula,
                    "Estado": f"Error API: {response.status_code}",
                    "Entidad": "No encontrado o Error de conexión"
                }

        except Exception as e:
            return {
                "Tipo Doc": tipo_doc,
                "Cedula": cedula,
                "Estado": "Error Excepción",
                "Entidad": f"Detalle: {str(e)}"
            }

    def process_massive(self, df, cedula_col, tipo_doc_col=None, default_tipo_doc="CC", progress_callback=None):
        results = []
        total = len(df)
        
        for i, row in df.iterrows():
            cedula = row[cedula_col]
            
            tipo_doc = default_tipo_doc
            if tipo_doc_col and tipo_doc_col in row and pd.notna(row[tipo_doc_col]):
                tipo_doc = str(row[tipo_doc_col]).strip()
            
            # Call API
            res = self.validate_cedula(cedula, tipo_doc=tipo_doc)
            results.append(res)
            
            if progress_callback:
                progress_callback(i + 1, total, message=f"Procesando: {tipo_doc} {cedula}...")
            
            # Small delay to be polite to the API
            time.sleep(0.5)
            
        return pd.DataFrame(results)

class ValidatorAdresWeb:
    def __init__(self, headless=False):
        self.url = "https://servicios.adres.gov.co/BDUA/Consulta-Afiliados-BDUA"
        self.driver = None
        self.headless = headless
        self.tipo_doc_map = {
            "CC": "1",
            "TI": "2",
            "CE": "3",
            "RC": "4",
            "PA": "5" # Pasaporte/Otros?
        }

    def start_driver(self):
        if not self.driver:
            options = webdriver.ChromeOptions()
            
            import sys
            # Forzar headless en servidores Linux (AWS) porque no tienen pantalla
            if self.headless or sys.platform.startswith('linux'):
                options.add_argument("--headless=new")
                
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")
            options.add_argument("--disable-gpu")
            options.add_argument("--window-size=1920,1080")
            
            # Important for stability
            options.add_experimental_option("excludeSwitches", ["enable-automation"])
            options.add_experimental_option('useAutomationExtension', False)
            
            from selenium.webdriver.chrome.service import Service
            from webdriver_manager.chrome import ChromeDriverManager
            from webdriver_manager.core.os_manager import ChromeType
            
            # If on Linux, use Chromium type to avoid version mismatch with installed Chromium
            if sys.platform.startswith('linux'):
                options.binary_location = "/usr/bin/chromium"
                service = Service(ChromeDriverManager(chrome_type=ChromeType.CHROMIUM).install())
            else:
                service = Service(ChromeDriverManager().install())
                
            self.driver = webdriver.Chrome(service=service, options=options)

    def close_driver(self):
        if self.driver:
            try:
                self.driver.quit()
            except:
                pass
            self.driver = None

    def validate_cedula(self, cedula, tipo_doc="CC", timeout=300):
        """
        Validates a single cedula using Selenium (requires manual CAPTCHA).
        Opens the site, inputs cedula, waits for user to solve captcha and submit.
        Detects result in NEW TAB/WINDOW.
        """
        self.start_driver()
        try:
            self.driver.get(self.url)
            
            # Check for iframes and switch to the one with the form
            wait = WebDriverWait(self.driver, 20)
            try:
                # Try to switch to the first iframe which usually contains the form
                wait.until(EC.frame_to_be_available_and_switch_to_it(0))
            except:
                # If failing, maybe it's not in an iframe or index changed
                self.driver.switch_to.default_content()
            
            # Wait for input
            inp_doc = wait.until(EC.element_to_be_clickable((By.ID, "txtNumDoc")))
            
            # Clear and enter cedula
            inp_doc.clear()
            inp_doc.send_keys(str(cedula))
            
            # Simulate pressing ENTER as requested
            inp_doc.send_keys(Keys.RETURN)
            
            # Select TipoDoc
            try:
                from selenium.webdriver.support.ui import Select
                select_elem = self.driver.find_element(By.ID, "tipoDoc")
                select = Select(select_elem)
                
                # Map value
                val_to_select = self.tipo_doc_map.get(str(tipo_doc).upper(), "1")
                select.select_by_value(val_to_select)
            except:
                pass
            
            # User will manually solve CAPTCHA and press Enter or Click Consultar
            # We wait for the NEW WINDOW to open
            
            original_window = self.driver.current_window_handle
            
            def check_new_window(d):
                return len(d.window_handles) > 1

            # Wait for new window to appear (Long timeout for manual captcha)
            WebDriverWait(self.driver, timeout).until(check_new_window)
            
            # Switch to new window
            new_window = [w for w in self.driver.window_handles if w != original_window][0]
            self.driver.switch_to.window(new_window)
            
            # Wait for content in new window (Tables)
            wait_new = WebDriverWait(self.driver, 20)
            wait_new.until(EC.presence_of_element_located((By.TAG_NAME, "table")))
            
            # Extract Data
            # Usually multiple tables. We want "Datos Básicos" and "Afiliación"
            # We can grab all text or specific IDs.
            # Let's try to parse tables into a dict
            
            page_source = self.driver.page_source
            tables = pd.read_html(page_source)
            
            data = {}
            
            for df in tables:
                # Based on user images, the tables are vertical (COLUMNAS | DATOS) or horizontal (ESTADO | ENTIDAD | ...)
                
                # Check for "COLUMNAS" and "DATOS" headers (First Table in image)
                cols = [str(c).upper().strip() for c in df.columns]
                
                if "COLUMNAS" in cols and "DATOS" in cols:
                    # It's the "Información Básica del Afiliado" table
                    for idx, row in df.iterrows():
                        key = str(row[cols.index("COLUMNAS")]).strip().upper()
                        val = str(row[cols.index("DATOS")]).strip()
                        data[key] = val
                        
                # Check for "ESTADO", "ENTIDAD", "REGIMEN" (Second Table in image)
                elif "ESTADO" in cols and "ENTIDAD" in cols:
                    # It's the "Datos de afiliación" table
                    if len(df) > 0:
                        row = df.iloc[0]
                        for c in df.columns:
                            data[str(c).upper().strip()] = str(row[c]).strip()

            # Map to our standard format
            # Keys based on the provided images:
            # Table 1 Keys: TIPO DE IDENTIFICACIÓN, NÚMERO DE IDENTIFICACION, NOMBRES, APELLIDOS, FECHA DE NACIMIENTO, DEPARTAMENTO, MUNICIPIO
            # Table 2 Keys: ESTADO, ENTIDAD, REGIMEN, FECHA DE AFILIACIÓN EFECTIVA, FECHA DE FINALIZACIÓN DE AFILIACIÓN, TIPO DE AFILIADO

            result = {
                "Tipo Doc": tipo_doc,
                "Cedula": str(cedula),
                "Nombres": f"{data.get('NOMBRES', '')} {data.get('APELLIDOS', '')}".strip(),
                "Apellidos": data.get("APELLIDOS", ""),
                "Fecha Nacimiento": data.get("FECHA DE NACIMIENTO", ""),
                "Departamento": data.get("DEPARTAMENTO", ""),
                "Municipio": data.get("MUNICIPIO", ""),
                "Estado": data.get("ESTADO", ""),
                "Entidad": data.get("ENTIDAD", ""),
                "Regimen": data.get("REGIMEN", ""),
                "Fecha Afiliacion": data.get("FECHA DE AFILIACIÓN EFECTIVA", ""),
                "Fecha Finalizacion": data.get("FECHA DE FINALIZACIÓN DE AFILIACIÓN", ""),
                "Tipo Afiliado": data.get("TIPO DE AFILIADO", "")
            }
            
            # Close new window and switch back
            self.driver.close()
            self.driver.switch_to.window(original_window)
            
            return result

        except TimeoutException:
            return {"Cedula": cedula, "Estado": "TimeOut (Captcha no resuelto)", "Entidad": "N/A"}
        except Exception as e:
            return {"Cedula": cedula, "Estado": "Error", "Entidad": str(e)}

    def process_massive(self, df, cedula_col, tipo_doc_col=None, default_tipo_doc="CC", progress_callback=None):
        results = []
        total = len(df)
        
        self.start_driver()
        
        try:
            for i, row in df.iterrows():
                cedula = row[cedula_col]
                
                tipo_doc = default_tipo_doc
                if tipo_doc_col and tipo_doc_col in row and pd.notna(row[tipo_doc_col]):
                    tipo_doc = str(row[tipo_doc_col]).strip()

                # Validate
                res = self.validate_cedula(cedula, tipo_doc=tipo_doc, timeout=120) # 2 mins per captcha
                results.append(res)
                
                if progress_callback:
                    progress_callback(i + 1, total, message=f"Procesando: {tipo_doc} {cedula} (Resuelva Captcha)...")
                    
        finally:
            self.close_driver()
            
        return pd.DataFrame(results)
