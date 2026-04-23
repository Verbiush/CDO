import os
import time
import base64
import urllib.parse
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

class OvidaValidator:
    def __init__(self, driver=None, headless=False):
        self.driver = driver
        self.headless = headless
        self.service = None

    def launch_browser(self):
        if self.driver:
            return self.driver
            
        import sys
        from webdriver_manager.core.os_manager import ChromeType
        
        options = webdriver.ChromeOptions()
        options.add_argument('--kiosk-printing')
        if self.headless or sys.platform.startswith('linux'):
             options.add_argument('--headless=new')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--disable-gpu')
        
        if sys.platform.startswith('linux'):
            options.binary_location = "/usr/bin/chromium"
            self.service = Service(ChromeDriverManager(chrome_type=ChromeType.CHROMIUM).install())
        else:
            self.service = Service(ChromeDriverManager().install())
            
        self.driver = webdriver.Chrome(service=self.service, options=options)
        return self.driver

    def go_to_login(self):
        if not self.driver: self.launch_browser()
        self.driver.get("https://ovidazs.siesacloud.com/ZeusSalud/ips/iniciando.php")

    def check_login_status(self):
        if not self.driver: return False
        try:
            curr = self.driver.current_url
            if "iniciando.php" not in curr or any(k in curr for k in ["menu.php", "index.php", "principal.php", "home", "dashboard"]):
                return True
        except:
            pass
        return False

    def process_massive(self, df, col_map, save_path, progress_callback=None):
        if not self.driver: raise Exception("Driver not initialized")
        
        descargados, errores, conflictos = 0, 0, 0
        total = len(df)
        
        for index, row in df.iterrows():
            if progress_callback:
                progress_callback(index + 1, total, f"Procesando {index + 1}/{total}")
                
            try:
                # Extraction logic
                nro_estudio = str(int(row[col_map['estudio']])).strip()
                # Date parsing
                fecha_ingreso_dt = pd.to_datetime(row[col_map['ingreso']])
                fecha_ingreso = fecha_ingreso_dt.strftime('%Y/%m/%d')
                fecha_egreso_dt = pd.to_datetime(row[col_map['egreso']])
                fecha_egreso = fecha_egreso_dt.strftime('%Y/%m/%d')
                nombre_carpeta = str(row[col_map['carpeta']]).strip()
                
                if not all([nro_estudio, fecha_ingreso, fecha_egreso, nombre_carpeta]):
                    errores += 1
                    continue
                    
                # URL construction
                base_url = "https://ovidazs.siesacloud.com/ZeusSalud/Reportes/Cliente//html/reporte_historia_general.php"
                params = {
                    'estudio': nro_estudio, 'fecha_inicio': fecha_ingreso, 'fecha_fin': fecha_egreso,
                    'verHC': 1, 'verEvo': 1, 'verPar': 1, 'ImprimirOrdenamiento': 1,
                    'ImprimirNotasPcte': 0, 'ImprimirSolOrdenesExt': 1, 'ImprimirGraficasHC': 1,
                    'ImprimirFormatos': 1, 'ImprimirRegistroAdmon': 1, 'ImprimirNovedad': 0,
                    'ImprimirRecomendaciones': 0, 'ImprimirDescripcionQX': 0, 'ImprimirNotasEnfermeria': 1,
                    'ImprimirSignosVitales': 0, 'ImprimirLog': 0, 'ImprimirEpicrisisSinHC': 0
                }
                full_url = f"{base_url}?{urllib.parse.urlencode(params)}"
                
                dest_folder = os.path.join(save_path, nombre_carpeta)
                os.makedirs(dest_folder, exist_ok=True)
                final_file_path = os.path.join(dest_folder, f"HC_{nro_estudio}.pdf")
                
                if os.path.exists(final_file_path):
                    conflictos += 1
                    continue
                
                self.driver.get(full_url)
                time.sleep(2)
                
                pdf_b64 = self.driver.execute_cdp_cmd("Page.printToPDF", {
                    "landscape": False, "printBackground": True,
                    "paperWidth": 8.5, "paperHeight": 11,
                    "marginTop": 0.4, "marginBottom": 0.4, "marginLeft": 0.4, "marginRight": 0.4
                })
                
                with open(final_file_path, 'wb') as f:
                    f.write(base64.b64decode(pdf_b64['data']))
                
                descargados += 1
                
            except Exception as e:
                errores += 1
                
        return {"descargados": descargados, "errores": errores, "conflictos": conflictos}
        
    def close(self):
        if self.driver:
            self.driver.quit()
            self.driver = None
