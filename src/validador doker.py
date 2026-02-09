#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ValidadorFEVRIPS.py
Interfaz gráfica para enviar archivos JSON al validador FEVRIPS (API del Ministerio).
Autor: Verbiush + ChatGPT (GPT-5)
"""

import os
import json
import csv
import requests
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime


class FEVRIPSApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Validador FEVRIPS - Generador de CUVs")
        self.geometry("880x560")
        self.minsize(820, 500)
        self.results = []

        self.create_widgets()

    # ---------------------------------------------------
    # GUI
    # ---------------------------------------------------
    def create_widgets(self):
        frame = ttk.Frame(self, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)

        row = 0
        ttk.Label(frame, text="URL del validador:").grid(column=0, row=row, sticky=tk.W)
        self.entry_url = ttk.Entry(frame, width=50)
        self.entry_url.insert(0, "https://localhost:9443")
        self.entry_url.grid(column=1, row=row, sticky=tk.EW, padx=5, columnspan=2)

        row += 1
        ttk.Label(frame, text="Endpoint validación:").grid(column=0, row=row, sticky=tk.W)
        self.entry_endpoint = ttk.Entry(frame, width=50)
        self.entry_endpoint.insert(0, "/api/FEVRIPS/validar")
        self.entry_endpoint.grid(column=1, row=row, sticky=tk.EW, padx=5, columnspan=2)

        row += 1
        ttk.Label(frame, text="Usuario:").grid(column=0, row=row, sticky=tk.W)
        self.entry_user = ttk.Entry(frame)
        self.entry_user.grid(column=1, row=row, sticky=tk.EW, padx=5)

        ttk.Label(frame, text="Contraseña:").grid(column=2, row=row, sticky=tk.W)
        self.entry_pass = ttk.Entry(frame, show="*")
        self.entry_pass.grid(column=3, row=row, sticky=tk.EW, padx=5)

        # Carpeta
        row += 1
        ttk.Label(frame, text="Carpeta con archivos JSON/XML:").grid(column=0, row=row, sticky=tk.W)
        self.entry_folder = ttk.Entry(frame)
        self.entry_folder.grid(column=1, row=row, sticky=tk.EW, padx=5, columnspan=2)
        ttk.Button(frame, text="Examinar...", command=self.browse_folder).grid(column=3, row=row, sticky=tk.E)

        # Botones
        row += 1
        ttk.Button(frame, text="Enviar al validador", command=self.process_files).grid(column=0, row=row, sticky=tk.W, pady=10)
        ttk.Button(frame, text="Exportar resultados CSV", command=self.export_csv).grid(column=1, row=row, sticky=tk.W, pady=10)

        # Status
        self.lbl_status = ttk.Label(frame, text="Listo.")
        self.lbl_status.grid(column=0, row=row+1, columnspan=4, sticky=tk.W)

        # Tabla de resultados
        columns = ("archivo", "estado", "mensaje", "cuv")
        self.tree = ttk.Treeview(frame, columns=columns, show="headings")
        self.tree.heading("archivo", text="Archivo")
        self.tree.heading("estado", text="Estado HTTP")
        self.tree.heading("mensaje", text="Mensaje / Respuesta")
        self.tree.heading("cuv", text="CUV")
        self.tree.column("archivo", width=300)
        self.tree.column("estado", width=80, anchor=tk.CENTER)
        self.tree.column("mensaje", width=350)
        self.tree.column("cuv", width=120, anchor=tk.CENTER)
        self.tree.grid(column=0, row=row+2, columnspan=4, sticky=tk.NSEW, pady=8)

        # Expandir tabla
        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(row+2, weight=1)

    # ---------------------------------------------------
    # Funciones GUI
    # ---------------------------------------------------
    def browse_folder(self):
        folder = filedialog.askdirectory(title="Selecciona la carpeta con archivos JSON")
        if folder:
            self.entry_folder.delete(0, tk.END)
            self.entry_folder.insert(0, folder)

    # ---------------------------------------------------
    # Lógica principal
    # ---------------------------------------------------
    def process_files(self):
        folder = self.entry_folder.get().strip()
        if not folder or not os.path.exists(folder):
            messagebox.showerror("Error", "Selecciona una carpeta válida.")
            return

        url_base = self.entry_url.get().strip().rstrip('/')
        endpoint = self.entry_endpoint.get().strip()
        user = self.entry_user.get().strip()
        password = self.entry_pass.get().strip()

        if not (url_base and user and password):
            messagebox.showerror("Campos requeridos", "Debes ingresar URL, usuario y contraseña.")
            return

        # Limpieza de tabla
        for i in self.tree.get_children():
            self.tree.delete(i)
        self.results = []

        self.lbl_status.config(text="Procesando archivos...")
        self.update_idletasks()

        # Login
        try:
            token = self.login_api(url_base, user, password)
        except Exception as e:
            messagebox.showerror("Error de autenticación", str(e))
            self.lbl_status.config(text="Error de login.")
            return

        headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

        # Recorre archivos .json en carpeta
        json_files = [f for f in os.listdir(folder) if f.lower().endswith(".json")]
        if not json_files:
            messagebox.showwarning("Sin archivos", "No se encontraron archivos .json en la carpeta.")
            return

        for file_name in json_files:
            full_path = os.path.join(folder, file_name)
            try:
                with open(full_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
            except Exception as e:
                self.add_result(file_name, "ERR", f"Error al leer JSON: {e}", "")
                continue

            # Envío al validador
            try:
                response = requests.post(f"{url_base}{endpoint}", headers=headers, json=data, verify=False, timeout=60)
                status = response.status_code
                try:
                    rjson = response.json()
                    
                    # --- GUARDAR RESULTADOS EXTRA ---
                    try:
                        # Extraer ID Prestador y Factura del JSON original
                        factura_num = data.get('numFactura', os.path.splitext(file_name)[0])
                        provider_id = data.get('numDocumentoIdentificacionObligado', '999')
                        
                        # 1. ResultadosLocales_nro_factura
                        f_loc_name = f"ResultadosLocales_{factura_num}.json"
                        with open(os.path.join(folder, f_loc_name), "w", encoding="utf-8") as f_out:
                            json.dump(rjson, f_out, indent=2, ensure_ascii=False)
                        
                        # 2. ResultadosMSPS_nro_factura_ID..._R
                        f_msps_name = f"ResultadosMSPS_{factura_num}_ID{provider_id}_R.json"
                        with open(os.path.join(folder, f_msps_name), "w", encoding="utf-8") as f_out:
                            json.dump(rjson, f_out, indent=2, ensure_ascii=False)
                    except Exception as e:
                        print(f"No se pudieron guardar los archivos extra: {e}")
                    # --------------------------------

                    cuv = rjson.get("cuv") or rjson.get("CUV") or ""
                    msg = json.dumps(rjson, ensure_ascii=False)[:300]
                except Exception:
                    cuv, msg = "", response.text[:300]
                self.add_result(file_name, status, msg, cuv)
            except Exception as e:
                self.add_result(file_name, "ERR", str(e), "")

        self.lbl_status.config(text=f"Procesados {len(json_files)} archivos.")
        self.update_idletasks()

    # ---------------------------------------------------
    # Login al validador (simulación genérica)
    # ---------------------------------------------------
    def login_api(self, url_base, user, password):
        """
        Realiza login en la API del validador FEVRIPS.
        Ajusta el endpoint si tu instalación usa otro.
        """
        login_url = f"{url_base}/api/Auth/Login"
        payload = {"usuario": user, "contrasena": password}
        try:
            resp = requests.post(login_url, json=payload, verify=False, timeout=30)
            if resp.status_code != 200:
                raise Exception(f"Error {resp.status_code}: {resp.text}")
            data = resp.json()
            token = data.get("access_token") or data.get("token") or ""
            if not token:
                raise Exception("Token no recibido.")
            return token
        except requests.exceptions.RequestException as e:
            raise Exception(f"Error de conexión: {e}")

    # ---------------------------------------------------
    # Guardar resultado en tabla y lista
    # ---------------------------------------------------
    def add_result(self, file_name, status, msg, cuv):
        self.tree.insert("", tk.END, values=(file_name, status, msg, cuv))
        self.results.append((file_name, status, msg, cuv))

    # ---------------------------------------------------
    # Exportar resultados CSV
    # ---------------------------------------------------
    def export_csv(self):
        if not self.results:
            messagebox.showwarning("Sin datos", "No hay resultados para exportar.")
            return

        fpath = filedialog.asksaveasfilename(
            title="Guardar resultados",
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")]
        )
        if not fpath:
            return

        try:
            with open(fpath, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(["archivo", "estado", "mensaje", "cuv", "fecha"])
                now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                for r in self.results:
                    writer.writerow([*r, now])
            messagebox.showinfo("Éxito", f"Resultados exportados a:\n{fpath}")
        except Exception as e:
            messagebox.showerror("Error al guardar", str(e))


# ---------------------------------------------------
# EJECUCIÓN
# ---------------------------------------------------
if __name__ == "__main__":
    app = FEVRIPSApp()
    app.mainloop()
