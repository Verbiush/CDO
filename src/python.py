
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import json
import pandas as pd
import os

class JSONtoExcelConverter:
    def __init__(self, master):
        self.master = master
        master.title("JSON a Excel Converter")
        master.geometry("600x400")

        # --- Frame principal ---
        self.main_frame = ttk.Frame(master, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        # --- Selección de carpeta ---
        self.folder_label = ttk.Label(self.main_frame, text="Carpeta con archivos JSON:")
        self.folder_label.grid(row=0, column=0, sticky=tk.W, pady=5)

        self.folder_path_var = tk.StringVar()
        self.folder_entry = ttk.Entry(self.main_frame, textvariable=self.folder_path_var, width=50)
        self.folder_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)

        self.browse_button = ttk.Button(self.main_frame, text="Buscar Carpeta", command=self.browse_folder)
        self.browse_button.grid(row=0, column=2, sticky=tk.W, pady=5)

        # --- Botón de convertir ---
        self.convert_button = ttk.Button(self.main_frame, text="Convertir a Excel", command=self.convert_json_to_excel)
        self.convert_button.grid(row=1, column=0, columnspan=3, pady=20)

        # --- Barra de progreso ---
        self.progress_label = ttk.Label(self.main_frame, text="Progreso:")
        self.progress_label.grid(row=2, column=0, sticky=tk.W, pady=5)

        self.progress_bar = ttk.Progressbar(self.main_frame, orient="horizontal", length=400, mode="determinate")
        self.progress_bar.grid(row=2, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5, padx=5)

        # --- Botón para guardar Excel ---
        self.save_excel_button = ttk.Button(self.main_frame, text="Guardar Archivo Excel", command=self.save_excel_file, state=tk.DISABLED)
        self.save_excel_button.grid(row=3, column=0, columnspan=3, pady=20)

        # --- Configuración de la cuadrícula para expansibilidad ---
        self.main_frame.columnconfigure(1, weight=1)
        self.main_frame.rowconfigure(2, weight=1)

        self.all_data = [] # Para almacenar todos los datos extraídos de los JSON

    def browse_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.folder_path_var.set(folder_selected)
            # Opcional: puedes quitar este messagebox si no quieres que aparezca cada vez
            # messagebox.showinfo("Carpeta Seleccionada", f"Carpeta: {folder_selected}")

    def extract_data_from_json(self, file_path):
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)

            extracted_records = []

            # Si es una lista de documentos, iterar sobre ellos
            if isinstance(data, list):
                for doc in data:
                    extracted_records.extend(self.process_document(doc))
            elif isinstance(data, dict):
                # Si es un solo objeto JSON, procesarlo directamente
                extracted_records.extend(self.process_document(data))
            else:
                print(f"Advertencia: El archivo {file_path} no contiene un formato JSON esperado (lista o diccionario).")
            
            return extracted_records

        except FileNotFoundError:
            messagebox.showerror("Error", f"Archivo no encontrado: {file_path}")
            return []
        except json.JSONDecodeError:
            messagebox.showerror("Error", f"Error al decodificar JSON en el archivo: {file_path}")
            return []
        except Exception as e:
            messagebox.showerror("Error Inesperado", f"Ocurrió un error al procesar {file_path}: {e}")
            return []

    def process_document(self, doc):
        """
        PROCESA UN ÚNICO OBJETO JSON DE DOCUMENTO Y EXTRAE LA INFORMACIÓN RELEVANTE.
        ESTA FUNCIÓN SE HA MODIFICADO PARA SER MÁS FLEXIBLE CON LA ESTRUCTURA DE LOS SERVICIOS.
        """
        records = []
        
        # Extracción de datos del nivel superior del JSON (ejemplo)
        num_documento_id_obligado = doc.get("numDocumentoIdObligado")
        num_factura = doc.get("numFactura")
        tipo_nota = doc.get("tipoNota")
        num_nota = doc.get("numNota")

        # Extracción de datos de la lista de usuarios
        usuarios = doc.get("usuarios", [])
        for i, usuario in enumerate(usuarios):
            tipo_documento_identificacion_usuario = usuario.get("tipoDocumentoIdentificacion")
            num_documento_identificacion_usuario = usuario.get("numDocumentoIdentificacion")
            tipo_usuario = usuario.get("tipoUsuario")
            fecha_nacimiento = usuario.get("fechaNacimiento")
            cod_sexo = usuario.get("codSexo")
            cod_pais_residencia = usuario.get("codPaisResidencia")
            cod_municipio_residencia = usuario.get("codMunicipioResidencia")
            cod_zona_territorial_residencia = usuario.get("codZonaTerritorialResidencia")
            incapacidad = usuario.get("incapacidad")
            
            # --- Lógica flexible para procesar los servicios ---
            servicios_data = usuario.get("servicios", {})
            servicios_a_procesar = []

            # Intentar extraer de la clave "consultas"
            if "consultas" in servicios_data and isinstance(servicios_data["consultas"], list):
                servicios_a_procesar = servicios_data["consultas"]
                servicio_type = "consulta"
            # Intentar extraer de la clave "procedimientos"
            elif "procedimientos" in servicios_data and isinstance(servicios_data["procedimientos"], list):
                servicios_a_procesar = servicios_data["procedimientos"]
                servicio_type = "procedimiento"
            else:
                print(f"Advertencia: No se encontraron 'consultas' ni 'procedimientos' en el usuario {num_documento_identificacion_usuario} del archivo.")

            # Procesar la lista de servicios encontrada (ya sea consultas o procedimientos)
            for j, servicio_item in enumerate(servicios_a_procesar):
                # Inicializar campos comunes
                cod_prestador = None
                fecha_inicio_atencion = None
                num_autorizacion = None
                cod_servicio_id = None # Este será el campo para codConsulta o codProcedimiento
                modalidad_grupo_servicio_tec_sal = None
                grupo_servicios = None
                finalidad_tecnologia_salud = None
                causa_motivo_atencion = None
                cod_diagnostico_principal = None
                cod_diagnostico_relacionado1 = None
                cod_diagnostico_relacionado2 = None
                cod_diagnostico_relacionado3 = None
                tipo_diagnostico_principal = None
                tipo_documento_identificacion_paciente = None
                num_documento_identificacion_paciente = None
                vr_servicio = None
                concepto_recaudo = None
                valor_pago_moderador = None
                num_fe_pago_moderador = None

                # Mapeo de campos según el tipo de servicio
                if servicio_type == "consulta":
                    cod_prestador = servicio_item.get("codPrestador")
                    fecha_inicio_atencion = servicio_item.get("fechaInicioAtencion")
                    num_autorizacion = servicio_item.get("numAutorizacion")
                    cod_servicio_id = servicio_item.get("codConsulta") # Campo específico para consulta
                    modalidad_grupo_servicio_tec_sal = servicio_item.get("modalidadGrupoServicioTecSal")
                    grupo_servicios = servicio_item.get("grupoServicios")
                    finalidad_tecnologia_salud = servicio_item.get("finalidadTecnologiaSalud")
                    causa_motivo_atencion = servicio_item.get("causaMotivoAtencion")
                    cod_diagnostico_principal = servicio_item.get("codDiagnosticoPrincipal")
                    cod_diagnostico_relacionado1 = servicio_item.get("codDiagnosticoRelacionado1")
                    cod_diagnostico_relacionado2 = servicio_item.get("codDiagnosticoRelacionado2")
                    cod_diagnostico_relacionado3 = servicio_item.get("codDiagnosticoRelacionado3")
                    tipo_diagnostico_principal = servicio_item.get("tipoDiagnosticoPrincipal")
                    tipo_documento_identificacion_paciente = servicio_item.get("tipoDocumentoIdentificacion") # Paciente de la consulta
                    num_documento_identificacion_paciente = servicio_item.get("numDocumentoIdentificacion") # Paciente de la consulta
                    vr_servicio = servicio_item.get("vrServicio")
                    concepto_recaudo = servicio_item.get("conceptoRecaudo")
                    valor_pago_moderador = servicio_item.get("valorPagoModerador")
                    num_fe_pago_moderador = servicio_item.get("numFEVPagoModerador")

                elif servicio_type == "procedimiento":
                    cod_prestador = servicio_item.get("codPrestador")
                    fecha_inicio_atencion = servicio_item.get("fechaInicioAtencion")
                    num_autorizacion = servicio_item.get("numAutorizacion")
                    cod_servicio_id = servicio_item.get("codProcedimiento") # Campo específico para procedimiento
                    via_ingreso_servicio_salud = servicio_item.get("viaIngresoServicioSalud") # Campo específico para procedimiento
                    modalidad_grupo_servicio_tec_sal = servicio_item.get("modalidadGrupoServicioTecSal")
                    grupo_servicios = servicio_item.get("grupoServicios")
                    cod_servicio = servicio_item.get("codServicio") # Campo específico para procedimiento
                    finalidad_tecnologia_salud = servicio_item.get("finalidadTecnologiaSalud")
                    causa_motivo_atencion = servicio_item.get("causaMotivoAtencion")
                    cod_diagnostico_principal = servicio_item.get("codDiagnosticoPrincipal")
                    cod_diagnostico_relacionado1 = servicio_item.get("codDiagnosticoRelacionado") # Campo diferente en procedimiento
                    #cod_diagnostico_relacionado2 = servicio_item.get("codDiagnosticoRelacionado2") # No presente en este ejemplo
                    #cod_diagnostico_relacionado3 = servicio_item.get("codDiagnosticoRelacionado3") # No presente en este ejemplo
                    tipo_diagnostico_principal = servicio_item.get("tipoDiagnosticoPrincipal")
                    tipo_documento_identificacion_paciente = servicio_item.get("tipoDocumentoIdentificacion") # Paciente del procedimiento
                    num_documento_identificacion_paciente = servicio_item.get("numDocumentoIdentificacion") # Paciente del procedimiento
                    vr_servicio = servicio_item.get("vrServicio")
                    concepto_recaudo = servicio_item.get("conceptoRecaudo")
                    valor_pago_moderador = servicio_item.get("valorPagoModerador")
                    num_fe_pago_moderador = servicio_item.get("numFEVPagoModerador")

                # Registra cada servicio como una fila en el Excel
                # Las claves de este diccionario serán las cabeceras de tu archivo Excel
                record_data = {
                    "numDocumentoIdObligado": num_documento_id_obligado,
                    "numFactura": num_factura,
                    "tipoNota": tipo_nota,
                    "numNota": num_nota,
                    "usuario_consecutivo": i + 1, # Para diferenciar usuarios si hay varios en un JSON
                    "usuario_tipoDocumentoIdentificacion": tipo_documento_identificacion_usuario,
                    "usuario_numDocumentoIdentificacion": num_documento_identificacion_usuario,
                    "usuario_tipoUsuario": tipo_usuario,
                    "usuario_fechaNacimiento": fecha_nacimiento,
                    "usuario_codSexo": cod_sexo,
                    "usuario_codPaisResidencia": cod_pais_residencia,
                    "usuario_codMunicipioResidencia": cod_municipio_residencia,
                    "usuario_codZonaTerritorialResidencia": cod_zona_territorial_residencia,
                    "usuario_incapacidad": incapacidad,
                    "servicio_consecutivo": j + 1, # Para diferenciar servicios si hay varios en un usuario
                    "servicio_tipo": servicio_type.capitalize(), # Indica si es Consulta o Procedimiento
                    "servicio_codPrestador": cod_prestador,
                    "servicio_fechaInicioAtencion": fecha_inicio_atencion,
                    "servicio_numAutorizacion": num_autorizacion,
                    # Aquí se mapea a un campo genérico "codServicioPrincipal" porque puede ser codConsulta o codProcedimiento
                    f"servicio_cod_{'Consulta' if servicio_type == 'consulta' else 'Procedimiento'}": cod_servicio_id,
                    "servicio_modalidadGrupoServicioTecSal": modalidad_grupo_servicio_tec_sal,
                    "servicio_grupoServicios": grupo_servicios,
                    # Mapeo condicional para codServicio si es procedimiento
                    f"servicio_cod_servicio_procedimiento": servicio_item.get("codServicio") if servicio_type == "procedimiento" else None,
                    "servicio_finalidadTecnologiaSalud": finalidad_tecnologia_salud,
                    "servicio_causaMotivoAtencion": causa_motivo_atencion,
                    "servicio_codDiagnosticoPrincipal": cod_diagnostico_principal,
                    "servicio_codDiagnosticoRelacionado1": cod_diagnostico_relacionado1,
                    "servicio_codDiagnosticoRelacionado2": cod_diagnostico_relacionado2,
                    "servicio_codDiagnosticoRelacionado3": cod_diagnostico_relacionado3,
                    "servicio_tipoDiagnosticoPrincipal": tipo_diagnostico_principal,
                    "servicio_tipoDocumentoIdentificacion_paciente": tipo_documento_identificacion_paciente,
                    "servicio_numDocumentoIdentificacion_paciente": num_documento_identificacion_paciente,
                    "servicio_vrServicio": vr_servicio,
                    "servicio_conceptoRecaudo": concepto_recaudo,
                    "servicio_valorPagoModerador": valor_pago_moderador,
                    "servicio_numFEVPagoModerador": num_fe_pago_moderador
                }
                
                # Si es un procedimiento, agregar campos específicos que no existen en consultas
                if servicio_type == "procedimiento":
                    record_data["servicio_viaIngresoServicioSalud"] = via_ingreso_servicio_salud
                    # Agregamos campos que podrían existir en procedimientos pero no en consultas y viceversa
                    # Para evitar que queden como None si el tipo no es el correcto
                    if "codDiagnosticoRelacionado2" not in servicio_item:
                         record_data["servicio_codDiagnosticoRelacionado2"] = None
                    if "codDiagnosticoRelacionado3" not in servicio_item:
                         record_data["servicio_codDiagnosticoRelacionado3"] = None


                records.append(record_data)
        return records

    def convert_json_to_excel(self):
        folder_path = self.folder_path_var.get()
        if not folder_path:
            messagebox.showwarning("Advertencia", "Por favor, selecciona una carpeta primero.")
            return

        self.all_data = [] # Limpiar datos anteriores si se vuelve a convertir
        json_files = [f for f in os.listdir(folder_path) if f.endswith('.json')]

        if not json_files:
            messagebox.showinfo("Información", "No se encontraron archivos JSON en la carpeta seleccionada.")
            return

        total_files = len(json_files)
        self.progress_bar["value"] = 0
        self.progress_bar["maximum"] = total_files
        self.master.update_idletasks() # Actualiza la GUI para mostrar la barra de progreso

        for i, filename in enumerate(json_files):
            file_path = os.path.join(folder_path, filename)
            print(f"Procesando: {filename}") # Mensaje en consola para depuración
            extracted_data = self.extract_data_from_json(file_path)
            self.all_data.extend(extracted_data)
            
            # Actualizar barra de progreso
            self.progress_bar["value"] = i + 1
            self.master.update_idletasks() # Importante para que la GUI se actualice en tiempo real

        if self.all_data:
            messagebox.showinfo("Conversión Completada", f"Se han procesado {total_files} archivos JSON. ¡Listo para guardar en Excel!")
            self.save_excel_button.config(state=tk.NORMAL) # Habilitar botón de guardar
        else:
            messagebox.showwarning("Advertencia", "No se extrajeron datos de los archivos JSON procesados.")
            self.save_excel_button.config(state=tk.DISABLED) # Asegurarse de que esté deshabilitado si no hay datos

    def save_excel_file(self):
        if not self.all_data:
            messagebox.showwarning("Advertencia", "No hay datos para guardar.")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*")],
            title="Guardar archivo Excel"
        )

        if not file_path: # Si el usuario cancela el diálogo de guardar
            return

        try:
            df = pd.DataFrame(self.all_data)
            df.to_excel(file_path, index=False)
            messagebox.showinfo("Guardado Exitoso", f"Archivo Excel guardado en: {file_path}")
        except Exception as e:
            messagebox.showerror("Error al Guardar", f"Ocurrió un error al guardar el archivo Excel: {e}")

# --- Inicialización de la aplicación ---
if __name__ == "__main__":
    root = tk.Tk()
    app = JSONtoExcelConverter(root)
    root.mainloop()