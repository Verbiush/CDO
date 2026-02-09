# Importar la clase del módulo RIPS
try:
    from rips_module import OrganizadorArchivosApp as RipsApp # type: ignore
except ImportError:
    RipsApp = None # Manejar el caso si el archivo no está presente
    print("Advertencia: No se encontró el módulo 'rips_module.py'. La pestaña RIPS no estará disponible.")

# ...

class MainApp: # Renombramos la clase principal para evitar conflicto
    def __init__(self, root):
        # ... (inicialización) ...
        self.create_widgets()
        # ...
    
    def create_widgets(self):
        # ...
        # ================== TAB RIPS (Nueva Pestaña) ==================
        if RipsApp: # Si el módulo RIPS se importó correctamente
            self._setup_rips_tab(notebook)
        else:
            # Opcional: Crear una pestaña de aviso si RipsApp no está disponible
            tab_rips_placeholder = ttk.Frame(notebook)
            notebook.add(tab_rips_placeholder, text="RIPS")
            ttk.Label(tab_rips_placeholder, text="El módulo RIPS no se encontró o no se pudo importar.").pack(padx=20, pady=20)
            ttk.Label(tab_rips_placeholder, text="Por favor, asegúrate de que 'rips_module.py' está en el mismo directorio y tiene las dependencias necesarias.").pack(padx=20)
        # ...

    def _setup_rips_tab(self, notebook):
        """Configura la pestaña RIPS para cargar la interfaz de OrganizadorArchivosApp."""
        tab_rips = ttk.Frame(notebook)
        notebook.add(tab_rips, text="RIPS")

        if RipsApp:
            rips_frame = ttk.Frame(tab_rips, padding=10)
            rips_frame.pack(fill="both", expand=True)
            
            # Instanciamos RipsApp dentro de este frame
            self.rips_instance = RipsApp(rips_frame) 
        else:
            ttk.Label(tab_rips, text="El módulo RIPS no está disponible. Asegúrate de que 'rips_module.py' se ha importado correctamente.").pack()