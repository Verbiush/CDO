
SESSION_FILE = "bot_session.json"

def guardar_sesion():
    global PASOS_MEMORIZADOS, FLUJOS_CONDICIONALES
    try:
        # Convertir claves int a str para JSON (aunque json dump lo hace solo, mejor ser explícito o dejarlo ser)
        # JSON keys must be strings.
        flujos_export = {str(k): v for k, v in FLUJOS_CONDICIONALES.items()}
        
        state = {
            "pasos_principales": PASOS_MEMORIZADOS,
            "flujos_condicionales": flujos_export
        }
        with open(SESSION_FILE, "w", encoding="utf-8") as f:
            json.dump(state, f, indent=4)
        return True, f"Sesión guardada en {SESSION_FILE}"
    except Exception as e:
        print(f"Error guardando sesión: {e}")
        return False, f"Error guardando sesión: {e}"

def cargar_sesion():
    global PASOS_MEMORIZADOS, FLUJOS_CONDICIONALES, PASOS_ALTERNATIVOS, CONDICION_EJECUCION
    try:
        if not os.path.exists(SESSION_FILE):
            return False, "No existe archivo de sesión anterior."
            
        with open(SESSION_FILE, "r", encoding="utf-8") as f:
            state = json.load(f)
            
        PASOS_MEMORIZADOS = state.get("pasos_principales", [])
        
        # Cargar flujos condicionales
        raw_flows = state.get("flujos_condicionales", {})
        # Convert keys back to int
        FLUJOS_CONDICIONALES = {int(k): v for k, v in raw_flows.items()}
        
        # Sync Legacy (Slot 0)
        if 0 in FLUJOS_CONDICIONALES:
            PASOS_ALTERNATIVOS = FLUJOS_CONDICIONALES[0].get("pasos", [])
            CONDICION_EJECUCION = FLUJOS_CONDICIONALES[0].get("condicion")
        else:
            PASOS_ALTERNATIVOS = []
            CONDICION_EJECUCION = None
            
        return True, f"Sesión restaurada ({len(PASOS_MEMORIZADOS)} pasos, {len(FLUJOS_CONDICIONALES)} flujos)."
    except Exception as e:
        return False, f"Error cargando sesión: {e}"
