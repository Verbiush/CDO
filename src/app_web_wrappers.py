
# --- TASK WRAPPERS FOR ANALYSIS ---
def run_analisis_sos_task(file_list, use_ai):
    from src.modules.analisis_sos import worker_analisis_sos
    result = worker_analisis_sos(file_list, use_ai=use_ai, silent_mode=True)
    if isinstance(result, tuple):
        out_xlsx, out_txt = result
        return {
            "files": [
                {"name": "Analisis_SOS.xlsx", "data": out_xlsx, "mime": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "label": "Excel"},
                {"name": "Analisis_SOS.txt", "data": out_txt, "mime": "text/csv", "label": "CSV/TXT"}
            ],
            "message": "Análisis SOS completado."
        }
    elif result:
        return {
            "files": [
                {"name": "Analisis_SOS.xlsx", "data": result, "mime": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "label": "Excel"}
            ],
            "message": "Análisis SOS completado."
        }
    return None

def run_analisis_historia_task(file_list):
    # This assumes worker_analisis_historia_clinica is defined in the global scope
    # Since it's defined in this file, it should be available.
    out = worker_analisis_historia_clinica(file_list)
    return out

def run_analisis_autorizacion_task(file_list):
    out = worker_analisis_autorizacion_nueva_eps(file_list)
    return out

def run_analisis_sanitas_task(file_list):
    out = worker_analisis_cargue_sanitas(file_list)
    return out

def run_analisis_retefuente_task(file_list):
    out = worker_analisis_retefuente(file_list)
    return out

def run_analisis_carpetas_task(path):
    out = worker_analisis_carpetas(path)
    return out
