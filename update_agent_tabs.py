import sys

content = open('src/tabs/tab_automated_actions.py', 'r', encoding='utf-8').read()

mover_coincidencia_code = '''def worker_mover_por_coincidencia(root_path, silent_mode=False, return_zip=False):
    is_native_mode = getattr(st, 'session_state', {}).get('force_native_mode', True) if hasattr(st, 'session_state') else False
    if is_native_mode:
        try:
            from src.agent_client import send_command, wait_for_result
            username = st.session_state.get("username", "default")
            if not silent_mode: st.info("Enviando tarea al Agente Local...")
            task_id = send_command(username, "mover_por_coincidencia", {"root_path": root_path})
            if task_id:
                res = wait_for_result(task_id, timeout=300)
                if isinstance(res, dict):
                    if "error" in res: return {"error": f"Error del Agente: {res['error']}"}
                    return {"message": res.get("message", "Operación completada por el agente.")}
                return {"error": "Respuesta inesperada del agente."}
            return {"error": "No se pudo conectar con el Agente Local."}
        except Exception as e:
            return {"error": f"Error comunicando con el agente: {e}"}

    if not silent_mode: st.info(f"Iniciando movimiento por coincidencia en: {root_path}")'''

content = content.replace('def worker_mover_por_coincidencia(root_path, silent_mode=False, return_zip=False):\n    if not silent_mode: st.info(f"Iniciando movimiento por coincidencia en: {root_path}")', mover_coincidencia_code)

feov_code = '''def worker_organizar_facturas_feov(root_path, target_path, silent_mode=False, return_zip=False):
    if not root_path or not target_path: return {"error": "Error: Rutas no válidas."}
    is_native_mode = getattr(st, 'session_state', {}).get('force_native_mode', True) if hasattr(st, 'session_state') else False
    if is_native_mode:
        try:
            from src.agent_client import send_command, wait_for_result
            username = st.session_state.get("username", "default")
            if not silent_mode: st.info("Enviando tarea al Agente Local...")
            task_id = send_command(username, "organizar_feov", {"root_path": root_path, "target_path": target_path})
            if task_id:
                res = wait_for_result(task_id, timeout=300)
                if isinstance(res, dict):
                    if "error" in res: return {"error": f"Error del Agente: {res['error']}"}
                    return {"message": res.get("message", "Operación completada por el agente.")}
                return {"error": "Respuesta inesperada del agente."}
            return {"error": "No se pudo conectar con el Agente Local."}
        except Exception as e: return {"error": f"Error comunicando con el agente: {e}"}

    if not silent_mode: st.info("Iniciando organización de facturas FEOV...")'''

content = content.replace('def worker_organizar_facturas_feov(root_path, target_path, silent_mode=False, return_zip=False):\n    if not root_path or not target_path:\n        return {"error": "Error: Rutas de origen o destino no válidas."}\n\n    if not silent_mode: st.info("Iniciando organización de facturas FEOV...")', feov_code)

map_sub_code = '''def worker_copiar_mapeo_subcarpetas(uploaded_file, sheet_name, col_src, col_dst, path_src_base, path_dst_base, use_filter=False, silent_mode=False):
    is_native_mode = getattr(st, 'session_state', {}).get('force_native_mode', True) if hasattr(st, 'session_state') else False
    if is_native_mode:
        try:
            from src.agent_client import send_command, wait_for_result
            import base64
            username = st.session_state.get("username", "default")
            if not silent_mode: st.info("Enviando tarea al Agente Local...")
            if hasattr(uploaded_file, "seek"): uploaded_file.seek(0)
            file_bytes = uploaded_file if isinstance(uploaded_file, bytes) else uploaded_file.getvalue()
            b64_file = base64.b64encode(file_bytes).decode('utf-8')
            
            task_id = send_command(username, "copiar_mapeo_subcarpetas", {
                "file_bytes_b64": b64_file,
                "sheet_name": sheet_name,
                "col_src": col_src,
                "col_dst": col_dst,
                "path_src_base": path_src_base,
                "path_dst_base": path_dst_base,
                "use_filter": use_filter
            })
            if task_id:
                res = wait_for_result(task_id, timeout=300)
                if isinstance(res, dict):
                    if "error" in res: return f"Error del Agente: {res['error']}"
                    return res.get("message", "Operación completada por el agente.")
                return "Respuesta inesperada del agente."
            return "No se pudo conectar con el Agente Local."
        except Exception as e: return f"Error comunicando con el agente: {e}"

    try:'''

content = content.replace('def worker_copiar_mapeo_subcarpetas(uploaded_file, sheet_name, col_src, col_dst, path_src_base, path_dst_base, use_filter=False, silent_mode=False):\n    try:', map_sub_code)


map_root_code = '''def worker_copiar_archivos_desde_raiz_mapeo(uploaded_file, sheet_name, col_id, col_folder, root_src, root_dst, use_filter=False, silent_mode=False):
    is_native_mode = getattr(st, 'session_state', {}).get('force_native_mode', True) if hasattr(st, 'session_state') else False
    if is_native_mode:
        try:
            from src.agent_client import send_command, wait_for_result
            import base64
            username = st.session_state.get("username", "default")
            if not silent_mode: st.info("Enviando tarea al Agente Local...")
            if hasattr(uploaded_file, "seek"): uploaded_file.seek(0)
            file_bytes = uploaded_file if isinstance(uploaded_file, bytes) else uploaded_file.getvalue()
            b64_file = base64.b64encode(file_bytes).decode('utf-8')
            
            task_id = send_command(username, "copiar_archivos_desde_raiz_mapeo", {
                "file_bytes_b64": b64_file,
                "sheet_name": sheet_name,
                "col_id": col_id,
                "col_folder": col_folder,
                "root_src": root_src,
                "root_dst": root_dst,
                "use_filter": use_filter
            })
            if task_id:
                res = wait_for_result(task_id, timeout=300)
                if isinstance(res, dict):
                    if "error" in res: return f"Error del Agente: {res['error']}"
                    return res.get("message", "Operación completada por el agente.")
                return "Respuesta inesperada del agente."
            return "No se pudo conectar con el Agente Local."
        except Exception as e: return f"Error comunicando con el agente: {e}"

    try:'''

content = content.replace('def worker_copiar_archivos_desde_raiz_mapeo(uploaded_file, sheet_name, col_id, col_folder, root_src, root_dst, use_filter=False, silent_mode=False):\n    try:', map_root_code)

with open('src/tabs/tab_automated_actions.py', 'w', encoding='utf-8') as f:
    f.write(content)
