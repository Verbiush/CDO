import threading
import uuid
import time
import streamlit as st
from streamlit.runtime.scriptrunner import add_script_run_ctx

def init_task_manager():
    """Inicializa la estructura de tareas en session_state si no existe."""
    if "bg_tasks" not in st.session_state:
        st.session_state.bg_tasks = {}

def submit_task(name, target_func, *args, **kwargs):
    """
    Envía una tarea a ejecución en segundo plano.
    
    Args:
        name (str): Nombre legible de la tarea.
        target_func (callable): Función a ejecutar. Debe devolver (bytes, msg) o bytes.
        *args, **kwargs: Argumentos para la función.
    """
    init_task_manager()
    task_id = str(uuid.uuid4())
    
    def task_wrapper():
        try:
            # Update status to running
            if task_id in st.session_state.bg_tasks:
                st.session_state.bg_tasks[task_id]["status"] = "RUNNING"
            
            # Run function
            result = target_func(*args, **kwargs)
            
            # Store result
            if task_id in st.session_state.bg_tasks:
                st.session_state.bg_tasks[task_id]["result"] = result
                st.session_state.bg_tasks[task_id]["status"] = "COMPLETED"
                st.session_state.bg_tasks[task_id]["end_time"] = time.time()
                st.session_state.bg_tasks[task_id]["seen"] = False # For notification
            
        except Exception as e:
            if task_id in st.session_state.bg_tasks:
                st.session_state.bg_tasks[task_id]["error"] = str(e)
                st.session_state.bg_tasks[task_id]["status"] = "FAILED"
                st.session_state.bg_tasks[task_id]["end_time"] = time.time()
                st.session_state.bg_tasks[task_id]["seen"] = False

    # Initialize task entry
    st.session_state.bg_tasks[task_id] = {
        "id": task_id,
        "name": name,
        "status": "PENDING",
        "start_time": time.time(),
        "progress": 0
    }
    
    t = threading.Thread(target=task_wrapper)
    add_script_run_ctx(t)
    t.start()
    return task_id

def show_task_notifications():
    """Muestra notificaciones toast para tareas terminadas recientemente."""
    if "bg_tasks" not in st.session_state:
        return

    # Check for completed tasks that haven't been seen
    # Usamos una lista auxiliar para no modificar el diccionario mientras iteramos si fuera necesario, 
    # aunque aquí solo modificamos valores internos.
    for tid, task in st.session_state.bg_tasks.items():
        if task.get("status") in ["COMPLETED", "FAILED"] and not task.get("seen", True):
            if task["status"] == "COMPLETED":
                st.toast(f"✅ Tarea terminada: {task['name']}", icon="✅")
            else:
                st.toast(f"❌ Error en tarea: {task['name']}", icon="❌")
            
            # Mark as seen
            task["seen"] = True

def render_task_center():
    """Renderiza el centro de tareas en la barra lateral."""
    if "bg_tasks" not in st.session_state:
        return

    # Auto-refresh si hay tareas corriendo (Check at the start or end of render)
    # We do it at the end of render so the user sees the "Running" status first.
    
    has_running = False
    if st.session_state.bg_tasks:
        with st.sidebar.expander("🔔 Centro de Tareas", expanded=True):
            # Sort by start time desc (newest first)
            tasks = sorted(st.session_state.bg_tasks.values(), key=lambda x: x["start_time"], reverse=True)
            
            # Limit to last 5 tasks to avoid clutter
            for task in tasks[:5]:
                st.markdown(f"**{task['name']}**")
                status = task["status"]
                
                if status == "RUNNING" or status == "PENDING":
                    has_running = True
                    if status == "RUNNING":
                        st.info("Ejecutando... Puedes cambiar de pestaña.", icon="⏳")
                    else:
                        st.text("Pendiente...")
                        
                elif status == "COMPLETED":
                    # Show result
                    res = task.get("result")
                    
                    # Support for standardized dict format: {"files": [{"name":..., "data":..., "label":...}], "message":...}
                    if isinstance(res, dict):
                        msg = res.get("message", "")
                        if msg: 
                            st.info(f"Resultado:\n{msg}")
                        
                        if "files" in res:
                            for i, file_info in enumerate(res["files"]):
                                data = file_info["data"]
                                if hasattr(data, "getvalue"): data = data.getvalue()
                                
                                st.download_button(
                                    label=f"📥 {file_info.get('label', 'Descargar')}",
                                    data=data,
                                    file_name=file_info["name"],
                                    mime=file_info.get("mime", "application/octet-stream"),
                                    key=f"dl_{task['id']}_{i}"
                                )
                        elif not msg:
                             # Si es un diccionario vacío o sin claves conocidas, mostrar algo genérico
                             st.write("✅ Finalizado")

                    else:
                        # Legacy support
                        file_data = None
                        msg = ""
                        
                        # Intentar desempaquetar (data, msg) si es tupla y el segundo es str
                        if isinstance(res, tuple) and len(res) == 2 and isinstance(res[1], str):
                            file_data, msg = res
                        else:
                            file_data = res
                            
                        if msg:
                            st.caption(f"Resultado: {msg}")
                            
                        if file_data:
                            if hasattr(file_data, "getvalue"):
                                file_data = file_data.getvalue()
                                
                            if isinstance(file_data, bytes):
                                st.download_button(
                                    label="📥 Descargar Resultado",
                                    data=file_data,
                                    file_name=f"Resultado_{task['name']}.xlsx", # Default name
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    key=f"dl_{task['id']}"
                                )
                            else:
                                 st.write("✅ Finalizado (Sin archivo descargable)")
                        else:
                            # Si no hay msg y no hay file_data (y no entró en el if dict de arriba)
                            st.write("✅ Finalizado")

                elif status == "FAILED":
                    st.error(f"Falló: {task.get('error')}")
                    
                if st.button("Limpiar", key=f"clr_{task['id']}", help="Eliminar de la lista"):
                    del st.session_state.bg_tasks[task['id']]
                    st.rerun()
                    
                st.divider()
            
    if has_running:
        time.sleep(1)
        st.rerun()
