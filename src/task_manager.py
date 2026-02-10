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
            # Give Streamlit a moment to render the UI response before starting heavy work
            time.sleep(1.0)
            
            # Update status to running
            if task_id in st.session_state.bg_tasks:
                st.session_state.bg_tasks[task_id]["status"] = "RUNNING"
            
            # Run function
            import inspect
            sig = inspect.signature(target_func)
            
            # Inject task_id or update_progress if requested
            func_kwargs = kwargs.copy()
            if "task_id" in sig.parameters:
                func_kwargs["task_id"] = task_id
            if "update_progress" in sig.parameters:
                def update_progress_callback(current, total, message=None):
                    if task_id in st.session_state.bg_tasks:
                        if total > 0:
                            p = int((current / total) * 100)
                        else:
                            p = 0
                        st.session_state.bg_tasks[task_id]["progress"] = p
                        if message:
                            st.session_state.bg_tasks[task_id]["status_message"] = message
                func_kwargs["update_progress"] = update_progress_callback
                
            result = target_func(*args, **func_kwargs)
            
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

def render_task_items(container, tasks):
    """Renderiza la lista de tareas en un contenedor dado."""
    for task in tasks:
        container.markdown(f"**{task['name']}**")
        status = task["status"]
        
        if status == "RUNNING" or status == "PENDING":
            if status == "RUNNING":
                container.info("Ejecutando... Puedes cambiar de pestaña.", icon="⏳")
                # Show status message if available
                status_msg = task.get("status_message")
                if status_msg:
                    container.caption(status_msg)
                # Show progress bar if available
                progress = task.get("progress", 0)
                container.progress(progress)
            else:
                container.text("Pendiente...")
                
        elif status == "COMPLETED":
            # Show result
            res = task.get("result")
            
            # Support for standardized dict format: {"files": [{"name":..., "data":..., "label":...}], "message":...}
            if isinstance(res, dict):
                msg = res.get("message", "")
                if msg: 
                    container.info(f"Resultado:\n{msg}")
                
                if "files" in res:
                    for i, file_info in enumerate(res["files"]):
                        data = file_info["data"]
                        if hasattr(data, "getvalue"): data = data.getvalue()
                        
                        container.download_button(
                            label=f"📥 {file_info.get('label', 'Descargar')}",
                            data=data,
                            file_name=file_info["name"],
                            mime=file_info.get("mime", "application/octet-stream"),
                            key=f"dl_{task['id']}_{i}"
                        )
                elif not msg:
                     # Si es un diccionario vacío o sin claves conocidas, mostrar algo genérico
                     container.write("✅ Finalizado")

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
                    container.caption(f"Resultado: {msg}")
                    
                if file_data:
                    if hasattr(file_data, "getvalue"):
                        file_data = file_data.getvalue()
                        
                    if isinstance(file_data, bytes):
                        container.download_button(
                            label="📥 Descargar Resultado",
                            data=file_data,
                            file_name=f"Resultado_{task['name']}.xlsx", # Default name
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"dl_{task['id']}"
                        )
                    else:
                         container.write("✅ Finalizado (Sin archivo descargable)")
                else:
                    # Si no hay msg y no hay file_data (y no entró en el if dict de arriba)
                    container.write("✅ Finalizado")

        elif status == "FAILED":
            container.error(f"Falló: {task.get('error')}")
            
        if container.button("Limpiar", key=f"clr_{task['id']}", help="Eliminar de la lista"):
            del st.session_state.bg_tasks[task['id']]
            st.rerun()
            
        container.divider()

def render_task_center():
    """Renderiza el centro de tareas en la barra lateral."""
    if "bg_tasks" not in st.session_state:
        return

    # Auto-refresh si hay tareas corriendo
    has_running = False
    if st.session_state.bg_tasks:
        # Sort by start time desc (newest first)
        tasks = sorted(st.session_state.bg_tasks.values(), key=lambda x: x["start_time"], reverse=True)
        
        # Check if any running to trigger rerun
        for task in tasks:
            if task["status"] in ["RUNNING", "PENDING"]:
                has_running = True
                break

        with st.sidebar.expander("🔔 Centro de Tareas", expanded=True):
            # Limit to last 5 tasks for sidebar to avoid clutter
            render_task_items(st.sidebar, tasks[:5])
            
    if has_running:
        time.sleep(1)
        st.rerun()
