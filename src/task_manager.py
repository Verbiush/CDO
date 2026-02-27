import time
import streamlit as st

def init_task_manager():
    """Inicializa la estructura de tareas en session_state si no existe."""
    if "bg_tasks" not in st.session_state:
        st.session_state.bg_tasks = {}



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


def render_task_items(container, tasks, key_prefix=""):
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
                            key=f"{key_prefix}dl_{task['id']}_{i}"
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
                            key=f"{key_prefix}dl_{task['id']}"
                        )
                    else:
                         container.write("✅ Finalizado (Sin archivo descargable)")
                else:
                    # Si no hay msg y no hay file_data (y no entró en el if dict de arriba)
                    container.write("✅ Finalizado")

        elif status == "FAILED":
            container.error(f"Falló: {task.get('error')}")
            
        if container.button("Limpiar", key=f"{key_prefix}clr_{task['id']}", help="Eliminar de la lista"):
            del st.session_state.bg_tasks[task['id']]
            st.rerun()
            
        container.divider()

def render_task_center(container=None, key_prefix=""):
    """Renderiza el centro de tareas en el contenedor dado (o sidebar por defecto)."""
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

        if container:
            # Render directly into the provided container
            render_task_items(container, tasks[:5], key_prefix=key_prefix)
        else:
            # Default to sidebar expander if no container provided
            with st.sidebar.expander("🔔 Centro de Tareas", expanded=True):
                render_task_items(st.sidebar, tasks[:5], key_prefix=f"{key_prefix}sidebar_")
        if has_running:
            time.sleep(0.1)
            st.rerun()
