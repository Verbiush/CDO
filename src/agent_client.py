import time
import json
import streamlit as st
try:
    import database
except ImportError:
    from src import database

# Timeout in seconds for waiting for agent response
AGENT_TIMEOUT = 60

def is_agent_active(username):
    """
    Checks if the user has an active agent.
    For now, we just check if there are recent tasks or if we want to trust the user.
    A better way would be a 'heartbeat' table, but for now we assume 
    if the user clicks 'Use Agent', they have it running.
    """
    return True

def send_command(username, command, params=None):
    """
    Creates a task in the database for the agent to execute.
    """
    success, task_id = database.create_task(username, command, params)
    return task_id if success else None

def wait_for_result(task_id, timeout=AGENT_TIMEOUT):
    """
    Polls the database for the result of a task.
    """
    start_time = time.time()
    while (time.time() - start_time) < timeout:
        task_data = database.get_task_result(task_id)
        if task_data:
            status = task_data.get("status")
            if status == "COMPLETED":
                return task_data.get("result")
            elif status == "ERROR":
                return {"error": "Agent reported an error"}
        
        time.sleep(1)
        
    return {"error": "Timeout waiting for agent response"}

def select_folder(username, title="Seleccionar Carpeta"):
    """
    Request the agent to open a folder selection dialog on the user's PC.
    """
    print(f"DEBUG: Requesting select_folder for user {username}")
    task_id = send_command(username, "browse_folder", {"title": title})
    if not task_id:
        print("DEBUG: Failed to create task (task_id is None)")
        return None
    
    print(f"DEBUG: Task created with ID {task_id}. Waiting for result...")
    with st.spinner("Esperando a que seleccione la carpeta en su PC..."):
        result = wait_for_result(task_id)
    
    print(f"DEBUG: Result received: {result}")
        
    if result and "path" in result:
        return result["path"]
    return None

def select_file(username, title="Seleccionar Archivo", file_types=None):
    """
    Request the agent to open a file selection dialog.
    """
    params = {"title": title}
    if file_types:
        params["file_types"] = file_types
        
    task_id = send_command(username, "browse_file", params)
    if not task_id:
        return None
    
    with st.spinner("Esperando a que seleccione el archivo en su PC..."):
        result = wait_for_result(task_id)
        
    if result and "path" in result:
        return result["path"]
    return None

def list_drives(username):
    """
    Request the agent to list drives.
    """
    task_id = send_command(username, "list_drives")
    if not task_id:
        return []
    
    result = wait_for_result(task_id, timeout=10)
    if result and "drives" in result:
        return result["drives"]
    return []

def list_files(username, path):
    """
    Request the agent to list files in a directory.
    """
    task_id = send_command(username, "list_files", {"path": path})
    if not task_id:
        return None
    
    # Wait up to 15 seconds
    result = wait_for_result(task_id, timeout=15)
    
    if result and "files" in result:
        return result["files"]
    return None
