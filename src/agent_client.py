import requests
import json
import os

AGENT_URL = "http://localhost:8989"

def is_agent_available():
    """Check if the local agent is running."""
    try:
        response = requests.get(f"{AGENT_URL}/health", timeout=0.5)
        return response.status_code == 200
    except requests.RequestException:
        return False

def select_folder():
    """Request the agent to open a folder selection dialog."""
    try:
        # Reduced timeout to avoid hanging the UI for too long if agent is unresponsive
        # The agent should respond quickly once the dialog is opened, but if the user
        # takes too long, we might timeout. Ideally, the agent returns "pending" or uses websockets,
        # but for simple HTTP, we'll use a moderate timeout.
        response = requests.post(f"{AGENT_URL}/select-folder", timeout=10) 
        if response.status_code == 200:
            data = response.json()
            if data.get("cancelled"):
                return None
            return data.get("path")
        return None
    except requests.RequestException:
        return None

def list_files(path):
    """Request the agent to list files in a directory."""
    try:
        response = requests.get(f"{AGENT_URL}/list-files", params={"path": path}, timeout=5)
        if response.status_code == 200:
            return response.json().get("files", [])
        return []
    except requests.RequestException:
        return []
