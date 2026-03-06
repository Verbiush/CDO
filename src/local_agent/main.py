import os
import sys
import uvicorn
import shutil
import platform
import subprocess
from fastapi import FastAPI, HTTPException, Body
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import List, Optional

# Add parent directory to path to allow imports from src
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)
if parent_dir not in sys.path:
    sys.path.insert(0, parent_dir)

try:
    from bot_zeus import abrir_navegador_inicial, obtener_driver
except ImportError:
    # Fallback if running standalone
    abrir_navegador_inicial = None
    obtener_driver = None

app = FastAPI(title="CDO Local Agent", version="1.0.0")

# Configure CORS to allow requests from the Web App (AWS)
origins = [
    "https://cdo-aws.com", # Replace with actual domain
    "http://localhost:8501", # Local Streamlit dev
    "*" # For testing, restrict later
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

class PathRequest(BaseModel):
    path: str

class BrowserRequest(BaseModel):
    url: Optional[str] = None
    headless: bool = False

@app.get("/")
def read_root():
    return {"status": "online", "agent": "CDO Local Agent", "system": platform.system()}

@app.get("/fs/drives")
def list_drives():
    """List available drives on Windows."""
    drives = []
    if platform.system() == "Windows":
        import string
        from ctypes import windll
        bitmask = windll.kernel32.GetLogicalDrives()
        for letter in string.ascii_uppercase:
            if bitmask & 1:
                drives.append(f"{letter}:\\")
            bitmask >>= 1
    else:
        drives.append("/")
    return {"drives": drives}

@app.post("/fs/list")
def list_files(request: PathRequest):
    """List files in a directory."""
    path = request.path
    if not os.path.exists(path):
        raise HTTPException(status_code=404, detail="Path not found")
    
    if not os.path.isdir(path):
        raise HTTPException(status_code=400, detail="Path is not a directory")
        
    try:
        items = []
        with os.scandir(path) as it:
            for entry in it:
                items.append({
                    "name": entry.name,
                    "is_dir": entry.is_dir(),
                    "path": entry.path
                })
        return {"items": items}
    except PermissionError:
        raise HTTPException(status_code=403, detail="Permission denied")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/browser/launch")
def launch_browser(request: BrowserRequest):
    """Launch local browser via Selenium."""
    if abrir_navegador_inicial is None:
        raise HTTPException(status_code=500, detail="Bot Zeus module not found")
    
    # Force native mode in a way compatible with bot_zeus logic
    # bot_zeus checks streamlit session state, which we don't have here.
    # We might need to monkeypatch or modify bot_zeus to accept params.
    # For now, let's assume obtaining driver works.
    
    try:
        # TODO: Refactor bot_zeus to accept headless param directly
        driver = obtener_driver() 
        if driver:
            if request.url:
                driver.get(request.url)
            return {"status": "success", "message": "Browser launched"}
        else:
            return {"status": "error", "message": "Failed to launch browser"}
    except Exception as e:
        return {"status": "error", "message": str(e)}

if __name__ == "__main__":
    # SSL is handled by uvicorn arguments or reverse proxy
    # Use 12345 as the agent port
    uvicorn.run(app, host="127.0.0.1", port=12345)
