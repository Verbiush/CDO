import sys
import io
import os

# Fix for PyInstaller --noconsole mode where sys.stdout/stderr are None
# We need a robust mock that includes isatty() because uvicorn checks it.
class NullWriter:
    def write(self, text):
        pass
    def flush(self):
        pass
    def isatty(self):
        return False
    def fileno(self):
        return -1

if sys.stdout is None:
    sys.stdout = NullWriter()
if sys.stderr is None:
    sys.stderr = NullWriter()

import uvicorn
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
import tkinter as tk
from tkinter import filedialog
import threading
import logging

# Configure logging
handlers = [logging.FileHandler("agent.log")]
# Only add StreamHandler if we are not in a frozen noconsole environment
if not getattr(sys, 'frozen', False) or hasattr(sys.stdout, 'isatty'):
    handlers.append(logging.StreamHandler(sys.stdout))

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=handlers
)
logger = logging.getLogger("LocalAgent")

app = FastAPI(title="Local File System Agent", version="1.0.0")

# Configure CORS - allow localhost
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:8501", "http://127.0.0.1:8501", "*"],  # Adjust as needed
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def select_folder_dialog():
    """Open a folder selection dialog in a thread-safe way."""
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    root.attributes('-topmost', True)  # Bring to front
    folder_path = filedialog.askdirectory()
    root.destroy()
    return folder_path

@app.get("/")
def read_root():
    return {"status": "running", "agent": "LocalFileSystemAgent"}

@app.get("/health")
def health_check():
    return {"status": "ok"}

@app.post("/select-folder")
def select_folder():
    """Opens a native folder picker dialog on the host machine."""
    try:
        # Run in a separate thread to not block the event loop
        # However, tkinter needs to run in the main thread usually, 
        # but since this script IS the main process, we might need care.
        # For simplicity in this agent, we'll try running it directly.
        # If it blocks the server, it's fine for a single user agent.
        folder = select_folder_dialog()
        if not folder:
            return {"cancelled": True}
        return {"path": folder, "cancelled": False}
    except Exception as e:
        logger.error(f"Error selecting folder: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/list-files")
def list_files(path: str):
    """List files in the specified directory."""
    if not os.path.exists(path):
        raise HTTPException(status_code=404, detail="Path not found")
    
    if not os.path.isdir(path):
        raise HTTPException(status_code=400, detail="Path is not a directory")

    try:
        files = []
        for entry in os.scandir(path):
            files.append({
                "name": entry.name,
                "is_dir": entry.is_dir(),
                "path": entry.path,
                "size": entry.stat().st_size if not entry.is_dir() else 0
            })
        return {"files": files}
    except Exception as e:
        logger.error(f"Error listing files: {e}")
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    # Run on a specific port, e.g., 8989
    uvicorn.run(app, host="127.0.0.1", port=8989)
