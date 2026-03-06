# Install Script for CDO Local Agent
# Run as Administrator in PowerShell

$ErrorActionPreference = "Stop"

Write-Host "Installing CDO Local Agent..." -ForegroundColor Cyan

# Get Current Directory (Absolute Path)
$currentDir = Get-Location
$srcDir = Join-Path $currentDir "src"
$agentDir = Join-Path $srcDir "local_agent"
$launcherPath = Join-Path $agentDir "launcher.py"

# 1. Install Dependencies
Write-Host "Installing Python dependencies..."
pip install fastapi uvicorn cryptography pystray pillow selenium webdriver-manager pandas

# 2. Generate Certificate
Write-Host "Generating SSL Certificate..."
# Ensure we run cert_gen from the right context or path
python "$agentDir\cert_gen.py"

# 3. Trust Certificate (Root Store)
$certPath = Join-Path $currentDir "server.crt"
if (Test-Path $certPath) {
    Write-Host "Adding Certificate to Trusted Root Store..."
    Import-Certificate -FilePath $certPath -CertStoreLocation Cert:\LocalMachine\Root
} else {
    Write-Host "Warning: Certificate not found at $certPath" -ForegroundColor Yellow
}

# 4. Create Launcher Script
$agentScript = @"
import os
import sys
import uvicorn
import pystray
from PIL import Image, ImageDraw
import threading
import webbrowser
import time

# Add src to path so we can import local_agent.main
# We assume this script is located in src/local_agent/
current_file = os.path.abspath(__file__)
agent_dir = os.path.dirname(current_file) # src/local_agent
src_dir = os.path.dirname(agent_dir)      # src
if src_dir not in sys.path:
    sys.path.insert(0, src_dir)

from local_agent.main import app

def run_server():
    # Run with SSL
    # Cert files are expected to be in the project root or relative to execution
    # We will try to find them
    project_root = os.path.dirname(src_dir)
    key_file = os.path.join(project_root, "server.key")
    cert_file = os.path.join(project_root, "server.crt")
    
    if not os.path.exists(key_file):
        print(f"Key file not found: {key_file}")
        return

    uvicorn.run(app, host="127.0.0.1", port=12345, ssl_keyfile=key_file, ssl_certfile=cert_file)

def on_quit(icon, item):
    icon.stop()
    os._exit(0)

def open_dashboard(icon, item):
    webbrowser.open("https://localhost:12345/docs")

def create_image():
    # Create an icon image
    width = 64
    height = 64
    image = Image.new('RGB', (width, height), (255, 255, 255))
    dc = ImageDraw.Draw(image)
    dc.rectangle((0, 0, width, height), fill="blue")
    dc.rectangle((10, 10, width-10, height-10), fill="white")
    return image

if __name__ == "__main__":
    # Start server in thread
    server_thread = threading.Thread(target=run_server, daemon=True)
    server_thread.start()
    
    # System Tray Icon
    icon = pystray.Icon("CDO Agent", create_image(), "CDO Local Agent", menu=pystray.Menu(
        pystray.MenuItem("Dashboard", open_dashboard),
        pystray.MenuItem("Exit", on_quit)
    ))
    icon.run()
"@

Set-Content -Path $launcherPath -Value $agentScript
Write-Host "Launcher created at $launcherPath"

# 5. Create Startup Shortcut
$startupDir = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup"
$shortcutPath = Join-Path $startupDir "CDO_Local_Agent.lnk"
$pythonPath = (Get-Command python).Source
# Use pythonw to avoid console window if possible, but python is safer for debugging initially.
# Let's use pythonw if available, else python.
if (Get-Command pythonw -ErrorAction SilentlyContinue) {
    $pythonExec = "pythonw.exe"
} else {
    $pythonExec = "python.exe"
}

$WshShell = New-Object -comObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut($shortcutPath)
$Shortcut.TargetPath = $pythonExec
$Shortcut.Arguments = "`"$launcherPath`""
$Shortcut.WorkingDirectory = $currentDir
$Shortcut.Description = "CDO Local Agent Service"
$Shortcut.Save()

Write-Host "Startup shortcut created at $shortcutPath"
Write-Host "Installation Complete!"
Write-Host "The agent will start automatically on next login."
Write-Host "To start it now, run: python $launcherPath"
