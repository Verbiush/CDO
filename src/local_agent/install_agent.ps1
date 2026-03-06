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
pip install requests selenium webdriver-manager pandas

# 2. Create Launcher Script (Wrapper to run main.py)
$agentScript = @"
import os
import sys
import subprocess
import time

# Wrapper to keep the window open or handle restarts
# For the new architecture, we just run main.py
# If it crashes, we might want to restart it.

current_file = os.path.abspath(__file__)
agent_dir = os.path.dirname(current_file) # src/local_agent
src_dir = os.path.dirname(agent_dir)      # src
if src_dir not in sys.path:
    sys.path.insert(0, src_dir)

from local_agent.main import main

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        pass
    except Exception as e:
        print(f"Critical Error: {e}")
        input("Press Enter to exit...")
"@

Set-Content -Path $launcherPath -Value $agentScript
Write-Host "Launcher created at $launcherPath"

# 3. Create Desktop Shortcut (for easy access/config)
$desktopDir = [Environment]::GetFolderPath("Desktop")
$shortcutPath = Join-Path $desktopDir "CDO_Agent.lnk"
$pythonPath = (Get-Command python).Source
$pythonExec = "python.exe" # Use console version so user can see prompts/logs

$WshShell = New-Object -comObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut($shortcutPath)
$Shortcut.TargetPath = $pythonExec
$Shortcut.Arguments = "`"$launcherPath`""
$Shortcut.WorkingDirectory = $agentDir
$Shortcut.Description = "CDO Local Agent"
$Shortcut.Save()

Write-Host "Desktop shortcut created at $shortcutPath"
Write-Host "Installation Complete!"
Write-Host "Please double-click the CDO_Agent shortcut on your desktop to configure and start the agent."
