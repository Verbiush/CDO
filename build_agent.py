
import PyInstaller.__main__
import os
import shutil
import sys

# Define paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SRC_DIR = os.path.join(BASE_DIR, 'src')
AGENT_SCRIPT = os.path.join(SRC_DIR, 'local_agent', 'main.py')
DIST_DIR = os.path.join(BASE_DIR, 'dist')
BUILD_DIR = os.path.join(BASE_DIR, 'build')

# Clean previous builds
if os.path.exists(DIST_DIR):
    shutil.rmtree(DIST_DIR)
if os.path.exists(BUILD_DIR):
    shutil.rmtree(BUILD_DIR)

print(f"Building agent from: {AGENT_SCRIPT}")

# Build arguments
args = [
    AGENT_SCRIPT,
    '--onefile',
    '--name=CDO_Agente',
    f'--paths={SRC_DIR}',  # Help PyInstaller find 'modules'
    '--clean',
    '--noconfirm',
    # '--noconsole', # Uncomment for production/silent mode
    '--hidden-import=modules',
    '--hidden-import=modules.ovida_validator',
    '--hidden-import=modules.registraduria_validator',
    '--hidden-import=modules.adres_validator',
    '--hidden-import=win32timezone', # Often needed for datetime
    '--hidden-import=src.tabs.tab_conversion',
    '--hidden-import=src.tabs.tab_automated_actions',
    '--hidden-import=src.modules.analisis_sos',
    '--hidden-import=src.gui_utils',
    '--exclude-module=streamlit'
]

# Run PyInstaller
try:
    PyInstaller.__main__.run(args)
    print("Build successful! Executable is in 'dist/CDO_Agente.exe'")
    
    # Copy to src/local_agent for easy access by web app
    dest = os.path.join(SRC_DIR, 'local_agent', 'CDO_Agente.exe')
    shutil.copy2(os.path.join(DIST_DIR, 'CDO_Agente.exe'), dest)
    print(f"Copied to: {dest}")
    
except Exception as e:
    print(f"Build failed: {e}")
    sys.exit(1)
