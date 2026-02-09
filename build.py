import os
import shutil
import subprocess
import sys

def run_command(command):
    print(f"Running: {command}")
    process = subprocess.run(command, shell=True)
    if process.returncode != 0:
        print(f"Error running command: {command}")
        sys.exit(1)

def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    src_dir = os.path.join(base_dir, "src")
    dist_dir = os.path.join(base_dir, "dist")
    build_dir = os.path.join(base_dir, "build")
    assets_dir = os.path.join(base_dir, "assets")
    streamlit_config = os.path.join(base_dir, ".streamlit", "config.toml")

    # Clean previous builds
    if os.path.exists(dist_dir):
        shutil.rmtree(dist_dir)
    if os.path.exists(build_dir):
        shutil.rmtree(build_dir)

    # Create a clean src directory for bundling
    clean_src_dir = os.path.join(build_dir, "src")
    if os.path.exists(clean_src_dir):
        shutil.rmtree(clean_src_dir)
    
    # Copy src to clean_src, excluding unwanted files
    shutil.copytree(src_dir, clean_src_dir, ignore=shutil.ignore_patterns('__pycache__', 'temp_sessions', 'temp_uploads', 'venv', '.git', '*.pyc', '*.zip', 'build_exe', 'dist_exe', '~$*', '*.tmp'))

    # Create a clean assets directory for bundling (exclude temp files like ~$*.xlsx)
    clean_assets_dir = os.path.join(build_dir, "assets")
    if os.path.exists(clean_assets_dir):
        shutil.rmtree(clean_assets_dir)
    
    shutil.copytree(assets_dir, clean_assets_dir, ignore=shutil.ignore_patterns('~$*', '*.tmp', '*.lock'))

    print("--- Building CDO_Cliente.exe (Native Client) ---")
    
    # Arguments for CDO_Cliente
    # We include the clean_src directory so app_web.py and its dependencies are available.
    # They will be placed in 'src' folder inside the EXE bundle.
    
    pyinstaller_cmd_client = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",
        "--onefile",
        "--uac-admin",
        "--windowed", # Start without console
        "--name", "CDO_Cliente",
        "--add-data", f"{clean_src_dir};src",
        "--add-data", f"{clean_assets_dir};assets",
        "--add-data", f"{streamlit_config};.streamlit",
        # Collect all streamlit dependencies
        "--collect-all", "streamlit",
        "--collect-all", "altair",
        "--collect-all", "pandas",
        "--collect-all", "pydeck",
        "--collect-all", "google.generativeai",
        "--collect-all", "google.ai",
        "--collect-all", "google.api_core",
        "--collect-all", "grpc",
        "--collect-all", "PIL",
        "--collect-all", "openpyxl",
        "--collect-all", "docx",
        "--collect-all", "pdf2docx",
        "--collect-all", "fitz", # PyMuPDF
        "--collect-all", "selenium",
        "--collect-all", "webdriver_manager",
        "--collect-all", "docx2pdf",
        "--collect-all", "plotly",
        "--collect-all", "streamlit_elements",
        "--hidden-import", "compileall",
        "--hidden-import", "send2trash",
        "--hidden-import", "streamlit_elements",
        "--copy-metadata", "streamlit",
        "--copy-metadata", "requests",
        "--copy-metadata", "packaging",
        "--hidden-import", "tornado",
        "--hidden-import", "watchdog",
        "--hidden-import", "smmap",
        "--hidden-import", "pydeck",
        "--hidden-import", "blinker",
        "--hidden-import", "cachetools",
        "--hidden-import", "rich",
        "--hidden-import", "tenacity",
        "--hidden-import", "pyarrow",
        "--hidden-import", "streamlit_option_menu",
        "--hidden-import", "st_aggrid",
        "--hidden-import", "pdfplumber",
        "--hidden-import", "docx",
        "--hidden-import", "openpyxl",
        "--hidden-import", "plotly",
        "--hidden-import", "requests",
        "--hidden-import", "packaging",
        os.path.join(src_dir, "run_native.py")
    ]
    
    run_command(" ".join(pyinstaller_cmd_client))

    client_exe = os.path.join(dist_dir, "CDO_Cliente.exe")
    if not os.path.exists(client_exe):
        print("Error: CDO_Cliente.exe was not created.")
        sys.exit(1)

    print("--- Building CDO_Instalador.exe (Installer) ---")

    pyinstaller_cmd_installer = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",
        "--onefile",
        "--uac-admin",
        "--windowed",
        "--name", "CDO_Instalador",
        "--add-data", f"{client_exe};.",
        "--add-data", f"{clean_assets_dir};assets",
        "--add-data", f"{os.path.join(base_dir, 'requirements.txt')};.",
        "--add-data", f"{os.path.join(base_dir, 'install_service.ps1')};.",
        os.path.join(base_dir, "setup_wizard.py")
    ]

    run_command(" ".join(pyinstaller_cmd_installer))

    print("--- Build Complete ---")
    print(f"Installer created at: {os.path.join(dist_dir, 'CDO_Instalador.exe')}")

if __name__ == "__main__":
    main()
