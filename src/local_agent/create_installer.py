import os
import subprocess
import shutil
import sys

def build_installer():
    print("=== Building CDO Local Agent Installer ===")
    
    current_dir = os.path.dirname(os.path.abspath(__file__))
    dist_dir = os.path.join(current_dir, "dist")
    build_dir = os.path.join(current_dir, "build")
    
    # Clean previous builds
    # if os.path.exists(dist_dir): shutil.rmtree(dist_dir)
    # if os.path.exists(build_dir):
    #     shutil.rmtree(build_dir)
        
    # --- Helper to Sign Code ---
    def sign_executable(file_path):
        print(f"Signing {file_path}...")
        sign_script = os.path.abspath(os.path.join(current_dir, "..", "..", "sign_code.ps1"))
        if os.path.exists(sign_script):
            try:
                subprocess.check_call([
                    "powershell",
                    "-ExecutionPolicy", "Bypass",
                    "-File", sign_script,
                    "-TargetFile", file_path
                ])
                print("Signature applied successfully.")
            except Exception as e:
                print(f"Warning: Failed to sign executable: {e}")
        else:
            print(f"Warning: Sign script not found at {sign_script}")

    print("1. Installing PyInstaller...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
    
    print("\n2. Building Agent Executable (CDO_Agente.exe)...")
    
    # Build main.py -> CDO_Agente.exe (noconsole)
    # Hidden imports for FastAPI/Uvicorn
    hidden_imports = [
        "--hidden-import=uvicorn.logging",
        "--hidden-import=uvicorn.loops",
        "--hidden-import=uvicorn.loops.auto",
        "--hidden-import=uvicorn.protocols",
        "--hidden-import=uvicorn.protocols.http",
        "--hidden-import=uvicorn.protocols.http.auto",
        "--hidden-import=uvicorn.lifespan.on",
        "--hidden-import=uvicorn.lifespan.off",
        "--hidden-import=anyio",
        "--hidden-import=starlette",
        "--hidden-import=fastapi",
        "--hidden-import=email_validator",
        "--hidden-import=pydantic",
        "--hidden-import=pydantic.deprecated.decorator",
        "--hidden-import=selenium",
        "--hidden-import=selenium.webdriver",
        "--hidden-import=selenium.webdriver.chrome.service",
        "--hidden-import=selenium.webdriver.common.by",
        "--hidden-import=selenium.webdriver.common.keys",
        "--hidden-import=selenium.webdriver.support.ui",
        "--hidden-import=selenium.webdriver.support.expected_conditions",
        "--hidden-import=selenium.common.exceptions",
        "--hidden-import=webdriver_manager",
        "--hidden-import=webdriver_manager.chrome",
        "--hidden-import=webdriver_manager.microsoft",
        "--hidden-import=pandas"
    ]
    
    # Add parent directory (src) to paths so bot_zeus.py can be found
    src_path = os.path.dirname(current_dir)
    
    cmd_agent = [
        sys.executable, "-m", "PyInstaller",
        "--noconsole",
        "--onefile",
        "--name=CDO_Agente",
        "--clean",
        f"--paths={src_path}",
        os.path.join(current_dir, "main.py")
    ] + hidden_imports
    
    subprocess.check_call(cmd_agent, cwd=current_dir)
    
    agent_exe = os.path.join(dist_dir, "CDO_Agente.exe")
    if not os.path.exists(agent_exe):
        print("Error: CDO_Agente.exe not found!")
        return

    # SIGN INNER EXE
    sign_executable(agent_exe)

    print("\n3. Building Installer (Instalador_Agente_CDO.exe)...")
    # Build setup_agent.py -> Instalador.exe
    # Include CDO_Agente.exe as data
    # On Windows, separator is ;
    add_data = f"{agent_exe};."
    
    cmd_installer = [
        sys.executable, "-m", "PyInstaller",
        "--noconsole",
        "--onefile",
        "--name=Instalador_Agente_CDO",
        "--clean",
        f"--add-data={add_data}",
        "--uac-admin",  # Request admin to write to registry/program files if needed
        os.path.join(current_dir, "setup_agent.py")
    ]
    
    subprocess.check_call(cmd_installer, cwd=current_dir)
    
    installer_exe = os.path.join(dist_dir, "Instalador_Agente_CDO.exe")
    
    if os.path.exists(installer_exe):
        # SIGN OUTER INSTALLER
        sign_executable(installer_exe)

        print(f"\nSUCCESS! Installer created at:\n{installer_exe}")
        
        # Move to root for easier access
        root_dir = os.path.abspath(os.path.join(current_dir, "..", ".."))
        final_path = os.path.join(root_dir, "Instalador_Agente_CDO.exe")
        shutil.copy2(installer_exe, final_path)
        print(f"Copied to root: {final_path}")
    else:
        print("Error: Installer creation failed.")

if __name__ == "__main__":
    build_installer()
