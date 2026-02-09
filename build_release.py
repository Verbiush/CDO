
import os
import sys
import shutil
import subprocess
import compileall

def run_command(command, cwd=None, capture_output=False):
    print(f"Running: {' '.join(command) if isinstance(command, list) else command}")
    try:
        result = subprocess.run(
            command, 
            cwd=cwd, 
            check=True, 
            capture_output=capture_output, 
            text=True if capture_output else False
        )
        return result
    except subprocess.CalledProcessError as e:
        print(f"Error executing command: {e}")
        if capture_output and e.stderr:
            print(f"STDERR: {e.stderr}")
        sys.exit(1)

def sign_exe(target_path, project_root):
    """Firma el ejecutable usando el script PowerShell del proyecto."""
    sign_script = os.path.join(project_root, "sign_code.ps1")
    if os.path.exists(sign_script) and os.path.exists(target_path):
        print(f"✍️ Firmando {os.path.basename(target_path)}...")
        try:
            # Run PowerShell script
            subprocess.run(
                ["powershell", "-ExecutionPolicy", "Bypass", "-File", sign_script, "-TargetFile", target_path], 
                check=True, 
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            print(f"✅ Firma aplicada a {os.path.basename(target_path)}")
        except Exception as e:
            print(f"⚠️ Advertencia: No se pudo firmar el código. {e}")
    else:
        print(f"⚠️ Script de firma no encontrado: {sign_script}")

def main():
    project_root = os.path.dirname(os.path.abspath(__file__))
    output_dir = os.path.join(project_root, "build_release_output")
    
    # Directorio temporal de construcción
    build_dir = os.path.join(output_dir, "build_temp")
    dist_dir = os.path.join(output_dir, "dist")
    
    # Limpieza previa
    for d in [build_dir, dist_dir]:
        if os.path.exists(d):
            try: shutil.rmtree(d)
            except: pass
    os.makedirs(build_dir, exist_ok=True)
    os.makedirs(dist_dir, exist_ok=True)

    print("--- Preparing Build Directory ---")

    # 1. Copiar Archivos Raíz necesarios para el build
    files_to_include_root = ["setup_wizard.py", "requirements.txt"]
    for f in files_to_include_root:
        src = os.path.join(project_root, f)
        dst = os.path.join(build_dir, f)
        if os.path.exists(src):
            shutil.copy2(src, dst)
    
    # 2. Copiar Assets
    assets_src = os.path.join(project_root, "assets")
    assets_dst = os.path.join(build_dir, "assets")
    if os.path.exists(assets_src):
        shutil.copytree(assets_src, assets_dst, ignore=shutil.ignore_patterns('~$*', '*.tmp'))

    # 3. Copiar y Compilar SRC
    src_origin = os.path.join(project_root, "src")
    src_dest = os.path.join(build_dir, "src")
    
    # Copiar todo src
    shutil.copytree(src_origin, src_dest, ignore=shutil.ignore_patterns('__pycache__', 'temp_sessions', 'temp_uploads', 'venv', '.git', '*.pyc', '*.zip', 'build_exe', 'dist_exe', '~$*', '*.tmp'))
    
    # Protección de Código: Compilar a .pyc y eliminar .py
    print("--- Applying Code Protection (Compile to .pyc) ---")
    app_web_path = os.path.join(src_dest, "app_web.py")
    if os.path.exists(app_web_path):
        core_path = os.path.join(src_dest, "_core_app.py")
        os.rename(app_web_path, core_path)
        
        compileall.compile_dir(src_dest, force=True, legacy=True, quiet=1)
        
        for root, dirs, files in os.walk(src_dest):
            for file in files:
                if file.endswith(".py"):
                    # Keep entry points
                    if file in ["run_native.py"]:
                        continue
                    os.remove(os.path.join(root, file))
        
        with open(app_web_path, "w") as f:
            f.write("# Launcher protegido\n")
            f.write("import _core_app\n")
    
    # 4. Compilar en TRES pasos:
    
    # Paso A: Compilar Agente Local (CDO_Agente.exe)
    print("--- Step 1/3: Compiling Local Agent (CDO_Agente.exe) ---")
    agent_dist_dir = os.path.join(output_dir, "dist_agent")
    
    # Check if we should use .py or .pyc
    agent_script = os.path.join(src_dest, "local_agent", "main.py")
    if not os.path.exists(agent_script):
         # If .py was compiled/removed, try .pyc (though PyInstaller usually needs .py entry point)
         # In our code protection step, we removed .py files. 
         # But for PyInstaller entry points, we need them or need to handle it.
         # Let's restore main.py for agent compilation if missing
         original_agent_script = os.path.join(src_origin, "local_agent", "main.py")
         if os.path.exists(original_agent_script):
             shutil.copy2(original_agent_script, agent_script)
         else:
             print(f"Error: Agent script source not found at {original_agent_script}")
             sys.exit(1)

    agent_cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",
        "--onefile",
        "--windowed", # Run as windowed (hidden console)
        "--name", "CDO_Agente",
        "--clean",
        "--workpath", os.path.join(output_dir, "build_agent"),
        "--distpath", agent_dist_dir,
        "--hidden-import", "fastapi",
        "--hidden-import", "uvicorn",
        "--hidden-import", "tkinter",
        agent_script
    ]
    
    run_command(agent_cmd, cwd=build_dir)
    
    agent_exe_path = os.path.join(agent_dist_dir, "CDO_Agente.exe")
    if not os.path.exists(agent_exe_path):
        print("Error: CDO_Agente.exe was not created.")
        sys.exit(1)
        
    sign_exe(agent_exe_path, project_root)

    # Paso A.2: Compilar Instalador del Agente (Instalar_Agente.exe) - Standalone
    print("--- Step 1.5/3: Compiling Agent Installer (Instalar_Agente.exe) ---")
    agent_setup_script = os.path.join(src_dest, "local_agent", "setup_agent.py")
    if not os.path.exists(agent_setup_script):
         original_setup_script = os.path.join(src_origin, "local_agent", "setup_agent.py")
         if os.path.exists(original_setup_script):
             shutil.copy2(original_setup_script, agent_setup_script)

    # Copiar CDO_Agente.exe al build dir para empaquetarlo
    shutil.copy2(agent_exe_path, os.path.join(output_dir, "build_agent", "CDO_Agente.exe"))
    
    setup_cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",
        "--onefile",
        "--console", # Console visible for status messages
        "--name", "Instalar_Agente",
        "--clean",
        "--workpath", os.path.join(output_dir, "build_agent_setup"),
        "--distpath", agent_dist_dir, # Same dist dir
        "--add-data", f"{os.path.join(output_dir, 'build_agent', 'CDO_Agente.exe')}{os.pathsep}.",
        agent_setup_script
    ]
    
    run_command(setup_cmd, cwd=build_dir)
    
    setup_exe_path = os.path.join(agent_dist_dir, "Instalar_Agente.exe")
    if os.path.exists(setup_exe_path):
        sign_exe(setup_exe_path, project_root)
        # Copiar a carpeta dist principal para fácil acceso
        shutil.copy2(setup_exe_path, os.path.join(dist_dir, "Instalar_Agente.exe"))
        # Copiar también a src root para que app_web lo encuentre en dev
        dev_path = os.path.join(project_root, "src", "Instalar_Agente.exe")
        shutil.copy2(setup_exe_path, dev_path)

    # Paso B: Compilar Cliente Nativo (CDO_Cliente.exe)
    print("--- Step 2/3: Compiling Native Client (CDO_Cliente.exe) ---")
    
    client_dist_dir = os.path.join(output_dir, "dist_client")
    client_script = os.path.join(src_dest, "run_native.py")
    
    # Asegurar que streamlit tenga sus metadatos
    client_cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",
        "--onefile",
        "--windowed",
        "--name", "CDO_Cliente",
        "--clean",
        "--workpath", os.path.join(output_dir, "build_client"),
        "--distpath", client_dist_dir,
        "--specpath", os.path.join(output_dir, "spec_client"),
        
        # Streamlit Critical Hooks
        "--copy-metadata", "streamlit",
        "--copy-metadata", "google-generativeai",
        "--recursive-copy-metadata", "streamlit",
        
        # Exclude Heavy Unused Modules
        "--exclude-module", "matplotlib",
        "--exclude-module", "scipy",
        "--exclude-module", "notebook",
        "--exclude-module", "jupyter",
        "--exclude-module", "ipython",
        "--exclude-module", "bokeh",
        "--exclude-module", "plotly",
        
        # Hidden Imports
        "--hidden-import", "streamlit",
        "--hidden-import", "streamlit.web.cli",
        "--hidden-import", "pandas",
        "--hidden-import", "altair",
        "--hidden-import", "pydeck",
        "--hidden-import", "rich",
        "--hidden-import", "watchdog",
        "--hidden-import", "tkinter",
        "--hidden-import", "PIL",
        "--hidden-import", "PIL.Image",
        
        # Data: Include SRC (app logic) and Assets
        "--add-data", f"{src_dest}{os.pathsep}src", 
        "--add-data", f"{assets_dst}{os.pathsep}assets",
        
        client_script
    ]
    
    run_command(client_cmd, cwd=build_dir)
        
    client_exe_path = os.path.join(client_dist_dir, "CDO_Cliente.exe")
    if not os.path.exists(client_exe_path):
        print("Error: CDO_Cliente.exe was not created.")
        sys.exit(1)
        
    # FIRMAR CLIENTE
    sign_exe(client_exe_path, project_root)

    # Paso C: Compilar Instalador
    print("--- Step 3/3: Compiling Installer (Instalador_CDO.exe) ---")
    
    # Mover Cliente compilado al directorio de build del instalador para empaquetarlo
    shutil.copy2(client_exe_path, os.path.join(build_dir, "CDO_Cliente.exe"))
    # Mover Agente compilado al directorio de build del instalador
    shutil.copy2(agent_exe_path, os.path.join(build_dir, "CDO_Agente.exe"))
    
    installer_script = os.path.join(build_dir, "setup_wizard.py")
    exe_name = "Instalador_CDO"
    
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",
        "--onefile",
        "--windowed",
        "--name", exe_name,
        "--clean",
        "--workpath", build_dir,
        "--distpath", dist_dir,
        "--specpath", build_dir,
        # Bundle the Client EXE inside the Installer
        "--add-data", f"CDO_Cliente.exe{os.pathsep}.",
        # Bundle the Agent EXE inside the Installer
        "--add-data", f"CDO_Agente.exe{os.pathsep}.",
        # Bundle assets for the installer UI itself
        "--add-data", f"assets{os.pathsep}assets",
        "--hidden-import", "tkinter",
        installer_script
    ]
    
    # Ejecutar compilación del instalador
    run_command(cmd, cwd=build_dir)
    
    exe_path = os.path.join(dist_dir, f"{exe_name}.exe")
    
    if not os.path.exists(exe_path):
        print("Error: Installer EXE was not created.")
        sys.exit(1)
        
    # FIRMAR INSTALADOR
    sign_exe(exe_path, project_root)
    
    print(f"--- Build Success! ---")
    print(f"Installer: {exe_path}")

if __name__ == "__main__":
    main()
