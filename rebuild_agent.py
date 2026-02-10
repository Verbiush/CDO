
import os
import sys
import shutil
import subprocess

def run_command(command, cwd=None):
    print(f"Running: {' '.join(command) if isinstance(command, list) else command}")
    try:
        subprocess.run(command, cwd=cwd, check=True)
    except subprocess.CalledProcessError as e:
        print(f"Error executing command: {e}")
        sys.exit(1)

def main():
    project_root = os.path.dirname(os.path.abspath(__file__))
    src_dir = os.path.join(project_root, "src")
    output_dir = os.path.join(project_root, "build_release_output")
    agent_dist_dir = os.path.join(output_dir, "dist_agent")
    
    # Ensure output directories exist
    os.makedirs(output_dir, exist_ok=True)
    os.makedirs(agent_dist_dir, exist_ok=True)
    
    # 1. Compile Local Agent (CDO_Agente.exe)
    print("--- 1. Compiling Local Agent (CDO_Agente.exe) ---")
    agent_script = os.path.join(src_dir, "local_agent", "main.py")
    modules_dir = os.path.join(src_dir, "modules")
    
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
        "--hidden-import", "pandas",
        "--hidden-import", "selenium",
        "--add-data", f"{modules_dir}{os.pathsep}modules",
        agent_script
    ]
    
    run_command(agent_cmd)
    
    agent_exe_path = os.path.join(agent_dist_dir, "CDO_Agente.exe")
    if not os.path.exists(agent_exe_path):
        print("Error: CDO_Agente.exe was not created.")
        sys.exit(1)

    # 2. Compile Agent Installer (Instalar_Agente.exe)
    print("--- 2. Compiling Agent Installer (Instalar_Agente.exe) ---")
    agent_setup_script = os.path.join(src_dir, "local_agent", "setup_agent.py")
    
    # Need to have CDO_Agente.exe available for bundling
    # PyInstaller needs the file to exist when --add-data is called
    # We use the one we just built
    
    setup_cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",
        "--onefile",
        "--noconsole", # GUI mode (Tkinter)
        "--name", "Instalar_Agente",
        "--clean",
        "--workpath", os.path.join(output_dir, "build_agent_setup"),
        "--distpath", agent_dist_dir,
        "--add-data", f"{agent_exe_path}{os.pathsep}.",
        agent_setup_script
    ]
    
    run_command(setup_cmd)
    
    setup_exe_path = os.path.join(agent_dist_dir, "Instalar_Agente.exe")
    if os.path.exists(setup_exe_path):
        print(f"Success! Installer created at: {setup_exe_path}")
        
        # Copy to src so the web app can find it
        dest_in_src = os.path.join(src_dir, "Instalar_Agente.exe")
        shutil.copy2(setup_exe_path, dest_in_src)
        print(f"Copied to {dest_in_src}")

if __name__ == "__main__":
    main()
