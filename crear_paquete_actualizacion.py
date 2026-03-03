import os
import zipfile
import datetime

def create_update_package():
    # Nombre del archivo con fecha
    date_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    zip_filename = f"actualizacion_cdo_{date_str}.zip"
    
    # Directorios y archivos a excluir
    exclude_dirs = {
        '.git', '.vscode', '__pycache__', 'venv', 'env', 'build', 'dist', 
        'node_modules', 'cdk.out', 'test-results', 'cdk.out'
    }
    exclude_extensions = {'.pyc', '.pyd', '.pyo', '.log', '.zip'}
    
    # Archivos específicos a excluir
    exclude_files = {
        'crear_paquete_actualizacion.py', 
        zip_filename,
        'users.db', # No sobrescribir la base de datos de usuarios en producción
        'user_preferences.json' # No sobrescribir preferencias locales
    }

    print(f"Creando paquete de actualización: {zip_filename}...")
    
    current_dir = os.getcwd()
    
    with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(current_dir):
            # Modificar dirs in-place para saltar los excluidos
            dirs[:] = [d for d in dirs if d not in exclude_dirs]
            
            for file in files:
                if file in exclude_files:
                    continue
                    
                _, ext = os.path.splitext(file)
                if ext in exclude_extensions:
                    continue
                
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, current_dir)
                
                try:
                    zipf.write(file_path, arcname)
                    # print(f"Agregado: {arcname}")
                except Exception as e:
                    print(f"Error agregando {file}: {e}")

    print(f"✅ Paquete creado exitosamente: {os.path.abspath(zip_filename)}")
    print("Copie este archivo a su servidor AWS y descomprímalo para actualizar.")

if __name__ == "__main__":
    create_update_package()
