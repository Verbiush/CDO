
import PyInstaller.__main__
import os
import shutil
import sys

# Define paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SRC_DIR = os.path.join(BASE_DIR, 'src')
SERVER_SCRIPT = os.path.join(SRC_DIR, 'server_api.py')
DIST_DIR = os.path.join(BASE_DIR, 'dist')
BUILD_DIR = os.path.join(BASE_DIR, 'build')

# Clean previous builds (optional, maybe keep agent build)
# if os.path.exists(DIST_DIR): shutil.rmtree(DIST_DIR)
# if os.path.exists(BUILD_DIR): shutil.rmtree(BUILD_DIR)

print(f"Building Server API from: {SERVER_SCRIPT}")

# Build arguments
args = [
    SERVER_SCRIPT,
    '--onefile',
    '--name=CDO_Servidor_API',
    f'--paths={SRC_DIR}',
    '--clean',
    '--noconfirm',
    # '--noconsole', # Keep console for server logs
    '--hidden-import=uvicorn.logging',
    '--hidden-import=uvicorn.loops',
    '--hidden-import=uvicorn.loops.auto',
    '--hidden-import=uvicorn.protocols',
    '--hidden-import=uvicorn.protocols.http',
    '--hidden-import=uvicorn.protocols.http.auto',
    '--hidden-import=uvicorn.protocols.websockets',
    '--hidden-import=uvicorn.protocols.websockets.auto',
    '--hidden-import=uvicorn.lifespan',
    '--hidden-import=uvicorn.lifespan.on',
    '--hidden-import=sqlite3',
    '--collect-all=uvicorn',
    '--collect-all=fastapi',
]

# Run PyInstaller
try:
    PyInstaller.__main__.run(args)
    print("Build successful! Executable is in 'dist/CDO_Servidor_API.exe'")
    
    # Copy existing database files if available
    for f in ['users.db', 'users.json']:
        src_f = os.path.join(SRC_DIR, f)
        if os.path.exists(src_f):
            shutil.copy2(src_f, os.path.join(DIST_DIR, f))
            print(f"Copied {f} to dist/")
            
except Exception as e:
    print(f"Build failed: {e}")
    sys.exit(1)
