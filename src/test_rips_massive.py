import os
import json
import pandas as pd
import io
import sys
import shutil

# Mock st and log
class MockSt:
    def progress(self, val, text=None):
        print(f"Progress: {val} - {text}")
        return self
    def empty(self):
        pass

sys.modules["streamlit"] = MockSt()
import streamlit as st

# Import functions from app_web.py by extracting them (since we can't easily import due to streamlit dependency at top level)
# Actually, I can use the same extraction technique as before or just copy the relevant parts if I want to be safe.
# But let's try to import app_web after mocking streamlit.

# Need to mock other st calls that might happen at module level
st.set_page_config = lambda **k: None
st.markdown = lambda *a, **k: None
st.title = lambda *a, **k: None
st.tabs = lambda *a, **k: [MockSt() for _ in range(7)]
st.session_state = {}
st.sidebar = MockSt()
st.sidebar.image = lambda *a, **k: None
st.sidebar.selectbox = lambda *a, **k: None
st.sidebar.button = lambda *a, **k: False

# Now try to import
sys.path.append(os.path.join(os.path.dirname(__file__), "OrganizadorArchivos", "src"))
# app_web is in d:/instalar/OrganizadorArchivos/src/app_web.py
# This script is likely in d:/instalar/OrganizadorArchivos/src/ or d:/instalar/
# Let's assume we run it from d:/instalar/OrganizadorArchivos/

try:
    from src.app_web import worker_consolidar_json_xlsx, worker_desconsolidar_xlsx_json, get_val_ci, clean_df_for_json
except ImportError:
    # If running from src directly
    sys.path.append(os.path.abspath("d:/instalar/OrganizadorArchivos/src"))
    from app_web import worker_consolidar_json_xlsx, worker_desconsolidar_xlsx_json, get_val_ci, clean_df_for_json

def test_massive_flow():
    # Setup temp dir
    test_dir = "test_massive_data"
    if os.path.exists(test_dir):
        shutil.rmtree(test_dir)
    os.makedirs(test_dir)
    
    # Create 2 JSON files
    json1 = {
        "numDocumentoIdObligado": "9001",
        "numFactura": "F1",
        "usuarios": [
            {
                "numDocumentoIdentificacion": "111",
                "consecutivo": 1,
                "servicios": {
                    "consultas": [{"codServicio": "C1", "consecutivo": 1}]
                }
            }
        ]
    }
    
    json2 = {
        "numDocumentoIdObligado": "9002",
        "numFactura": "F2",
        "usuarios": [
            {
                "numDocumentoIdentificacion": "222",
                "consecutivo": 1, # Same consecutivo, different file
                "servicios": {
                    "consultas": [{"codServicio": "C2", "consecutivo": 1}]
                }
            }
        ]
    }
    
    with open(os.path.join(test_dir, "file1.json"), "w") as f:
        json.dump(json1, f)
    with open(os.path.join(test_dir, "file2.json"), "w") as f:
        json.dump(json2, f)
        
    print("--- Testing Consolidar (JSON -> XLSX) ---")
    xlsx_bytes, error = worker_consolidar_json_xlsx(test_dir)
    
    if error:
        print(f"Error converting: {error}")
        sys.exit(1)
        
    # Check Excel Structure
    xls = pd.ExcelFile(io.BytesIO(xlsx_bytes))
    print(f"Sheets: {xls.sheet_names}")
    
    if "Transaccion" not in xls.sheet_names or "Usuarios" not in xls.sheet_names:
        print("FAIL: Missing Transaccion or Usuarios sheet")
        sys.exit(1)
        
    df_t = pd.read_excel(xls, sheet_name="Transaccion")
    print("Transaccion columns:", df_t.columns.tolist())
    if "archivo_origen" not in df_t.columns:
        print("FAIL: Missing archivo_origen in Transaccion")
        sys.exit(1)
        
    print("--- Testing Desconsolidar (XLSX -> JSON) ---")
    out_dir = os.path.join(test_dir, "output")
    os.makedirs(out_dir)
    
    success, msg = worker_desconsolidar_xlsx_json(io.BytesIO(xlsx_bytes), out_dir)
    if not success:
        print(f"Error desconsolidating: {msg}")
        sys.exit(1)
        
    print(f"Result: {msg}")
    
    # Verify outputs
    files = os.listdir(out_dir)
    print(f"Generated files: {files}")
    
    if len(files) != 2:
        print("FAIL: Expected 2 files")
        sys.exit(1)
        
    with open(os.path.join(out_dir, "file1.json"), "r") as f:
        d1 = json.load(f)
        if d1["numFactura"] != "F1":
            print("FAIL: File1 content mismatch")
            sys.exit(1)
            
    print("SUCCESS: Massive flow verified with Normalized structure")
    
    # Cleanup
    shutil.rmtree(test_dir)

if __name__ == "__main__":
    test_massive_flow()
