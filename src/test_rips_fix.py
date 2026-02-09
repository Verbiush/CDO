import sys
import os
import ast
import pandas as pd
import json
import io
import unicodedata
import warnings

# Suppress warnings
warnings.filterwarnings("ignore")

# Read app_web.py
app_web_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app_web.py")
with open(app_web_path, "r", encoding="utf-8") as f:
    source = f.read()

# Parse AST
tree = ast.parse(source)

# Extract specific functions
functions_to_extract = [
    "clean_df_for_json", 
    "normalize_key", 
    "get_val_ci",
    "worker_json_a_xlsx_ind",
    "worker_xlsx_a_json_ind"
]

extracted_code = ""
# Helper to get source segment (python 3.8+) or manual slice
for node in tree.body:
    if isinstance(node, ast.FunctionDef) and node.name in functions_to_extract:
        # Simple slicing based on line numbers
        start = node.lineno - 1
        end = node.end_lineno
        lines = source.splitlines()[start:end]
        extracted_code += "\n".join(lines) + "\n\n"

# Execute extracted code in a separate namespace
exec_globals = {
    "pd": pd,
    "json": json,
    "io": io,
    "unicodedata": unicodedata,
    "os": os,
    "st": None # Mock st if needed, but workers don't use it directly except for logging which we might need to mock or remove
}

# Mocking log function or st.write if used
extracted_code = extracted_code.replace("log(", "print(")
extracted_code = extracted_code.replace("st.write(", "print(")

print("Compiling extracted code...")
try:
    exec(extracted_code, exec_globals)
except Exception as e:
    print(f"Error compiling code: {e}")
    sys.exit(1)

print("Code compiled successfully.")

# Define sample data with MIXED CASE keys to test robustness
sample_json = {
    "numDocumentoIdObligado": "900438792",
    "numFactura": "FEOV67636",
    "tipoNota": None,
    "numNota": None,
    "usuarios": [
        {
            "tipoDocumentoIdentificacion": "CC",
            "numDocumentoIdentificacion": "123456",
            "tipoUsuario": "1",
            "consecutivo": 1,
            # Test mixed case key for robust search
            "Fechanacimiento": "1990-01-01", 
            "servicios": {
                "consultas": [
                    {
                        "codServicio": "890201",
                        "valorPagoModerador": 0,
                        "consecutivo": 1
                    }
                ]
            }
        }
    ]
}

# Write sample to file
test_file = "test_sample.json"
with open(test_file, "w", encoding="utf-8") as f:
    json.dump(sample_json, f)

# 1. Test JSON -> XLSX
print("Testing JSON -> XLSX...")
with open(test_file, "r", encoding="utf-8") as f:
    xlsx_bytes, error = exec_globals["worker_json_a_xlsx_ind"](f)

if error:
    print(f"Error converting JSON to XLSX: {error}")
    sys.exit(1)

# 2. Test XLSX -> JSON
print("Testing XLSX -> JSON...")
file_obj = io.BytesIO(xlsx_bytes)

# DEBUG: Check sheets
xls_debug = pd.ExcelFile(io.BytesIO(xlsx_bytes))
print("DEBUG: Sheets found:", xls_debug.sheet_names)

json_out_str, error = exec_globals["worker_xlsx_a_json_ind"](file_obj)
if error:
    print(f"Error converting XLSX to JSON: {error}")
    sys.exit(1)

json_out = json.loads(json_out_str)

print("DEBUG: json_out keys:", list(json_out.keys()))
# print("DEBUG: json_out content:", json.dumps(json_out, indent=2))

# 3. Verify
print("Verifying data...")
header_match = (
    str(json_out.get("numDocumentoIdObligado")) == str(sample_json["numDocumentoIdObligado"]) and
    str(json_out.get("numFactura")) == str(sample_json["numFactura"])
)
print(f"Header match: {header_match}")
if not header_match:
    print(f"Header Expected: {sample_json['numDocumentoIdObligado']}, Got: {json_out.get('numDocumentoIdObligado')}")

user_match = (
    str(json_out["usuarios"][0]["numDocumentoIdentificacion"]) == str(sample_json["usuarios"][0]["numDocumentoIdentificacion"])
)
print(f"User match: {user_match}")

# Verify fechaNacimiento recovery from mixed case input
dob_match = (
    str(json_out["usuarios"][0].get("fechaNacimiento")) == str(sample_json["usuarios"][0]["Fechanacimiento"])
)
print(f"DOB match (Mixed Case): {dob_match}")
if not dob_match:
    print(f"DOB Expected: {sample_json['usuarios'][0]['Fechanacimiento']}, Got: {json_out['usuarios'][0].get('fechaNacimiento')}")

service_match = False
if "consultas" in json_out["usuarios"][0]["servicios"]:
    service_match = (
        len(json_out["usuarios"][0]["servicios"]["consultas"]) == 1 and
        str(json_out["usuarios"][0]["servicios"]["consultas"][0]["codServicio"]) == "890201"
    )
print(f"Service match: {service_match}")

if header_match and user_match and service_match and dob_match:
    print("SUCCESS: Roundtrip verified!")
else:
    print("FAILURE: Data mismatch")
    
# Clean up
if os.path.exists(test_file):
    os.remove(test_file)
