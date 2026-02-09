
import os
import sys
import struct

print(f"Python Version: {sys.version}")
print(f"Python Architecture: {struct.calcsize('P') * 8}-bit")

try:
    import pytesseract
    from PIL import Image
    print("pytesseract and PIL imported successfully.")
except ImportError as e:
    print(f"ImportError: {e}")

tesseract_path = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
print(f"Checking {tesseract_path}...")
exists = os.path.exists(tesseract_path)
print(f"Exists: {exists}")

if exists:
    try:
        pytesseract.pytesseract.tesseract_cmd = tesseract_path
        print(f"tesseract_cmd set to {tesseract_path}")
        print(f"Tesseract Version: {pytesseract.get_tesseract_version()}")
    except Exception as e:
        print(f"Error getting version: {e}")

# Check environment variables
print(f"ProgramFiles: {os.environ.get('ProgramFiles')}")
print(f"ProgramFiles(x86): {os.environ.get('ProgramFiles(x86)')}")
print(f"ProgramW6432: {os.environ.get('ProgramW6432')}")
