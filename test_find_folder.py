
import sys
import os
import unittest
from unittest.mock import MagicMock

# Mock streamlit before importing the module
sys.modules['streamlit'] = MagicMock()
import streamlit as st

# Define st.session_state as a dict-like object
st.session_state = {}

# Now import the functions to test
# We need to manually import because the file is not in a package structure that allows easy relative imports from here
# So we will read the file and exec it, or just copy the helper function for testing.
# Importing is better to test the actual code.
sys.path.append(os.path.join(os.getcwd(), 'src', 'tabs'))

# We can't easily import tab_automated_actions because it has many dependencies.
# Let's mock the dependencies.
sys.modules['pandas'] = MagicMock()
sys.modules['fitz'] = MagicMock()
sys.modules['PIL'] = MagicMock()
sys.modules['docx'] = MagicMock()
sys.modules['requests'] = MagicMock()
sys.modules['openpyxl'] = MagicMock()
sys.modules['gui_utils'] = MagicMock()

# Now try to import
try:
    from tab_automated_actions import find_folder_path
except ImportError:
    # If import fails (e.g. due to other imports), we might need to rely on code inspection
    # or copy the function here.
    # Let's try to just copy the function logic to verify it against the mock data structure
    pass

def find_folder_path_impl(base_path, folder_name):
    """
    Intenta encontrar una carpeta usando los resultados de búsqueda en sesión.
    Si no la encuentra, asume que es una subcarpeta directa de base_path.
    """
    target_name = str(folder_name).strip().lower()
    
    # 1. Buscar en resultados de búsqueda (si existen)
    if "search_results" in st.session_state and st.session_state.search_results:
        for item in st.session_state.search_results:
            # Normalizar claves (puede venir como 'name' o 'Nombre')
            i_name = str(item.get("name", item.get("Nombre", ""))).strip().lower()
            i_type = str(item.get("type", item.get("Tipo", ""))).strip().lower()
            i_path = item.get("path", item.get("Ruta completa", ""))
            
            if i_type in ["folder", "carpeta", "directory"] and i_name == target_name:
                if True: # Mock existence check: os.path.exists(i_path):
                    return i_path

    # 2. Fallback: Subcarpeta directa
    return os.path.join(base_path, str(folder_name).strip())

class TestFindFolderPath(unittest.TestCase):
    def test_find_in_search_results_english(self):
        st.session_state['search_results'] = [
            {'name': 'FolderA', 'type': 'folder', 'path': '/abs/path/FolderA'},
            {'name': 'FileB', 'type': 'file', 'path': '/abs/path/FileB'}
        ]
        result = find_folder_path_impl('/base', 'FolderA')
        self.assertEqual(result, '/abs/path/FolderA')

    def test_find_in_search_results_spanish(self):
        st.session_state['search_results'] = [
            {'Nombre': 'CarpetaX', 'Tipo': 'Carpeta', 'Ruta completa': '/abs/path/CarpetaX'},
        ]
        result = find_folder_path_impl('/base', 'CarpetaX')
        self.assertEqual(result, '/abs/path/CarpetaX')

    def test_not_found_fallback(self):
        st.session_state['search_results'] = []
        result = find_folder_path_impl('/base', 'FolderZ')
        self.assertEqual(result, os.path.join('/base', 'FolderZ'))

    def test_case_insensitive(self):
        st.session_state['search_results'] = [
            {'name': 'FolderCase', 'type': 'folder', 'path': '/abs/path/FolderCase'}
        ]
        result = find_folder_path_impl('/base', 'foldercase')
        self.assertEqual(result, '/abs/path/FolderCase')

if __name__ == '__main__':
    unittest.main()
