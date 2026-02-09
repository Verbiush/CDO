
import streamlit as st
import os

APP_FILE = r"d:\instalar\OrganizadorArchivos\src\app_web.py"

# 1. Logic to remove 'seleccionar_carpeta' (Tkinter) and restore 'browse_modal' call
# 2. Logic to add Search functionality to browse_modal

with open(APP_FILE, "r", encoding="utf-8") as f:
    content = f.read()

# --- REMOVE Tkinter Function ---
start_marker = "def seleccionar_carpeta():"
end_marker = "def funcion_no_implementada(nombre):"

s_idx = content.find(start_marker)
e_idx = content.find(end_marker)

if s_idx != -1 and e_idx != -1:
    # Remove the function body, keep the end marker
    content = content[:s_idx] + content[e_idx:]
    print("Removed seleccionar_carpeta function.")

# --- UPDATE Button to use browse_modal ---
# Find the button block we added previously
old_button_block = r'''        if st.button("📂 Examinar", help="Abrir selector de carpetas (Nativo)"):
            folder = seleccionar_carpeta()
            if folder:
                st.session_state.current_path = folder
                st.session_state.path_input = folder
                if st.session_state.get("username"):
                    update_user_last_path(st.session_state.username, folder)
                st.rerun()'''

new_button_block = r'''        if st.button("📂 Examinar", help="Abrir gestor de ubicaciones"):
            browse_modal()'''

if old_button_block in content:
    content = content.replace(old_button_block, new_button_block)
    print("Restored browse_modal button.")
else:
    # Try fuzzy match or just look for the button line
    print("Could not find exact button block to replace. Attempting fallback.")
    # (Optional: Implement fallback if needed, but the previous Write was exact)

# --- IMPLEMENT SEARCH IN browse_modal ---
# We need to find where 'items = os.listdir(current_path)' is and inject filtering.

search_code_original = r'''                    items = os.listdir(current_path)
                    folders = [i for i in items if os.path.isdir(os.path.join(current_path, i))]'''

search_code_new = r'''                    items = os.listdir(current_path)
                    
                    # --- FILTRO DE BUSQUEDA ---
                    search_term = st.session_state.get("modal_search_box", "").lower()
                    if search_term:
                        items = [i for i in items if search_term in i.lower()]
                    # --------------------------

                    folders = [i for i in items if os.path.isdir(os.path.join(current_path, i))]'''

if search_code_original in content:
    content = content.replace(search_code_original, search_code_new)
    print(" injected search logic.")
else:
    print("Could not inject search logic. Check indentation.")

# --- UPDATE SEARCH INPUT KEY ---
# We need to give the search input a key so we can read it.
# Original: st.text_input("Buscar", placeholder=f"Buscar en ...", label_visibility="collapsed")
# We need to use regex because the placeholder is dynamic f-string in the source? 
# Actually in the file it is: 
# st.text_input("Buscar", placeholder=f"Buscar en {os.path.basename(current_path) if current_path != DRIVES_VIEW else 'Este equipo'}", label_visibility="collapsed")

# Let's try to replace the whole line.
# It's line 207 in the read output.
# st.text_input("Buscar", placeholder=f"Buscar en {os.path.basename(current_path) if current_path != DRIVES_VIEW else 'Este equipo'}", label_visibility="collapsed")

# Since it has an f-string, exact string replacement might be tricky if I don't match exactly.
# Let's use a unique part of the line.
search_input_part = 'st.text_input("Buscar", placeholder='
search_input_replacement = 'st.text_input("Buscar", key="modal_search_box", placeholder='

if search_input_part in content:
    content = content.replace(search_input_part, search_input_replacement)
    print("Updated search input key.")
else:
    print("Could not update search input key.")


with open(APP_FILE, "w", encoding="utf-8") as f:
    f.write(content)
