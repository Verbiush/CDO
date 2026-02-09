
import os

APP_FILE = r"d:\instalar\OrganizadorArchivos\src\app_web.py"

# 1. New definition for seleccionar_carpeta
NEW_FUNC = r'''def seleccionar_carpeta():
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        folder = filedialog.askdirectory()
        root.destroy()
        return folder
    except Exception as e:
        st.error(f"Error al abrir diálogo nativo: {e}")
        return None'''

# 2. New button logic
NEW_BUTTON = r'''        if st.button("📂 Examinar", help="Abrir selector de carpetas (Nativo)"):
            folder = seleccionar_carpeta()
            if folder:
                st.session_state.current_path = folder
                st.session_state.path_input = folder
                if st.session_state.get("username"):
                    update_user_last_path(st.session_state.username, folder)
                st.rerun()'''

with open(APP_FILE, "r", encoding="utf-8") as f:
    content = f.read()

# Replace definition of seleccionar_carpeta
# It currently looks like:
# def seleccionar_carpeta():
#     st.toast("⚠️ La selección de carpetas nativa no está disponible en modo web. Usa el navegador de servidor o ingresa la ruta manualmente.")
#     return None

start_def = content.find("def seleccionar_carpeta():")
if start_def != -1:
    # Find the end of the function (next def or reasonable length)
    # The current function is short (3 lines)
    end_def = content.find("def funcion_no_implementada", start_def)
    if end_def != -1:
        # Replace the function body
        content = content[:start_def] + NEW_FUNC + "\n\n" + content[end_def:]
        print("Updated seleccionar_carpeta definition")
    else:
        print("Could not find end of seleccionar_carpeta")

# Replace the button logic
# Search for: if st.button("📂 Examinar", help="Abrir gestor de ubicaciones (Local/Red)"):
#             browse_modal()

search_btn = 'if st.button("📂 Examinar", help="Abrir gestor de ubicaciones (Local/Red)"):            browse_modal()'
# The spaces might vary, let's try a regex or split approach if exact match fails
# But my previous Write tool used exact indentation.
# Let's try exact string first, but be careful with newlines.

# In previous tool call I wrote:
#         with col_path_2:
#             if st.button("📂 Examinar", help="Abrir gestor de ubicaciones (Local/Red)"):
#                 browse_modal()

# Let's try to match slightly more context
search_ctx = 'if st.button("📂 Examinar", help="Abrir gestor de ubicaciones (Local/Red)"): \n            browse_modal()'
# Actually, the file likely has correct indentation (4 spaces or 8 spaces)
# Let's read the file content around line 4696 again to be sure of indentation
pass
