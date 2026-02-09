
import os
import shutil
import datetime
from datetime import datetime

APP_FILE = r"d:\instalar\OrganizadorArchivos\src\app_web.py"

NEW_BROWSE_MODAL = r'''@st.dialog("Abrir", width="large")
def browse_modal():
    # --- CSS para simular estilo Windows ---
    st.markdown("""
        <style>
        div[data-testid="stDialog"] div[role="dialog"] {
            width: 85vw;
            max-width: 1100px;
            height: 80vh;
        }
        .file-row:hover {
            background-color: #cce8ff;
            cursor: pointer;
        }
        </style>
    """, unsafe_allow_html=True)

    if "modal_path" not in st.session_state:
        st.session_state.modal_path = st.session_state.current_path
    
    # Estado para archivo seleccionado
    if "selected_file" not in st.session_state:
        st.session_state.selected_file = ""

    # Constante para la vista "Este equipo"
    DRIVES_VIEW = "::DRIVES::"

    # Validacion inicial
    if st.session_state.modal_path != DRIVES_VIEW and not os.path.exists(st.session_state.modal_path):
         st.session_state.modal_path = DRIVES_VIEW
         
    current_path = st.session_state.modal_path
    
    # --- 1. BARRA DE DIRECCIONES Y NAVEGACIÓN ---
    col_nav_btns, col_address, col_search = st.columns([1.5, 6, 2.5])
    
    with col_nav_btns:
        c_back, c_up, c_refresh = st.columns(3)
        # Back (Simulado, no tenemos historial real facil aqui, usaremos Up como principal)
        c_back.button("⬅", disabled=True, key="nav_back", help="Atrás (No disponible)")
        
        if c_up.button("⬆", help="Subir a «" + (os.path.dirname(current_path) if current_path != DRIVES_VIEW else "Escritorio") + "»"):
            if current_path == DRIVES_VIEW:
                pass
            else:
                parent = os.path.dirname(current_path)
                if parent == current_path: # Root
                     st.session_state.modal_path = DRIVES_VIEW
                elif parent and os.path.exists(parent):
                     # UNC Check
                     if current_path.startswith(r"\\") and len(current_path.strip("\\").split("\\")) <= 2:
                         st.session_state.modal_path = DRIVES_VIEW
                     else:
                         st.session_state.modal_path = parent
                st.session_state.selected_file = "" # Limpiar seleccion al cambiar dir
                st.rerun()

        if c_refresh.button("↻", help="Actualizar"):
             st.rerun()

    with col_address:
        display_path = "Este equipo" if current_path == DRIVES_VIEW else current_path
        
        def on_addr_change():
            p = st.session_state.addr_input
            if p == "Este equipo": st.session_state.modal_path = DRIVES_VIEW
            elif os.path.exists(p) and os.path.isdir(p): st.session_state.modal_path = p
        
        st.text_input("Dirección", value=display_path, key="addr_input", label_visibility="collapsed", on_change=on_addr_change)

    with col_search:
        st.text_input("Buscar", placeholder=f"Buscar en {os.path.basename(current_path) if current_path != DRIVES_VIEW else 'Este equipo'}", label_visibility="collapsed")

    st.markdown("---")

    # --- 2. CUERPO PRINCIPAL (SIDEBAR + FILE LIST) ---
    col_sidebar, col_main = st.columns([2, 8])
    
    # === PANEL DE NAVEGACIÓN (IZQUIERDA) ===
    with col_sidebar:
        # Favoritos (Quick Access)
        st.caption("Acceso rápido")
        if st.button("📌 Escritorio", use_container_width=True):
             # Try standard desktop paths
             desktop = os.path.join(os.path.expanduser("~"), "Desktop")
             if os.path.exists(desktop): 
                 st.session_state.modal_path = desktop
                 st.session_state.selected_file = ""
                 st.rerun()
        
        if st.session_state.get("username"):
            u_conf = get_user_config(st.session_state.username)
            favs = u_conf.get("favorites", [])
            for fav in favs:
                name = os.path.basename(fav) or fav
                if st.button(f"⭐ {name}", key=f"side_fav_{fav}", use_container_width=True):
                    st.session_state.modal_path = fav
                    st.session_state.selected_file = ""
                    st.rerun()
        
        st.caption("Este equipo")
        if st.button("💻 Este equipo", use_container_width=True):
            st.session_state.modal_path = DRIVES_VIEW
            st.session_state.selected_file = ""
            st.rerun()

        st.caption("Red")
        if st.button("🌐 Conectar a Red...", use_container_width=True):
            st.session_state.show_connect_pc = not st.session_state.get("show_connect_pc", False)
            st.rerun()

    # === LISTA DE ARCHIVOS (DERECHA) ===
    with col_main:
        # Panel de conexión de red (overlay)
        if st.session_state.get("show_connect_pc", False):
            with st.container(border=True):
                st.write("Conectar unidad de red")
                pc = st.text_input("PC", placeholder="\\192.168.1.X")
                share = st.text_input("Carpeta", placeholder="Compartida")
                if st.button("Conectar"):
                    path = f"\\\\{pc}\\{share}".replace("\\\\\\\\", "\\\\")
                    if os.path.exists(path):
                        st.session_state.modal_path = path
                        st.session_state.show_connect_pc = False
                        st.rerun()
                    else:
                        st.error("No encontrado")

        elif current_path == DRIVES_VIEW:
            st.subheader("Dispositivos y unidades")
            drives = []
            if os.name == 'nt':
                import string
                try: drives = [f"{d}:\\" for d in string.ascii_uppercase if os.path.exists(f"{d}:\\")]
                except: pass
            else: drives = ["/"]
            
            for d in drives:
                c_img, c_info, c_btn = st.columns([1, 6, 2])
                c_img.write("💽")
                with c_info:
                    st.write(f"**Disco Local ({d})**")
                    try:
                        total, used, free = shutil.disk_usage(d)
                        st.progress(used/total)
                    except: pass
                if c_btn.button("Explorar", key=f"open_{d}"):
                    st.session_state.modal_path = d
                    st.rerun()

        else:
            # Encabezados de columnas
            h_name, h_date, h_type = st.columns([5, 3, 2])
            h_name.markdown("**Nombre**")
            h_date.markdown("**Fecha de modificación**")
            h_type.markdown("**Tipo**")
            st.divider()

            # Contenido desplazable (simulado con container)
            with st.container(height=400):
                try:
                    items = os.listdir(current_path)
                    folders = [i for i in items if os.path.isdir(os.path.join(current_path, i))]
                    files = [i for i in items if not os.path.isdir(os.path.join(current_path, i))]
                    folders.sort()
                    files.sort()
                    
                    # Carpetas
                    for f in folders:
                        fp = os.path.join(current_path, f)
                        c1, c2, c3 = st.columns([5, 3, 2])
                        if c1.button(f"📁 {f}", key=f"dir_{f}", use_container_width=True):
                            st.session_state.modal_path = fp
                            st.session_state.selected_file = ""
                            st.rerun()
                        
                        dt = datetime.fromtimestamp(os.path.getmtime(fp)).strftime("%d/%m/%Y %H:%M")
                        c2.caption(dt)
                        c3.caption("Carpeta de archivos")

                    # Archivos
                    # Filtro por tipo seleccionado (si se implementa)
                    target_exts = st.session_state.get("filter_ext", "Todos los archivos (*.*)")
                    
                    for f in files:
                        fp = os.path.join(current_path, f)
                        ext = os.path.splitext(f)[1].lower()
                        
                        # Icono
                        icon = "📄"
                        if ext == '.pdf': icon = "📕"
                        elif ext in ['.xls','.xlsx']: icon = "📊"
                        elif ext in ['.doc','.docx']: icon = "📝"
                        
                        # Highlight si está seleccionado
                        is_selected = (st.session_state.selected_file == f)
                        btn_label = f"{'✅ ' if is_selected else ''}{icon} {f}"
                        
                        c1, c2, c3 = st.columns([5, 3, 2])
                        
                        # Logica de seleccion
                        if c1.button(btn_label, key=f"file_{f}", use_container_width=True, type="primary" if is_selected else "secondary"):
                            st.session_state.selected_file = f
                            st.rerun()
                            
                        dt = datetime.fromtimestamp(os.path.getmtime(fp)).strftime("%d/%m/%Y %H:%M")
                        c2.caption(dt)
                        
                        # Tipo legible
                        type_str = "Archivo"
                        if ext == '.xlsx': type_str = "Hoja de cálculo"
                        elif ext == '.docx': type_str = "Documento de Word"
                        elif ext == '.pdf': type_str = "Documento PDF"
                        c3.caption(type_str)

                except Exception as e:
                    st.error(f"Error de acceso: {e}")

    st.markdown("---")

    # --- 3. BARRA INFERIOR (INPUTS Y BOTONES) ---
    c_label, c_input, c_type_sel = st.columns([2, 6, 3])
    with c_label:
        st.write("Nombre de archivo:")
    with c_input:
        # El input muestra el archivo seleccionado
        final_name = st.text_input("Filename", value=st.session_state.selected_file, label_visibility="collapsed", key="final_filename_input")
        # Si el usuario escribe manualmente, actualizamos selected_file
        if final_name != st.session_state.selected_file:
            st.session_state.selected_file = final_name
            
    with c_type_sel:
        st.selectbox("Filter", ["Todos los archivos (*.*)", "Documentos PDF (*.pdf)", "Excel (*.xls;*.xlsx)"], label_visibility="collapsed", key="filter_ext")

    c_empty, c_open, c_cancel = st.columns([7, 2, 2])
    with c_open:
        # Texto del botón dinámico
        btn_text = "Seleccionar Ruta"
        if st.session_state.selected_file:
            btn_text = "Abrir Archivo"
            
        if st.button(btn_text, type="primary", use_container_width=True):
            # Determinar ruta final
            if current_path == DRIVES_VIEW:
                st.warning("Selecciona una unidad o carpeta primero.")
            else:
                target = os.path.join(current_path, st.session_state.selected_file) if st.session_state.selected_file else current_path
                
                if os.path.exists(target):
                    st.session_state.current_path = target
                    st.session_state.path_input = target
                    
                    # Guardar en favoritos/historial de usuario
                    if st.session_state.get("username"):
                        update_user_last_path(st.session_state.username, target)
                        
                    st.rerun()
                else:
                    st.error("Ruta no encontrada: " + target)

    with c_cancel:
        if st.button("Cancelar", use_container_width=True):
            st.rerun()
'''

with open(APP_FILE, "r", encoding="utf-8") as f:
    content = f.read()

# Replace block
start_marker = '@st.dialog("Abrir", width="large")'
end_marker = 'import xml.etree.ElementTree'

start_idx = content.find(start_marker)
end_idx = content.find(end_marker)

if start_idx != -1 and end_idx != -1:
    new_content = content[:start_idx] + NEW_BROWSE_MODAL + "\n\n\n" + content[end_idx:]
    with open(APP_FILE, "w", encoding="utf-8") as f:
        f.write(new_content)
    print("Updated browse_modal v4 successfully")
else:
    print("Markers not found")
