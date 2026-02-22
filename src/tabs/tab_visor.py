import streamlit as st
import json
import xml.etree.ElementTree as ET
import xml.dom.minidom
import io
import time
import re
try:
    from src.tabs.tab_automated_actions import recursive_update_notes
except ImportError:
    # Fallback if import fails (e.g. during testing or different path structure)
    def recursive_update_notes(data, target_text, new_note):
        count = 0
        if isinstance(data, dict):
            for k, v in data.items():
                if isinstance(v, str) and target_text in v:
                    data[k] = new_note
                    count += 1
                elif isinstance(v, (dict, list)):
                    count += recursive_update_notes(v, target_text, new_note)
        elif isinstance(data, list):
            for item in data:
                count += recursive_update_notes(item, target_text, new_note)
        return count


# Try to import code_editor
try:
    from code_editor import code_editor
    HAS_CODE_EDITOR = True
except ImportError:
    HAS_CODE_EDITOR = False

def format_json(content):
    try:
        parsed = json.loads(content)
        return json.dumps(parsed, indent=4, ensure_ascii=False)
    except Exception:
        return content

def format_xml(content):
    try:
        # Parse logic
        dom = xml.dom.minidom.parseString(content)
        pretty = dom.toprettyxml(indent="  ")
        # Remove excessive blank lines if present
        return "\n".join([line for line in pretty.split('\n') if line.strip()])
    except Exception:
        return content

def get_markers(content, search_term):
    markers = []
    if not search_term:
        return markers
    
    lines = content.split('\n')
    for i, line in enumerate(lines):
        # Case insensitive search
        for match in re.finditer(re.escape(search_term), line, re.IGNORECASE):
            start_col = match.start()
            end_col = match.end()
            markers.append({
                "startRow": i,
                "startCol": start_col,
                "endRow": i,
                "endCol": end_col,
                "className": "ace_highlight-marker",
                "type": "text"
            })
    return markers

def render(container=None):
    if container is None:
        container = st.container()
        
    # Inject CSS for markers
    st.markdown("""
    <style>
    .ace_highlight-marker {
        position: absolute;
        background: rgba(255, 255, 0, 0.4);
        z-index: 20;
    }
    </style>
    """, unsafe_allow_html=True)
        
    with container:
        st.header("📄 Visor y Editor (JSON/XML)")
        
        uploaded_file = st.file_uploader("Subir archivo JSON o XML", type=["json", "xml"], key="visor_uploader")
        
        if uploaded_file:
            # Leer contenido inicial
            if "visor_file_content" not in st.session_state or st.session_state.get("visor_filename") != uploaded_file.name:
                raw_content = uploaded_file.getvalue().decode("utf-8", errors="ignore")
                
                # Auto-formatear al cargar
                formatted_content = raw_content
                if uploaded_file.name.lower().endswith(".json"):
                    formatted_content = format_json(raw_content)
                elif uploaded_file.name.lower().endswith(".xml"):
                    formatted_content = format_xml(raw_content)
                
                st.session_state.visor_file_content = formatted_content
                st.session_state.visor_filename = uploaded_file.name

            content = st.session_state.visor_file_content
            file_type = "json" if uploaded_file.name.lower().endswith(".json") else "xml"
            lang_mode = "json" if file_type == "json" else "xml"

            # --- BARRA DE BÚSQUEDA ---
            search_col1, search_col2 = st.columns([0.8, 0.2])
            with search_col1:
                search_term = st.text_input("🔍 Buscar en el contenido:", key="visor_search_term")
            
            markers = []
            with search_col2:
                # Mostrar conteo de resultados y generar marcadores
                if search_term:
                    markers = get_markers(content, search_term)
                    count = len(markers)
                    st.info(f"{count} resultados")
                else:
                    st.write("") # Spacer

            # --- NOTAS DE AJUSTE ---
            with st.expander("🛠️ Notas de Ajuste (Actualización Masiva)", expanded=False):
                st.info("Esta herramienta permite actualizar notas masivamente en estructuras JSON recursivas.")
                na_col1, na_col2, na_col3 = st.columns([0.4, 0.4, 0.2])
                with na_col1:
                    target_note_text = st.text_input("Texto a buscar en notas:", key="visor_na_target")
                with na_col2:
                    new_note_text = st.text_input("Nueva nota:", key="visor_na_new")
                with na_col3:
                    st.write("") # Spacer for alignment
                    st.write("") 
                    if st.button("Actualizar Notas", key="visor_na_btn"):
                        if file_type != "json":
                            st.error("Esta función solo está disponible para archivos JSON.")
                        elif not target_note_text:
                            st.warning("Debe ingresar el texto a buscar.")
                        else:
                            try:
                                json_data = json.loads(content)
                                count = recursive_update_notes(json_data, target_note_text, new_note_text)
                                if count > 0:
                                    st.session_state.visor_file_content = json.dumps(json_data, indent=4, ensure_ascii=False)
                                    st.success(f"Se actualizaron {count} notas correctamente.")
                                    time.sleep(1) # Give time to read message
                                    st.rerun()
                                else:
                                    st.warning("No se encontraron coincidencias para actualizar.")
                            except json.JSONDecodeError:
                                st.error("El contenido actual no es un JSON válido.")
                            except Exception as e:
                                st.error(f"Error al actualizar notas: {e}")

            # Custom buttons for editor (only for right side)
            editor_buttons = [{
                "name": "Guardar Cambios",
                "feather": "Save",
                "primary": True,
                "hasText": True,
                "alwaysOn": True,
                "commands": ["submit"],
                "style": {"bottom": "0.46rem", "right": "0.4rem"}
            }]
            
            # Configuración para Vista Previa (Solo Lectura)
            preview_options = {
                "readOnly": True,
                "showLineNumbers": True,
                "wrap": True,
                "highlightActiveLine": True,
                "highlightSelectedWord": True
            }

            # Configuración para Editor (Escritura)
            editor_options = {
                "readOnly": False,
                "showLineNumbers": True,
                "wrap": True
            }
            
            # Props para pasar los marcadores
            editor_props = {"markers": markers} if markers else {}

            col1, col2 = st.columns([1, 1])
            
            # --- VISTA PREVIA (IZQUIERDA) ---
            with col1:
                st.subheader("Vista Previa (Lectura)")
                if HAS_CODE_EDITOR:
                    code_editor(
                        content,
                        lang=lang_mode,
                        height="600px",
                        options=preview_options,
                        props=editor_props,
                        key="visor_preview_component"
                    )
                else:
                    st.code(content, language=lang_mode)

            # --- EDITOR (DERECHA) ---
            with col2:
                c_title, c_fmt = st.columns([0.7, 0.3])
                with c_title:
                    st.subheader("Editor de Contenido")
                with c_fmt:
                    if st.button("✨ Dar Formato", key=f"fmt_{lang_mode}"):
                        formatted = format_json(content) if file_type == "json" else format_xml(content)
                        if formatted != content:
                            st.session_state.visor_file_content = formatted
                            st.rerun()

                if HAS_CODE_EDITOR:
                    response_dict = code_editor(
                        content, 
                        lang=lang_mode, 
                        height="600px",
                        buttons=editor_buttons,
                        options=editor_options,
                        props=editor_props,
                        key=f"visor_{lang_mode}_editor_component"
                    )
                    
                    if response_dict['type'] == "submit" and len(response_dict['text']) > 0:
                        if response_dict['text'] != content:
                            st.session_state.visor_file_content = response_dict['text']
                            st.rerun()
                else:
                    new_content = st.text_area(
                        f"Edite el {file_type.upper()} aquí:", 
                        value=content, 
                        height=600,
                        key=f"visor_{lang_mode}_editor_text",
                        label_visibility="collapsed"
                    )
                    if new_content != content:
                        st.session_state.visor_file_content = new_content
                        st.rerun()

                st.download_button(
                    label=f"💾 Descargar {file_type.upper()} Modificado",
                    data=content,
                    file_name=f"modificado_{uploaded_file.name}",
                    mime=f"application/{file_type}"
                )
