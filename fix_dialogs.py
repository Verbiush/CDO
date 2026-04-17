import sys, re

with open('src/tabs/tab_automated_actions.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Replace open_auto_dialog definition
old_open = """    for k in keys_to_clear:
        if k in st.session_state:
            del st.session_state[k]
    st.session_state["active_auto_dialog"] = dialog_name
    st.rerun()"""
    
new_open = """    for k in keys_to_clear:
        if k in st.session_state:
            del st.session_state[k]
    
    if callable(dialog_name):
        dialog_name()"""

content = content.replace(old_open, new_open)

# Replace dialog string calls with function calls
mapping = {
    '"organizar_feov"': 'dialog_organizar_feov',
    '"copiar_mapeo_subcarpetas"': 'dialog_copiar_mapeo',
    '"copiar_mapeo_raiz"': 'dialog_copiar_raiz',
    '"exportar_nombres"': 'dialog_exportar_renombrado',
    '"aplicar_nombres"': 'dialog_aplicar_renombrado',
    '"sufijo_archivos"': 'dialog_sufijo',
    '"renombrar_excel"': 'dialog_renombrar_mapeo_excel',
    '"modif_docx"': 'dialog_modif_docx_completo',
    '"insertar_firma_docx"': 'dialog_insertar_firma_docx',
    '"crear_carpetas"': 'dialog_crear_carpetas_excel',
    '"descargar_firmas"': 'dialog_descargar_firmas',
    '"descargar_ovida"': 'dialog_descargar_historias_ovida',
    '"descargar_zeus"': 'dialog_descargar_zeus_adjuntos',
    '"crear_firma"': 'dialog_crear_firma',
    '"distribuir_base"': 'dialog_distribuir_base',
    '"cruzar_excels"': 'dialog_cruzar_excels',
    '"autorizacion_docx"': 'dialog_autorizacion_docx',
    '"regimen_docx"': 'dialog_regimen_docx',
    '"mover_coincidencia"': 'dialog_organizar_feov_avanzado'
}

for k, v in mapping.items():
    content = content.replace(f'open_auto_dialog({k})', f'open_auto_dialog({v})')

# Remove the bottom block
# Move dialog triggers to the root scope to avoid "Only one dialog" exception
pattern = r'    # Move dialog triggers to the root scope to avoid "Only one dialog" exception.*?del st\.session_state\["active_auto_dialog"\]\s*'
content = re.sub(pattern, '', content, flags=re.DOTALL)

with open('src/tabs/tab_automated_actions.py', 'w', encoding='utf-8') as f:
    f.write(content)
