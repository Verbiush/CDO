import sys
import re

with open('src/tabs/tab_automated_actions.py', 'r', encoding='utf-8') as f:
    content = f.read()

helper = '''
def _should_delegate(path_or_list):
    import os
    if not path_or_list: return False
    path = path_or_list
    if isinstance(path_or_list, list):
        path = path_or_list[0]
    if isinstance(path, dict):
        path = path.get("Ruta completa", "")
    return not os.path.exists(path)
'''

if '_should_delegate' not in content:
    content = content.replace('import json', 'import json\n' + helper)

pattern = r"(is_native_mode = st\.session_state\.get\('force_native_mode', True\)\s+)if is_native_mode and not silent_mode:"

# We need to replace manually or dynamically because some use root_path and some use file_list
def repl(match):
    return match.group(1) + 'if is_native_mode and not silent_mode and _should_delegate(root_path if "root_path" in locals() else file_list):'

# Wait, `locals()` inside the worker function? Yes! But it's evaluated at runtime!
# Actually, let's just use regex to find the function signature and replace it.

new_content = content
for func in ['worker_analisis_carpetas']:
    new_content = new_content.replace(
        "if is_native_mode and not silent_mode:\n        if not silent_mode: st.info(f\"Delegando análisis de carpetas al Agente Local...\")",
        "if is_native_mode and not silent_mode and _should_delegate(root_path):\n        if not silent_mode: st.info(f\"Delegando análisis de carpetas al Agente Local...\")"
    )

for func in ['worker_analisis_historia_clinica']:
    new_content = new_content.replace(
        "if is_native_mode and not silent_mode:\n        if not silent_mode: st.info(f\"Delegando análisis HC al Agente Local...\")",
        "if is_native_mode and not silent_mode and _should_delegate(file_list):\n        if not silent_mode: st.info(f\"Delegando análisis HC al Agente Local...\")"
    )

for func in ['worker_leer_pdf_retefuente']:
    new_content = new_content.replace(
        "if is_native_mode and not silent_mode:\n        if not silent_mode: st.info(f\"Delegando análisis Retefuente al Agente Local...\")",
        "if is_native_mode and not silent_mode and _should_delegate(file_list):\n        if not silent_mode: st.info(f\"Delegando análisis Retefuente al Agente Local...\")"
    )

for func in ['worker_analisis_emssanar']:
    new_content = new_content.replace(
        "if is_native_mode and not silent_mode:\n        if not silent_mode: st.info(f\"Delegando análisis Emssanar al Agente Local...\")",
        "if is_native_mode and not silent_mode and _should_delegate(file_list):\n        if not silent_mode: st.info(f\"Delegando análisis Emssanar al Agente Local...\")"
    )

for func in ['worker_analisis_medicina_legal']:
    new_content = new_content.replace(
        "if is_native_mode and not silent_mode:\n        if not silent_mode: st.info(f\"Delegando análisis Medicina Legal al Agente Local...\")",
        "if is_native_mode and not silent_mode and _should_delegate(file_list):\n        if not silent_mode: st.info(f\"Delegando análisis Medicina Legal al Agente Local...\")"
    )

for func in ['worker_analisis_sos']:
    new_content = new_content.replace(
        "if is_native_mode and not silent_mode:\n        if not silent_mode: st.info(f\"Delegando análisis SOS al Agente Local...\")",
        "if is_native_mode and not silent_mode and _should_delegate(file_list):\n        if not silent_mode: st.info(f\"Delegando análisis SOS al Agente Local...\")"
    )

for func in ['worker_analisis_autorizacion_nueva_eps']:
    new_content = new_content.replace(
        "if is_native_mode and not silent_mode:\n        if not silent_mode: st.info(f\"Delegando análisis Nueva EPS al Agente Local...\")",
        "if is_native_mode and not silent_mode and _should_delegate(file_list):\n        if not silent_mode: st.info(f\"Delegando análisis Nueva EPS al Agente Local...\")"
    )

for func in ['worker_analisis_cargue_sanitas']:
    new_content = new_content.replace(
        "if is_native_mode and not silent_mode:\n        if not silent_mode: st.info(f\"Delegando análisis Sanitas al Agente Local...\")",
        "if is_native_mode and not silent_mode and _should_delegate(file_list):\n        if not silent_mode: st.info(f\"Delegando análisis Sanitas al Agente Local...\")"
    )

with open('src/tabs/tab_automated_actions.py', 'w', encoding='utf-8') as f:
    f.write(new_content)

print('Done.')
