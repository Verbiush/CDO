import sys, re

with open('src/tabs/tab_search_actions.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Replace open_action_dialog definition
old_open = """    st.session_state["active_action_dialog"] = dialog_name
    st.rerun()"""
    
new_open = """    if callable(dialog_name):
        dialog_name()"""

content = content.replace(old_open, new_open)

# Replace dialog string calls with function calls
mapping = {
    '"rename"': 'dialog_renombrar',
    '"edit_text"': 'dialog_editar_texto',
    '"copy"': 'dialog_copiar',
    '"move"': 'dialog_mover',
    '"delete"': 'dialog_eliminar',
    '"zip"': 'dialog_comprimir',
    '"zip_individual"': 'dialog_comprimir_individual'
}

for k, v in mapping.items():
    content = content.replace(f'open_action_dialog({k})', f'open_action_dialog({v})')

# Remove the bottom block
pattern = r'    active_dialog = st\.session_state\.get\("active_action_dialog"\).*?del st\.session_state\["active_action_dialog"\]\s*'
content = re.sub(pattern, '', content, flags=re.DOTALL)

with open('src/tabs/tab_search_actions.py', 'w', encoding='utf-8') as f:
    f.write(content)
