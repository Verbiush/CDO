import sys

def fix_file():
    with open('src/tabs/tab_automated_actions.py', 'r', encoding='utf-8') as f:
        content = f.read()

    old_str = "is_native_mode = st.session_state.get('force_native_mode', True)"
    new_str = "is_native_mode = getattr(st, 'session_state', {}).get('force_native_mode', True) if hasattr(st, 'session_state') else False"
    
    content = content.replace(old_str, new_str)

    with open('src/tabs/tab_automated_actions.py', 'w', encoding='utf-8') as f:
        f.write(content)
        
    print("Fixed tab_automated_actions.py")

if __name__ == '__main__':
    fix_file()
