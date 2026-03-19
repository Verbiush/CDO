import streamlit as st

st.session_state["force_native_mode"] = True

def render_path_selector(label, key, default_path=""):
    use_custom = True
    target_path = default_path
    if key in st.session_state:
        target_path = st.session_state[key]
    else:
        st.session_state[key] = target_path

    input_key = f"input_{key}"
    if input_key not in st.session_state:
        st.session_state[input_key] = target_path

    st.text_input(label, key=input_key, on_change=lambda: st.session_state.update({key: st.session_state[input_key]}))
    
    if use_custom and input_key in st.session_state:
        st.session_state[key] = st.session_state[input_key]
        
    return st.session_state.get(key, target_path)

path_an = render_path_selector("Carpeta", "tab_an_folder", default_path="C:\\Users")
st.write(f"path_an is: '{path_an}'")

if st.button("Test"):
    if path_an:
        st.success("Valid!")
    else:
        st.warning("Seleccione una carpeta válida.")
