import streamlit as st

st.title("Test Sync")

if "my_path" not in st.session_state:
    st.session_state["my_path"] = "C:\\Default"

target_path = st.session_state["my_path"]
input_key = "input_my_path"

# simulate native mode text input
if input_key not in st.session_state:
    st.session_state[input_key] = target_path

st.text_input("Path", key=input_key, on_change=lambda: st.session_state.update({"my_path": st.session_state[input_key]}))

if st.button("Simulate Folder Select"):
    selected = "C:\\NewFolder"
    st.session_state["my_path"] = selected
    if input_key in st.session_state:
        del st.session_state[input_key]
    st.rerun()

if input_key in st.session_state:
    st.session_state["my_path"] = st.session_state[input_key]

st.write(f"Current my_path: {st.session_state['my_path']}")
