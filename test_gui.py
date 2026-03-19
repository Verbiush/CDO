import streamlit as st
import os
import sys

# Add src to path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), 'src')))

from gui_utils import render_path_selector

st.title("Test Path Selector")

st.session_state['force_native_mode'] = True

default_path = st.session_state.get("current_path", os.path.expanduser("~"))

path_an = render_path_selector(
    label="Carpeta de Análisis",
    key="tab_an_folder",
    default_path=default_path
)

st.write(f"Returned path_an: '{path_an}'")

if st.button("Test"):
    if path_an:
        st.success("Path is valid")
    else:
        st.warning("Seleccione una carpeta válida.")
