import streamlit as st
if 'test_input' not in st.session_state:
    st.session_state['test_input'] = 'Initial Value'
val = st.text_input('Label', key='test_input')
with open('result.txt', 'w') as f:
    f.write(repr(val))
