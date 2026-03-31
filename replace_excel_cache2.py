import re

with open('src/tabs/tab_automated_actions.py', 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace("""            xl = pd.ExcelFile(uploaded)
            sheet_name = st.selectbox("Seleccione la Hoja", xl.sheet_names, key="firmas_sheet_sel")
            df_preview = pd.read_excel(uploaded, sheet_name=sheet_name, nrows=1)""",
"""            file_bytes = uploaded.getvalue()
            sheet_names = _get_excel_sheet_names(file_bytes)
            sheet_name = st.selectbox("Seleccione la Hoja", sheet_names, key="firmas_sheet_sel")
            df_preview = _get_excel_preview(file_bytes, sheet_name, nrows=1)""")

content = content.replace("""            xl = pd.ExcelFile(uploaded)
            sheet_name = st.selectbox("Seleccione la Hoja", xl.sheet_names, key="ovida_sheet_sel")
            df_preview = pd.read_excel(uploaded, sheet_name=sheet_name, nrows=1)""",
"""            file_bytes = uploaded.getvalue()
            sheet_names = _get_excel_sheet_names(file_bytes)
            sheet_name = st.selectbox("Seleccione la Hoja", sheet_names, key="ovida_sheet_sel")
            df_preview = _get_excel_preview(file_bytes, sheet_name, nrows=1)""")

content = content.replace("""                xl = pd.ExcelFile(uploaded)
                sheet_name = st.selectbox("Hoja (ReteFuente)", xl.sheet_names, key="rete_sheet_sel")
                df_preview = pd.read_excel(uploaded, sheet_name=sheet_name, nrows=1)""",
"""                file_bytes = uploaded.getvalue()
                sheet_names = _get_excel_sheet_names(file_bytes)
                sheet_name = st.selectbox("Hoja (ReteFuente)", sheet_names, key="rete_sheet_sel")
                df_preview = _get_excel_preview(file_bytes, sheet_name, nrows=1)""")

with open('src/tabs/tab_automated_actions.py', 'w', encoding='utf-8') as f:
    f.write(content)
print("Done replacements")
