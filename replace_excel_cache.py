import re

with open('src/tabs/tab_automated_actions.py', 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace("""            uploaded_excel.seek(0)
            xls = pd.ExcelFile(uploaded_excel)
            sheet = st.selectbox("Hoja", xls.sheet_names, key="dist_base_sheet")
            if sheet:
                df_preview = pd.read_excel(uploaded_excel, sheet_name=sheet, nrows=5)""",
"""            file_bytes = uploaded_excel.getvalue()
            sheet_names = _get_excel_sheet_names(file_bytes)
            sheet = st.selectbox("Hoja", sheet_names, key="dist_base_sheet")
            if sheet:
                df_preview = _get_excel_preview(file_bytes, sheet, nrows=5)""")

content = content.replace("""            uploaded.seek(0)
            xls = pd.ExcelFile(uploaded)
            sheet = st.selectbox("Hoja", xls.sheet_names, key="create_fold_sheet")
            if sheet:
                df_preview = pd.read_excel(uploaded, sheet_name=sheet, nrows=5)""",
"""            file_bytes = uploaded.getvalue()
            sheet_names = _get_excel_sheet_names(file_bytes)
            sheet = st.selectbox("Hoja", sheet_names, key="create_fold_sheet")
            if sheet:
                df_preview = _get_excel_preview(file_bytes, sheet, nrows=5)""")

content = content.replace("""            xl = pd.ExcelFile(uploaded)
            sheet = st.selectbox("Seleccione la Hoja", xl.sheet_names, key="suf_sheet")
            df_preview = pd.read_excel(uploaded, sheet_name=sheet, nrows=1)""",
"""            file_bytes = uploaded.getvalue()
            sheet_names = _get_excel_sheet_names(file_bytes)
            sheet = st.selectbox("Seleccione la Hoja", sheet_names, key="suf_sheet")
            df_preview = _get_excel_preview(file_bytes, sheet, nrows=1)""")

content = content.replace("""            if hasattr(uploaded, 'seek'):
                uploaded.seek(0)
            xls = pd.ExcelFile(uploaded)
            sheet = st.selectbox("Nombre Hoja", xls.sheet_names, key="ren_map_sheet")
            if sheet:
                df_preview = pd.read_excel(uploaded, sheet_name=sheet, nrows=5)""",
"""            file_bytes = uploaded.getvalue()
            sheet_names = _get_excel_sheet_names(file_bytes)
            sheet = st.selectbox("Nombre Hoja", sheet_names, key="ren_map_sheet")
            if sheet:
                df_preview = _get_excel_preview(file_bytes, sheet, nrows=5)""")

content = content.replace("""            uploaded.seek(0)
            xls = pd.ExcelFile(uploaded)
            sheet = st.selectbox("Nombre Hoja", xls.sheet_names, key="mod_full_sheet")""",
"""            file_bytes = uploaded.getvalue()
            sheet_names = _get_excel_sheet_names(file_bytes)
            sheet = st.selectbox("Nombre Hoja", sheet_names, key="mod_full_sheet")""")

content = content.replace("""            excel_file.seek(0)
            xl = pd.ExcelFile(excel_file)
            sheet = st.selectbox("Hoja de Excel", xl.sheet_names, key="uni_sheet")
            if sheet:
                df_preview = pd.read_excel(excel_file, sheet_name=sheet, nrows=5)""",
"""            file_bytes = excel_file.getvalue()
            sheet_names = _get_excel_sheet_names(file_bytes)
            sheet = st.selectbox("Hoja de Excel", sheet_names, key="uni_sheet")
            if sheet:
                df_preview = _get_excel_preview(file_bytes, sheet, nrows=5)""")

content = content.replace("""            excel_file.seek(0)
            xl = pd.ExcelFile(excel_file)
            sheet = st.selectbox("Seleccione la Hoja", xl.sheet_names, key="renombrar_apply_sheet")
            if sheet:
                df_preview = pd.read_excel(excel_file, sheet_name=sheet, nrows=1)""",
"""            file_bytes = excel_file.getvalue()
            sheet_names = _get_excel_sheet_names(file_bytes)
            sheet = st.selectbox("Seleccione la Hoja", sheet_names, key="renombrar_apply_sheet")
            if sheet:
                df_preview = _get_excel_preview(file_bytes, sheet, nrows=1)""")

with open('src/tabs/tab_automated_actions.py', 'w', encoding='utf-8') as f:
    f.write(content)
print("Done replacements")
