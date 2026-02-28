import streamlit as st
import pandas as pd
try:
    import plotly.express as px
    HAS_PLOTLY = True
except ImportError:
    HAS_PLOTLY = False
import time
import datetime
import os
import sys
import json

# Try importing database
try:
    import database as db
except ImportError:
    try:
        from src import database as db
    except ImportError:
        # Fallback if running from tabs dir
        sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
        import database as db

# Try importing db_gestion for specific admin tasks
try:
    import db_gestion
except ImportError:
    try:
        from src import db_gestion
    except ImportError:
        # Fallback if running from tabs dir
        sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
        import db_gestion

def render(*args, **kwargs):
    role = st.session_state.get("role", "user")
    
    if role not in ["admin", "manager"]:
        st.error("⛔ Acceso Denegado. Se requieren permisos de Administrador o Manager para ver esta sección.")
        return

    # Unified header for both roles since user management is now in a separate tab
    st.header("📊 Gestión de Información")
    
    tab_data, tab_backup, tab_reports, tab_sql = st.tabs([
        "🗄️ Explorador BD", "📦 Respaldo BD", "📊 Informes y Gráficos", "🛠️ Admin SQL"
    ])

    # --- SHARED TABS (Admin & Manager) ---
    with tab_data:
        st.subheader("🗄️ Explorador de Base de Datos")
        st.info("Visualice y gestione los datos de las tablas del sistema (Usuarios, Facturas, Tareas, etc).")
        
        try:
            # Use a new connection for safety
            conn = db.get_connection()
            
            # Get tables
            tables_df = pd.read_sql("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name", conn)
            table_list = tables_df['name'].tolist() if not tables_df.empty else []
            
            selected_table = st.selectbox("Seleccionar Tabla", table_list, index=0 if table_list else None)
            
            if selected_table:
                st.markdown(f"### Tabla: `{selected_table}`")
                
                # Load data
                df = pd.read_sql(f"SELECT * FROM {selected_table}", conn)
                st.dataframe(df, use_container_width=True)
                
                st.divider()
                col_d1, col_d2 = st.columns(2)
                
                with col_d1:
                    st.markdown("**📥 Exportar Datos**")
                    # CSV
                    csv = df.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        "⬇️ Descargar como CSV",
                        csv,
                        f"{selected_table}_export.csv",
                        "text/csv",
                        key=f"dl_csv_{selected_table}"
                    )
                    
                with col_d2:
                    st.markdown("**🗑️ Gestión de Registros**")
                    
                    # Delete by ID (if applicable)
                    if 'id' in df.columns:
                        with st.expander("Eliminar Registro por ID"):
                            id_to_del = st.number_input("ID a eliminar", min_value=0, step=1, key=f"del_id_{selected_table}")
                            if st.button("❌ Eliminar Registro", key=f"btn_del_id_{selected_table}", type="primary"):
                                try:
                                    cursor = conn.cursor()
                                    cursor.execute(f"DELETE FROM {selected_table} WHERE id = ?", (id_to_del,))
                                    conn.commit()
                                    st.success(f"Registro ID {id_to_del} eliminado.")
                                    time.sleep(1)
                                    # st.rerun()
                                except Exception as e:
                                    st.error(f"Error al eliminar: {e}")
                    
                    # Delete All (Dangerous)
                    with st.expander("⚠️ Zona de Peligro (Eliminar Todo)"):
                        st.warning("Esta acción eliminará TODOS los registros de la tabla y no se puede deshacer.")
                        confirm_del = st.checkbox(f"Confirmar vaciado de tabla '{selected_table}'", key=f"chk_del_{selected_table}")
                        if confirm_del:
                            if st.button("🔥 VACIAR TABLA COMPLETA", key=f"btn_truncate_{selected_table}", type="primary"):
                                try:
                                    cursor = conn.cursor()
                                    cursor.execute(f"DELETE FROM {selected_table}")
                                    conn.commit()
                                    st.success(f"Tabla {selected_table} vaciada completamente.")
                                    time.sleep(1)
                                    # st.rerun()
                                except Exception as e:
                                    st.error(f"Error al vaciar tabla: {e}")

        except Exception as e:
            st.error(f"Error accediendo a la base de datos: {e}")
        finally:
            if 'conn' in locals():
                conn.close()

    with tab_backup:
        st.subheader("📦 Respaldo de Base de Datos")
        st.info("Descargue una copia de seguridad de la base de datos actual (users.db).")
        
        db_path = db.get_db_path()
        if os.path.exists(db_path):
            with open(db_path, "rb") as f:
                db_bytes = f.read()
            
            # Filename with timestamp
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            file_name = f"backup_users_db_{timestamp}.db"
            
            st.download_button(
                label="⬇️ Descargar Copia de Seguridad",
                data=db_bytes,
                file_name=file_name,
                mime="application/x-sqlite3",
                key="btn_backup_db"
            )
        else:
            st.error("No se encontró el archivo de base de datos.")

    with tab_reports:
        st.subheader("📊 Informes y Estadísticas")
        
        # Key Metrics
        try:
            # 1. Fetch ALL data first
            all_invoices = db.get_all_invoices()
            
            if not all_invoices:
                st.info("No hay datos para mostrar.")
            else:
                df_all = pd.DataFrame(all_invoices)
                
                # --- FILTROS ---
                with st.expander("🔍 Filtros de Búsqueda", expanded=True):
                    col_f1, col_f2, col_f3 = st.columns(3)
                    
                    # A. Filter by Date Range
                    with col_f1:
                        st.markdown("**📅 Rango de Fechas**")
                        # Convert to datetime for filtering
                        df_all['fecha_dt'] = pd.to_datetime(df_all['fecha_factura'], errors='coerce')
                        
                        min_date = df_all['fecha_dt'].min()
                        max_date = df_all['fecha_dt'].max()
                        
                        if pd.isnull(min_date): min_date = datetime.date.today()
                        if pd.isnull(max_date): max_date = datetime.date.today()
                        
                        date_range = st.date_input(
                            "Seleccionar Rango",
                            value=(min_date, max_date),
                            key="filter_date_range"
                        )
                    
                    # B. Filter by EPS
                    with col_f2:
                        st.markdown("**🏥 EPS**")
                        if "eps" in df_all.columns:
                            eps_options = sorted(df_all['eps'].dropna().unique().tolist())
                            selected_eps = st.multiselect("Filtrar por EPS", eps_options, default=[], key="filter_eps")
                        else:
                            selected_eps = []
                            
                    # C. Filter by Regimen
                    with col_f3:
                        st.markdown("**📋 Régimen**")
                        if "regimen" in df_all.columns:
                            reg_options = sorted(df_all['regimen'].dropna().unique().tolist())
                            selected_regimen = st.multiselect("Filtrar por Régimen", reg_options, default=[], key="filter_regimen")
                        else:
                            selected_regimen = []

                # --- APPLY FILTERS ---
                df_filtered = df_all.copy()
                
                # Apply Date Filter
                if isinstance(date_range, tuple) and len(date_range) == 2:
                    start_date, end_date = date_range
                    # Ensure we compare date objects
                    mask_date = (df_filtered['fecha_dt'].dt.date >= start_date) & (df_filtered['fecha_dt'].dt.date <= end_date)
                    df_filtered = df_filtered[mask_date]
                
                # Apply EPS Filter
                if selected_eps:
                    df_filtered = df_filtered[df_filtered['eps'].isin(selected_eps)]
                    
                # Apply Regimen Filter
                if selected_regimen:
                    df_filtered = df_filtered[df_filtered['regimen'].isin(selected_regimen)]
                
                # --- CALCULATE METRICS ON FILTERED DATA ---
                # Define subsets based on filtered data
                # Pending: status == 'PENDING' (case insensitive check just in case)
                pending_mask = df_filtered['status'].str.upper() == 'PENDING'
                df_pending = df_filtered[pending_mask]
                
                # Radicado: radicado is not null and not empty
                radicado_mask = df_filtered['radicado'].notna() & (df_filtered['radicado'] != '')
                df_rad = df_filtered[radicado_mask]
                
                st.divider()
                
                col_m1, col_m2, col_m3 = st.columns(3)
                col_m1.metric("Total Facturas (Filtradas)", len(df_filtered))
                col_m2.metric("Pendientes", len(df_pending))
                col_m3.metric("Con Radicado", len(df_rad))
                
                st.divider()
                
                report_type = st.radio("Seleccionar Informe", ["Facturas Pendientes", "Facturas con Radicado", "Análisis General"], horizontal=True, key="report_type_radio")
                
                if report_type == "Facturas Pendientes":
                    st.markdown("#### Facturas Pendientes de Pago/Trámite")
                    if not df_pending.empty:
                        # Select useful columns if available
                        desired_cols = ["no_factura", "fecha_factura", "total", "status", "eps", "regimen", "nombre_completo", "no_doc"]
                        cols = [c for c in desired_cols if c in df_pending.columns]
                        st.dataframe(df_pending[cols] if cols else df_pending, use_container_width=True)
                    else:
                        st.info("No hay facturas pendientes en la selección actual.")
                        
                elif report_type == "Facturas con Radicado":
                    st.markdown("#### Facturas con Radicado Asignado")
                    if not df_rad.empty:
                        desired_cols = ["no_factura", "radicado", "fecha_radicado", "total", "eps", "regimen", "nombre_completo"]
                        cols = [c for c in desired_cols if c in df_rad.columns]
                        st.dataframe(df_rad[cols] if cols else df_rad, use_container_width=True)
                    else:
                        st.info("No hay facturas con radicado en la selección actual.")
                        
                elif report_type == "Análisis General":
                    st.markdown("#### Análisis General")
                    if not df_filtered.empty:
                        
                        col_g1, col_g2 = st.columns(2)
                        
                        with col_g1:
                            # Status Distribution
                            if "status" in df_filtered.columns:
                                st.markdown("**Distribución por Estado**")
                                status_counts = df_filtered['status'].value_counts().reset_index()
                                status_counts.columns = ['Estado', 'Cantidad']
                                
                                if HAS_PLOTLY:
                                    fig = px.pie(status_counts, values='Cantidad', names='Estado', title='Estado de Facturas')
                                    st.plotly_chart(fig, use_container_width=True)
                                else:
                                    st.bar_chart(df_filtered['status'].value_counts())
                        
                        with col_g2:
                            # EPS Distribution
                            if "eps" in df_filtered.columns:
                                st.markdown("**Distribución por EPS**")
                                eps_counts = df_filtered['eps'].fillna("Sin EPS").value_counts().reset_index()
                                eps_counts.columns = ['EPS', 'Cantidad']
                                
                                if HAS_PLOTLY:
                                    fig = px.pie(eps_counts, values='Cantidad', names='EPS', title='Facturas por EPS')
                                    st.plotly_chart(fig, use_container_width=True)
                                else:
                                    st.bar_chart(df_filtered['eps'].value_counts())

                        # Regimen Distribution (New Row)
                        if "regimen" in df_filtered.columns:
                            st.markdown("**Distribución por Régimen**")
                            reg_counts = df_filtered['regimen'].fillna("Sin Régimen").value_counts().reset_index()
                            reg_counts.columns = ['Régimen', 'Cantidad']
                            
                            if HAS_PLOTLY:
                                fig = px.pie(reg_counts, values='Cantidad', names='Régimen', title='Facturas por Régimen')
                                st.plotly_chart(fig, use_container_width=True)
                            else:
                                st.bar_chart(df_filtered['regimen'].value_counts())
                        
                        # Monthly Billing
                        if "fecha_factura" in df_filtered.columns and "total" in df_filtered.columns:
                            try:
                                # Clean total (remove currency symbols, commas, convert to float)
                                df_filtered['total_clean'] = df_filtered['total'].astype(str).str.replace(r'[^\d]', '', regex=True)
                                df_filtered['total_num'] = pd.to_numeric(df_filtered['total_clean'], errors='coerce').fillna(0)
                                
                                # Use already parsed date
                                # Drop invalid dates
                                df_billing = df_filtered.dropna(subset=['fecha_dt'])
                                
                                if not df_billing.empty:
                                    df_billing['mes_año'] = df_billing['fecha_dt'].dt.strftime('%Y-%m')
                                    monthly_billing = df_billing.groupby('mes_año')['total_num'].sum()
                                    
                                    st.markdown("**Facturación Mensual**")
                                    st.bar_chart(monthly_billing)
                            except Exception as e:
                                st.warning(f"No se pudo generar gráfico mensual: {e}")
                    else:
                        st.info("No hay datos suficientes en la selección actual.")
                    
        except Exception as e:
            st.error(f"Error generando informes: {e}")
            import traceback
            st.text(traceback.format_exc())



    with tab_sql:
        st.subheader("🛠️ Administración de Base de Datos SQL")
        
        st.markdown("""
        Esta sección permite gestionar la base de datos **SQLite** (Motor SQL Integrado).
        
        **Estado Actual:**
        - Motor: SQLite 3
        - Esquema: Relacional (Pacientes -> Atenciones -> Facturas)
        - Archivo: `src/users.db`
        """)
        
        col_admin_1, col_admin_2 = st.columns(2)
        
        with col_admin_1:
            st.info("Estructura definida en `src/schema.sql`")
            
            # Reset DB Button
            if st.button("🔄 Reiniciar Base de Datos (Borrar Todo)", type="primary"):
                pass
            
            confirm_reset = st.checkbox("Estoy seguro de BORRAR TODOS LOS DATOS y reiniciar la estructura.")
            
            if confirm_reset:
                if st.button("⚠️ CONFIRMAR REINICIO ⚠️"):
                    success, msg = db_gestion.reset_database()
                    if success:
                        st.success(f"✅ Base de datos reiniciada: {msg}")
                        import time
                        time.sleep(1)
                        # st.rerun()
                    else:
                        st.error(f"❌ Error al reiniciar: {msg}")

        with col_admin_2:
             st.markdown("### 🗑️ Eliminación Rápida")
             st.caption("Eliminar registro por ID (Tabla: Facturas)")
             id_to_delete = st.number_input("ID Factura a eliminar", min_value=0, step=1, key="sql_tab_delete_id")
             if st.button("🗑️ Eliminar Factura", key="sql_tab_delete_btn"):
                 if db_gestion.delete_document_record(id_to_delete):
                     st.success(f"Registro {id_to_delete} eliminado.")
                     time.sleep(1)
                     # st.rerun()
                 else:
                     st.error("No se pudo eliminar el registro (ID no encontrado).")

        st.divider()
        st.subheader("📟 Consola SQL")
        sql_query = st.text_area("Ejecutar consulta SQL manual", "SELECT * FROM facturas ORDER BY id DESC LIMIT 5;")
        
        if st.button("▶️ Ejecutar SQL"):
            if not sql_query.strip():
                st.warning("Escriba una consulta.")
            else:
                try:
                    conn = db.get_connection()
                    # If SELECT, show dataframe
                    if sql_query.strip().upper().startswith("SELECT") or sql_query.strip().upper().startswith("PRAGMA"):
                        df = pd.read_sql_query(sql_query, conn)
                        st.dataframe(df)
                    else:
                        cursor = conn.cursor()
                        cursor.execute(sql_query)
                        conn.commit()
                        st.success(f"Consulta ejecutada. Filas afectadas: {cursor.rowcount}")
                    conn.close()
                except Exception as e:
                    st.error(f"Error SQL: {e}")
