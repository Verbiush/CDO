import streamlit as st
import pandas as pd
import time
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

def render(*args, **kwargs):
    role = st.session_state.get("role", "user")
    
    if role != "admin":
        st.error("⛔ Acceso Denegado. Se requieren permisos de Administrador para ver esta sección.")
        return

    st.header("👥 Gestión de Usuarios")
    
    tab_list, tab_create, tab_edit = st.tabs([
        "Listar / Eliminar", "Crear Nuevo", "✏️ Editar Permisos"
    ])

    with tab_list:
        users = db.get_all_users()
        df_data = [{"Usuario": u, "Rol": d.get("role", "user"), "Bot Zeus": d.get("permissions", {}).get("bot_zeus", "full")} for u, d in users.items()]
        st.dataframe(df_data, use_container_width=True)
        
        st.divider()
        st.subheader("🗑️ Eliminar Usuario")
        user_to_delete = st.selectbox("Seleccionar usuario a eliminar", [u for u in users.keys() if u != "admin"])
        
        if st.button("Eliminar Usuario Seleccionado", type="primary"):
            if user_to_delete:
                ok, msg = db.delete_user(user_to_delete)
                if ok:
                    st.success(msg)
                    time.sleep(1)
                    st.rerun()
                else:
                    st.error(msg)
    
    with tab_create:
        st.subheader("➕ Nuevo Usuario")
        new_user = st.text_input("Nombre de usuario")
        new_pass = st.text_input("Contraseña", type="password")
        new_role = st.selectbox("Rol", ["user", "manager", "admin"])
        
        if st.button("Crear Usuario"):
            if new_user and new_pass:
                # Create with default permissions
                ok, msg = db.create_user(new_user, new_pass, role=new_role)
                if ok:
                    st.success(msg)
                    time.sleep(1)
                    st.rerun()
                else:
                    st.error(msg)
            else:
                st.warning("Complete todos los campos")

    with tab_edit:
        st.subheader("✏️ Editar Permisos y Roles")
        users = db.get_all_users()
        user_to_edit = st.selectbox("Seleccionar usuario", list(users.keys()))
        
        if user_to_edit:
            user_data = users[user_to_edit]
            current_role = user_data.get("role", "user")
            current_perms = user_data.get("permissions", {})
            current_bot = current_perms.get("bot_zeus", "full") # default full for backward compat
            current_tabs = current_perms.get("allowed_tabs", ["*"])
            
            # Form
            role_options = ["user", "manager", "admin"]
            try:
                role_index = role_options.index(current_role)
            except:
                role_index = 0
                
            new_role_edit = st.selectbox("Rol del Usuario", role_options, index=role_index, key=f"edit_role_{user_to_edit}")
            
            st.divider()
            st.markdown("**Permisos Específicos**")
            
            # Bot Zeus Permissions
            bot_options = ["full", "edit", "execute", "none"]
            bot_labels = ["Completo (Crear/Editar/Ejecutar)", "Edición (Editar/Ejecutar)", "Solo Ejecución", "Sin Acceso"]
            
            try:
                bot_index = bot_options.index(current_bot)
            except:
                bot_index = 0
                
            new_bot_perm = st.selectbox(
                "🤖 Permisos Bot Zeus", 
                bot_options, 
                format_func=lambda x: bot_labels[bot_options.index(x)],
                index=bot_index,
                key=f"edit_bot_{user_to_edit}"
            )
            
            # Tab Visibility
            all_tabs_available = [
                "🔎 Búsqueda y Acciones",
                "⚙️ Acciones Automatizadas",
                "🔄 Conversión de Archivos",
                "📄 Visor (JSON/XML)",
                "RIPS",
                "📂 Gestión Documental",
                "👤 Validación Usuario",
                "🤖 Asistente IA (Gemini)",
                "🤖 Bot Zeus Salud",
                "📊 Gestión de Información",
                "👥 Gestión de Usuarios"
            ]
            
            # Helper for multiselect default
            default_tabs = all_tabs_available if "*" in current_tabs else [t for t in current_tabs if t in all_tabs_available]
            
            new_tabs = st.multiselect(
                "👁️ Pestañas Visibles",
                all_tabs_available,
                default=default_tabs,
                key=f"edit_tabs_{user_to_edit}"
            )
            
            if st.button("💾 Guardar Cambios de Permisos", key=f"btn_save_{user_to_edit}"):
                # Update role
                db.update_user_role(user_to_edit, new_role_edit)
                
                # Update permissions
                new_perms = current_perms.copy()
                new_perms["bot_zeus"] = new_bot_perm
                
                # Enforce: Only admins can see "Gestión de Usuarios"
                if new_role_edit != "admin":
                    if "👥 Gestión de Usuarios" in new_tabs:
                        new_tabs.remove("👥 Gestión de Usuarios")
                        st.toast("⚠️ 'Gestión de Usuarios' removido (Requiere Admin)", icon="🛡️")

                # If all selected, save as "*"
                if len(new_tabs) == len(all_tabs_available):
                    new_perms["allowed_tabs"] = ["*"]
                else:
                    new_perms["allowed_tabs"] = new_tabs
                    
                # Use the dedicated function to update the permissions column
                db.update_user_permissions(user_to_edit, new_perms)
                st.success(f"Permisos actualizados para {user_to_edit}")
                time.sleep(1)
                st.rerun()
