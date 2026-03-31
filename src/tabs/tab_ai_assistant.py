import streamlit as st
import os
import google.generativeai as genai
import fitz  # PyMuPDF
from docx import Document

def extract_text_from_file(file_path):
    """Extrae texto de PDF, DOCX o TXT para análisis con IA."""
    ext = os.path.splitext(file_path)[1].lower()
    text = ""
    try:
        if ext == ".pdf":
            doc = fitz.open(file_path)
            for page in doc:
                text += page.get_text() + "\n"
        elif ext == ".docx":
            doc = Document(file_path)
            for para in doc.paragraphs:
                text += para.text + "\n"
        elif ext in [".txt", ".csv", ".json", ".xml", ".py", ".js", ".html"]:
            with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                text = f.read()
    except Exception as e:
        return f"Error leyendo archivo: {e}"
    return text

def worker_consultar_gemini(prompt, file_context=None):
    # Intentar obtener de config, sino usar fallback proporcionado
    config = getattr(st.session_state, "app_config", {})
    api_key = config.get("gemini_api_key")
    
    if not api_key:
        api_key = "AIzaSyAvnswqCLSWzrUctGKdZ2Un_AKYB8Gfc1w"

    if not api_key:
        return "⚠️ Por favor configura tu API Key de Google Gemini en el panel lateral."
    
    try:
        # Limpiar API Key de espacios accidentales
        api_key = api_key.strip()
        genai.configure(api_key=api_key)
        
        model_name = st.session_state.app_config.get("gemini_model", "gemini-flash-latest")
        
        # Función interna para llamar al modelo con reintentos
        def call_gemini(model_n, p_prompt, p_context=None):
             # Nota: genai.GenerativeModel suele aceptar tanto 'gemini-pro' como 'models/gemini-pro'
             # pero para mayor robustez probamos ambos si falla.
             try:
                 model = genai.GenerativeModel(model_n)
                 full_prompt = p_prompt
                 if p_context:
                     full_prompt = f"Contexto del archivo:\n{p_context}\n\nPregunta:\n{p_prompt}"
                 response = model.generate_content(full_prompt)
                 return response.text
             except Exception as e_inner:
                 # Fallback: Si falla con 404 y tiene/no tiene prefijo, intentar la inversa
                 err_msg = str(e_inner)
                 if "404" in err_msg:
                     alt_name = None
                     if model_n.startswith("models/"):
                         alt_name = model_n.replace("models/", "")
                     else:
                         alt_name = f"models/{model_n}"
                     
                     if alt_name and alt_name != model_n:
                         model = genai.GenerativeModel(alt_name)
                         full_prompt = p_prompt
                         if p_context:
                             full_prompt = f"Contexto del archivo:\n{p_context}\n\nPregunta:\n{p_prompt}"
                         response = model.generate_content(full_prompt)
                         return response.text
                 raise e_inner

        try:
            return call_gemini(model_name, prompt, file_context)
            
        except Exception as e_model:
            # Si falla por cuota (429) o no encontrado (404), intentar con la versión estable más reciente
            err_msg = str(e_model)
            if "429" in err_msg or "quota" in err_msg.lower() or "404" in err_msg:
                # Intentar con gemini-flash-latest si no era ese el que falló
                if model_name != "gemini-flash-latest":
                    return call_gemini("gemini-flash-latest", prompt, file_context)
            
            raise e_model

    except Exception as e:
        msg = str(e)
        if "API_KEY_INVALID" in msg or "403" in msg:
            return "⛔ Error de Autenticación: Tu API Key no es válida o ha expirado. Verifícala en Google AI Studio."
        if "404" in msg:
            return f"⛔ Modelo no encontrado o no disponible: {model_name}. Intenta seleccionar otro modelo."
        return f"❌ Error consultando a Gemini: {msg}"

def render(tab_container):
    with tab_container:
        st.header("🤖 Asistente IA (Gemini)")

        col_chat, col_context = st.columns([2, 1])

        with col_context:
            st.markdown("### 📄 Contexto")
            st.info("Sube un archivo o usa el contenido seleccionado para que la IA lo analice.")
    
            context_option = st.radio("Fuente de contexto:", ["Ninguno", "Subir Archivo", "Texto Manual"])
    
            file_context = None
    
            if context_option == "Subir Archivo":
                uploaded_file = st.file_uploader("Sube un archivo (PDF, DOCX, TXT)", type=["pdf", "docx", "txt", "csv", "json", "xml"], key="up_ai_assist")
                if uploaded_file:
                    # Guardar temporalmente para extraer texto
                    try:
                        with open("temp_context_file", "wb") as f:
                            f.write(uploaded_file.getbuffer())
                
                        file_context = extract_text_from_file("temp_context_file")
                        st.success(f"Archivo cargado ({len(file_context)} caracteres)")
                    except Exception as e:
                        st.error(f"Error cargando archivo: {e}")
                    finally:
                        if os.path.exists("temp_context_file"):
                            os.remove("temp_context_file")
    
            elif context_option == "Texto Manual":
                file_context = st.text_area("Pega el texto aquí:", height=300)

        with col_chat:
            st.markdown("### 💬 Chat")
    
            if "chat_history" not in st.session_state:
                st.session_state.chat_history = []
        
            # Mostrar historial
            for message in st.session_state.chat_history:
                role_icon = "👤" if message["role"] == "user" else "🤖"
                with st.chat_message(message["role"], avatar=role_icon):
                    st.markdown(message["content"])
            
            # Input de chat
            if prompt := st.chat_input("Escribe tu pregunta para Gemini..."):
                # Agregar mensaje de usuario
                st.session_state.chat_history.append({"role": "user", "content": prompt})
                with st.chat_message("user", avatar="👤"):
                    st.markdown(prompt)
        
                # Respuesta de IA
                with st.chat_message("assistant", avatar="🤖"):
                    message_placeholder = st.empty()
                    
                    with st.spinner("Gemini está pensando..."):
                        response_text = worker_consultar_gemini(prompt, file_context)
                
                    message_placeholder.markdown(response_text)
                    st.session_state.chat_history.append({"role": "assistant", "content": response_text})
    
            if st.button("🗑️ Borrar Historial"):
                st.session_state.chat_history = []
                st.rerun()
