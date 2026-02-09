import streamlit as st
import pdfplumber
import re
import pandas as pd
import io

def extraer_datos_sos(pdf_file):
    """
    Extrae datos del archivo PDF (file-like object o path).
    Soporta formato estándar y "Validador Web".
    """
    datos_extraidos = {}
    
    try:
        with pdfplumber.open(pdf_file) as pdf:
            # Seleccionamos la primera página
            if not pdf.pages:
                return {"error": "El PDF está vacío"}
            
            pagina = pdf.pages[0]
            texto = pagina.extract_text() or ""
            
            # 1. Extraer datos generales usando Expresiones Regulares (Regex)
            # Adaptado para soportar múltiples layouts (Informe Autorización vs Validador Web)
            patrones = {
                "fecha_consulta": r"Fecha Consulta:\s*([\d/]+)",
                "identificacion": r"Identificación:\s*([\d\.]+)",
                # Afiliado puede terminar en Plan, Identificación o salto de línea
                "afiliado": r"Afiliado\s*:\s*(.+?)(?=\s+Identificación|\s+Plan|\n|$)",
                "plan": r"Plan:\s*(.+?)(?=\s+Rango|\n|$)",
                "rango_salarial": r"Rango Salarial:\s*(.+?)(?=\n|$)",
                "derecho": r"Derecho:\s*(.+?)(?=\s+Ambito|\s+IPS Primaria|\n|$)",
                "ambito": r"Ambito:\s*(.+?)(?=\n|$)",
                "ips_primaria": r"IPS Primaria:\s*(.+?)(?=\s+IPS Solicitante|\n|$)",
                "ips_solicitante": r"IPS Solicitante:\s*(.+?)(?=\n|$)"
            }

            for clave, patron in patrones.items():
                match = re.search(patron, texto, re.IGNORECASE)
                if match:
                    val = match.group(1).strip()
                    # Limpieza extra por si el regex capturó basura del siguiente campo
                    val = re.sub(r'\s+(Identificación|Plan|Rango|Ambito|IPS).*$', '', val, flags=re.IGNORECASE)
                    datos_extraidos[clave] = val

            # 2. Extraer la tabla de prestaciones
            # Estrategia: Buscar tablas y detectar estructura por número de columnas o encabezados
            tablas = pagina.extract_tables()
            
            datos_extraidos["items"] = [] # Lista para soportar múltiples items si los hubiera
            
            if tablas:
                for tabla in tablas:
                    for fila in tabla:
                        if not fila: continue
                        
                        # Limpiar fila
                        fila_str = [str(c).strip() if c else "" for c in fila]
                        
                        # Ignorar encabezados
                        row_text = "".join(fila_str).lower()
                        if "código" in row_text or "autorización" in row_text:
                            if "nro." in row_text or "p-autorización" in row_text:
                                continue # Es header
                            if "nombre" in row_text and "cantidad" in row_text:
                                continue # Es header estándar
                        
                        item = {}
                        
                        # DETECCIÓN DE FORMATO
                        
                        # Formato 1: Validador Web (4 columnas)
                        # [Nro. P-Autorización, Aprobada, Justificación Resultado, Nro. Autorización]
                        if len(fila) == 4:
                            # Validar que parezca datos (ej. columna 0 o 3 tiene números)
                            if (fila_str[0].isdigit() or fila_str[3].isdigit()):
                                item = {
                                    "codigo_prestacion": fila_str[0], # Nro. P-Autorización
                                    "nombre_prestacion": "Ver P-Autorización", # No hay nombre explícito en este formato
                                    "cantidad": "1", # Asumimos 1
                                    "respuesta": fila_str[1], # Aprobada (SI/NO)
                                    "justificacion": fila_str[2],
                                    "no_autorizacion": fila_str[3]
                                }

                        # Formato 2: Estándar SOS (8+ columnas)
                        # [Codigo, Nombre, Cantidad, Respuesta, ..., No. Autorizacion]
                        elif len(fila) >= 8:
                            if fila_str[0].isdigit() and len(fila_str[0]) >= 5:
                                item = {
                                    "codigo_prestacion": fila_str[0],
                                    "nombre_prestacion": fila_str[1].replace("\n", " "),
                                    "cantidad": fila_str[2],
                                    "respuesta": fila_str[3],
                                    "no_autorizacion": fila_str[7]
                                }
                        
                        if item:
                            datos_extraidos["items"].append(item)
                            # Actualizar datos planos con el primer item encontrado (para compatibilidad)
                            if "no_autorizacion" not in datos_extraidos:
                                datos_extraidos.update(item)
                        
    except Exception as e:
        return {"error": str(e)}

    return datos_extraidos

# --- INTERFAZ GRÁFICA CON STREAMLIT ---
st.set_page_config(page_title="Analizador SOS", page_icon="🏥", layout="wide")

st.title("🏥 Analizador de Autorizaciones SOS")
st.write("Sube tu archivo PDF (Autorización o Validador Web) para extraer la información.")

uploaded_file = st.file_uploader("Cargar PDF", type="pdf")

if uploaded_file is not None:
    with st.spinner("Analizando documento..."):
        # Extraer datos
        datos = extraer_datos_sos(uploaded_file)
        
        if "error" in datos:
            st.error(f"Error al procesar el archivo: {datos['error']}")
        else:
            st.success("✅ Análisis completado")
            
            # Mostrar datos principales
            st.subheader("📋 Datos del Paciente")
            
            # Crear layout dinámico
            col1, col2, col3 = st.columns(3)
            with col1:
                st.markdown(f"**Paciente:** {datos.get('afiliado', '-')}")
                st.markdown(f"**ID:** {datos.get('identificacion', '-')}")
                st.markdown(f"**Rango Salarial:** {datos.get('rango_salarial', '-')}")
            
            with col2:
                st.markdown(f"**Plan:** {datos.get('plan', '-')}")
                st.markdown(f"**Fecha:** {datos.get('fecha_consulta', '-')}")
                st.markdown(f"**Derecho:** {datos.get('derecho', '-')}")

            with col3:
                st.markdown(f"**IPS Primaria:** {datos.get('ips_primaria', '-')}")
                st.markdown(f"**IPS Solicitante:** {datos.get('ips_solicitante', '-')}")

            st.divider()

            # Mostrar detalle prestación
            st.subheader("💊 Detalle de Autorización")
            
            items = datos.get("items", [])
            if items:
                df_detalle = pd.DataFrame(items)
                st.dataframe(df_detalle, use_container_width=True)
            else:
                st.warning("No se encontraron tablas de autorización legibles.")
            
            # Opciones de descarga
            st.subheader("📥 Descargar Resultados")
            
            # Preparar CSV (Flatten data)
            flat_data = []
            if items:
                for item in items:
                    row = datos.copy()
                    row.pop("items", None) # Remove list
                    row.update(item) # Merge item details
                    flat_data.append(row)
            else:
                row = datos.copy()
                row.pop("items", None)
                flat_data.append(row)

            df_full = pd.DataFrame(flat_data)
            csv = df_full.to_csv(index=False).encode('utf-8-sig')
            
            st.download_button(
                label="Descargar CSV",
                data=csv,
                file_name="analisis_sos_completo.csv",
                mime="text/csv",
            )
            
            # JSON Preview
            with st.expander("Ver JSON crudo"):
                st.json(datos)

# Footer
st.markdown("---")
st.caption("Soporta formato 'Informe de Autorización' y 'Validador Web'.")
