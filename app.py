import streamlit as st
import pandas as pd
import google.generativeai as genai
from docx import Document
from io import BytesIO
import re
import requests
from bs4 import BeautifulSoup
import PyPDF2
import os
import numpy as np

# Configuración inicial de la API de Gemini
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
genai.configure(api_key=GEMINI_API_KEY)
model = genai.GenerativeModel('gemini-1.5-pro-latest')

# Función para validar URLs
def is_valid_url(url):
    regex = re.compile(
        r'^(?:http|ftp)s?://'  # http:// or https://
        r'(?:(?:[A-Z0-9](?:[A-Z0-9-]{0,61}[A-Z0-9])?\.)+(?:[A-Z]{2,6}\.?|[A-Z0-9-]{2,}\.?)|'
        r'localhost|'  # localhost...
        r'\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})'  # ...or ip
        r'(?::\d+)?'  # optional port
        r'(?:/?|[/?]\S+)$', re.IGNORECASE)
    return re.match(regex, url) is not None

# Funciones para leer archivos
def read_excel(file):
    try:
        return pd.read_excel(file, engine='openpyxl'), "excel/csv"
    except Exception as e:
        st.error(f"Error al leer Excel: {e}")
        return pd.DataFrame(), "none"

def read_csv(file):
    try:
        return pd.read_csv(file, encoding='latin1', on_bad_lines='skip'), "excel/csv"
    except Exception as e:
        st.error(f"Error al leer CSV: {e}")
        return pd.DataFrame(), "none"

def read_docx(file):
    document = Document(file)
    text = "\n".join([paragraph.text for paragraph in document.paragraphs])
    return pd.DataFrame([text], columns=['text']), "word"

def read_pdf(file):
    try:
        pdf_reader = PyPDF2.PdfReader(file)
        text = "\n".join([page.extract_text() for page in pdf_reader.pages if page.extract_text()])
        return pd.DataFrame([text], columns=['text']), "pdf"
    except Exception as e:
        st.error(f"Error al leer PDF: {e}")
        return pd.DataFrame(), "none"

def fetch_web_content(url):
    try:
        response = requests.get(url)
        response.raise_for_status()
        soup = BeautifulSoup(response.content, 'html.parser')
        text = soup.get_text(separator='\n')
        return pd.DataFrame([text], columns=['text']), "web"
    except Exception as e:
        st.error(f"Error al leer URL: {e}")
        return pd.DataFrame(), "none"

# Función para calcular cuartiles (Q1 Mejor, Q4 Peor)
def calculate_quartiles(df, column):
    if column in df.columns and pd.api.types.is_numeric_dtype(df[column]):
        # Dividir en cuartiles, etiquetando Q1 como el mejor y Q4 como el peor
        quartiles = pd.qcut(df[column], q=4, labels=["Q4 (Peor)", "Q3", "Q2", "Q1 (Mejor)"], duplicates='drop')
        return df.assign(Quartil=quartiles)
    else:
        st.error(f"La columna '{column}' no es válida o no es numérica.")
        return df

# Generar prompts según el nivel de análisis
def generate_prompt(level, context_text, df_display):
    prompts = {
        "Análisis de Ranking": f"""
            Realiza un análisis exhaustivo del ranking de agentes en el call center basado en métricas clave.
            **Contexto:** {context_text}.
            **Datos:** {df_display.to_string()}.
            Aspectos a cubrir:
            - Rendimiento individual: Identifica fortalezas y áreas de mejora.
            - Comparativa: Analiza diferencias entre agentes destacados y de bajo rendimiento.
            - Recomendaciones: Propón estrategias para mejorar el desempeño.
            - Top 10 mejores y peores: Describe características clave.
            - Cuartilización: Los datos han sido divididos en cuartiles, donde Q1 (Mejor) representa el 25% superior y Q4 (Peor) representa el 25% inferior. Analiza las diferencias entre los cuartiles.
            """,
        "Tiempos Productivos, Hold, Baño, Break": f"""
            Analiza tiempos productivos, hold, pausas y descansos en el call center.
            **Contexto:** {context_text}.
            **Datos:** {df_display.to_string()}.
            Aspectos a cubrir:
            - Tiempo productivo: Evalúa eficiencia operativa.
            - Tiempos de espera (Hold): Optimiza patrones.
            - Pausas (Baño, Break): Analiza frecuencia y duración.
            - Recomendaciones: Propón mejoras para maximizar productividad.
            """,
        "Tableros de Incidencias": f"""
            Analiza incidencias reportadas en el call center.
            **Contexto:** {context_text}.
            **Datos:** {df_display.to_string()}.
            Aspectos a cubrir:
            - Identificación de incidencias recurrentes: Detecta problemas frecuentes.
            - Frecuencia e impacto: Evalúa cómo afectan la operación.
            - Resolución: Analiza procedimientos actuales.
            - Recomendaciones: Propón estrategias para reducir incidencias.
            """,
        "Satisfacción del Cliente": f"""
            Evalúa métricas de satisfacción del cliente como NPS, CSAT y comentarios.
            **Contexto:** {context_text}.
            **Datos:** {df_display.to_string()}.
            Aspectos a cubrir:
            - Métricas clave: Evalúa tendencias y patrones.
            - Factores influyentes: Identifica qué impulsa la satisfacción o insatisfacción.
            - Recomendaciones: Propón estrategias para mejorar la experiencia del cliente.
            """,
        "Costos y Rentabilidad": f"""
            Analiza costos operativos, márgenes de rentabilidad y eficiencia en el uso de recursos.
            **Contexto:** {context_text}.
            **Datos:** {df_display.to_string()}.
            Aspectos a cubrir:
            - Costos operativos: Evalúa estructura y tendencias.
            - Rentabilidad: Analiza márgenes y drivers clave.
            - Recomendaciones: Propón estrategias para optimizar costos y aumentar rentabilidad.
            """,
        "Libre": f"""
            Realiza un análisis libre basado exclusivamente en el contexto proporcionado.
            **Contexto:** {context_text}.
            Aspectos a cubrir:
            - Análisis detallado: Extrae insights clave.
            - Recomendaciones: Propón acciones concretas.
            - Observaciones: Identifica oportunidades implícitas.
            """
    }
    return prompts[level]

# Interfaz de usuario
uploaded_file = st.file_uploader("Carga tu archivo Excel, CSV, Word o PDF", type=["xls", "xlsx", "csv", "docx", "pdf"])
web_url = st.text_input("Ingresa la URL de la Página Web")
context_text = st.text_area("Ingresa el contexto para el análisis")

# Procesamiento de datos
df, data_type = pd.DataFrame(), "none"
if uploaded_file:
    file_type = uploaded_file.type
    if file_type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
        df, data_type = read_excel(uploaded_file)
    elif file_type == "text/csv":
        df, data_type = read_csv(uploaded_file)
    elif file_type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
        df, data_type = read_docx(uploaded_file)
    elif file_type == "application/pdf":
        df, data_type = read_pdf(uploaded_file)
    else:
        st.error("Tipo de archivo no soportado.")
elif web_url and is_valid_url(web_url):
    df, data_type = fetch_web_content(web_url)

# Mostrar datos cargados
if data_type != "none" and not df.empty:
    st.write("Archivo cargado exitosamente!")
    for col in df.columns:
        if df[col].dtype == 'object':
            df[col] = df[col].astype(str)
    st.dataframe(df)

    selected_columns = st.multiselect("Selecciona las columnas a analizar", df.columns.tolist())
    if selected_columns:
        df_with_quartiles = calculate_quartiles(df, selected_columns[0])
        st.dataframe(df_with_quartiles)
        df_display = df_with_quartiles.head(100)
    else:
        df_display = df.head(100)

    level = st.selectbox("Selecciona el nivel del análisis", [
        "Análisis de Ranking", 
        "Tiempos Productivos, Hold, Baño, Break", 
        "Tableros de Incidencias", 
        "Satisfacción del Cliente", 
        "Costos y Rentabilidad", 
        "Libre"
    ])

    if st.button("Empezar el análisis 🚀"):
        if df.empty:
            st.error("No hay datos para analizar. Por favor, carga un archivo o ingresa una URL.")
        else:
            prompt = generate_prompt(level, context_text, df_display)
            try:
                contents = [
                    {"role": "user", "parts": ["Eres un analista senior. Aplica un análisis riguroso y multifacético."]},
                    {"role": "user", "parts": [prompt]}
                ]
                response = model.generate_content(contents)
                informe = response.text
                st.write("--- Informe Generado ---")
                lines = informe.splitlines()
                for line in lines:
                    if line.startswith(("Título principal:", "Subtítulo")):
                        st.markdown(f"<h2 style='color: blue;'>{line}</h2>", unsafe_allow_html=True)
                    else:
                        st.write(line)

                # Generar documento Word
                document = Document()
                document.add_heading('Informe Generado por CAT-AI', 0)
                for line in lines:
                    document.add_paragraph(line)
                docx_stream = BytesIO()
                document.save(docx_stream)
                docx_stream.seek(0)
                st.download_button(
                    label="Descargar informe en Word",
                    data=docx_stream,
                    file_name="informe.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"Error al generar el informe: {e}")
else:
    st.info("Por favor, carga un archivo o ingresa una URL válida.")