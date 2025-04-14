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

# Configuración inicial de la API de Gemini
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
genai.configure(api_key=GEMINI_API_KEY)
model = genai.GenerativeModel('gemini-1.5-pro-latest')

# Validar URLs
def is_valid_url(url):
    regex = re.compile(
        r'^(?:http|ftp)s?://'  # http:// or https://
        r'(?:(?:[A-Z0-9](?:[A-Z0-9-]{0,61}[A-Z0-9])?\.)+(?:[A-Z]{2,6}\.?|[A-Z0-9-]{2,}\.?)|'
        r'localhost|'  # localhost
        r'\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})'
        r'(?::\d+)?(?:/?|[/?]\S+)$', re.IGNORECASE)
    return re.match(regex, url) is not None

# Función para leer archivos
def read_file(file, file_type):
    try:
        if file_type == "xlsx" or file_type == "xls":  # Para Excel
            return pd.read_excel(file, engine='openpyxl'), "excel/csv"
        elif file_type == "csv":  # Para CSV
            return pd.read_csv(file, encoding='latin1', on_bad_lines='skip'), "excel/csv"
        elif file_type == "docx":  # Para Word
            document = Document(file)
            text = "\n".join([paragraph.text for paragraph in document.paragraphs])
            return pd.DataFrame([text], columns=['text']), "word"
        elif file_type == "pdf":  # Para PDF
            pdf_reader = PyPDF2.PdfReader(file)
            text = "\n".join([page.extract_text() for page in pdf_reader.pages if page.extract_text()])
            return pd.DataFrame([text], columns=['text']), "pdf"
        elif file_type == "web":  # Para URL
            response = requests.get(file)
            response.raise_for_status()
            soup = BeautifulSoup(response.content, 'html.parser')
            text = soup.get_text(separator='\n')
            return pd.DataFrame([text], columns=['text']), "web"
    except Exception as e:
        st.error(f"Error al leer el archivo: {e}")
        return pd.DataFrame(), "none"

# Función para filtrar cada columna por su tipo de dato
def filter_columns(df):
    filtered_df = df.copy()

    # Iterar sobre cada columna para aplicar un filtro
    for column in df.columns:
        if pd.api.types.is_numeric_dtype(df[column]):  # Filtro para columnas numéricas
            min_value = df[column].min()
            max_value = df[column].max()
            filter_value = st.slider(f"Selecciona el rango para filtrar '{column}'", min_value=min_value, max_value=max_value, value=(min_value, max_value))
            filtered_df = filtered_df[(filtered_df[column] >= filter_value[0]) & (filtered_df[column] <= filter_value[1])]
        
        elif pd.api.types.is_string_dtype(df[column]):  # Filtro para columnas de texto
            unique_values = df[column].dropna().unique()
            selected_values = st.multiselect(f"Selecciona los valores de la columna '{column}'", unique_values)
            if selected_values:
                filtered_df = filtered_df[filtered_df[column].isin(selected_values)]
        
        elif pd.api.types.is_datetime64_any_dtype(df[column]):  # Filtro para columnas de fecha
            min_date = df[column].min()
            max_date = df[column].max()
            filter_date = st.date_input(f"Selecciona las fechas para filtrar '{column}'", min_value=min_date, max_value=max_date, value=(min_date, max_date))
            filtered_df = filtered_df[(filtered_df[column] >= filter_date[0]) & (filtered_df[column] <= filter_date[1])]
    
    return filtered_df

# Función para generar el informe
def generate_prompt(level, context_text, df_display):
    prompts = {
        "Análisis de Ranking": f"""
            Realiza un análisis exhaustivo del ranking de agentes en el call center basado en métricas clave.
            **Contexto:** {context_text}.
            **Datos:** {df_display.to_string()}.
            - Rendimiento individual: fortalezas y debilidades.
            - Comparación entre agentes destacados y bajos.
            - Recomendaciones estratégicas.
            - Top 10 mejores y peores agentes.
            - Análisis por cuartiles: Q1 (Mejor), Q4 (Peor).
        """,
        "Tiempos Productivos, Hold, Baño, Break": f"""
            Analiza la gestión del tiempo en el call center.
            **Contexto:** {context_text}.
            **Datos:** {df_display.to_string()}.
            - Evaluación de eficiencia.
            - Identificación de tiempos de espera ociosos.
            - Recomendaciones de mejora.
        """,
        "Tableros de Incidencias": f"""
            Analiza las incidencias en la operación del call center.
            **Contexto:** {context_text}.
            **Datos:** {df_display.to_string()}.
            - Frecuencia de incidencias.
            - Problemas recurrentes.
            - Recomendaciones para prevención y resolución.
        """,
        "Satisfacción del Cliente": f"""
            Analiza métricas de satisfacción como NPS, CSAT y comentarios.
            **Contexto:** {context_text}.
            **Datos:** {df_display.to_string()}.
            - Tendencias de satisfacción.
            - Factores que afectan la experiencia del cliente.
            - Acciones de mejora.
        """,
        "Costos y Rentabilidad": f"""
            Analiza costos operativos y márgenes del call center.
            **Contexto:** {context_text}.
            **Datos:** {df_display.to_string()}.
            - Identificación de costos relevantes.
            - Análisis de rentabilidad.
            - Oportunidades de optimización.
        """,
        "Libre": f"""
            Realiza un análisis libre y detallado según el contexto proporcionado.
            **Contexto:** {context_text}.
            **Datos:** {df_display.to_string()}.
            - Insights clave.
            - Recomendaciones estratégicas.
            - Observaciones relevantes.
        """
    }
    return prompts[level]

# Interfaz de usuario
uploaded_file = st.file_uploader("Carga tu archivo (Excel, CSV, Word, PDF)", type=["xls", "xlsx", "csv", "docx", "pdf"])
web_url = st.text_input("O ingresa una URL")
context_text = st.text_area("Describe el contexto del análisis")

df, data_type = pd.DataFrame(), "none"

# Gestión de la carga de archivos
if uploaded_file:
    file_extension = uploaded_file.name.split('.')[-1].lower()  # Obtener la extensión del archivo cargado
    df, data_type = read_file(uploaded_file, file_extension)  # Llamar a la función con la extensión correcta
elif web_url and is_valid_url(web_url):
    df, data_type = read_file(web_url, "web")

# Mostrar los datos cargados
if data_type != "none" and not df.empty:
    st.success("Datos cargados correctamente.")
    df = df.astype(str) if data_type in ["pdf", "word", "web"] else df
    st.dataframe(df)

    # Filtrar las columnas según su tipo de datos y actualizar el DataFrame
    filtered_df = filter_columns(df)
    st.write(f"Datos después del filtrado: {filtered_df.shape[0]} filas restantes")

    # Si el dataframe filtrado está vacío, mostrar un mensaje
    if filtered_df.empty:
        st.warning("Los filtros aplicados han dejado el conjunto de datos vacío. Intenta ajustar los filtros.")

    else:
        # Mostrar los primeros 100 datos para análisis
        df_display = filtered_df.head(100)

        level = st.selectbox("Selecciona el tipo de análisis", [
            "Análisis de Ranking", 
            "Tiempos Productivos, Hold, Baño, Break",
            "Tableros de Incidencias", 
            "Satisfacción del Cliente", 
            "Costos y Rentabilidad", 
            "Libre"
        ])

        if st.button("Empezar el análisis 🚀"):
            prompt = generate_prompt(level, context_text, df_display)
            try:
                contents = [
                    {"role": "user", "parts": ["Eres un analista senior. Aplica un análisis riguroso y multifacético."]},
                    {"role": "user", "parts": [prompt]}
                ]
                response = model.generate_content(contents)
                informe = response.text
                st.write("## Informe Generado")
                for line in informe.splitlines():
                    st.write(line)

                # Guardar como Word
                doc = Document()
                doc.add_heading('Informe Generado por CAT-AI', 0)
                for line in informe.splitlines():
                    doc.add_paragraph(line)
                buffer = BytesIO()
                doc.save(buffer)
                buffer.seek(0)
                st.download_button("Descargar informe Word", buffer, "informe.docx",
                                   mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            except Exception as e:
                st.error(f"Error al generar informe: {e}")
else:
    st.info("Cargá un archivo o ingresá una URL para comenzar.")
