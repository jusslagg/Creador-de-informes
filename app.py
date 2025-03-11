import streamlit as st
import pandas as pd
import google.generativeai as genai
from docx import Document
from docx.shared import Inches
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from io import BytesIO
import matplotlib.pyplot as plt
import tempfile
import requests
from bs4 import BeautifulSoup
import os
import re
from docx.shared import RGBColor
import PyPDF2

# Configura la API de Gemini
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
genai.configure(api_key=GEMINI_API_KEY)

# Modelo Gemini
model = genai.GenerativeModel('gemini-1.5-pro-001')

st.title("CAT-AI")

uploaded_file = st.file_uploader("Carga tu archivo Excel, CSV, Word o PDF", type=["xls", "xlsx", "csv", "docx", "pdf"])
web_url = st.text_input("Ingresa la URL de la Página Web")
context_text = st.text_area("Ingresa el contexto para el análisis")
uploaded_images = st.file_uploader("Carga tus imágenes", type=["png", "jpg", "jpeg"], accept_multiple_files=True)

def is_valid_url(url):
    # Regex para validar una URL
    regex = re.compile(
        r'^(?:http|ftp)s?://' # http:// or https://
        r'(?:(?:[A-Z0-9](?:[A-Z0-9-]{0,61}[A-Z0-9])?\.)+(?:[A-Z]{2,6}\.?|[A-Z0-9-]{2,}\.?)|' #domain...
        r'localhost|' #localhost...
        r'\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})' # ...or ip
        r'(?::\d+)?' # optional port
        r'(?:/?|[/?]\S+)$', re.IGNORECASE)
    return re.match(regex, url) is not None

if uploaded_file is not None:
    file_type = uploaded_file.type
    try:
        if file_type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
            try:
                df = pd.read_excel(uploaded_file, engine='openpyxl')
                data_type = "excel/csv"
            except Exception as e:
                st.write(f"No se pudo leer el archivo Excel. Error: {e}. Por favor, verifica que el archivo Excel esté correctamente formateado.")
                df = pd.DataFrame()
                data_type = "excel/csv"
        elif file_type == "text/csv":
            try:
                df = pd.read_csv(uploaded_file, encoding='latin1', on_bad_lines='skip')
                data_type = "excel/csv"
            except Exception as e:
                st.write(f"No se pudo leer el archivo CSV. Error: {e}. Por favor, verifica que el archivo CSV esté correctamente formateado.")
                df = pd.DataFrame()
                data_type = "excel/csv"
        elif file_type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
            document = Document(uploaded_file)
            text = ""
            for paragraph in document.paragraphs:
                text += paragraph.text + "\n"
            df = pd.DataFrame([text], columns=['text'])
            data_type = "word"
        elif file_type == "application/pdf":
            try:
                pdf_reader = PyPDF2.PdfReader(uploaded_file)
                text = ""
                for page in pdf_reader.pages:
                    text += page.extract_text() + "\n"
                df = pd.DataFrame([text], columns=['text'])
                data_type = "pdf"
            except Exception as e:
                st.write(f"No se pudo leer el archivo PDF. Error: {e}")
                df = pd.DataFrame()
                data_type = "none"
        else:
            st.write(f"Tipo de archivo no soportado: {file_type}")
            df = pd.DataFrame()
            data_type = "none"
    except Exception as e:
        st.write(f"No se pudo leer el archivo. Error: {e}")
        df = pd.DataFrame()
        data_type = "none"
elif web_url and is_valid_url(web_url):
    try:
        response = requests.get(web_url)
        response.raise_for_status()
        soup = BeautifulSoup(response.content, 'html.parser')
        text = soup.get_text(separator='\n')
        df = pd.DataFrame([text], columns=['text'])
        data_type = "web"
    except Exception as e:
        st.write(f"No se pudo leer la URL. Error: {e}")
        df = pd.DataFrame()
        data_type = "none"
else:
    df = pd.DataFrame()
    data_type = "none"

if data_type != "none":
    st.write("Archivo cargado exitosamente!")

    # Convierte las columnas con tipo 'object' a tipo 'string'
    for col in df.columns:
        if df[col].dtype == 'object':
            df[col] = df[col].astype(str)

    st.dataframe(df)

    # Permite al usuario seleccionar las columnas a analizar
    selected_columns = st.multiselect("Selecciona las columnas a analizar", df.columns.tolist())

    # Limita el tamaño del dataset y selecciona las columnas
    if selected_columns:
        df = df[selected_columns].head(100)
    else:
        df = df.head(100)

    # Permite al usuario seleccionar el nivel del análisis
    level = st.selectbox(
        "Selecciona el nivel del análisis",
        ["Análisis de Ranking", "Tiempos Productivos, Hold, Baño, Break", "Tableros de Incidencias", "Libre"]
    )

    if st.button("Empezar el análisis 🚀"):
        if df.empty:
            st.write("No hay datos para analizar. Por favor, carga un archivo o ingresa una URL.")
        else:
            if level == "Análisis de Ranking":
                prompt = f"""Como analista experto en control de gestión (CAT-AI), realiza un análisis exhaustivo y **altamente enfocado** del ranking de agentes en el call center, basado en métricas clave como tiempo de respuesta, resolución de problemas, y satisfacción del cliente. **Es fundamental que el análisis se centre en el siguiente contexto proporcionado por el usuario, y que se le dé la máxima prioridad en la generación del informe: {context_text}. Por favor, asegúrate de que el informe refleje este contexto de manera precisa y detallada.** El informe debe detallar los siguientes aspectos:

                - **Rendimiento de los agentes:** Evalúa el rendimiento de cada agente en función de las métricas establecidas. ¿Qué agentes se destacan positivamente? ¿Cuáles presentan áreas de mejora?
                - **Comparativa entre agentes:** Realiza un análisis comparativo entre los agentes, destacando las mejores prácticas que los agentes con mejores resultados siguen y las posibles áreas de mejora para los de bajo rendimiento.
                - **Recomendaciones de mejora:** Propón estrategias concretas para mejorar el rendimiento de los agentes con menor ranking, considerando entrenamientos adicionales, ajustes en las herramientas de trabajo, o cambios en los procesos operativos.
                - **Top 10:** Identifica y describe el top 10 de los agentes con mejor rendimiento, destacando sus fortalezas y mejores prácticas.
                - **Peores 10:** Identifica y describe los 10 agentes con peor rendimiento, señalando las áreas de mejora y posibles causas de su bajo desempeño.
                - **Cuartilización:** Divide a los agentes en cuartiles según su rendimiento y analiza las características de cada cuartil. ¿Qué diferencias existen entre los agentes de los diferentes cuartiles? ¿Qué estrategias se pueden implementar para ayudar a los agentes a subir de cuartil?

                A continuación, se muestra el contenido del archivo de desempeño de los agentes: {df.to_string()}.
                El informe debe estar en español. Genera un informe original, no copies contenido existente. No utilices asteriscos ni numerales en el informe. No incluyas sugerencias de gráficos. Utiliza los siguientes encabezados para los títulos y subtítulos:
                Título principal: [Título principal]
                Subtítulo 1: Análisis de rendimiento de los agentes
                Subtítulo 2: Comparativa de agentes
                Subtítulo 3: Estrategias de mejora
                Subtítulo 4: Top 10 más destacados
                Subtítulo 5: Top 10 menos destacados
                Subtítulo 6: Análisis de cuartilización"""

            elif level == "Tiempos Productivos, Hold, Baño, Break":
                prompt = f"""Como analista experto en control de gestión (CAT-AI), realiza un análisis detallado sobre los tiempos productivos, los tiempos en espera (hold), los descansos (baño, break) y otros periodos de inactividad de los agentes en el call center. **Es fundamental que el análisis se centre en el siguiente contexto proporcionado por el usuario, y que se le dé la máxima prioridad en la generación del informe: {context_text}. Por favor, asegúrate de que el informe refleje este contexto de manera precisa y detallada.** El informe debe abordar lo siguiente:

                - **Tiempo productivo:** Analiza el tiempo que los agentes están efectivamente atendiendo llamadas o gestionando tareas operativas. ¿Está siendo aprovechado al máximo? ¿Cuánto tiempo de la jornada laboral se dedica a actividades productivas?
                - **Tiempos de espera (Hold):** Examina los periodos en los que los agentes mantienen a los clientes en espera. ¿Están dentro de los límites esperados? ¿Cómo se pueden optimizar estos tiempos?
                - **Pausas (Baño y Break):** Evalúa la frecuencia y duración de las pausas que toman los agentes. ¿Se están tomando en los momentos adecuados? ¿Hay alguna mejora en la gestión de los descansos para asegurar que los agentes mantengan su productividad sin afectar la calidad del servicio?
                - **Recomendaciones:** Proporciona sugerencias para optimizar los tiempos de inactividad, ajustando las pausas o el tiempo en espera sin sacrificar la calidad del servicio al cliente.

                A continuación, se muestra el contenido del archivo con los registros de los tiempos: {df.to_string()}.
                El informe debe estar en español. Genera un informe original, no copies contenido existente. No utilices asteriscos ni numerales en el informe. No incluyas sugerencias de gráficos. Utiliza los siguientes encabezados para los títulos y subtítulos:
                Título principal: [Título principal]
                Subtítulo 1: Análisis de tiempos productivos
                Subtítulo 2: Evaluación de tiempos en espera (Hold)
                Subtítulo 3: Optimización de pausas y descansos"""

            elif level == "Tableros de Incidencias":
                prompt = f"""Como analista experto en control de gestión (CAT-AI), realiza un análisis exhaustivo de las incidencias reportadas en el call center, utilizando los tableros de incidencias. **Es fundamental que el análisis se centre en el siguiente contexto proporcionado por el usuario, y que se le dé la máxima prioridad en la generación del informe: {context_text}. Por favor, asegúrate de que el informe refleje este contexto de manera precisa y detallada.** El informe debe cubrir los siguientes aspectos:

                - **Identificación de incidencias recurrentes:** Analiza las incidencias más comunes reportadas, tanto a nivel de clientes como de agentes. ¿Qué problemas están afectando más a la operación? ¿Cuáles son las áreas más problemáticas?
                - **Frecuencia de incidencias:** Evalúa la frecuencia de las incidencias y su impacto en la eficiencia operativa. ¿Las incidencias están afectando los tiempos de atención al cliente? ¿Cómo se distribuyen las incidencias a lo largo del día o semana?
                - **Resolución de incidencias:** Revisa cómo se están gestionando las incidencias. ¿Se están resolviendo a tiempo? ¿Existen procedimientos estandarizados para resolver problemas recurrentes?
                - **Recomendaciones para la gestión de incidencias:** Propón acciones o procedimientos que puedan reducir la cantidad de incidencias y mejorar la velocidad de resolución.

                A continuación, se muestra el contenido de los tableros de incidencias: {df.to_string()}.
                El informe debe estar en español. Genera un informe original, no copies contenido existente. No utilices asteriscos ni numerales en el informe. No incluyas sugerencias de gráficos. Utiliza los siguientes encabezados para los títulos y subtítulos:
                Título principal: [Título principal]
                Subtítulo 1: Identificación y análisis de incidencias recurrentes
                Subtítulo 2: Frecuencia e impacto de las incidencias
                Subtítulo 3: Estrategias para mejorar la gestión de incidencias"""

            elif level == "Libre":
                prompt = f"""Como analista experto en control de gestión (CAT-AI), realiza un análisis exhaustivo de los datos proporcionados en el archivo, sin limitaciones de un área específica. **Es fundamental que el análisis se centre en el siguiente contexto proporcionado por el usuario, y que se le dé la máxima prioridad en la generación del informe: {context_text}. Por favor, asegúrate de que el informe refleje este contexto de manera precisa y detallada.** El informe debe abordar lo siguiente:

                - **Análisis detallado de la información:** Realiza un análisis completo de todos los datos disponibles, identificando patrones clave y áreas de mejora.
                - **Recomendaciones de optimización:** Proporciona recomendaciones para mejorar las operaciones generales del call center basándote en los datos presentados.
                - **Oportunidades de mejora:** Identifica y describe oportunidades de mejora en cualquier área que consideres relevante para optimizar la eficiencia, calidad del servicio, y productividad en el call center.

                A continuación, se muestra el contenido del archivo con los datos proporcionados: {df.to_string()}.
                El informe debe estar en español. Genera un informe original, no copies contenido existente. No utilices asteriscos ni numerales en el informe. No incluyas sugerencias de gráficos. Utiliza los siguientes encabezados para los títulos y subtítulos:
                Título principal: [Título principal]
                Subtítulo 1: Análisis general
                Subtítulo 2: Identificación de oportunidades
                Subtítulo 3: Estrategias de optimización"""

            response = model.generate_content(prompt)
            informe = response.text

            st.write("Informe generado por Gemini:")

            # Divide el informe en líneas
            lines = informe.splitlines()

            # Itera sobre las líneas y da formato a los títulos y subtítulos
            for line in lines:
                if line.startswith("Título principal:"):
                    st.markdown(f"<h1 style='color: blue; font-weight: bold;'>{line[17:].replace('*', '').replace('#', '')}</h1>", unsafe_allow_html=True)
                elif line.startswith("Subtítulo 1:"):
                    st.markdown(f"<h2 style='color: blue; font-weight: bold;'>{line[13:].replace('*', '').replace('#', '')}</h2>", unsafe_allow_html=True)
                elif line.startswith("Subtítulo 2:"):
                    st.markdown(f"<h3 style='color: blue; font-weight: bold;'>{line[13:].replace('*', '').replace('#', '')}</h3>", unsafe_allow_html=True)
                elif line.startswith("Subtítulo 3:"):
                    st.markdown(f"<h4 style='color: blue; font-weight: bold;'>{line[13:].replace('*', '').replace('#', '')}</h4>", unsafe_allow_html=True)
                else:
                    st.write(line.replace('*', '').replace('#', ''))

                    # Genera gráficos
                    #if len(df.select_dtypes(include=['number', 'datetime']).columns) > 0:
                    #    fig, ax = plt.subplots()
                    #    df.hist(ax=ax)
                    #    plt.tight_layout()
                
                    #    # Guarda el gráfico en un archivo temporal
                    #    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmpfile:
                    #        fig.savefig(tmpfile.name, format="png")
                    #        temp_filename = tmpfile.name
                    #elif len(df.columns) > 0:
                    #    # Si no hay columnas numéricas o de fecha y hora, genera un gráfico de barras con la frecuencia de los nombres
                    #    fig, ax = plt.subplots()
                    #    nombres = df.iloc[:, 0].value_counts()
                    #    nombres.plot(kind='bar', ax=ax)
                    #    plt.tight_layout()
                
                    #    # Guarda el gráfico en un archivo temporal
                    #    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmpfile:
                    #        fig.savefig(tmpfile.name, format="png")
                    #        temp_filename = tmpfile.name
                    #else:
            temp_filename = None

        # Genera el informe en Word
        document = Document()
        document.add_heading('Informe Generado por CAT-AI', 0)
        #document.add_paragraph(informe)

        # Divide el informe en líneas
        lines = informe.splitlines()

        # Itera sobre las líneas y da formato a los títulos y subtítulos
        for line in lines:
            if line.startswith("Título principal:"):
                document.add_paragraph(line[17:])
            elif line.startswith("Subtítulo 1:"):
                document.add_paragraph(line[13:])
            elif line.startswith("Subtítulo 2:"):
                document.add_paragraph(line[13:])
            elif line.startswith("Subtítulo 3:"):
                document.add_paragraph(line[13:])
            else:
                document.add_paragraph(line)

                
                #if temp_filename:
                #    document.add_picture(temp_filename, width=Inches(6))
            
                # Guarda el documento en memoria
        docx_stream = BytesIO()
        document.save(docx_stream)
        docx_stream.seek(0)

        st.download_button(
            label="Descargar informe en Word",
            data=docx_stream,
            file_name="informe.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
