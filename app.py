import streamlit as st
import pandas as pd
import google.generativeai as genai
from docx import Document
from docx.shared import Inches
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from io import BytesIO
import PyPDF2
import matplotlib.pyplot as plt
import tempfile
import requests
from bs4 import BeautifulSoup
import os
import re
from docx.shared import RGBColor

# Configura la API de Gemini
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
genai.configure(api_key=GEMINI_API_KEY)

# Modelo Gemini
model = genai.GenerativeModel('gemini-1.5-pro-001')

st.title("CAT-AI")

uploaded_file = st.file_uploader("Carga tu archivo Excel, CSV, Word o PDF", type=["xls", "xlsx", "csv", "docx", "pdf"])
web_url = st.text_input("Ingresa la URL de la Página Web")

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
        df = pd.DataFrame()  # Define df como un DataFrame vacío
        #st.stop()
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
        ("Jefe de Site", "Gerente", "Director"),
    )

    if st.button("Generar"):
        with st.spinner("Generando informe..."):
            # Define prompts específicos para cada nivel de análisis
            if data_type == "excel/csv":
                if level == "Jefe de Site":
                    prompt = f"""Como analista experto en control de gestión (CAT-AI), realiza un análisis exhaustivo de los datos presentes en el archivo proporcionado. Genera un informe profesional y extremadamente detallado que describa lo siguiente, enfocándote en la eficiencia operativa, los costos, la productividad y los problemas del día a día en el sitio:

                    Análisis de eficiencia operativa: Identifica las áreas donde se pueden reducir costos y mejorar la eficiencia en el sitio.
                    Análisis de productividad: Evalúa la productividad de los recursos y propone mejoras para optimizar el rendimiento.
                    Identificación de problemas del día a día: Detecta los problemas operativos más comunes y sugiere soluciones prácticas.
                    A continuación, se muestra el contenido del archivo: {df.to_string()}.
                    El informe debe estar en español. Genera un informe original, no copies contenido existente. No utilices asteriscos ni numerales en el informe. No incluyas sugerencias de gráficos. Utiliza los siguientes encabezados para los títulos y subtítulos:
                    Título principal: [Título principal]
                    Subtítulo 1: [Subtítulo 1]
                    Subtítulo 2: [Subtítulo 2]
                    Subtítulo 3: [Subtítulo 3]"""
                elif level == "Gerente":
                    prompt = f"""Como analista experto en control de gestión (CAT-AI), realiza un análisis exhaustivo de los datos presentes en el archivo proporcionado. Genera un informe profesional y extremadamente detallado que describa lo siguiente, enfocándote en un enfoque balanceado entre lo operativo y lo comercial:

                    Análisis de rentabilidad: Evalúa la rentabilidad de las operaciones y propone estrategias para aumentarla.
                    Análisis de costos: Identifica los principales costos y sugiere medidas para reducirlos sin afectar la calidad.
                    Análisis de eficiencia operativa: Evalúa la eficiencia de los procesos y propone mejoras para optimizar el rendimiento.
                    Análisis de ventas: Analiza las ventas y propone estrategias para aumentar los ingresos.
                    A continuación, se muestra el contenido del archivo: {df.to_string()}.
                    El informe debe estar en español. Genera un informe original, no copies contenido existente. No utilices asteriscos ni numerales en el informe. No incluyas sugerencias de gráficos. Utiliza los siguientes encabezados para los títulos y subtítulos:
                    Título principal: [Título principal]
                    Subtítulo 1: [Subtítulo 1]
                    Subtítulo 2: [Subtítulo 2]
                    Subtítulo 3: [Subtítulo 3]"""
                elif level == "Director":
                    prompt = f"""Como analista experto en control de gestión (CAT-AI), realiza un análisis exhaustivo de los datos presentes en el archivo proporcionado. Genera un informe profesional y extremadamente detallado que describa lo siguiente, enfocándote en los aspectos comerciales, como las ventas, la rentabilidad, el market share y las estrategias de crecimiento:

                    Análisis de ventas: Evalúa el rendimiento de las ventas y propone estrategias para aumentar los ingresos y la cuota de mercado.
                    Análisis de rentabilidad: Identifica los productos o servicios más rentables y sugiere estrategias para maximizar su contribución.
                    Análisis de market share: Evalúa la posición de la empresa en el mercado y propone estrategias para aumentar la cuota de mercado.
                    Análisis de estrategias de crecimiento: Evalúa las estrategias de crecimiento actuales y propone nuevas estrategias para expandir el negocio.
                    A continuación, se muestra el contenido del archivo: {df.to_string()}.
                    El informe debe estar en español. Genera un informe original, no copies contenido existente. No utilices asteriscos ni numerales en el informe. No incluyas sugerencias de gráficos. Utiliza los siguientes encabezados para los títulos y subtítulos:
                    Título principal: [Título principal]
                    Subtítulo 1: [Subtítulo 1]
                    Subtítulo 2: [Subtítulo 2]
                    Subtítulo 3: [Subtítulo 3]"""
                else:
                    prompt = f"""Como analista experto en control de gestión (CAT-AI), realiza un análisis exhaustivo de los datos presentes en el archivo proporcionado. Genera un informe profesional y extremadamente detallado que describa lo siguiente:

                    Análisis de tendencias y focos de atención: Basándote EXCLUSIVAMENTE en los resultados específicos que se encuentran en el archivo proporcionado, realiza un análisis de tendencias en profundidad, comparando datos semejantes cuando sea necesario para identificar puntos de mejora. Destaca los focos de atención principales que impactan la eficiencia y rentabilidad de la organización. Desarrolla la información al máximo, profundizando en los detalles y proporcionando indicaciones claras y concisas sobre dónde se debe hacer foco para optimizar las operaciones y aumentar la rentabilidad. Proporciona al menos 5 oportunidades de mejora específicas y accionables para cada aspecto analizado, desde la perspectiva de un analista experto en control de gestión. Para cada área analizada, identifica y describe lo que más se hace, lo que más se destaca y lo que menos se hace. No te limites en la cantidad de información proporcionada, sé lo más exhaustivo y detallado posible, incluyendo todos los puntos relevantes, tanto positivos como negativos.
                    A continuación, se muestra el contenido del archivo: {df.to_string()}.
                    El informe debe estar en español. Genera un informe original, no copies contenido existente. No utilices asteriscos ni numerales en el informe. No incluyas sugerencias de gráficos. Utiliza los siguientes encabezados para los títulos y subtítulos:
                    Título principal: [Título principal]
                    Subtítulo 1: [Subtítulo 1]
                    Subtítulo 2: [Subtítulo 2]
                    Subtítulo 3: [Subtítulo 3]"""
            elif data_type == "web":
                if level == "Jefe de Site":
                    prompt = f"""Como analista experto en control de gestión (CAT-AI), realiza un análisis exhaustivo del texto proporcionado. Genera un informe profesional y extremadamente detallado que describa lo siguiente, enfocándote en la eficiencia operativa, los costos, la productividad y los problemas del día a día en el sitio:

                    Análisis de contenido: Explica el tema principal del texto y los subtemas que se tratan.
                    Identificación de ideas clave: Resume las ideas más importantes del texto.
                    Identificación de entidades: Identifica las personas, lugares, organizaciones y otros elementos relevantes que se mencionan en el texto.
                    Análisis de tendencias: Realiza un análisis de tendencias, comparando datos semejantes y haciendo indicaciones sobre dónde se debe hacer foco en la eficiencia operativa, los costos, la productividad y los problemas del día a día en el sitio.
                    A continuación, se muestra el contenido del texto: {df['text'].iloc[0]}.
                    El informe debe estar en español. Genera un informe original, no copies contenido existente. No utilices asteriscos ni numerales en el informe. No incluyas sugerencias de gráficos. Utiliza los siguientes encabezados para los títulos y subtítulos:
                    Título principal: [Título principal]
                    Subtítulo 1: [Subtítulo 1]
                    Subtítulo 2: [Subtítulo 2]
                    Subtítulo 3: [Subtítulo 3]"""
                elif level == "Gerente":
                    prompt = f"""Como analista experto en control de gestión (CAT-AI), realiza un análisis exhaustivo del texto proporcionado. Genera un informe profesional y extremadamente detallado que describa lo siguiente, enfocándote en un enfoque balanceado entre lo operativo y lo comercial:

                    Análisis de contenido: Explica el tema principal del texto y los subtemas que se tratan.
                    Identificación de ideas clave: Resume las ideas más importantes del texto.
                    Identificación de entidades: Identifica las personas, lugares, organizaciones y otros elementos relevantes que se mencionan en el texto.
                    Análisis de tendencias: Realiza un análisis de tendencias, comparando datos semejantes y haciendo indicaciones sobre dónde se debe hacer foco en un enfoque balanceado entre lo operativo y lo comercial.
                    A continuación, se muestra el contenido del texto: {df['text'].iloc[0]}.
                    El informe debe estar en español. Genera un informe original, no copies contenido existente. No utilices asteriscos ni numerales en el informe. No incluyas sugerencias de gráficos. Utiliza los siguientes encabezados para los títulos y subtítulos:
                    Título principal: [Título principal]
                    Subtítulo 1: [Subtítulo 1]
                    Subtítulo 2: [Subtítulo 2]
                    Subtítulo 3: [Subtítulo 3]"""
                elif level == "Director":
                    prompt = f"""Como analista experto en control de gestión (CAT-AI), realiza un análisis exhaustivo del texto proporcionado. Genera un informe profesional y extremadamente detallado que describa lo siguiente, enfocándote en los aspectos comerciales, como las ventas, la rentabilidad, el market share y las estrategias de crecimiento:

                    Análisis de contenido: Explica el tema principal del texto y los subtemas que se tratan.
                    Identificación de ideas clave: Resume las ideas más importantes del texto.
                    Identificación de entidades: Identifica las personas, lugares, organizaciones y otros elementos relevantes que se mencionan en el texto.
                    Análisis de tendencias: Realiza un análisis de tendencias, comparando datos semejantes y haciendo indicaciones sobre dónde se debe hacer foco en los aspectos comerciales, como las ventas, la rentabilidad, el market share y las estrategias de crecimiento.
                    A continuación, se muestra el contenido del texto: {df['text'].iloc[0]}.
                    El informe debe estar en español. Genera un informe original, no copies contenido existente. No utilices asteriscos ni numerales en el informe. No incluyas sugerencias de gráficos. Utiliza los siguientes encabezados para los títulos y subtítulos:
                    Título principal: [Título principal]
                    Subtítulo 1: [Subtítulo 1]
                    Subtítulo 2: [Subtítulo 2]
                    Subtítulo 3: [Subtítulo 3]"""
                else:
                    prompt = f"""Como analista experto en control de gestión (CAT-AI), realiza un análisis exhaustivo del texto proporcionado. Genera un informe profesional y extremadamente detallado que describa lo siguiente:

                    Análisis de contenido: Explica el tema principal del texto y los subtemas que se tratan.
                    Identificación de ideas clave: Resume las ideas más importantes del texto.
                    Identificación de entidades: Identifica las personas, lugares, organizaciones y otros elementos relevantes que se mencionan en el texto.
                    Análisis de tendencias: Realiza un análisis de tendencias, comparando datos semejantes y haciendo indicaciones sobre dónde se debe hacer foco.
                    A continuación, se muestra el contenido del texto: {df['text'].iloc[0]}.
                    El informe debe estar en español. Genera un informe original, no copies contenido existente. No utilices asteriscos ni numerales en el informe. No incluyas sugerencias de gráficos. Utiliza los siguientes encabezados para los títulos y subtítulos:
                    Título principal: [Título principal]
                    Subtítulo 1: [Subtítulo 1]
                    Subtítulo 2: [Subtítulo 2]
                    Subtítulo 3: [Subtítulo 3]"""
            else:
                if level == "Jefe de Site":
                    prompt = f"""Como analista experto en control de gestión (CAT-AI), realiza un análisis exhaustivo del texto proporcionado. Genera un informe profesional y extremadamente detallado que describa lo siguiente, enfocándote en la eficiencia operativa, los costos, la productividad y los problemas del día a día en el sitio:

                    Análisis de contenido: Explica el tema principal del texto y los subtemas que se tratan.
                    Identificación de ideas clave: Resume las ideas más importantes del texto.
                    Identificación de entidades: Identifica las personas, lugares, organizaciones y otros elementos relevantes que se mencionan en el texto.
                    Análisis de tendencias: Realiza un análisis de tendencias, comparando datos semejantes y haciendo indicaciones sobre dónde se debe hacer foco en la eficiencia operativa, los costos, la productividad y los problemas del día a día en el sitio.
                    A continuación, se muestra el contenido del texto: {df['text'].iloc[0]}.
                    El informe debe estar en español. Genera un informe original, no copies contenido existente. No utilices asteriscos ni numerales en el informe. No incluyas sugerencias de gráficos. Utiliza los siguientes encabezados para los títulos y subtítulos:
                    Título principal: [Título principal]
                    Subtítulo 1: [Subtítulo 1]
                    Subtítulo 2: [Subtítulo 2]
                    Subtítulo 3: [Subtítulo 3]"""
                elif level == "Gerente":
                    prompt = f"""Como analista experto en control de gestión (CAT-AI), realiza un análisis exhaustivo del texto proporcionado. Genera un informe profesional y extremadamente detallado que describa lo siguiente, enfocándote en un enfoque balanceado entre lo operativo y lo comercial:

                    Análisis de contenido: Explica el tema principal del texto y los subtemas que se tratan.
                    Identificación de ideas clave: Resume las ideas más importantes del texto.
                    Identificación de entidades: Identifica las personas, lugares, organizaciones y otros elementos relevantes que se mencionan en el texto.
                    Análisis de tendencias: Realiza un análisis de tendencias, comparando datos semejantes y haciendo indicaciones sobre dónde se debe hacer foco en un enfoque balanceado entre lo operativo y lo comercial.
                    A continuación, se muestra el contenido del texto: {df['text'].iloc[0]}.
                    El informe debe estar en español. Genera un informe original, no copies contenido existente. No utilices asteriscos ni numerales en el informe. No incluyas sugerencias de gráficos. Utiliza los siguientes encabezados para los títulos y subtítulos:
                    Título principal: [Título principal]
                    Subtítulo 1: [Subtítulo 1]
                    Subtítulo 2: [Subtítulo 2]
                    Subtítulo 3: [Subtítulo 3]"""
                elif level == "Director":
                    prompt = f"""Como analista experto en control de gestión (CAT-AI), realiza un análisis exhaustivo del texto proporcionado. Genera un informe profesional y extremadamente detallado que describa lo siguiente, enfocándote en los aspectos comerciales, como las ventas, la rentabilidad, el market share y las estrategias de crecimiento:

                    Análisis de contenido: Explica el tema principal del texto y los subtemas que se tratan.
                    Identificación de ideas clave: Resume las ideas más importantes del texto.
                    Identificación de entidades: Identifica las personas, lugares, organizaciones y otros elementos relevantes que se mencionan en el texto.
                    Análisis de tendencias: Realiza un análisis de tendencias, comparando datos semejantes y haciendo indicaciones sobre dónde se debe hacer foco en los aspectos comerciales, como las ventas, la rentabilidad, el market share y las estrategias de crecimiento.
                    A continuación, se muestra el contenido del texto: {df['text'].iloc[0]}.
                    El informe debe estar en español. Genera un informe original, no copies contenido existente. No utilices asteriscos ni numerales en el informe. No incluyas sugerencias de gráficos. Utiliza los siguientes encabezados para los títulos y subtítulos:
                    Título principal: [Título principal]
                    Subtítulo 1: [Subtítulo 1]
                    Subtítulo 2: [Subtítulo 2]
                    Subtítulo 3: [Subtítulo 3]"""
                else:
                    prompt = f"""Como analista experto en control de gestión (CAT-AI), realiza un análisis exhaustivo del texto proporcionado. Genera un informe profesional y extremadamente detallado que describa lo siguiente:

                    Análisis de contenido: Explica el tema principal del texto y los subtemas que se tratan.
                    Identificación de ideas clave: Resume las ideas más importantes del texto.
                    Identificación de entidades: Identifica las personas, lugares, organizaciones y otros elementos relevantes que se mencionan en el texto.
                    Análisis de tendencias: Realiza un análisis de tendencias, comparando datos semejantes y haciendo indicaciones sobre dónde se debe hacer foco.
                    A continuación, se muestra el contenido del texto: {df['text'].iloc[0]}.
                    El informe debe estar en español. Genera un informe original, no copies contenido existente. No utilices asteriscos ni numerales en el informe. No incluyas sugerencias de gráficos. Utiliza los siguientes encabezados para los títulos y subtítulos:
                    Título principal: [Título principal]
                    Subtítulo 1: [Subtítulo 1]
                    Subtítulo 2: [Subtítulo 2]
                    Subtítulo 3: [Subtítulo 3]"""
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
                    #heading.style.font.color.rgb = RGBColor(0x00, 0x00, 0xFF)
                    #heading.style.font.bold = True
                elif line.startswith("Subtítulo 1:"):
                    document.add_paragraph(line[13:])
                    #heading.style.font.color.rgb = RGBColor(0x00, 0x00, 0xFF)
                    #heading.style.font.bold = True
                elif line.startswith("Subtítulo 2:"):
                    document.add_paragraph(line[13:])
                    #heading.style.font.color.rgb = RGBColor(0x00, 0x00, 0xFF)
                    #heading.style.font.bold = True
                elif line.startswith("Subtítulo 3:"):
                    document.add_paragraph(line[13:])
                    #heading.style.font.color.rgb = RGBColor(0x00, 0x00, 0xFF)
                    #heading.style.font.bold = True
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
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
