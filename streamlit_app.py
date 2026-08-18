import streamlit as st
import streamlit.components.v1 as components
import os
import base64
import tempfile
import json
import generar_reporte

# Configuración de página de Streamlit
st.set_page_config(
    page_title="SIM-Q RELIARISK Costa Afuera",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS premium para la interfaz de Streamlit (STORM & Reliarisk Palette)
st.markdown("""
<style>
    /* Estilos globales */
    .stApp {
        background-color: #f7f9fb;
    }
    
    /* Encabezados y títulos */
    .main-title-container {
        background: linear-gradient(135deg, #012743 0%, #1d3d5a 100%);
        color: white;
        padding: 24px 32px;
        border-radius: 12px;
        margin-bottom: 24px;
        box-shadow: 0 4px 20px rgba(1, 39, 67, 0.08);
        border: 1px solid rgba(255, 255, 255, 0.1);
    }
    .main-title-container h1 {
        margin: 0;
        font-family: 'Space Grotesk', sans-serif;
        font-size: 32px;
        font-weight: 700;
        letter-spacing: -0.5px;
        color: #ffffff !important;
    }
    .main-title-container p {
        margin: 6px 0 0 0;
        font-size: 14px;
        color: #e2e8f0;
        opacity: 0.9;
    }

    /* Pestañas personalizadas */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: transparent;
        padding: 0 4px;
    }
    .stTabs [data-baseweb="tab"] {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 8px 8px 0 0;
        padding: 10px 20px;
        font-weight: 600;
        color: #64748b;
        transition: all 0.2s;
    }
    .stTabs [data-baseweb="tab"]:hover {
        color: #012743;
        border-color: #cbd5e1;
    }
    .stTabs [aria-selected="true"] {
        background-color: #012743 !important;
        color: white !important;
        border-color: #012743 !important;
    }

    /* Tarjetas y contenedores */
    .panel-card {
        background-color: #ffffff;
        padding: 24px;
        border-radius: 12px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 2px 8px rgba(0,0,0,0.02);
        margin-bottom: 20px;
    }
    
    .panel-card h3 {
        color: #012743;
        font-family: 'Space Grotesk', sans-serif;
        font-weight: 700;
        margin-top: 0;
        margin-bottom: 12px;
        border-bottom: 2px solid #f1f5f9;
        padding-bottom: 8px;
    }

    /* Botones personalizados */
    .stButton>button {
        background-color: #012743;
        color: white;
        font-weight: 600;
        border: none;
        border-radius: 6px;
        padding: 10px 24px;
        transition: all 0.2s;
    }
    .stButton>button:hover {
        background-color: #1d3d5a;
        color: white;
        box-shadow: 0 4px 12px rgba(1, 39, 67, 0.2);
    }
    
    /* Ocultar barra de menú de streamlit por defecto */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# Helper para codificar imágenes locales en Base64 Data URI
def get_base64_image(image_path):
    if not os.path.exists(image_path):
        return ""
    with open(image_path, "rb") as img_file:
        encoded = base64.b64encode(img_file.read()).decode('utf-8')
    mime_type = "image/jpeg"
    if image_path.lower().endswith(".png"):
        mime_type = "image/png"
    return f"data:{mime_type};base64,{encoded}"

# Compilar aplicación HTML/JS/CSS en un archivo único autocontenido para el iframe
@st.cache_data
def compile_interactive_app_v4():
    html_path = "app/index.html"
    css_path = "app/style.css"
    js_path = "app/app.js"
    logo_mi_path = "app/Mi_Logo.jpeg"
    logo_app_path = "app/Logo_App.jpg"
    
    # Validar archivos mínimos
    if not os.path.exists(html_path) or not os.path.exists(css_path) or not os.path.exists(js_path):
        return None
        
    # Leer contenido original
    with open(html_path, "r", encoding="utf-8") as f:
        html = f.read()
        
    with open(css_path, "r", encoding="utf-8") as f:
        css = f.read()
        
    with open(js_path, "r", encoding="utf-8") as f:
        js = f.read()
        
    # Reemplazar recursos locales externos
    html = html.replace('<link rel="stylesheet" href="style.css">', f'<style>\n{css}\n</style>')
    html = html.replace('<script src="app.js?v=14"></script>', f'<script>\n{js}\n</script>')
    
    # Incorporar logos como base64 Data URIs
    logo_mi_b64 = get_base64_image(logo_mi_path)
    logo_app_b64 = get_base64_image(logo_app_path)
    
    if logo_mi_b64:
        html = html.replace('src="Mi_Logo.jpeg"', f'src="{logo_mi_b64}"')
    if logo_app_b64:
        html = html.replace('src="Logo_App.jpg"', f'src="{logo_app_b64}"')
        
    return html

# --- ENCABEZADO PRINCIPAL ---
st.markdown("""
<div class="main-title-container">
    <h1>🛡️ SIM-Q RELIARISK</h1>
    <p>Sistema de Gestión de Integridad Estructural y Matriz de Riesgo 3x3 | Complejo Ek-Balam (API RP 2SIM)</p>
</div>
""", unsafe_allow_html=True)

# Crear pestañas (2 pestañas: simulador y generador de reporte)
tab_simulator, tab_report = st.tabs([
    "📊 Simulador Interactivo", 
    "📝 Generador de Reporte Técnico"
])

# --- PESTAÑA 1: SIMULADOR INTERACTIVO ---
with tab_simulator:
    html_compiled = compile_interactive_app_v4()
    
    if html_compiled is None:
        st.error("Error al compilar los archivos de la aplicación interactiva. Asegúrate de que la carpeta 'app/' con index.html, style.css y app.js exista.")
    else:
        # Renderizar la aplicación en un iframe de ancho completo
        components.html(html_compiled, height=880, scrolling=True)

# --- PESTAÑA 2: GENERADOR DE REPORTE TÉCNICO ---
with tab_report:
    st.markdown("""
    <div class="panel-card">
        <h3>📝 Generación de Reporte Word (.docx)</h3>
        <p>Esta utilidad permite procesar la base de datos de respuestas de la flota de plataformas y generar el documento técnico en Word (.docx) completamente formateado bajo los lineamientos y cálculos de la metodología <strong>API RP 2SIM</strong> ajustada a la Matriz de Criticidad de 3x3.</p>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2 = st.columns([1, 1], gap="medium")
    
    with col1:
        st.markdown("""
        <div class="panel-card" style="height: 100%;">
            <h3>📥 Plantilla de Entrada</h3>
            <p>Descarga la base de datos Excel modelo para registrar o modificar las respuestas de tu flota antes de procesarlas:</p>
        </div>
        """, unsafe_allow_html=True)
        
        excel_template_path = "Base de datos Plataformas EK Balam.xlsx"
        if os.path.exists(excel_template_path):
            with open(excel_template_path, "rb") as f:
                bytes_data = f.read()
            st.download_button(
                label="📥 Descargar Base de Datos Excel Plantilla",
                data=bytes_data,
                file_name="Base de datos Plataformas EK Balam.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No se encontró el archivo de base de datos Excel por defecto en el servidor.")
            
        st.markdown("""
        <br>
        <div style="font-size: 13px; color: #64748b; line-height: 1.4;">
            <strong>Instrucciones:</strong><br>
            1. Abre el archivo Excel descargado.<br>
            2. Cada pestaña corresponde a una plataforma marina.<br>
            3. Escribe las respuestas (<code>a</code>, <code>b</code>, o <code>c</code>) en la columna de respuestas de cada pregunta.<br>
            4. Guarda el archivo y súbelo en la sección de la derecha.
        </div>
        """, unsafe_allow_html=True)

    with col2:
        st.markdown("""
        <div class="panel-card" style="height: 100%;">
            <h3>📤 Cargar y Procesar</h3>
            <p>Sube tu base de datos de respuestas en Excel para generar el reporte de integridad correspondiente:</p>
        </div>
        """, unsafe_allow_html=True)
        
        uploaded_file = st.file_uploader(
            "Cargar Base de Datos Excel (.xlsx)", 
            type=["xlsx"],
            help="Sube el archivo Excel con las respuestas de la flota"
        )
        
        btn_generate = st.button("⚡ Procesar y Generar Reporte Word")
        
        if btn_generate:
            with st.spinner("⏳ Procesando respuestas y estructurando reporte en Word..."):
                try:
                    # Crear archivos temporales seguros para evitar colisiones entre usuarios
                    temp_excel_fd, temp_excel_path = tempfile.mkstemp(suffix=".xlsx")
                    os.close(temp_excel_fd)
                    
                    temp_docx_fd, temp_docx_path = tempfile.mkstemp(suffix=".docx")
                    os.close(temp_docx_fd)
                    
                    # Guardar archivo subido en ruta temporal o usar la plantilla por defecto si no se cargó nada
                    if uploaded_file is not None:
                        with open(temp_excel_path, "wb") as f:
                            f.write(uploaded_file.getbuffer())
                    else:
                        if os.path.exists(excel_template_path):
                            with open(temp_excel_path, "wb") as f:
                                with open(excel_template_path, "rb") as template_f:
                                    f.write(template_f.read())
                        else:
                            raise FileNotFoundError("No se encontró base de datos Excel para procesar.")
                    
                    # Llamar al generador de reportes pasándole las rutas dinámicas
                    generar_reporte.main(
                        excel_path=temp_excel_path,
                        json_path="preguntas_y_opciones.json",
                        output_filename=temp_docx_path
                    )
                    
                    # Leer el reporte generado
                    with open(temp_docx_path, "rb") as f:
                        docx_bytes = f.read()
                        
                    # Limpiar archivos temporales
                    os.remove(temp_excel_path)
                    os.remove(temp_docx_path)
                    
                    st.success("✅ ¡Reporte técnico generado con éxito!")
                    st.download_button(
                        label="💾 Descargar Reporte Técnico (.docx)",
                        data=docx_bytes,
                        file_name="Reporte_Integridad_Estructural_Ek_Balam.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                except Exception as e:
                    st.error(f"❌ Ocurrió un error al generar el reporte: {str(e)}")

