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
    
    /* Encabezados y títulos: Barra azul unificada con selector de idioma */
    div[data-testid="stHorizontalBlock"]:first-of-type {
        background: linear-gradient(135deg, #012743 0%, #1d3d5a 100%);
        color: white;
        padding: 16px 28px;
        border-radius: 12px;
        margin-bottom: 20px;
        box-shadow: 0 4px 20px rgba(1, 39, 67, 0.08);
        border: 1px solid rgba(255, 255, 255, 0.1);
        align-items: center;
    }
    
    .top-title-area h1 {
        margin: 0;
        font-family: 'Space Grotesk', sans-serif;
        font-size: 30px;
        font-weight: 700;
        letter-spacing: -0.5px;
        color: #ffffff !important;
    }
    
    .top-title-area p {
        margin: 4px 0 0 0;
        font-size: 13.5px;
        color: #e2e8f0;
        opacity: 0.95;
    }
    
    /* Selector de idioma estilizado dentro de la barra azul */
    .top-lang-wrapper div[data-baseweb="select"] {
        background-color: rgba(255, 255, 255, 0.15) !important;
        border: 1px solid rgba(255, 255, 255, 0.35) !important;
        border-radius: 8px !important;
        backdrop-filter: blur(6px);
        transition: all 0.2s ease;
    }
    
    .top-lang-wrapper div[data-baseweb="select"]:hover {
        background-color: rgba(255, 255, 255, 0.25) !important;
        border-color: rgba(255, 255, 255, 0.7) !important;
    }
    
    .top-lang-wrapper div[data-baseweb="select"] * {
        color: #ffffff !important;
        font-weight: 600 !important;
        font-size: 13px !important;
    }
    
    .top-lang-wrapper div[data-baseweb="select"] svg {
        fill: #ffffff !important;
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

# Inicializar estado de idioma global en Streamlit
if 'app_lang' not in st.session_state:
    st.session_state.app_lang = 'es'

# Diccionario de textos bilingües para el entorno de Streamlit
texts = {
    'main_title': '🛡️ SIM-Q RELIARISK',
    'main_subtitle': {
        'es': 'Sistema de Gestión de Integridad Estructural y Matriz de Riesgo 3x3 | Complejo Ek-Balam (API RP 2SIM)',
        'en': 'Structural Integrity Management System & 3x3 Risk Matrix | Ek-Balam Complex (API RP 2SIM)'
    },
    'tab_simulator': {
        'es': '📊 Simulador Interactivo',
        'en': '📊 Interactive Simulator'
    },
    'tab_report': {
        'es': '📝 Generador de Reporte Técnico',
        'en': '📝 Technical Report Generator'
    },
    'report_card_title': {
        'es': '📝 Generación de Reporte Word (.docx)',
        'en': '📝 Word Technical Report Generation (.docx)'
    },
    'report_card_desc': {
        'es': 'Esta utilidad permite procesar la base de datos de respuestas de la flota de plataformas y generar el documento técnico en Word (.docx) completamente formateado bajo los lineamientos y cálculos de la metodología <strong>API RP 2SIM</strong> ajustada a la Matriz de Criticidad de 3x3.',
        'en': 'This utility processes the platform fleet response database and generates the fully formatted technical Word document (.docx) following the guidelines and calculations of the <strong>API RP 2SIM</strong> methodology adjusted to the 3x3 Criticality Matrix.'
    },
    'template_title': {
        'es': '📥 Plantilla de Entrada',
        'en': '📥 Input Template'
    },
    'template_desc': {
        'es': 'Descarga la base de datos Excel modelo para registrar o modificar las respuestas de tu flota antes de procesarlas:',
        'en': 'Download the Excel database template to register or modify your fleet responses before processing them:'
    },
    'template_btn': {
        'es': '📥 Descargar Base de Datos Excel Plantilla',
        'en': '📥 Download Excel Database Template'
    },
    'template_missing': {
        'es': 'No se encontró el archivo de base de datos Excel por defecto en el servidor.',
        'en': 'Default Excel database file was not found on the server.'
    },
    'instructions_title': {
        'es': '<strong>Instrucciones:</strong>',
        'en': '<strong>Instructions:</strong>'
    },
    'instructions_body': {
        'es': '1. Abre el archivo Excel descargado.<br>2. Cada pestaña corresponde a una plataforma marina.<br>3. Escribe las respuestas (<code>a</code>, <code>b</code>, o <code>c</code>) en la columna de respuestas de cada pregunta.<br>4. Guarda el archivo y súbelo en la sección de la derecha.',
        'en': '1. Open the downloaded Excel file.<br>2. Each sheet corresponds to a marine platform.<br>3. Enter the answers (<code>a</code>, <code>b</code>, or <code>c</code>) in the response column for each question.<br>4. Save the file and upload it in the section on the right.'
    },
    'upload_title': {
        'es': '📤 Cargar y Procesar',
        'en': '📤 Upload & Process'
    },
    'upload_desc': {
        'es': 'Sube tu base de datos de respuestas en Excel para generar el reporte de integridad correspondiente:',
        'en': 'Upload your Excel response database to generate the corresponding integrity report:'
    },
    'upload_label': {
        'es': 'Cargar Base de Datos Excel (.xlsx)',
        'en': 'Upload Excel Database (.xlsx)'
    },
    'upload_help': {
        'es': 'Sube el archivo Excel con las respuestas de la flota',
        'en': 'Upload the Excel file with the fleet responses'
    },
    'btn_generate': {
        'es': '⚡ Procesar y Generar Reporte Word',
        'en': '⚡ Process & Generate Word Report'
    },
    'spinner': {
        'es': '⏳ Procesando respuestas y estructurando reporte en Word...',
        'en': '⏳ Processing responses and structuring Word report...'
    },
    'success': {
        'es': '✅ ¡Reporte técnico generado con éxito!',
        'en': '✅ Technical report generated successfully!'
    },
    'download_report_btn': {
        'es': '💾 Descargar Reporte Técnico (.docx)',
        'en': '💾 Download Technical Report (.docx)'
    },
    'error_prefix': {
        'es': '❌ Ocurrió un error al generar el reporte:',
        'en': '❌ An error occurred while generating the report:'
    },
    'error_no_excel': {
        'es': 'No se encontró base de datos Excel para procesar.',
        'en': 'No Excel database found to process.'
    },
    'error_compile': {
        'es': "Error al compilar los archivos de la aplicación interactiva. Asegúrate de que la carpeta 'app/' con index.html, style.css y app.js exista.",
        'en': "Error compiling the interactive application files. Make sure the 'app/' folder with index.html, style.css and app.js exists."
    }
}

def t(key):
    val = texts[key]
    if isinstance(val, dict):
        return val[st.session_state.app_lang]
    return val

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
def compile_interactive_app_v7(lang):
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
        
    # Inyectar inicialización de idioma desde Streamlit
    lang_init_code = f"""
    <script>
        try {{
            localStorage.setItem('simq_lang', '{lang}');
            currentLang = '{lang}';
        }} catch(e) {{}}
    </script>
    """
    
    # Reemplazar recursos locales externos
    html = html.replace('<link rel="stylesheet" href="style.css">', f'<style>\n{css}\n</style>')
    html = html.replace('<script src="app.js?v=14"></script>', f'{lang_init_code}\n<script>\n{js}\n</script>')
    
    # Incorporar logos como base64 Data URIs
    logo_mi_b64 = get_base64_image(logo_mi_path)
    logo_app_b64 = get_base64_image(logo_app_path)
    
    if logo_mi_b64:
        html = html.replace('src="Mi_Logo.jpeg"', f'src="{logo_mi_b64}"')
    if logo_app_b64:
        html = html.replace('src="Logo_App.jpg"', f'src="{logo_app_b64}"')
        
    return html

# --- ENCABEZADO PRINCIPAL (BARRA AZUL CON SELECTOR DE IDIOMA INTEGRADO) ---
header_col1, header_col2 = st.columns([3.8, 1.2], vertical_alignment="center")

with header_col1:
    st.markdown(f"""
    <div class="top-title-area">
        <h1>🛡️ SIM-Q RELIARISK</h1>
        <p>{t('main_subtitle')}</p>
    </div>
    """, unsafe_allow_html=True)

with header_col2:
    st.markdown('<div class="top-lang-wrapper">', unsafe_allow_html=True)
    lang_choice = st.selectbox(
        "Idioma / Language",
        options=["🇪🇸 Español", "🇺🇸 English"],
        index=0 if st.session_state.app_lang == 'es' else 1,
        label_visibility="collapsed",
        key="top_bar_language_selector"
    )
    new_lang = "en" if "English" in lang_choice else "es"
    if new_lang != st.session_state.app_lang:
        st.session_state.app_lang = new_lang
        st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

# Renderizar directamente el Simulador Interactivo (única pestaña activa)

# --- SIMULADOR INTERACTIVO (vista principal) ---
html_compiled = compile_interactive_app_v7(st.session_state.app_lang)

if html_compiled is None:
    st.error(t('error_compile'))
else:
    # Renderizar la aplicación en un iframe de ancho completo
    components.html(html_compiled, height=880, scrolling=True)

