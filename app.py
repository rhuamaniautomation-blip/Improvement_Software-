#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
================================================================================
SISTEMA DE GESTION DOCUMENTAL - MoC | Mejora A3 | Simple Kaizen
Version 7.5.0 - Diagnóstico API y Contexto Estricto
================================================================================
Diseñado por: CAVA - Especialistas en Robotica y Automatizacion
Desarrollador: Roger Huamani
Version: 7.5.0
Fecha: Agosto 2026
================================================================================
"""
import streamlit as st
import os
import json
import re
import uuid
import shutil
import subprocess
import tempfile
import base64
import hashlib
from datetime import datetime
from pathlib import Path
from copy import deepcopy
from io import BytesIO
# Librerías para documentos
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from docx import Document
from docx.shared import Inches as DocxInches
from docx.shared import Pt as DocxPt
from docx.shared import RGBColor as DocxRGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
# Corrector ortográfico
SPELLCHECKER_AVAILABLE = False
try:
    from spellchecker import SpellChecker
    SPELLCHECKER_AVAILABLE = True
except ImportError:
    pass
# Intentar importar reportlab para PDF fallback
try:
    from reportlab.lib.pagesizes import A4, letter
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch, cm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage, PageBreak
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

# =============================================================================
# CONFIGURACION INICIAL DE PAGINA
# =============================================================================
st.set_page_config(
    page_title="Gestión Documental - MoC | A3 | Kaizen",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =============================================================================
# RUTAS Y DIRECTORIOS DE PERSISTENCIA
# =============================================================================
BASE_DIR = Path(__file__).parent if "__file__" in dir() else Path(".")
DATA_DIR = BASE_DIR / "data"
TEMPLATES_DIR = BASE_DIR / "templates"
HISTORY_FILE = DATA_DIR / "history.json"
CONFIG_FILE = DATA_DIR / "config.json"
DATA_DIR.mkdir(exist_ok=True)
TEMPLATES_DIR.mkdir(exist_ok=True)

# =============================================================================
# CSS PERSONALIZADO EMPRESARIAL
# =============================================================================
CUSTOM_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
html, body, [class*="css"] { font-family: 'Inter', 'Segoe UI', sans-serif !important; }
.main-header {
    background: linear-gradient(135deg, #1a5f7a 0%, #2e8bc0 100%);
    padding: 2rem; border-radius: 12px; color: white;
    margin-bottom: 2rem; box-shadow: 0 4px 20px rgba(26, 95, 122, 0.3);
}
.main-header h1 { color: white !important; font-weight: 700 !important; margin-bottom: 0.5rem !important; }
.main-header p { color: rgba(255,255,255,0.9) !important; font-size: 1.1rem !important; }
.doc-card {
    background: white; border-radius: 16px; padding: 2rem;
    border: 2px solid #e2e8f0; transition: all 0.3s ease;
    cursor: pointer; height: 100%;
}
.doc-card:hover {
    border-color: #1a5f7a; transform: translateY(-4px);
    box-shadow: 0 12px 40px rgba(26, 95, 122, 0.15);
}
.doc-card-moc { border-left: 5px solid #1a5f7a; }
.doc-card-a3 { border-left: 5px solid #10b981; }
.doc-card-kaizen { border-left: 5px solid #f59e0b; }
.stButton > button {
    border-radius: 10px !important; font-weight: 600 !important;
    padding: 0.75rem 2rem !important; transition: all 0.2s ease !important;
}
.stButton > button:hover {
    transform: translateY(-2px); box-shadow: 0 4px 12px rgba(0,0,0,0.15);
}
.stTextInput > div > div > input,
.stTextArea > div > div > textarea,
.stSelectbox > div > div > div {
    border-radius: 8px !important; border: 1px solid #cbd5e1 !important;
    font-size: 15px !important;
}
.section-header {
    background: #f1f5f9; padding: 1rem 1.5rem;
    border-radius: 10px; margin: 1.5rem 0 1rem 0;
    border-left: 4px solid #1a5f7a;
}
.section-header h3 { margin: 0 !important; color: #1e293b !important; font-weight: 600 !important; }
.field-card {
    background: white; border: 1px solid #e2e8f0;
    border-radius: 10px; padding: 1rem; margin: 0.5rem 0;
}
.field-card:hover {
    border-color: #1a5f7a; box-shadow: 0 2px 8px rgba(0,0,0,0.05);
}
.gemini-badge {
    display: inline-block; background: #e0e7ff; color: #4338ca;
    padding: 0.25rem 0.75rem; border-radius: 20px;
    font-size: 12px; font-weight: 600; margin-left: 0.5rem;
}
.history-item {
    background: white; border: 1px solid #e2e8f0;
    border-radius: 10px; padding: 1rem; margin: 0.5rem 0;
    transition: all 0.2s;
}
.history-item:hover {
    border-color: #1a5f7a; box-shadow: 0 2px 8px rgba(0,0,0,0.05);
}
.app-footer {
    text-align: center; padding: 2rem; margin-top: 3rem;
    border-top: 1px solid #e2e8f0; color: #64748b;
}
.auto-correct-badge {
    display: inline-flex; align-items: center;
    background: #dcfce7; color: #166534;
    padding: 0.25rem 0.75rem; border-radius: 20px;
    font-size: 11px; font-weight: 600; margin-bottom: 0.5rem;
}
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
[data-testid="stSidebar"] { background: #1e293b !important; }
[data-testid="stSidebar"] .stMarkdown { color: #94a3b8 !important; }
[data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3 { color: #f8fafc !important; }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# =============================================================================
# PERSISTENCIA LOCAL
# =============================================================================
class LocalStorage:
    @staticmethod
    def save_config(config):
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            st.error(f"Error guardando configuración: {e}")
            return False

    @staticmethod
    def load_config():
        try:
            if CONFIG_FILE.exists():
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    return json.load(f)
        except Exception as e:
            st.warning(f"Error cargando configuración: {e}")
            return None

    @staticmethod
    def save_history(history):
        try:
            with open(HISTORY_FILE, 'w', encoding='utf-8') as f:
                json.dump(history, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            st.error(f"Error guardando historial: {e}")
            return False

    @staticmethod
    def load_history():
        try:
            if HISTORY_FILE.exists():
                with open(HISTORY_FILE, 'r', encoding='utf-8') as f:
                    return json.load(f)
        except Exception as e:
            st.warning(f"Error cargando historial: {e}")
            return {"documents": []}

    @staticmethod
    def save_template_bytes(template_name, file_bytes):
        try:
            template_path = TEMPLATES_DIR / f"{template_name}.bin"
            with open(template_path, 'wb') as f:
                f.write(file_bytes)
            return True
        except Exception as e:
            st.error(f"Error guardando template {template_name}: {e}")
            return False

    @staticmethod
    def load_template_bytes(template_name):
        try:
            template_path = TEMPLATES_DIR / f"{template_name}.bin"
            if template_path.exists():
                with open(template_path, 'rb') as f:
                    return f.read()
        except Exception as e:
            st.warning(f"Error cargando template {template_name}: {e}")
            return None

# =============================================================================
# UTILIDADES
# =============================================================================
class Utils:
    @staticmethod
    def format_date():
        meses = {1:"ENERO",2:"FEBRERO",3:"MARZO",4:"ABRIL",5:"MAYO",6:"JUNIO",
                 7:"JULIO",8:"AGOSTO",9:"SEPTIEMBRE",10:"OCTUBRE",11:"NOVIEMBRE",12:"DICIEMBRE"}
        now = datetime.now()
        return f"{meses[now.month]} {now.year}"

    @staticmethod
    def format_date_short():
        now = datetime.now()
        return now.strftime("%d/%m/%Y")

    @staticmethod
    def generate_doc_number(doc_type):
        config = st.session_state.config
        key = f"last_{doc_type}_number"
        config[key] = config.get(key, 0) + 1
        st.session_state.config = config
        LocalStorage.save_config(config)
        now = datetime.now()
        prefix = {"moc": "MOC", "a3": "A3", "kaizen": "KZN"}
        return f"{prefix.get(doc_type, 'DOC')}-{now.year}{now.month:02d}{now.day:02d}-{config[key]:04d}"

    @staticmethod
    def sanitize_filename(filename):
        return re.sub(r'[<>":/\\|?*]', '_', filename)[:50]

    @staticmethod
    def add_to_history(doc_info):
        history = st.session_state.history
        doc_info["id"] = str(uuid.uuid4())
        doc_info["timestamp"] = datetime.now().isoformat()
        history["documents"].insert(0, doc_info)
        st.session_state.history = history
        LocalStorage.save_history(history)

    @staticmethod
    def get_history(doc_type=None):
        docs = st.session_state.history.get("documents", [])
        if doc_type:
            docs = [d for d in docs if d.get("type") == doc_type]
        return docs

    @staticmethod
    def delete_from_history(doc_id):
        history = st.session_state.history
        history["documents"] = [d for d in history["documents"] if d.get("id") != doc_id]
        st.session_state.history = history
        LocalStorage.save_history(history)

    @staticmethod
    def correct_spelling_basic(text):
        """Corrector ortográfico robusto con diccionario técnico industrial"""
        if not text or not text.strip():
            return text
        corrections = {
            "tecnico": "técnico", "Tecnico": "Técnico", "TECNICO": "TÉCNICO",
            "tecnica": "técnica", "Tecnica": "Técnica", "TECNICA": "TÉCNICA",
            "tecnologia": "tecnología", "Tecnologia": "Tecnología", "TECNOLOGIA": "TECNOLOGÍA",
            "produccion": "producción", "Produccion": "Producción", "PRODUCCION": "PRODUCCIÓN",
            "implementacion": "implementación", "Implementacion": "Implementación",
            "evaluacion": "evaluación", "Evaluacion": "Evaluación",
            "operacion": "operación", "Operacion": "Operación",
            "condicion": "condición", "Condicion": "Condición",
            "modificacion": "modificación", "Modificacion": "Modificación",
            "verificacion": "verificación", "Verificacion": "Verificación",
            "capacitacion": "capacitación", "Capacitacion": "Capacitación",
            "socializacion": "socialización", "Socializacion": "Socialización",
            "documentacion": "documentación", "Documentacion": "Documentación",
            "estandarizacion": "estandarización", "Estandarizacion": "Estandarización",
            "optimizacion": "optimización", "Optimizacion": "Optimización",
            "identificacion": "identificación", "Identificacion": "Identificación",
            "clasificacion": "clasificación", "Clasificacion": "Clasificación",
            "notificacion": "notificación", "Notificacion": "Notificación",
            "coordinacion": "coordinación", "Coordinacion": "Coordinación",
            "aprobacion": "aprobación", "Aprobacion": "Aprobación",
            "revison": "revisión", "Revison": "Revisión",
            "ejecucion": "ejecución", "Ejecucion": "Ejecución",
            "inspeccion": "inspección", "Inspeccion": "Inspección",
            "proteccion": "protección", "Proteccion": "Protección",
            "deteccion": "detección", "Deteccion": "Detección",
            "prevencion": "prevención", "Prevencion": "Prevención",
            "intervencion": "intervención", "Intervencion": "Intervención",
            "supervision": "supervisión", "Supervision": "Supervisión",
            "comunicacion": "comunicación", "Comunicacion": "Comunicación",
            "organizacion": "organización", "Organizacion": "Organización",
            "planificacion": "planificación", "Planificacion": "Planificación",
            "calificacion": "calificación", "Calificacion": "Calificación",
            "certificacion": "certificación", "Certificacion": "Certificación",
            "validacion": "validación", "Validacion": "Validación",
            "calibracion": "calibración", "Calibracion": "Calibración",
            "configuracion": "configuración", "Configuracion": "Configuración",
            "programacion": "programación", "Programacion": "Programación",
            "automatizacion": "automatización", "Automatizacion": "Automatización",
            "integracion": "integración", "Integracion": "Integración",
            "funcion": "función", "Funcion": "Función",
            "relacion": "relación", "Relacion": "Relación",
            "conexion": "conexión", "Conexion": "Conexión",
            "direccion": "dirección", "Direccion": "Dirección",
            "seleccion": "selección", "Seleccion": "Selección",
            "proyeccion": "proyección", "Proyeccion": "Proyección",
            "restriccion": "restricción", "Restriccion": "Restricción",
            "distribucion": "distribución", "Distribucion": "Distribución",
            "construccion": "construcción", "Construccion": "Construcción",
            "destruccion": "destrucción", "Destruccion": "Destrucción",
            "instruccion": "instrucción", "Instruccion": "Instrucción",
            "conduccion": "conducción", "Conduccion": "Conducción",
            "introduccion": "introducción", "Introduccion": "Introducción",
            "reduccion": "reducción", "Reduccion": "Reducción",
            "reproduccion": "reproducción", "Reproduccion": "Reproducción",
            "traduccion": "traducción", "Traduccion": "Traducción",
            "deduccion": "deducción", "Deduccion": "Deducción",
            "induccion": "inducción", "Induccion": "Inducción",
            "seduccion": "seducción", "Seduccion": "Seducción",
            "maquina": "máquina", "Maquina": "Máquina", "MAQUINA": "MÁQUINA",
            "maquinas": "máquinas", "Maquinas": "Máquinas",
            "podria": "podría", "Podria": "Podría",
            "podrian": "podrían", "Podrian": "Podrían",
            "habria": "habría", "Habria": "Habría",
            "seria": "sería", "Seria": "Sería",
            "tendria": "tendría", "Tendria": "Tendría",
            "haria": "haría", "Haria": "Haría",
            "daria": "daría", "Daria": "Daría",
            "estaria": "estaría", "Estaria": "Estaría",
            "tendrian": "tendrían", "Tendrian": "Tendrían",
            "habrian": "habrían", "Habrian": "Habrían",
            "serian": "serían", "Serian": "Serían",
            "harian": "harían", "Harian": "Harían",
            "darian": "darían", "Darian": "Darían",
            "estarian": "estarían", "Estarian": "Estarían",
            "deberia": "debería", "Deberia": "Debería",
            "deberian": "deberían", "Deberian": "Deberían",
            "mas": "más", "Mas": "Más", "MAS": "MÁS",
            "aun": "aún", "Aun": "Aún",
            "solo": "solo",
            "tambien": "también", "Tambien": "También",
            "asi": "así", "Asi": "Así",
            "aqui": "aquí", "Aqui": "Aquí",
            "alli": "allí", "Alli": "Allí",
            "alla": "allá", "Alla": "Allá",
            "despues": "después", "Despues": "Después",
            "antes": "antes",
            "ademas": "además", "Ademas": "Además",
            "aunque": "aunque",
            "mientras": "mientras",
            "durante": "durante",
            "segun": "según", "Segun": "Según",
            "numero": "número", "Numero": "Número", "NUMERO": "NÚMERO",
            "maximo": "máximo", "Maximo": "Máximo", "MAXIMO": "MÁXIMO",
            "minimo": "mínimo", "Minimo": "Mínimo", "MINIMO": "MÍNIMO",
            "optimo": "óptimo", "Optimo": "Óptimo", "OPTIMO": "ÓPTIMO",
            "ultimo": "último", "Ultimo": "Último", "ULTIMO": "ÚLTIMO",
            "periodo": "período", "Periodo": "Período", "PERIODO": "PERÍODO",
            "epoca": "época", "Epoca": "Época", "EPOCA": "ÉPOCA",
            "decada": "década", "Decada": "Década", "DECADA": "DÉCADA",
            "area": "área", "Area": "Área", "AREA": "ÁREA",
            "dia": "día", "Dia": "Día", "DIA": "DÍA",
            "manana": "mañana", "Manana": "Mañana", "MANANA": "MAÑANA",
            "proximo": "próximo", "Proximo": "Próximo", "PROXIMO": "PRÓXIMO",
            "analisis": "análisis", "Analisis": "Análisis", "ANALISIS": "ANÁLISIS",
            "sintesis": "síntesis", "Sintesis": "Síntesis", "SINTESIS": "SÍNTESIS",
            "crisis": "crisis",
            "tesis": "tesis",
            "hipotesis": "hipótesis", "Hipotesis": "Hipótesis", "HIPOTESIS": "HIPÓTESIS",
            "parentesis": "paréntesis", "Parentesis": "Paréntesis", "PARENTESIS": "PARÉNTESIS",
            "sinopsis": "sinopsis",
            "axis": "axis",
            "praxis": "praxis",
            "metodo": "método", "Metodo": "Método", "METODO": "MÉTODO",
            "parametro": "parámetro", "Parametro": "Parámetro", "PARAMETRO": "PARÁMETRO",
            "parametros": "parámetros", "Parametros": "Parámetros", "PARAMETROS": "PARÁMETROS",
            "caracteristica": "característica", "Caracteristica": "Característica",
            "caracteristicas": "características", "Caracteristicas": "Características",
            "especifico": "específico", "Especifico": "Específico",
            "especifica": "específica", "Especifica": "Específica",
            "generico": "genérico", "Generico": "Genérico",
            "generica": "genérica", "Generica": "Genérica",
            "atomico": "atómico", "Atomico": "Atómico",
            "atomica": "atómica", "Atomica": "Atómica",
            "ionico": "iónico", "Ionico": "Iónico",
            "ionica": "iónica", "Ionico": "Iónica",
            "electronico": "electrónico", "Electronico": "Electrónico",
            "electronica": "electrónica", "Electronica": "Electrónica",
            "electrico": "eléctrico", "Electrico": "Eléctrico",
            "electrica": "eléctrica", "Electrica": "Eléctrica",
            "hidraulico": "hidráulico", "Hidraulico": "Hidráulico",
            "hidraulica": "hidráulica", "Hidraulica": "Hidráulica",
            "neumatico": "neumático", "Neumatico": "Neumático",
            "neumatica": "neumática", "Neumatica": "Neumática",
            "termico": "térmico", "Termico": "Térmico",
            "termica": "térmica", "Termica": "Térmica",
            "optico": "óptico", "Optico": "Óptico",
            "optica": "óptica", "Optica": "Óptica",
            "acustico": "acústico", "Acustico": "Acústico",
            "acustica": "acústica", "Acustica": "Acústica",
            "magnetico": "magnético", "Magnetico": "Magnético",
            "magnetica": "magnética", "Magnetica": "Magnética",
            "quimico": "químico", "Quimico": "Químico",
            "quimica": "química", "Quimica": "Química",
            "fisico": "físico", "Fisico": "Físico",
            "fisica": "física", "Fisica": "Física",
            "biologico": "biológico", "Biologico": "Biológico",
            "biologica": "biológica", "Biologica": "Biológica",
            "geologico": "geológico", "Geologico": "Geológico",
            "geologica": "geológica", "Geologica": "Geológica",
            "ecologico": "ecológico", "Ecologico": "Ecológico",
            "ecologica": "ecológica", "Ecologica": "Ecológica",
            "psicologico": "psicológico", "Psicologico": "Psicológico",
            "psicologica": "psicológica", "Psicologica": "Psicológica",
            "sociologico": "sociológico", "Sociologico": "Sociológico",
            "sociologica": "sociológica", "Sociologica": "Sociológica",
            "antropologico": "antropológico", "Antropologico": "Antropológico",
            "antropologica": "antropológica", "Antropologica": "Antropológica",
            "arqueologico": "arqueológico", "Arqueologico": "Arqueológico",
            "arqueologica": "arqueológica", "Arqueologica": "Arqueológica",
            "filosofico": "filosófico", "Filosofico": "Filosófico",
            "filosofica": "filosófica", "Filosofica": "Filosófica",
            "historico": "histórico", "Historico": "Histórico",
            "historica": "histórica", "Historica": "Histórica",
            "economico": "económico", "Economico": "Económico",
            "economica": "económica", "Economica": "Económica",
            "politico": "político", "Politico": "Político",
            "politica": "política", "Politica": "Política",
            "juridico": "jurídico", "Juridico": "Jurídico",
            "juridica": "jurídica", "Juridica": "Jurídica",
            "artistico": "artístico", "Artistico": "Artístico",
            "artistica": "artística", "Artistica": "Artística",
            "literario": "literario",
            "literaria": "literaria",
            "musical": "musical",
            "plastico": "plástico", "Plastico": "Plástico",
            "plastica": "plástica", "Plastica": "Plástica",
            "grafico": "gráfico", "Grafico": "Gráfico",
            "grafica": "gráfica", "Grafica": "Gráfica",
            "geografico": "geográfico", "Geografico": "Geográfico",
            "geografica": "geográfica", "Geografica": "Geográfica",
            "topografico": "topográfico", "Topografico": "Topográfico",
            "topografica": "topográfica", "Topografica": "Topográfica",
            "cartografico": "cartográfico", "Cartografico": "Cartográfico",
            "cartografica": "cartográfica", "Cartografica": "Cartográfica",
            "fotografico": "fotográfico", "Fotografico": "Fotográfico",
            "fotografica": "fotográfica", "Fotografica": "Fotográfica",
            "radiografico": "radiográfico", "Radiografico": "Radiográfico",
            "radiografica": "radiográfica", "Radiografica": "Radiográfica",
            "cinematografico": "cinematográfico", "Cinematografico": "Cinematográfico",
            "cinematografica": "cinematográfica", "Cinematografica": "Cinematográfica",
            "autobiografico": "autobiográfico", "Autobiografico": "Autobiográfico",
            "autobiografica": "autobiográfica", "Autobiografica": "Autobiográfica",
            "bibliografico": "bibliográfico", "Bibliografico": "Bibliográfico",
            "bibliografica": "bibliográfica", "Bibliografica": "Bibliográfica",
            "discografico": "discográfico", "Discografico": "Discográfico",
            "discografica": "discográfica", "Discografica": "Discográfica",
            "lexicografico": "lexicográfico", "Lexicografico": "Lexicográfico",
            "lexicografica": "lexicográfica", "Lexicografica": "Lexicográfica",
            "ortografico": "ortográfico", "Ortografico": "Ortográfico",
            "ortografica": "ortográfica", "Ortografica": "Ortográfica",
            "estenografico": "estenográfico", "Estenografico": "Estenográfico",
            "estenografica": "estenográfica", "Estenografica": "Estenográfica",
            "estilografico": "estilográfico", "Estilografico": "Estilográfico",
            "estilografica": "estilográfica", "Estilografica": "Estilográfica",
            "monografico": "monográfico", "Monografico": "Monográfico",
            "monografica": "monográfica", "Monografica": "Monográfica",
            "poligrafo": "polígrafo", "Poligrafo": "Polígrafo",
            "poligrafa": "polígrafa", "Poligrafa": "Polígrafa",
            "paragrafo": "párrafo", "Paragrafo": "Párrafo",
            "paragrafos": "párrafos", "Paragrafos": "Párrafos",
            "telegrafo": "telégrafo", "Telegrafo": "Telégrafo",
            "telegrafos": "telégrafos", "Telegrafos": "Telégrafos",
            "telegrama": "telegrama",
            "programa": "programa",
            "programas": "programas",
            "programatico": "programático", "Programatico": "Programático",
            "programatica": "programática", "Programatica": "Programática",
            "programador": "programador",
            "programadora": "programadora",
            "programacion": "programación", "Programacion": "Programación",
            "programable": "programable",
            "reprogramable": "reprogramable",
            "desprogramar": "desprogramar",
            "reprogramar": "reprogramar",
            "compilador": "compilador",
            "compiladora": "compiladora",
            "compilacion": "compilación", "Compilacion": "Compilación",
            "interpretador": "interpretador",
            "interpretadora": "interpretadora",
            "interpretacion": "interpretación", "Interpretacion": "Interpretación",
            "traductor": "traductor",
            "traductora": "traductora",
            "traduccion": "traducción", "Traduccion": "Traducción",
            "traducible": "traducible",
            "intraducible": "intraducible",
            "version": "versión", "Version": "Versión", "VERSION": "VERSIÓN",
            "reversion": "reversión", "Reversion": "Reversión",
            "conversion": "conversión", "Conversion": "Conversión",
            "inversion": "inversión", "Inversion": "Inversión",
            "diversion": "diversión", "Diversion": "Diversión",
            "aversion": "aversión", "Aversion": "Aversión",
            "perversion": "perversión", "Perversion": "Perversión",
            "subversion": "subversión", "Subversion": "Subversión",
            "introversion": "introversión", "Introversion": "Introversión",
            "extroversion": "extroversión", "Extroversion": "Extroversión",
            "retroversion": "retroversión", "Retroversion": "Retroversión",
            "controversion": "controversión", "Controversion": "Controversión",
            "adversion": "adversión", "Adversion": "Adversión",
            "trabagar": "trabajar",
            "podra": "podrá",
            "configura": "configura",
            "maxima": "máxima",
            "limite": "límite",
            "Habilitacion": "Habilitación",
            "habilitacion": "habilitación",
            "Velocidad": "Velocidad",
            "Prensa": "Prensa",
            "Casquillos": "Casquillos",
            "comentan": "comentan",
            "operadores": "operadores",
            "regular": "regular",
            "velocidad": "velocidad",
            "panel": "panel",
            "tiene": "tiene",
            "seguridad": "seguridad",
            "incrementando": "incrementando",
            "puede": "puede",
            "llegar": "llegar",
            "frecuencia": "frecuencia",
            "motor": "motor",
            "equipo": "equipo",
            "esta": "está",
            "peligroso": "peligroso",
            "funcione": "funcione",
            "romper": "romper",
            "algunas": "algunas",
            "piezas": "piezas",
            "normalmente": "normalmente",
            "debe": "debe",
            "trabajar": "trabajar",
            "variador": "variador",
            "forma": "forma",
            "girar": "girar",
            "descripcion": "descripción",
            "solucion": "solución",
            "implementada": "implementada",
            "beneficios": "beneficios",
            "proximos": "próximos",
            "pasos": "pasos",
            "desperdicio": "desperdicio",
            "impacto": "impacto",
            "bto": "BTO",
            "safe": "Safe",
            "sustainable": "Sustainable",
            "people": "People",
            "culture": "Culture",
            "network": "Network",
            "optimisation": "Optimisation",
            "supply": "Supply",
            "chain": "Chain",
            "manufacturing": "Manufacturing",
            "excellence": "Excellence",
            "motion": "Motion",
            "skills": "Skills",
            "inventory": "Inventory",
            "transportation": "Transportation",
            "over production": "Over Production",
            "over processing": "Over Processing",
            "waiting": "Waiting",
            "defects": "Defects",
            "opportunity": "Opportunity",
            "improvement": "Improvement",
            "benefit": "Benefit",
            "leader": "Leader",
            "team": "Team",
            "members": "Members",
            "plant": "Plant",
            "date": "Date",
            "name": "Name",
            "simple": "Simple",
            "kaizen": "Kaizen",
            "moc": "MoC",
            "mejora": "Mejora",
            "a3": "A3",
            "management": "Management",
            "change": "Change",
            "naturaleza": "naturaleza",
            "originador": "originador",
            "produccion": "producción",
            "specialist": "Specialist",
            "shes": "SHES",
            "mantenimiento": "mantenimiento",
            "revisores": "revisores",
            "enablon": "Enablon",
            "revisor": "revisor",
            "aprobador": "aprobador",
            "final": "final",
            "experto": "experto",
            "revision": "revisión",
            "especialistas": "especialistas",
            "expertos": "expertos",
            "problema": "problema",
            "condicion": "condición",
            "actual": "actual",
            "propuesta": "propuesta",
            "razones": "razones",
            "cambio": "cambio",
            "alternativas": "alternativas",
            "consideradas": "consideradas",
            "plan": "plan",
            "retorno": "retorno",
            "recursos": "recursos",
            "disponibles": "disponibles",
            "implementacion": "implementación",
            "tiempo": "tiempo",
            "dura": "dura",
            "resultado": "resultado",
            "evaluacion": "evaluación",
            "estudio": "estudio",
            "riesgos": "riesgos",
            "identificado": "identificado",
            "controles": "controles",
            "recomendados": "recomendados",
            "medidas": "medidas",
            "control": "control",
            "propuestos": "propuestos",
            "plazo": "plazo",
            "fin": "fin",
            "presentacion": "presentación",
            "autor": "autor",
            "miembros": "miembros",
            "equipo": "equipo",
            "antecedentes": "antecedentes",
            "situacion": "situación",
            "objetivos": "objetivos",
            "causa": "causa",
            "raiz": "raíz",
            "contramedidas": "contramedidas",
            "resultados": "resultados",
            "esperados": "esperados",
            "seguimiento": "seguimiento",
            "lecciones": "lecciones",
            "aprendidas": "aprendidas",
            "estandarizacion": "estandarización",
            "exelente": "excelente", "Exelente": "Excelente",
            "exelencia": "excelencia", "Exelencia": "Excelencia",
            "deficiente": "deficiente",
            "suficiente": "suficiente",
            "insuficiente": "insuficiente",
            "necesario": "necesario",
            "innecesario": "innecesario",
            "obligatorio": "obligatorio",
            "voluntario": "voluntario",
            "opcional": "opcional",
            "requerido": "requerido",
            "requerimiento": "requerimiento",
            "requisito": "requisito",
            "especificacion": "especificación",
            "especifico": "específico",
            "generico": "genérico",
            "particular": "particular",
            "general": "general",
            "especial": "especial",
            "especifica": "específica",
            "generica": "genérica",
            "atomico": "atómico",
            "atomica": "atómica",
            "ionico": "iónico",
            "ionica": "iónica",
            "electronico": "electrónico",
            "electronica": "electrónica",
            "electrico": "eléctrico",
            "electrica": "eléctrica",
            "hidraulico": "hidráulico",
            "hidraulica": "hidráulica",
            "neumatico": "neumático",
            "neumatica": "neumática",
            "termico": "térmico",
            "termica": "térmica",
            "optico": "óptico",
            "optica": "óptica",
            "acustico": "acústico",
            "acustica": "acústica",
            "magnetico": "magnético",
            "magnetica": "magnética",
            "quimico": "químico",
            "quimica": "química",
            "fisico": "físico",
            "fisica": "física",
            "biologico": "biológico",
            "biologica": "biológica",
            "geologico": "geológico",
            "geologica": "geológica",
            "ecologico": "ecológico",
            "ecologica": "ecológica",
            "psicologico": "psicológico",
            "psicologica": "psicológica",
            "sociologico": "sociológico",
            "sociologica": "sociológica",
            "antropologico": "antropológico",
            "antropologica": "antropológica",
            "arqueologico": "arqueológico",
            "arqueologica": "arqueológica",
            "filosofico": "filosófico",
            "filosofica": "filosófica",
            "historico": "histórico",
            "historica": "histórica",
            "economico": "económico",
            "economica": "económica",
            "politico": "político",
            "politica": "política",
            "juridico": "jurídico",
            "juridica": "jurídica",
            "artistico": "artístico",
            "artistica": "artística",
            "literario": "literario",
            "literaria": "literaria",
            "musical": "musical",
            "plastico": "plástico",
            "plastica": "plástica",
            "grafico": "gráfico",
            "grafica": "gráfica",
            "geografico": "geográfico",
            "geografica": "geográfica",
            "topografico": "topográfico",
            "topografica": "topográfica",
            "cartografico": "cartográfico",
            "cartografica": "cartográfica",
            "fotografico": "fotográfico",
            "fotografica": "fotográfica",
            "radiografico": "radiográfico",
            "radiografica": "radiográfica",
            "cinematografico": "cinematográfico",
            "cinematografica": "cinematográfica",
            "autobiografico": "autobiográfico",
            "autobiografica": "autobiográfica",
            "bibliografico": "bibliográfico",
            "bibliografica": "bibliográfica",
            "discografico": "discográfico",
            "discografica": "discográfica",
            "lexicografico": "lexicográfico",
            "lexicografica": "lexicográfica",
            "ortografico": "ortográfico",
            "ortografica": "ortográfica",
            "estenografico": "estenográfico",
            "estenografica": "estenográfica",
            "estilografico": "estilográfico",
            "estilografica": "estilográfica",
            "monografico": "monográfico",
            "monografica": "monográfica",
            "poligrafo": "polígrafo",
            "poligrafa": "polígrafa",
            "paragrafo": "párrafo",
            "paragrafos": "párrafos",
            "telegrafo": "telégrafo",
            "telegrafos": "telégrafos",
            "telegrama": "telegrama",
            "programa": "programa",
            "programas": "programas",
            "programatico": "programático",
            "programatica": "programática",
            "programador": "programador",
            "programadora": "programadora",
            "programacion": "programación",
            "programable": "programable",
            "reprogramable": "reprogramable",
            "desprogramar": "desprogramar",
            "reprogramar": "reprogramar",
            "compilador": "compilador",
            "compiladora": "compiladora",
            "compilacion": "compilación",
            "interpretador": "interpretador",
            "interpretadora": "interpretadora",
            "interpretacion": "interpretación",
            "traductor": "traductor",
            "traductora": "traductora",
            "traduccion": "traducción",
            "traducible": "traducible",
            "intraducible": "intraducible",
            "version": "versión",
            "reversion": "reversión",
            "conversion": "conversión",
            "inversion": "inversión",
            "diversion": "diversión",
            "aversion": "aversión",
            "perversion": "perversión",
            "subversion": "subversión",
            "introversion": "introversión",
            "extroversion": "extroversión",
            "retroversion": "retroversión",
            "controversion": "controversión",
            "adversion": "adversión",
        }
        result = text
        for wrong, correct in corrections.items():
            result = result.replace(wrong, correct)
        result = re.sub(r'  +', ' ', result)
        result = re.sub(r' ([.,;:!?])', r'\1', result)
        return result

# =============================================================================
# SERVICIO GEMINI API - CON DIAGNÓSTICO MEJORADO
# =============================================================================
class GeminiService:
    MODELS = {
        "gemini-1.5-pro": {"name": "Gemini 1.5 Pro", "desc": "Máxima calidad y razonamiento"},
        "gemini-1.5-flash": {"name": "Gemini 1.5 Flash", "desc": "Rápido y eficiente"},
    }

    def __init__(self, api_key="", model="gemini-1.5-pro"):
        self.api_key = api_key
        self.model = model
        self.base_url = "https://generativelanguage.googleapis.com/v1beta/models"

    def _call_api(self, prompt, temperature=0.3, max_tokens=4096):
        if not self.api_key:
            raise ValueError("API Key no configurada")
        import requests
        url = f"{self.base_url}/{self.model}:generateContent?key={self.api_key}"
        payload = {
            "contents": [{"parts": [{"text": prompt}]}],
            "generationConfig": {"temperature": temperature, "maxOutputTokens": max_tokens}
        }
        try:
            response = requests.post(url, json=payload, timeout=120)
            response.raise_for_status()
            result = response.json()
            if "candidates" in result and len(result["candidates"]) > 0:
                return result["candidates"][0]["content"]["parts"][0]["text"]
            return ""
        except requests.exceptions.HTTPError as e:
            if e.response.status_code == 404:
                raise Exception("Error 404: La 'Generative Language API' NO está habilitada en tu proyecto de Google Cloud, o el nombre del modelo es incorrecto. Ve a Google Cloud Console > APIs y Servicios > Biblioteca > busca 'Generative Language API' y habilítala.")
            elif e.response.status_code == 403:
                raise Exception("Error 403: Acceso denegado. La API Key puede estar restringida (por IP o Referer) o no tener permisos de uso.")
            else:
                raise Exception(f"Error HTTP {e.response.status_code}: {e.response.text}")

    def _extract_json(self, text):
        import json
        json_match = re.search(r'```json\s*(.*?)\s*```', text, re.DOTALL)
        if json_match:
            try:
                return json.loads(json_match.group(1))
            except json.JSONDecodeError:
                pass
        
        json_match = re.search(r'\{.*\}', text, re.DOTALL)
        if json_match:
            try:
                return json.loads(json_match.group())
            except json.JSONDecodeError:
                pass
                
        return {"generated_text": text, "error": "No se pudo extraer JSON válido"}

    def generate_moc(self, problem, context="", equipo=""):
        if not self.api_key:
            st.error("❌ API Key no configurada. Configure en Configuración > API Gemini")
            return None
        
        prompt = f"""Eres un ingeniero senior de seguridad industrial con 20 años de experiencia, especializado en Management of Change (MoC) bajo normas ISO 45001, ISO 9001 e ISO 13849.

REGLAS CRÍTICAS OBLIGATORIAS (LEER DETENIDAMENTE):
1. CONTEXTO EXCLUSIVO: Toda la redacción debe basarse ÚNICA Y EXCLUSIVAMENTE en el problema reportado por el usuario a continuación.
2. PROHIBICIÓN DE TEXTO GENÉRICO: Está TERMINANTEMENTE PROHIBIDO usar frases genéricas como "degradación progresiva de componentes", "parámetros fuera de rango", "desviaciones del proceso" o "condición técnica que afecta la continuidad" a menos que el usuario las haya escrito explícitamente.
3. EXTRACCIÓN DE ENTIDADES: Debes identificar y usar los elementos específicos del usuario: equipos (estaciones de espera, máquina), componentes (interlocks, compuertas, sensores de seguridad, PLC), riesgos (atrapamiento, material energético) y normas (ISO 13849).
4. REDACCIÓN HUMANIZADA Y TÉCNICA: Escribe como un ingeniero senior. Usa voz activa, conectores lógicos y párrafos bien estructurados.
5. ORTOGRAFÍA IMPECABLE: Tildes correctas en todas las palabras.

PROBLEMA REPORTADO POR EL USUARIO:
{problem}

CONTEXTO ADICIONAL:
{context}

EQUIPO INVOLUCRADO:
{equipo}

Genera en ESPAÑOL formato JSON con esta estructura EXACTA:
{{
  "moc_title": "Título técnico conciso (máx. 12 palabras, basado en el problema del usuario)",
  "descripcion_problema": "Descripción técnica detallada del problema reportado. Menciona EXPLÍCITAMENTE las estaciones de espera, la falta de interlocks, el riesgo de atrapamiento, el material energético y la necesidad de cumplir con ISO 13849. Mínimo 250 palabras.",
  "condicion_actual": "Descripción técnica exhaustiva del estado actual. Explica EXPLÍCITAMENTE que las compuertas pueden abrirse con la máquina en funcionamiento, la falta de sensores en puntos de acceso expuestos y la ausencia de enclavamiento de seguridad en el PLC.",
  "condicion_propuesta": "Descripción detallada de la solución propuesta. Explica EXPLÍCITAMENTE la instalación de interlocks en cada estación de espera, la detención automática de la máquina al abrirse las compuertas, la habilitación previa desde el panel de control y la integración al sistema de enclavamiento del PLC.",
  "razones_cambio": "Lista de 4-6 razones técnicas usando viñetas '❖' que justifiquen el cambio basándose en el problema del usuario (ej: prevención de atrapamiento, cumplimiento ISO 13849, protección del material energético).",
  "alternativas_retorno": "Análisis de 2 alternativas evaluadas con pros/contras específicos para este problema. Incluye un plan de retorno detallado para desinstalar los interlocks y restaurar la operación manual segura si falla la implementación.",
  "recursos": "Listado exhaustivo de recursos humanos, materiales (sensores de seguridad, cableado, módulos de seguridad para PLC, interlocks), técnicos y EPP específico requeridos.",
  "plan_implementacion": "Plan detallado por fases para instalar interlocks: instalación física, cableado, programación del PLC para el enclavamiento, pruebas de funcionamiento, validación de seguridad.",
  "tiempo_duracion": "Estimación detallada del tiempo total para instalar interlocks con desglose por fase.",
  "riesgos_controles": [{{"riesgo": "Riesgo específico de calidad/técnico relacionado con la integración del PLC", "control": "Medida de control específica"}}],
  "riesgos_shes": [{{"riesgo": "Riesgo SHES específico (ej: atrapamiento por compuertas, energía residual)", "control": "Plan de acción específico", "plazo": "Plazo"}}]
}}

Responde SOLO con el JSON válido, sin comentarios adicionales."""
        try:
            text = self._call_api(prompt, temperature=0.4, max_tokens=8192)
            result = self._extract_json(text)
            for key in result:
                if isinstance(result[key], str):
                    result[key] = Utils.correct_spelling_basic(result[key])
                elif isinstance(result[key], list):
                    for item in result[key]:
                        if isinstance(item, dict):
                            for k in item:
                                if isinstance(item[k], str):
                                    item[k] = Utils.correct_spelling_basic(item[k])
            return result
        except Exception as e:
            st.error(f"❌ Error API: {e}")
            return None

    def generate_a3(self, problem, context=""):
        if not self.api_key:
            st.error("❌ API Key no configurada")
            return None
        prompt = f"""Eres un experto senior en metodología A3 Lean. Redactas documentos con redacción humanizada, técnica y profesional.
INSTRUCCIONES: Usa EXCLUSIVAMENTE el contexto del problema. NO inventes problemas genéricos. Ortografía impecable.
PROBLEMA: {problem}
CONTEXTO: {context}
Genera JSON con: titulo, antecedentes, problema_actual, analisis_situacion, objetivos, analisis_causa_raiz, contramedidas, resultados_esperados, plan_seguimiento, lecciones_aprendidas, estandarizacion."""
        try:
            text = self._call_api(prompt, temperature=0.4, max_tokens=8192)
            result = self._extract_json(text)
            for key in result:
                if isinstance(result[key], str):
                    result[key] = Utils.correct_spelling_basic(result[key])
            return result
        except Exception as e:
            st.error(f"❌ Error API: {e}")
            return None

    def generate_kaizen(self, activity, context=""):
        if not self.api_key:
            st.error("❌ API Key no configurada")
            return None
        prompt = f"""Eres un experto en Kaizen. Redactas registros con redacción humanizada y práctica.
INSTRUCCIONES: Usa EXCLUSIVAMENTE el contexto. NO inventes problemas. Ortografía impecable.
ACTIVIDAD: {activity}
CONTEXTO: {context}
Genera JSON con: titulo, area, descripcion_problema, solucion, beneficios, tipo_desperdicio, impacto_bto, proximos_pasos, leader, team_members."""
        try:
            text = self._call_api(prompt, temperature=0.4, max_tokens=4096)
            result = self._extract_json(text)
            for key in result:
                if isinstance(result[key], str):
                    result[key] = Utils.correct_spelling_basic(result[key])
            return result
        except Exception as e:
            st.error(f"❌ Error API: {e}")
            return None

    def translate_document(self, data):
        if not self.api_key:
            return data
        prompt = f"""Traduce del español al inglés profesional industrial: {json.dumps(data, ensure_ascii=False, indent=2)}. Responde SOLO el JSON traducido."""
        try:
            text = self._call_api(prompt, temperature=0.2, max_tokens=8192)
            return self._extract_json(text)
        except:
            return data

    def correct_spelling(self, text):
        if not self.api_key or not text.strip():
            return Utils.correct_spelling_basic(text)
        prompt = f"""Corrige ortografía, gramática y puntuación. Mantén el significado técnico. Asegura tildes correctas. Devuelve SOLO el texto corregido.\nTEXTO:\n{text}"""
        try:
            corrected = self._call_api(prompt, temperature=0.2, max_tokens=4096).strip()
            return Utils.correct_spelling_basic(corrected)
        except:
            return Utils.correct_spelling_basic(text)

# =============================================================================
# REEMPLAZO INTELIGENTE DE TEXTO EN POWERPOINT Y WORD
# =============================================================================
def replace_text_in_shape(shape, old_text, new_text):
    if not shape.has_text_frame:
        return False
    text_frame = shape.text_frame
    text = text_frame.text
    if old_text not in text:
        return False
    for paragraph in text_frame.paragraphs:
        paragraph_text = paragraph.text
        if old_text in paragraph_text:
            for run in paragraph.runs:
                if old_text in run.text:
                    run.text = run.text.replace(old_text, new_text)
                    return True
            if paragraph.runs:
                paragraph.runs[0].text = paragraph_text.replace(old_text, new_text)
                for run in paragraph.runs[1:]:
                    run.text = ""
                return True
    return False

def replace_all_text_in_presentation(prs, replacements):
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for old_text, new_text in replacements.items():
                    replace_text_in_shape(shape, old_text, new_text)
            if shape.has_table:
                table = shape.table
                for row in table.rows:
                    for cell in row.cells:
                        for old_text, new_text in replacements.items():
                            if old_text in cell.text:
                                for paragraph in cell.text_frame.paragraphs:
                                    for run in paragraph.runs:
                                        if old_text in run.text:
                                            run.text = run.text.replace(old_text, new_text)
                                            break
                                else:
                                    if old_text in paragraph.text:
                                        paragraph.text = paragraph.text.replace(old_text, new_text)

def fill_table_cell(cell, text):
    if cell.text_frame.paragraphs:
        first_para = cell.text_frame.paragraphs[0]
        if first_para.runs:
            first_run = first_para.runs[0]
            first_run.text = str(text)
        else:
            first_para.text = str(text)
    else:
        cell.text = str(text)

def replace_text_in_docx_preserve_runs(doc, old_text, new_text):
    def replace_in_paragraph(paragraph):
        if old_text not in paragraph.text:
            return False
        for run in paragraph.runs:
            if old_text in run.text:
                run.text = run.text.replace(old_text, new_text)
                return True
        full_text = paragraph.text
        if old_text in full_text:
            if paragraph.runs:
                first_run = paragraph.runs[0]
                first_run.text = full_text.replace(old_text, new_text)
                for run in paragraph.runs[1:]:
                    run.text = ""
                return True
        return False
    for para in doc.paragraphs:
        replace_in_paragraph(para)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    replace_in_paragraph(para)
    for section in doc.sections:
        for para in section.header.paragraphs:
            replace_in_paragraph(para)
        for table in section.header.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        replace_in_paragraph(para)
        for para in section.footer.paragraphs:
            replace_in_paragraph(para)
        for table in section.footer.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        replace_in_paragraph(para)

# =============================================================================
# GENERADOR DE DOCUMENTOS
# =============================================================================
class DocumentGenerator:
    def generate_moc(self, data, images=None, template_bytes=None):
        if template_bytes is None:
            st.error("❌ Template MoC no cargado. Vaya a Configuración > Templates.")
            return None
        prs = Presentation(BytesIO(template_bytes))
        replacements = {
            "MOC:  OPTIMIZACIÓN DEL SISTEMA DE ALIMENTACIÓN DE RETARDOS EN ST08": f"MOC:  {data.get('moc_title', '')}",
            "MOC: OPTIMIZACIÓN DEL SISTEMA DE ALIMENTACIÓN DE RETARDOS EN ST08": f"MOC:  {data.get('moc_title', '')}",
            "FECHA 29/05/2026": f"FECHA {data.get('fecha', Utils.format_date_short())}",
            "FECHA 29/05/2025": f"FECHA {data.get('fecha', Utils.format_date_short())}",
            "NÚMERO DE LA MOC: 2026-MOC00307017": f"NÚMERO DE LA MOC: {data.get('moc_number', '')}",
            "NUMERO DE LA MOC: 2026-MOC00307017": f"NÚMERO DE LA MOC: {data.get('moc_number', '')}",
            "nÚmero DE LA MOC: XXXXXXXXXXXXX": f"NÚMERO DE LA MOC: {data.get('moc_number', '')}",
            "NATURALEZA DE LA MOC: PERMANENTE": f"NATURALEZA DE LA MOC: {data.get('naturaleza', 'PERMANENTE').upper()}",
            "NATURALEZA DE LA MOC: PERMANENTE ": f"NATURALEZA DE LA MOC: {data.get('naturaleza', 'PERMANENTE').upper()}",
            "Naturaleza de la moc: permanente": f"Naturaleza de la moc: {data.get('naturaleza', 'permanente')}",
            "ORIGINADOR DE LA MOC: DANIEL VICENTE/ ERNESTO RAMIREZ/ ROGER HUAMANI": f"ORIGINADOR DE LA MOC: {data.get('originador', '')}",
            "ORIGINADOR DE LA MOC: ROGER HUAMANI": f"ORIGINADOR DE LA MOC: {data.get('originador', '')}",
            "ORIGINADOR DE LA MOC:": f"ORIGINADOR DE LA MOC: {data.get('originador', '')}",
        }
        replace_all_text_in_presentation(prs, replacements)

        if len(prs.slides) > 1:
            slide2 = prs.slides[1]
            equipo_replacements = {
                "Daniel Vicente": data.get('produccion', ''),
                "Viviana Hillon": data.get('specialist_shes', ''),
                "Ernesto Ramirez/ Roger Huamani": data.get('mantenimiento', ''),
                "Ernesto Ramirez": data.get('mantenimiento', ''),
                "Roger Huamani": data.get('mantenimiento', ''),
                "Maribel Burgos": data.get('revisor1', ''),
                "Roberto Lopez": data.get('revisor2', ''),
                "Victor Florian": data.get('revisor3', ''),
                "Max Huaman": data.get('revisor4', ''),
                "Hector Montoya": data.get('aprobador_final', ''),
                "Gary Davies/ Roberto Lopez": data.get('expertos', ''),
                "Gary Davies": data.get('experto1', ''),
            }
            for shape in slide2.shapes:
                if shape.has_text_frame:
                    for old_text, new_text in equipo_replacements.items():
                        if new_text and old_text in shape.text_frame.text:
                            replace_text_in_shape(shape, old_text, new_text)

        if len(prs.slides) > 2:
            slide3 = prs.slides[2]
            for shape in slide3.shapes:
                if shape.has_table:
                    table = shape.table
                    if len(table.rows) >= 2 and len(table.columns) >= 2:
                        fill_table_cell(table.cell(1, 0), data.get('condicion_actual', ''))
                        fill_table_cell(table.cell(1, 1), data.get('condicion_propuesta', ''))

        if len(prs.slides) > 3:
            slide4 = prs.slides[3]
            for shape in slide4.shapes:
                if shape.has_text_frame:
                    text = shape.text_frame.text
                    if "Razones del cambio" in text or "Eliminación de paradas" in text:
                        for paragraph in shape.text_frame.paragraphs:
                            para_text = paragraph.text
                            if "Eliminación" in para_text or "Adaptación" in para_text or "Reducción" in para_text or "Mejora" in para_text:
                                for run in paragraph.runs:
                                    run.text = ""
                        if shape.text_frame.paragraphs:
                            first_para = shape.text_frame.paragraphs[0]
                            if first_para.runs:
                                first_para.runs[0].text = data.get('razones_cambio', '')
                            else:
                                first_para.text = data.get('razones_cambio', '')
                    if "Alternativas consideradas" in text or "Mantener sistema actual" in text:
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                run.text = ""
                        if shape.text_frame.paragraphs:
                            first_para = shape.text_frame.paragraphs[0]
                            if first_para.runs:
                                first_para.runs[0].text = data.get('alternativas_retorno', '')
                            else:
                                first_para.text = data.get('alternativas_retorno', '')
                    if "Plan de retorno" in text or "Reinstalación del sistema actual" in text:
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                run.text = ""
                        if shape.text_frame.paragraphs:
                            first_para = shape.text_frame.paragraphs[0]
                            if first_para.runs:
                                first_para.runs[0].text = data.get('alternativas_retorno', '')
                            else:
                                first_para.text = data.get('alternativas_retorno', '')

        if len(prs.slides) > 4:
            slide5 = prs.slides[4]
            for shape in slide5.shapes:
                if shape.has_text_frame:
                    text = shape.text_frame.text
                    if "Descripción del Problema" in text or "Al procesar elementos" in text or "Cuando un elemento" in text:
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                run.text = ""
                        if shape.text_frame.paragraphs:
                            first_para = shape.text_frame.paragraphs[0]
                            if first_para.runs:
                                first_para.runs[0].text = data.get('descripcion_problema', '')
                            else:
                                first_para.text = data.get('descripcion_problema', '')

        if images and len(prs.slides) > 5:
            for idx, img_info in enumerate(images):
                if idx < 2 and len(prs.slides) > 5 + idx:
                    target_slide = prs.slides[5 + idx]
                    try:
                        img_path = img_info["path"] if isinstance(img_info, dict) else img_info
                        img = Image.open(img_path)
                        img_w, img_h = img.size
                        aspect = img_w / img_h
                        max_w = Inches(8)
                        max_h = Inches(5)
                        if aspect > (8/5):
                            w = max_w
                            h = w / aspect
                        else:
                            h = max_h
                            w = h * aspect
                        img_left = (Inches(10) - w) / 2
                        target_slide.shapes.add_picture(img_path, img_left, Inches(1.5), w, h)
                    except Exception as e:
                        st.warning(f"Error con imagen {idx+1}: {e}")
                else:
                    blank_layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]
                    new_slide = prs.slides.add_slide(blank_layout)
                    try:
                        img_path = img_info["path"] if isinstance(img_info, dict) else img_info
                        img = Image.open(img_path)
                        img_w, img_h = img.size
                        aspect = img_w / img_h
                        max_w = Inches(9)
                        max_h = Inches(6)
                        if aspect > (9/6):
                            w = max_w
                            h = w / aspect
                        else:
                            h = max_h
                            w = h * aspect
                        img_left = (Inches(10) - w) / 2
                        new_slide.shapes.add_picture(img_path, img_left, Inches(1), w, h)
                        desc = img_info.get("desc", f"Figura {idx+1}") if isinstance(img_info, dict) else f"Figura {idx+1}"
                        desc_box = new_slide.shapes.add_textbox(Inches(0.5), Inches(7.3), Inches(9), Inches(0.5))
                        dtf = desc_box.text_frame
                        dtf.text = desc
                        for p in dtf.paragraphs:
                            p.font.size = Pt(11)
                            p.font.italic = True
                            p.alignment = PP_ALIGN.CENTER
                    except Exception as e:
                        st.warning(f"Error con imagen {idx+1}: {e}")

        if len(prs.slides) > 7:
            slide8 = prs.slides[7]
            for shape in slide8.shapes:
                if shape.has_text_frame:
                    text = shape.text_frame.text
                    if "Recursos" in text and "Mecánicos" in text:
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                run.text = ""
                        if shape.text_frame.paragraphs:
                            first_para = shape.text_frame.paragraphs[0]
                            if first_para.runs:
                                first_para.runs[0].text = f"Recursos\n{data.get('recursos', '')}"
                            else:
                                first_para.text = f"Recursos\n{data.get('recursos', '')}"
                    if "Plan de Implementación" in text or "Pruebas de deslizamiento" in text:
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                run.text = ""
                        if shape.text_frame.paragraphs:
                            first_para = shape.text_frame.paragraphs[0]
                            if first_para.runs:
                                first_para.runs[0].text = f"Plan de Implementación\n{data.get('plan_implementacion', '')}"
                            else:
                                first_para.text = f"Plan de Implementación\n{data.get('plan_implementacion', '')}"
                    if "tiempo estimado" in text.lower() or "1.5 horas" in text:
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                run.text = ""
                        if shape.text_frame.paragraphs:
                            first_para = shape.text_frame.paragraphs[0]
                            if first_para.runs:
                                first_para.runs[0].text = data.get('tiempo_duracion', '')
                            else:
                                first_para.text = data.get('tiempo_duracion', '')

        if len(prs.slides) > 8:
            slide9 = prs.slides[8]
            # Aquí iría el checklist 360 si el template lo tiene como tabla

        if len(prs.slides) > 9:
            slide10 = prs.slides[9]
            # Aquí iría la tabla de documentos impactados

        if len(prs.slides) > 11:
            slide12 = prs.slides[11]
            riesgos_shes = data.get('riesgos_shes', [])
            riesgos_calidad = data.get('riesgos_controles', [])
            all_risks = riesgos_calidad + riesgos_shes
            for shape in slide12.shapes:
                if shape.has_table:
                    table = shape.table
                    for i, risk in enumerate(all_risks):
                        row_idx = i + 1
                        if row_idx < len(table.rows):
                            if len(table.columns) > 0:
                                fill_table_cell(table.cell(row_idx, 0), str(i + 1))
                            if len(table.columns) > 1:
                                fill_table_cell(table.cell(row_idx, 1), risk.get('riesgo', ''))
                            if len(table.columns) > 2:
                                fill_table_cell(table.cell(row_idx, 2), risk.get('control', ''))
                            if len(table.columns) > 3:
                                fill_table_cell(table.cell(row_idx, 3), risk.get('plazo', ''))

        output_buffer = BytesIO()
        prs.save(output_buffer)
        output_buffer.seek(0)
        return output_buffer

    def generate_a3(self, data, images=None, template_bytes=None):
        if template_bytes is None:
            st.error("❌ Template A3 no cargado. Vaya a Configuración > Templates.")
            return None
        doc = Document(BytesIO(template_bytes))
        replacements = {
            "Autor:": f"Autor: {data.get('autor', '')}",
            "Miembros del equipo:": f"Miembros del equipo: {data.get('miembros_equipo', '')}",
        }
        for para in doc.paragraphs:
            for old, new in replacements.items():
                if old in para.text:
                    para.text = new
                    for run in para.runs:
                        run.font.name = 'Calibri'
                        run.font.size = DocxPt(11)
        doc.add_page_break()
        heading = doc.add_heading(data.get('titulo', 'Mejora A3'), level=1)
        for run in heading.runs:
            run.font.color.rgb = DocxRGBColor(0x1a, 0x5f, 0x7a)
            run.font.name = 'Calibri'
        sections = [
            ("ANTECEDENTES", "antecedentes"), ("PROBLEMA ACTUAL", "problema_actual"),
            ("ANÁLISIS DE LA SITUACIÓN", "analisis_situacion"), ("OBJETIVOS", "objetivos"),
            ("ANÁLISIS DE CAUSA RAÍZ", "analisis_causa_raiz"), ("CONTRAMEDIDAS", "contramedidas"),
            ("RESULTADOS ESPERADOS", "resultados_esperados"), ("PLAN DE SEGUIMIENTO", "plan_seguimiento"),
            ("LECCIONES APRENDIDAS", "lecciones_aprendidas"), ("ESTANDARIZACIÓN", "estandarizacion"),
        ]
        for section_title, key in sections:
            h = doc.add_heading(section_title, level=2)
            for run in h.runs:
                run.font.color.rgb = DocxRGBColor(0x1a, 0x5f, 0x7a)
                run.font.name = 'Calibri'
            content = data.get(key, '')
            if content:
                p = doc.add_paragraph(content)
                for run in p.runs:
                    run.font.name = 'Calibri'
                    run.font.size = DocxPt(11)
        if images:
            doc.add_page_break()
            doc.add_heading('IMÁGENES DE SOPORTE', level=1)
            for idx, img_info in enumerate(images, 1):
                doc.add_paragraph()
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run(f"Figura {idx}")
                run.bold = True
                run.font.size = DocxPt(12)
                run.font.color.rgb = DocxRGBColor(0x1a, 0x5f, 0x7a)
                try:
                    img_path = img_info["path"] if isinstance(img_info, dict) else img_info
                    doc.add_picture(img_path, width=DocxInches(5.5))
                    last_paragraph = doc.paragraphs[-1]
                    last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    desc = img_info.get("desc", f"Figura {idx}") if isinstance(img_info, dict) else f"Figura {idx}"
                    caption = doc.add_paragraph()
                    caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    cap_run = caption.add_run(desc)
                    cap_run.italic = True
                    cap_run.font.size = DocxPt(10)
                    cap_run.font.color.rgb = DocxRGBColor(0x47, 0x55, 0x69)
                except Exception as e:
                    doc.add_paragraph(f"[Error imagen {idx}: {e}]")
        output_buffer = BytesIO()
        doc.save(output_buffer)
        output_buffer.seek(0)
        return output_buffer

    def generate_kaizen(self, data, images=None, template_bytes=None):
        if template_bytes is None:
            st.error("❌ Template Kaizen no cargado. Vaya a Configuración > Templates.")
            return None
        prs = Presentation(BytesIO(template_bytes))
        if len(prs.slides) > 0:
            slide = prs.slides[0]
            for shape in slide.shapes:
                if shape.has_text_frame:
                    text = shape.text_frame.text
                    if "Name:" in text:
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                if "Name:" in run.text:
                                    run.text = f"Name: {data.get('titulo', '')}"
                    if "Plant/Area:" in text:
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                if "Plant/Area:" in run.text:
                                    run.text = f"Plant/Area: {data.get('area', '')}"
                    if "Date:" in text:
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                if "Date:" in run.text:
                                    run.text = f"Date: {data.get('fecha', '')}"
                    if "Leader" in text:
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                if "Leader" in run.text:
                                    run.text = f"Leader: {data.get('leader', '')}"
                    if "Opportunity" in text:
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                if "Opportunity" in run.text:
                                    run.text = f"Opportunity:\n{data.get('descripcion_problema', '')}"
                    if "Improvement" in text:
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                if "Improvement" in run.text:
                                    run.text = f"Improvement:\n{data.get('solucion', '')}"
                    if "Benefit" in text:
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                if "Benefit" in run.text:
                                    run.text = f"Benefit:\n{data.get('beneficios', '')}"
        output_buffer = BytesIO()
        prs.save(output_buffer)
        output_buffer.seek(0)
        return output_buffer

# =============================================================================
# EXPORTADOR A PDF
# =============================================================================
class PDFExporter:
    @staticmethod
    def pptx_to_pdf_libreoffice(pptx_bytes, output_filename):
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                tmpdir = Path(tmpdir)
                input_path = tmpdir / "input.pptx"
                with open(input_path, 'wb') as f:
                    f.write(pptx_bytes)
                result = subprocess.run(
                    ['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', str(tmpdir), str(input_path)],
                    capture_output=True, text=True, timeout=60
                )
                output_path = tmpdir / "input.pdf"
                if output_path.exists():
                    with open(output_path, 'rb') as f:
                        return f.read()
                result = subprocess.run(
                    ['soffice', '--headless', '--convert-to', 'pdf', '--outdir', str(tmpdir), str(input_path)],
                    capture_output=True, text=True, timeout=60
                )
                if output_path.exists():
                    with open(output_path, 'rb') as f:
                        return f.read()
        except Exception as e:
            st.warning(f"Conversión LibreOffice falló: {e}")
            return None

    @staticmethod
    def docx_to_pdf_libreoffice(docx_bytes, output_filename):
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                tmpdir = Path(tmpdir)
                input_path = tmpdir / "input.docx"
                with open(input_path, 'wb') as f:
                    f.write(docx_bytes)
                result = subprocess.run(
                    ['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', str(tmpdir), str(input_path)],
                    capture_output=True, text=True, timeout=60
                )
                output_path = tmpdir / "input.pdf"
                if output_path.exists():
                    with open(output_path, 'rb') as f:
                        return f.read()
                result = subprocess.run(
                    ['soffice', '--headless', '--convert-to', 'pdf', '--outdir', str(tmpdir), str(input_path)],
                    capture_output=True, text=True, timeout=60
                )
                if output_path.exists():
                    with open(output_path, 'rb') as f:
                        return f.read()
        except Exception as e:
            st.warning(f"Conversión LibreOffice falló: {e}")
            return None

    @staticmethod
    def generate_pdf_from_data(data, doc_type, meta, images=None):
        if not REPORTLAB_AVAILABLE:
            return None
        try:
            buffer = BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=A4,
                                    rightMargin=72, leftMargin=72,
                                    topMargin=72, bottomMargin=18)
            styles = getSampleStyleSheet()
            story = []
            title_style = ParagraphStyle(
                'CustomTitle', parent=styles['Heading1'],
                fontSize=20, textColor=colors.HexColor('#1a5f7a'),
                spaceAfter=30, alignment=TA_CENTER, fontName='Helvetica-Bold'
            )
            heading_style = ParagraphStyle(
                'CustomHeading', parent=styles['Heading2'],
                fontSize=14, textColor=colors.HexColor('#1a5f7a'),
                spaceAfter=12, spaceBefore=12, fontName='Helvetica-Bold'
            )
            body_style = ParagraphStyle(
                'CustomBody', parent=styles['BodyText'],
                fontSize=10, leading=14, alignment=TA_JUSTIFY, fontName='Helvetica'
            )
            type_names = {"moc": "Management of Change (MoC)", "a3": "Mejora A3", "kaizen": "Simple Kaizen"}
            doc_title = type_names.get(doc_type, "Documento")
            story.append(Paragraph(f"<b>{doc_title}</b>", title_style))
            story.append(Spacer(1, 20))
            meta_text = ""
            for key, value in meta.items():
                if value and key not in ['id', 'timestamp']:
                    meta_text += f"<b>{key.replace('_', ' ').title()}:</b> {value}<br/>"
            if meta_text:
                story.append(Paragraph(meta_text, body_style))
            story.append(Spacer(1, 20))
            if doc_type == "moc":
                sections = [
                    ("Descripción del Problema", "descripcion_problema"),
                    ("Condición Actual", "condicion_actual"),
                    ("Condición Propuesta", "condicion_propuesta"),
                    ("Razones del Cambio", "razones_cambio"),
                    ("Alternativas y Plan de Retorno", "alternativas_retorno"),
                    ("Recursos", "recursos"),
                    ("Plan de Implementación", "plan_implementacion"),
                    ("Tiempo de Duración", "tiempo_duracion"),
                ]
                for title, key in sections:
                    story.append(Paragraph(f"<b>{title}</b>", heading_style))
                    content = data.get(key, '').replace('\n', '<br/>')
                    story.append(Paragraph(content, body_style))
                    story.append(Spacer(1, 10))
            elif doc_type == "a3":
                sections = [
                    ("Antecedentes", "antecedentes"), ("Problema Actual", "problema_actual"),
                    ("Análisis de la Situación", "analisis_situacion"), ("Objetivos", "objetivos"),
                    ("Análisis de Causa Raíz", "analisis_causa_raiz"), ("Contramedidas", "contramedidas"),
                    ("Resultados Esperados", "resultados_esperados"), ("Plan de Seguimiento", "plan_seguimiento"),
                    ("Lecciones Aprendidas", "lecciones_aprendidas"), ("Estandarización", "estandarizacion"),
                ]
                for title, key in sections:
                    story.append(Paragraph(f"<b>{title}</b>", heading_style))
                    content = data.get(key, '').replace('\n', '<br/>')
                    story.append(Paragraph(content, body_style))
                    story.append(Spacer(1, 10))
            elif doc_type == "kaizen":
                sections = [
                    ("Descripción del Problema", "descripcion_problema"),
                    ("Solución Implementada", "solucion"),
                    ("Beneficios", "beneficios"),
                    ("Próximos Pasos", "proximos_pasos"),
                ]
                for title, key in sections:
                    story.append(Paragraph(f"<b>{title}</b>", heading_style))
                    content = data.get(key, '').replace('\n', '<br/>')
                    story.append(Paragraph(content, body_style))
                    story.append(Spacer(1, 10))
            doc.build(story)
            buffer.seek(0)
            return buffer.getvalue()
        except Exception as e:
            st.error(f"Error generando PDF con ReportLab: {e}")
            return None

# =============================================================================
# INICIALIZACION DE SESSION STATE
# =============================================================================
def init_session_state():
    saved_config = LocalStorage.load_config()
    saved_history = LocalStorage.load_history()
    defaults = {
        "page": "inicio",
        "config": saved_config or {
            "gemini_api_key": "",
            "gemini_model": "gemini-1.5-pro",
            "company_name": "",
            "department": "",
            "default_author": "",
            "default_area": "",
            "header_text": "",
            "footer_text": "",
            "last_moc_number": 0,
            "last_a3_number": 0,
            "last_kaizen_number": 0,
            "spell_check": True,
            "thinking_level": "Estándar",
            "auto_correct": True,
        },
        "history": saved_history or {"documents": []},
        "generated_data": {},
        "doc_meta": {},
        "doc_images": [],
        "doc_type": None,
        "templates_uploaded": False,
        "template_moc_bytes": LocalStorage.load_template_bytes("moc"),
        "template_a3_bytes": LocalStorage.load_template_bytes("a3"),
        "template_kaizen_bytes": LocalStorage.load_template_bytes("kaizen"),
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value
    if saved_config is None:
        LocalStorage.save_config(st.session_state.config)
    if saved_history is None:
        LocalStorage.save_history(st.session_state.history)

init_session_state()

# =============================================================================
# INTERFAZ DE USUARIO
# =============================================================================
def render_sidebar():
    config = st.session_state.config
    st.sidebar.markdown("""
<div style="text-align: center; padding: 1rem 0;">
<h2 style="color: #f8fafc; margin: 0; font-size: 1.3rem;">📋 GESTIÓN</h2>
<h2 style="color: #f8fafc; margin: 0; font-size: 1.3rem;">DOCUMENTAL</h2>
<p style="color: #94a3b8; margin-top: 0.5rem; font-size: 0.85rem;">MoC · A3 · Kaizen</p>
</div>
<hr style="border-color: #334155; margin: 1rem 0;">
""", unsafe_allow_html=True)
    st.sidebar.markdown("### 🧭 Navegación")
    nav_items = [
        ("🏠 Inicio", "inicio"),
        ("📋 Nueva MoC", "nueva_moc"),
        ("📊 Nueva Mejora A3", "nueva_a3"),
        ("⚡ Nuevo Kaizen", "nuevo_kaizen"),
        ("📁 Historial", "historial"),
        ("⚙️ Configuración", "configuracion"),
    ]
    for label, page_key in nav_items:
        if st.sidebar.button(label, key=f"nav_{page_key}", use_container_width=True):
            st.session_state.page = page_key
            st.rerun()
    st.sidebar.markdown("<hr style='border-color: #334155; margin: 1rem 0;'>", unsafe_allow_html=True)
    model_name = GeminiService.MODELS.get(config.get("gemini_model", "gemini-1.5-pro"), {}).get("name", "Gemini 1.5 Pro")
    st.sidebar.markdown(f"""
<div style="text-align: center; color: #64748b; font-size: 0.75rem;">
<p>Modelo IA: <span class="gemini-badge">{model_name}</span></p>
<p>v7.5.0 · Agosto 2026</p>
</div>
""", unsafe_allow_html=True)
    st.sidebar.markdown("""
<hr style="border-color: #334155; margin: 1rem 0;">
<div style="text-align: center; padding: 0.5rem;">
<p style="color: #64748b; font-size: 0.75rem; margin: 0;">
<strong style="color: #94a3b8;">CAVA</strong><br>
Especialistas en Robótica<br>
y Automatización<br><br>
Diseñado por<br>
<strong style="color: #2e8bc0;">Roger Huamani</strong>
</p>
</div>
""", unsafe_allow_html=True)

def render_header():
    st.markdown("""
<div class="main-header">
<h1>🎯 Sistema de Gestión Documental</h1>
<p>Automatización inteligente de documentos MoC, Mejora A3 y Simple Kaizen con IA</p>
</div>
""", unsafe_allow_html=True)

def render_welcome():
    render_header()
    templates_ok = all([
        st.session_state.get("template_moc_bytes"),
        st.session_state.get("template_a3_bytes"),
        st.session_state.get("template_kaizen_bytes")
    ])
    if not templates_ok:
        st.warning("⚠️ **Templates no cargados.** Vaya a Configuración > Templates para subir los formatos oficiales.")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("""
<div class="doc-card doc-card-moc">
<h3 style="color: #1a5f7a; margin-top: 0;">📋 Management of Change</h3>
<p style="color: #64748b; font-size: 0.9rem;">Formato oficial MDET con análisis integral.</p>
<ul style="color: #475569; font-size: 0.85rem; padding-left: 1.2rem;">
<li>12 slides estandarizados</li><li>Checklist 360° automático</li><li>15 documentos impactados</li><li>Riesgos SHES detallados</li>
</ul>
</div>
""", unsafe_allow_html=True)
        if st.button("➕ Crear MoC", key="btn_moc", use_container_width=True):
            st.session_state.page = "nueva_moc"
            st.rerun()
    with col2:
        st.markdown("""
<div class="doc-card doc-card-a3">
<h3 style="color: #10b981; margin-top: 0;">📊 Mejora A3</h3>
<p style="color: #64748b; font-size: 0.9rem;">Formato estructurado con análisis 5 Porqués y contramedidas SMART.</p>
<ul style="color: #475569; font-size: 0.85rem; padding-left: 1.2rem;">
<li>Análisis 5 Porqués</li><li>Contramedidas priorizadas</li><li>Plan de seguimiento</li><li>Estandarización</li>
</ul>
</div>
""", unsafe_allow_html=True)
        if st.button("➕ Crear A3", key="btn_a3", use_container_width=True):
            st.session_state.page = "nueva_a3"
            st.rerun()
    with col3:
        st.markdown("""
<div class="doc-card doc-card-kaizen">
<h3 style="color: #f59e0b; margin-top: 0;">⚡ Simple Kaizen</h3>
<p style="color: #64748b; font-size: 0.9rem;">Registro rápido de mejoras con clasificación de desperdicios Lean.</p>
<ul style="color: #475569; font-size: 0.85rem; padding-left: 1.2rem;">
<li>8 Desperdicios (Wastes)</li><li>Impacto BTO</li><li>Beneficios medibles</li><li>Replicabilidad</li>
</ul>
</div>
""", unsafe_allow_html=True)
        if st.button("➕ Crear Kaizen", key="btn_kaizen", use_container_width=True):
            st.session_state.page = "nuevo_kaizen"
            st.rerun()
    st.markdown("<br>", unsafe_allow_html=True)
    docs = st.session_state.history.get("documents", [])
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("📋 MoC", len([d for d in docs if d.get("type") == "moc"]))
    with col2:
        st.metric("📊 A3", len([d for d in docs if d.get("type") == "a3"]))
    with col3:
        st.metric("⚡ Kaizen", len([d for d in docs if d.get("type") == "kaizen"]))
    with col4:
        st.metric("📁 Total", len(docs))

def auto_correct_text_input(label, value, key, height=100, help_text=""):
    config = st.session_state.config
    if config.get("auto_correct", True):
        st.markdown('<div class="auto-correct-badge">✨ Corrección ortográfica automática activa</div>', unsafe_allow_html=True)
    text = st.text_area(label, value=value, height=height, key=key, help=help_text)
    if config.get("auto_correct", True) and text.strip():
        corrected = Utils.correct_spelling_basic(text)
        if corrected != text:
            st.info(f"🔧 Texto corregido automáticamente")
            return corrected
    return text

def render_moc_form():
    config = st.session_state.config
    st.markdown('<div class="section-header"><h3>📋 Nueva Management of Change (MoC)</h3></div>', unsafe_allow_html=True)
    st.info("💡 Complete la información y describa el problema con detalle. La IA generará automáticamente los slides basándose EXCLUSIVAMENTE en su problema específico.")
    if not st.session_state.get("template_moc_bytes"):
        st.error("❌ **Template MoC no cargado.** Vaya a Configuración > Templates.")
        if st.button("Ir a Configuración", key="go_config_moc"):
            st.session_state.page = "configuracion"
            st.rerun()
        return
    st.markdown("#### 1. Información General")
    col1, col2 = st.columns(2)
    with col1:
        moc_title = st.text_input("Título de la MoC:", placeholder="Ej: INSTALACIÓN DE INTERLOCKS DE SEGURIDAD EN ESTACIONES DE ESPERA")
        moc_number = st.text_input("Número:", value=Utils.generate_doc_number("moc"), disabled=True)
    with col2:
        naturaleza = st.selectbox("Naturaleza:", ["permanente", "temporal", "emergencia"])
        originador = st.text_input("Originador:", value=config.get("default_author", ""))
        fecha = st.text_input("Fecha:", value=Utils.format_date_short(), disabled=True)
    st.markdown("#### 2. Equipo de Revisión")
    col1, col2, col3 = st.columns(3)
    with col1:
        produccion = st.text_input("Producción:")
        specialist_shes = st.text_input("Specialist SHES:")
    with col2:
        mantenimiento = st.text_input("Mantenimiento:")
        revisores = st.text_input("Revisores Enablon:")
    with col3:
        experto_aprobador = st.text_input("Experto Aprobador:")
    st.markdown("#### 3. Descripción del Problema/Cambio (SEA LO MÁS DETALLADO POSIBLE)")
    st.warning("⚠️ **IMPORTANTE:** Describa el problema con el mayor detalle posible. La IA usará EXCLUSIVAMENTE esta información. Incluya: equipos específicos, componentes, riesgos, normas aplicables, solución propuesta.")
    problem_desc = auto_correct_text_input(
        "Describa el problema o cambio con sus palabras:",
        "",
        "moc_problem_desc",
        height=300,
        help_text="Ejemplo: Actualmente las estaciones de espera no tienen interlocks de seguridad. Los operadores pueden abrir las compuertas con la máquina en funcionamiento, exponiéndose a riesgos de atrapamiento. Se propone instalar interlocks que detengan la máquina al abrir las compuertas, integrados al PLC según norma ISO 13849."
    )
    st.markdown("#### 4. Contexto Adicional (Opcional pero recomendado)")
    context = auto_correct_text_input(
        "Información adicional:",
        "",
        "moc_context",
        height=120,
        help_text="TAG del equipo, área específica, normativas aplicables, fechas relevantes, datos numéricos, planos de referencia, etc."
    )
    st.markdown("#### 5. Imágenes de Soporte (Opcional)")
    uploaded_images = st.file_uploader("Seleccione imágenes (Vista general, planos, etc.):", type=["png", "jpg", "jpeg"], accept_multiple_files=True)
    image_paths = []
    if uploaded_images:
        for idx, img_file in enumerate(uploaded_images, 1):
            img_path = f"/tmp/temp_moc_img_{moc_number}_{idx}.png"
            with open(img_path, "wb") as f:
                f.write(img_file.getbuffer())
            image_paths.append({"path": img_path, "desc": f"Figura {idx} - {img_file.name}"})
        st.success(f"📷 {len(image_paths)} imagen(es) cargada(s)")
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🤖 Generar Documento MoC con IA", type="primary", use_container_width=True):
        if not problem_desc.strip():
            st.error("❌ Describa el problema antes de generar.")
            return
        with st.spinner("🧠 La IA está generando el documento basándose en su problema específico..."):
            gemini = GeminiService(config.get("gemini_api_key", ""), config.get("gemini_model", "gemini-1.5-pro"))
            equipo_data = {
                "produccion": produccion, "specialist_shes": specialist_shes,
                "mantenimiento": mantenimiento, "revisores": revisores,
                "experto_aprobador": experto_aprobador
            }
            result = gemini.generate_moc(problem_desc, context, json.dumps(equipo_data))
            if result is None:
                st.error("❌ No se pudo generar el documento. Verifique su API Key en Configuración.")
                return
            st.session_state.generated_data = result
            st.session_state.doc_meta = {
                "moc_title": moc_title, "moc_number": moc_number, "naturaleza": naturaleza,
                "originador": originador, "fecha": fecha, **equipo_data
            }
            st.session_state.doc_images = image_paths
            st.session_state.doc_type = "moc"
            st.session_state.page = "revisar"
            st.rerun()

def render_a3_form():
    config = st.session_state.config
    st.markdown('<div class="section-header"><h3>📊 Nueva Mejora A3</h3></div>', unsafe_allow_html=True)
    st.info("💡 Describa el problema con detalle y la IA generará el documento A3 completo basado en su problema específico.")
    if not st.session_state.get("template_a3_bytes"):
        st.error("❌ **Template A3 no cargado.** Vaya a Configuración > Templates.")
        if st.button("Ir a Configuración", key="go_config_a3"):
            st.session_state.page = "configuracion"
            st.rerun()
        return
    st.markdown("#### 1. Información General")
    col1, col2 = st.columns(2)
    with col1:
        a3_title = st.text_input("Título:", placeholder="Ej: Reducción de tiempo de cambio de formato")
        area = st.text_input("Área:", value=config.get("default_area", ""))
    with col2:
        autor = st.text_input("Autor:", value=config.get("default_author", ""))
        doc_number = st.text_input("Número:", value=Utils.generate_doc_number("a3"), disabled=True)
        fecha = st.text_input("Fecha:", value=Utils.format_date(), disabled=True)
    st.markdown("#### 2. Descripción del Problema (Detallada)")
    problem_desc = auto_correct_text_input(
        "Describa el problema actual:",
        "",
        "a3_problem_desc",
        height=250,
        help_text="¿Qué está pasando? ¿Impacto? ¿Desde cuándo? ¿Datos cuantitativos? ¿Frecuencia?"
    )
    context = auto_correct_text_input(
        "Contexto adicional:",
        "",
        "a3_context",
        height=100,
        help_text="Herramientas Lean aplicables, benchmarks, áreas relacionadas, etc."
    )
    st.markdown("#### 3. Imágenes de Soporte")
    uploaded_images = st.file_uploader("Seleccione imágenes:", type=["png", "jpg", "jpeg"], accept_multiple_files=True)
    image_paths = []
    if uploaded_images:
        for idx, img_file in enumerate(uploaded_images, 1):
            img_path = f"/tmp/temp_a3_img_{doc_number}_{idx}.png"
            with open(img_path, "wb") as f:
                f.write(img_file.getbuffer())
            image_paths.append({"path": img_path, "desc": f"Figura {idx} - {img_file.name}"})
        st.success(f"📷 {len(image_paths)} imagen(es) cargada(s)")
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🤖 Generar Documento A3 con IA", type="primary", use_container_width=True):
        if not problem_desc.strip():
            st.error("❌ Describa el problema antes de generar.")
            return
        with st.spinner("🧠 Generando documento A3 con análisis detallado..."):
            gemini = GeminiService(config.get("gemini_api_key", ""), config.get("gemini_model", "gemini-1.5-pro"))
            result = gemini.generate_a3(problem_desc, context)
            if result is None:
                st.error("❌ No se pudo generar el documento. Verifique su API Key en Configuración.")
                return
            st.session_state.generated_data = result
            st.session_state.doc_meta = {"titulo": a3_title, "area": area, "autor": autor, "doc_number": doc_number, "fecha": fecha}
            st.session_state.doc_images = image_paths
            st.session_state.doc_type = "a3"
            st.session_state.page = "revisar"
            st.rerun()

def render_kaizen_form():
    config = st.session_state.config
    st.markdown('<div class="section-header"><h3>⚡ Nuevo Simple Kaizen</h3></div>', unsafe_allow_html=True)
    st.info("💡 Describa la actividad de mejora realizada con detalle.")
    if not st.session_state.get("template_kaizen_bytes"):
        st.error("❌ **Template Kaizen no cargado.** Vaya a Configuración > Templates.")
        if st.button("Ir a Configuración", key="go_config_kzn"):
            st.session_state.page = "configuracion"
            st.rerun()
        return
    st.markdown("#### 1. Información General")
    col1, col2 = st.columns(2)
    with col1:
        kaizen_title = st.text_input("Título (Name):", placeholder="Ej: Organización de área de herramientas")
        area = st.text_input("Plant/Area:", value=config.get("default_area", ""))
    with col2:
        leader = st.text_input("Leader:", value=config.get("default_author", ""))
        doc_number = st.text_input("Número:", value=Utils.generate_doc_number("kaizen"), disabled=True)
        fecha = st.text_input("Date:", value=Utils.format_date(), disabled=True)
        team_members = st.text_input("Team Members (separados por coma):", placeholder="Ej: Juan Pérez, María García")
    st.markdown("#### 2. Descripción de la Actividad (Detallada)")
    activity_desc = auto_correct_text_input(
        "Describa la mejora realizada:",
        "",
        "kzn_activity_desc",
        height=250,
        help_text="Describa el antes, durante y después. Incluya datos de tiempo, movimientos, cantidades."
    )
    st.markdown("#### 3. Clasificación")
    col1, col2 = st.columns(2)
    with col1:
        tipo_desp = st.multiselect("Tipo de Desperdicio eliminado:",
                                   ["Motion", "Skills", "Inventory", "Transportation",
                                    "Over Production", "Over Processing", "Waiting", "Defects"])
    with col2:
        impacto_bto = st.selectbox("Impacto BTO:",
                                   ["Safe and Sustainable", "People & Culture",
                                    "Network Optimisation", "Supply Chain and Manufacturing Excellence"])
    st.markdown("#### 4. Beneficios")
    beneficios = auto_correct_text_input(
        "Describa los beneficios obtenidos:",
        "",
        "kzn_beneficios",
        height=120,
        help_text="Incluya datos cuantitativos: tiempos antes/después, porcentajes, ahorros."
    )
    st.markdown("#### 5. Imágenes Antes/Después")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**Imagen ANTES:**")
        img_antes = st.file_uploader("Subir imagen ANTES:", type=["png", "jpg", "jpeg"], key="img_antes")
    with col2:
        st.markdown("**Imagen DESPUÉS:**")
        img_despues = st.file_uploader("Subir imagen DESPUÉS:", type=["png", "jpg", "jpeg"], key="img_despues")
    image_paths = []
    if img_antes:
        img_path = f"/tmp/temp_kzn_antes_{doc_number}.png"
        with open(img_path, "wb") as f:
            f.write(img_antes.getbuffer())
        image_paths.append({"path": img_path, "desc": "ANTES - Estado inicial"})
    if img_despues:
        img_path = f"/tmp/temp_kzn_despues_{doc_number}.png"
        with open(img_path, "wb") as f:
            f.write(img_despues.getbuffer())
        image_paths.append({"path": img_path, "desc": "DESPUÉS - Estado final"})
    context = auto_correct_text_input(
        "Contexto adicional:",
        "",
        "kzn_context",
        height=80,
        help_text="Costos, tiempos medidos, materiales utilizados, etc."
    )
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🤖 Generar Documento Kaizen con IA", type="primary", use_container_width=True):
        if not activity_desc.strip():
            st.error("❌ Describa la actividad antes de generar.")
            return
        with st.spinner("🧠 Generando documento Kaizen..."):
            gemini = GeminiService(config.get("gemini_api_key", ""), config.get("gemini_model", "gemini-1.5-pro"))
            result = gemini.generate_kaizen(activity_desc, context)
            if result is None:
                st.error("❌ No se pudo generar el documento. Verifique su API Key en Configuración.")
                return
            result["tipo_desperdicio"] = ", ".join(tipo_desp) if tipo_desp else result.get("tipo_desperdicio", "")
            result["impacto_bto"] = impacto_bto
            result["leader"] = leader
            result["team_members"] = team_members
            result["beneficios"] = beneficios if beneficios else result.get("beneficios", "")
            st.session_state.generated_data = result
            st.session_state.doc_meta = {
                "titulo": kaizen_title, "area": area, "leader": leader,
                "team_members": team_members, "doc_number": doc_number, "fecha": fecha
            }
            st.session_state.doc_images = image_paths
            st.session_state.doc_type = "kaizen"
            st.session_state.page = "revisar"
            st.rerun()

def render_review():
    doc_type = st.session_state.doc_type
    data = st.session_state.get("generated_data", {})
    meta = st.session_state.get("doc_meta", {})
    images = st.session_state.get("doc_images", [])
    config = st.session_state.config
    type_names = {"moc": "MoC", "a3": "Mejora A3", "kaizen": "Simple Kaizen"}
    type_name = type_names.get(doc_type, "Documento")
    st.markdown(f'<div class="section-header"><h3>👁️ Revisar y Editar {type_name}</h3></div>', unsafe_allow_html=True)
    st.info("💡 Revise cada campo, realice correcciones manuales si es necesario y genere el documento final.")
    if doc_type == "moc":
        _render_moc_review(data, meta, images, config)
    elif doc_type == "a3":
        _render_a3_review(data, meta, images, config)
    elif doc_type == "kaizen":
        _render_kaizen_review(data, meta, images, config)

def _spell_check_field(label, value, key_prefix, gemini):
    col1, col2 = st.columns([6, 1])
    with col1:
        text = st.text_area(label, value=value, height=120, key=f"{key_prefix}_field")
    with col2:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("✨ Corregir", key=f"{key_prefix}_spell"):
            with st.spinner("Corrigiendo..."):
                corrected = gemini.correct_spelling(text)
                st.session_state[f"{key_prefix}_corrected"] = corrected
                st.rerun()
    if st.session_state.get(f"{key_prefix}_corrected"):
        text = st.session_state[f"{key_prefix}_corrected"]
        del st.session_state[f"{key_prefix}_corrected"]
    return text

def _render_moc_review(data, meta, images, config):
    gemini = GeminiService(config.get("gemini_api_key", ""), config.get("gemini_model", "gemini-1.5-pro"))
    tabs = st.tabs(["📋 General", "📝 Contenido", "📊 Riesgos", "📷 Imágenes", "⚙️ Generar"])
    with tabs[0]:
        st.markdown("#### Información del Documento")
        meta["moc_title"] = st.text_input("Título:", value=meta.get("moc_title", ""), key="moc_rev_title")
        meta["moc_number"] = st.text_input("Número:", value=meta.get("moc_number", ""), disabled=True)
        meta["naturaleza"] = st.selectbox("Naturaleza:", ["permanente", "temporal", "emergencia"],
                                          index=["permanente", "temporal", "emergencia"].index(meta.get("naturaleza", "permanente")), key="moc_rev_nat")
        meta["originador"] = st.text_input("Originador:", value=meta.get("originador", ""), key="moc_rev_orig")
        st.markdown("#### Equipo de Revisión")
        meta["produccion"] = st.text_input("Producción:", value=meta.get("produccion", ""), key="moc_rev_prod")
        meta["specialist_shes"] = st.text_input("Specialist SHES:", value=meta.get("specialist_shes", ""), key="moc_rev_shes")
        meta["mantenimiento"] = st.text_input("Mantenimiento:", value=meta.get("mantenimiento", ""), key="moc_rev_mant")
        meta["revisores"] = st.text_input("Revisores:", value=meta.get("revisores", ""), key="moc_rev_rev")
        meta["experto_aprobador"] = st.text_input("Experto Aprobador:", value=meta.get("experto_aprobador", ""), key="moc_rev_exp")
    with tabs[1]:
        st.markdown("#### Descripción del Problema")
        data["descripcion_problema"] = _spell_check_field("", data.get("descripcion_problema", ""), "moc_desc", gemini)
        st.markdown("#### Condición Actual")
        data["condicion_actual"] = _spell_check_field("", data.get("condicion_actual", ""), "moc_actual", gemini)
        st.markdown("#### Condición Propuesta")
        data["condicion_propuesta"] = _spell_check_field("", data.get("condicion_propuesta", ""), "moc_prop", gemini)
        st.markdown("#### Razones del Cambio")
        data["razones_cambio"] = _spell_check_field("", data.get("razones_cambio", ""), "moc_raz", gemini)
        st.markdown("#### Alternativas y Plan de Retorno")
        data["alternativas_retorno"] = _spell_check_field("", data.get("alternativas_retorno", ""), "moc_alt", gemini)
        st.markdown("#### Recursos")
        data["recursos"] = _spell_check_field("", data.get("recursos", ""), "moc_rec", gemini)
        st.markdown("#### Plan de Implementación")
        data["plan_implementacion"] = _spell_check_field("", data.get("plan_implementacion", ""), "moc_plan", gemini)
        st.markdown("#### Tiempo de Duración")
        data["tiempo_duracion"] = _spell_check_field("", data.get("tiempo_duracion", ""), "moc_tiempo", gemini)
    with tabs[2]:
        st.markdown("#### Riesgos y Controles")
        risks = data.get("riesgos_controles", [])
        updated_risks = []
        for i, risk in enumerate(risks):
            st.markdown(f"**Riesgo {i+1}**")
            col1, col2 = st.columns(2)
            with col1:
                r_riesgo = st.text_input(f"Riesgo {i+1}:", value=risk.get("riesgo", ""), key=f"risk_{i}")
            with col2:
                r_control = st.text_input(f"Control {i+1}:", value=risk.get("control", ""), key=f"ctrl_{i}")
            updated_risks.append({"riesgo": r_riesgo, "control": r_control})
        if st.button("➕ Agregar Riesgo", key="add_risk"):
            updated_risks.append({"riesgo": "", "control": ""})
        data["riesgos_controles"] = updated_risks
        st.markdown("#### Riesgos SHES")
        risks_shes = data.get("riesgos_shes", [])
        updated_shes = []
        for i, risk in enumerate(risks_shes):
            st.markdown(f"**Riesgo SHES {i+1}**")
            col1, col2, col3 = st.columns(3)
            with col1:
                s_riesgo = st.text_input(f"Riesgo S{i+1}:", value=risk.get("riesgo", ""), key=f"shes_r_{i}")
            with col2:
                s_control = st.text_input(f"Control S{i+1}:", value=risk.get("control", ""), key=f"shes_c_{i}")
            with col3:
                s_plazo = st.text_input(f"Plazo S{i+1}:", value=risk.get("plazo", ""), key=f"shes_p_{i}")
            updated_shes.append({"riesgo": s_riesgo, "control": s_control, "plazo": s_plazo})
        if st.button("➕ Agregar Riesgo SHES", key="add_shes"):
            updated_shes.append({"riesgo": "", "control": "", "plazo": ""})
        data["riesgos_shes"] = updated_shes
    with tabs[3]:
        st.markdown("#### Imágenes Cargadas")
        if images:
            for idx, img_info in enumerate(images, 1):
                st.image(img_info["path"], caption=f"Figura {idx}: {img_info['desc']}", width=400)
        else:
            st.info("No se cargaron imágenes")
    with tabs[4]:
        st.markdown("#### Generar Documento Final")
        st.success("✅ Documento listo para generar")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            if st.button("🇪🇸 PPTX Español", type="primary", use_container_width=True):
                _finalize_document(data, meta, images, "es", "moc", "pptx")
        with col2:
            if st.button("🇺🇸 PPTX Inglés", type="primary", use_container_width=True):
                _finalize_document(data, meta, images, "en", "moc", "pptx")
        with col3:
            if st.button("📄 PDF Español", type="secondary", use_container_width=True):
                _finalize_document(data, meta, images, "es", "moc", "pdf")
        with col4:
            if st.button("🔄 Regenerar", use_container_width=True):
                st.session_state.page = "nueva_moc"
                st.rerun()
    st.session_state.generated_data = data
    st.session_state.doc_meta = meta

def _render_a3_review(data, meta, images, config):
    gemini = GeminiService(config.get("gemini_api_key", ""), config.get("gemini_model", "gemini-1.5-pro"))
    tabs = st.tabs(["📋 General", "📝 Contenido", "📷 Imágenes", "⚙️ Generar"])
    with tabs[0]:
        meta["titulo"] = st.text_input("Título:", value=meta.get("titulo", ""), key="a3_rev_title")
        meta["area"] = st.text_input("Área:", value=meta.get("area", ""), key="a3_rev_area")
        meta["autor"] = st.text_input("Autor:", value=meta.get("autor", ""), key="a3_rev_autor")
        meta["doc_number"] = st.text_input("Número:", value=meta.get("doc_number", ""), disabled=True)
        meta["fecha"] = st.text_input("Fecha:", value=meta.get("fecha", ""), disabled=True)
    with tabs[1]:
        sections = [
            ("Antecedentes", "antecedentes"), ("Problema Actual", "problema_actual"),
            ("Análisis de Situación", "analisis_situacion"), ("Objetivos", "objetivos"),
            ("Análisis Causa Raíz", "analisis_causa_raiz"), ("Contramedidas", "contramedidas"),
            ("Resultados Esperados", "resultados_esperados"), ("Plan de Seguimiento", "plan_seguimiento"),
            ("Lecciones Aprendidas", "lecciones_aprendidas"), ("Estandarización", "estandarizacion"),
        ]
        for label, key in sections:
            st.markdown(f"**{label}**")
            data[key] = _spell_check_field("", data.get(key, ""), f"a3_{key}", gemini)
    with tabs[2]:
        if images:
            for idx, img_info in enumerate(images, 1):
                st.image(img_info["path"], caption=f"Figura {idx}: {img_info['desc']}", width=400)
        else:
            st.info("No se cargaron imágenes")
    with tabs[3]:
        col1, col2, col3 = st.columns(3)
        with col1:
            if st.button("✅ DOCX Español", type="primary", use_container_width=True):
                _finalize_document(data, meta, images, "es", "a3", "docx")
        with col2:
            if st.button("📄 PDF Español", type="secondary", use_container_width=True):
                _finalize_document(data, meta, images, "es", "a3", "pdf")
        with col3:
            if st.button("🔄 Regenerar", use_container_width=True):
                st.session_state.page = "nueva_a3"
                st.rerun()
    st.session_state.generated_data = data
    st.session_state.doc_meta = meta

def _render_kaizen_review(data, meta, images, config):
    gemini = GeminiService(config.get("gemini_api_key", ""), config.get("gemini_model", "gemini-1.5-pro"))
    tabs = st.tabs(["📋 General", "📝 Contenido", "📷 Imágenes", "⚙️ Generar"])
    with tabs[0]:
        meta["titulo"] = st.text_input("Título (Name):", value=meta.get("titulo", ""), key="kzn_rev_title")
        meta["area"] = st.text_input("Plant/Area:", value=meta.get("area", ""), key="kzn_rev_area")
        meta["leader"] = st.text_input("Leader:", value=meta.get("leader", ""), key="kzn_rev_leader")
        meta["team_members"] = st.text_input("Team Members:", value=meta.get("team_members", ""), key="kzn_rev_team")
        meta["doc_number"] = st.text_input("Número:", value=meta.get("doc_number", ""), disabled=True)
        meta["fecha"] = st.text_input("Date:", value=meta.get("fecha", ""), disabled=True)
    with tabs[1]:
        st.markdown("**Descripción del Problema (Opportunity)**")
        data["descripcion_problema"] = _spell_check_field("", data.get("descripcion_problema", ""), "kzn_desc", gemini)
        st.markdown("**Solución Implementada (Improvement)**")
        data["solucion"] = _spell_check_field("", data.get("solucion", ""), "kzn_sol", gemini)
        st.markdown("**Beneficios (Benefit)**")
        data["beneficios"] = _spell_check_field("", data.get("beneficios", ""), "kzn_ben", gemini)
        st.markdown("**Tipo de Desperdicio**")
        data["tipo_desperdicio"] = st.text_input("", value=data.get("tipo_desperdicio", ""), key="kzn_desp")
        st.markdown("**Impacto BTO**")
        data["impacto_bto"] = st.text_input("", value=data.get("impacto_bto", ""), key="kzn_bto")
        st.markdown("**Próximos Pasos**")
        data["proximos_pasos"] = _spell_check_field("", data.get("proximos_pasos", ""), "kzn_next", gemini)
    with tabs[2]:
        st.markdown("#### Imágenes Cargadas")
        if images:
            for idx, img_info in enumerate(images, 1):
                st.image(img_info["path"], caption=f"{img_info['desc']}", width=400)
        else:
            st.info("No se cargaron imágenes")
    with tabs[3]:
        col1, col2, col3 = st.columns(3)
        with col1:
            if st.button("✅ PPTX Español", type="primary", use_container_width=True):
                _finalize_document(data, meta, images, "es", "kaizen", "pptx")
        with col2:
            if st.button("📄 PDF Español", type="secondary", use_container_width=True):
                _finalize_document(data, meta, images, "es", "kaizen", "pdf")
        with col3:
            if st.button("🔄 Regenerar", use_container_width=True):
                st.session_state.page = "nuevo_kaizen"
                st.rerun()
    st.session_state.generated_data = data
    st.session_state.doc_meta = meta

def _finalize_document(data, meta, images, language, doc_type, output_format="pptx"):
    config = st.session_state.config
    gemini = GeminiService(config.get("gemini_api_key", ""), config.get("gemini_model", "gemini-1.5-pro"))
    with st.spinner(f"📄 Generando documento..."):
        final_data = {**meta, **data}
        if language == "en" and doc_type == "moc":
            st.info("🌐 Traduciendo documento al inglés...")
            final_data = gemini.translate_document(final_data)
        generator = DocumentGenerator()
        pdf_exporter = PDFExporter()
        if doc_type == "moc":
            if output_format == "pdf":
                pptx_buffer = generator.generate_moc(final_data, images, st.session_state.get("template_moc_bytes"))
                if pptx_buffer:
                    pdf_bytes = pdf_exporter.pptx_to_pdf_libreoffice(pptx_buffer.getvalue(), "moc.pdf")
                    if pdf_bytes:
                        buffer = BytesIO(pdf_bytes)
                        ext = "pdf"
                        mime = "application/pdf"
                    else:
                        st.warning("No se pudo generar PDF. Descargando PPTX.")
                        buffer = pptx_buffer
                        ext = "pptx"
                        mime = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
                else:
                    return
            else:
                buffer = generator.generate_moc(final_data, images, st.session_state.get("template_moc_bytes"))
                ext = "pptx"
                mime = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        elif doc_type == "a3":
            if output_format == "pdf":
                docx_buffer = generator.generate_a3(final_data, images, st.session_state.get("template_a3_bytes"))
                if docx_buffer:
                    pdf_bytes = pdf_exporter.docx_to_pdf_libreoffice(docx_buffer.getvalue(), "a3.pdf")
                    if pdf_bytes:
                        buffer = BytesIO(pdf_bytes)
                        ext = "pdf"
                        mime = "application/pdf"
                    else:
                        st.warning("No se pudo generar PDF. Descargando DOCX.")
                        buffer = docx_buffer
                        ext = "docx"
                        mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                else:
                    return
            else:
                buffer = generator.generate_a3(final_data, images, st.session_state.get("template_a3_bytes"))
                ext = "docx"
                mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        else:
            if output_format == "pdf":
                pptx_buffer = generator.generate_kaizen(final_data, images, st.session_state.get("template_kaizen_bytes"))
                if pptx_buffer:
                    pdf_bytes = pdf_exporter.pptx_to_pdf_libreoffice(pptx_buffer.getvalue(), "kaizen.pdf")
                    if pdf_bytes:
                        buffer = BytesIO(pdf_bytes)
                        ext = "pdf"
                        mime = "application/pdf"
                    else:
                        st.warning("No se pudo generar PDF. Descargando PPTX.")
                        buffer = pptx_buffer
                        ext = "pptx"
                        mime = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
                else:
                    return
            else:
                buffer = generator.generate_kaizen(final_data, images, st.session_state.get("template_kaizen_bytes"))
                ext = "pptx"
                mime = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        if buffer is None:
            return
        filename = f"{meta.get('moc_number', meta.get('doc_number', 'DOC'))}_{language}.{ext}"
        doc_info = {
            "type": doc_type,
            "title": meta.get("moc_title", meta.get("titulo", "Sin título")),
            "number": meta.get("moc_number", meta.get("doc_number", "")),
            "language": language,
            "format": ext,
            "filename": filename
        }
        Utils.add_to_history(doc_info)
        st.success(f"✅ Documento generado exitosamente: {filename}")
        st.download_button(
            label=f"📥 Descargar {ext.upper()}",
            data=buffer,
            file_name=filename,
            mime=mime,
            use_container_width=True
        )

def render_history():
    st.markdown('<div class="section-header"><h3>📁 Historial de Documentos Generados</h3></div>', unsafe_allow_html=True)
    docs = st.session_state.history.get("documents", [])
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("#### 💾 Exportar Configuración")
        export_data = {
            "config": st.session_state.config,
            "history": st.session_state.history,
            "export_date": datetime.now().isoformat(),
            "version": "7.5.0"
        }
        export_json = json.dumps(export_data, indent=2, ensure_ascii=False)
        st.download_button(
            label="📥 Descargar backup completo (JSON)",
            data=export_json,
            file_name=f"gestion_documental_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
            mime="application/json",
            use_container_width=True
        )
    with col2:
        st.markdown("#### 📂 Importar Configuración")
        uploaded_backup = st.file_uploader("Seleccione archivo de backup (.json):", type=["json"], key="import_backup")
        if uploaded_backup:
            try:
                backup_data = json.loads(uploaded_backup.read())
                if "config" in backup_data:
                    st.session_state.config = backup_data["config"]
                    LocalStorage.save_config(backup_data["config"])
                if "history" in backup_data:
                    st.session_state.history = backup_data["history"]
                    LocalStorage.save_history(backup_data["history"])
                st.success("✅ Configuración e historial restaurados correctamente.")
                st.rerun()
            except Exception as e:
                st.error(f"❌ Error al importar: {e}")
    st.markdown("<hr>", unsafe_allow_html=True)
    if not docs:
        st.info("📭 No hay documentos generados aún.")
        return
    col1, col2, col3 = st.columns(3)
    with col1:
        filter_type = st.selectbox("Filtrar:", ["Todos", "MoC", "A3", "Kaizen"], key="hist_filter_type")
    with col2:
        filter_lang = st.selectbox("Idioma:", ["Todos", "Español", "Inglés"], key="hist_filter_lang")
    with col3:
        search = st.text_input("Buscar:", placeholder="Título o número...", key="hist_search")
    filtered = docs
    if filter_type != "Todos":
        type_map = {"MoC": "moc", "A3": "a3", "Kaizen": "kaizen"}
        filtered = [d for d in filtered if d.get("type") == type_map.get(filter_type)]
    if filter_lang != "Todos":
        lang_map = {"Español": "es", "Inglés": "en"}
        filtered = [d for d in filtered if d.get("language", "es") == lang_map.get(filter_lang)]
    if search:
        filtered = [d for d in filtered if search.lower() in d.get("title", "").lower()
                    or search.lower() in d.get("number", "").lower()]
    st.markdown(f"**Mostrando {len(filtered)} documento(s) de {len(docs)} total(es)**")
    for doc in filtered:
        type_emoji = {"moc": "📋", "a3": "📊", "kaizen": "⚡"}.get(doc.get("type"), "📄")
        type_label = {"moc": "MoC", "a3": "A3", "kaizen": "Kaizen"}.get(doc.get("type"), "Doc")
        lang_flag = "🇪🇸" if doc.get("language") == "es" else "🇺🇸"
        fmt_icon = "📄" if doc.get("format") == "pdf" else "📑"
        st.markdown(f"""
<div class="history-item">
<h4 style="margin: 0; color: #1e293b;">{type_emoji} {doc.get('title', 'Sin título')}</h4>
<p style="margin: 0.25rem 0; color: #64748b; font-size: 0.9rem;">
{type_label} · {doc.get('number', '')} · {lang_flag} · {fmt_icon} · {doc.get('timestamp', '')[:10]}
</p>
</div>
""", unsafe_allow_html=True)
        if st.button("🗑️ Eliminar", key=f"del_{doc.get('id', 'x')}"):
            Utils.delete_from_history(doc.get('id'))
            st.rerun()

def render_settings():
    st.markdown('<div class="section-header"><h3>⚙️ Configuración del Sistema</h3></div>', unsafe_allow_html=True)
    config = st.session_state.config
    tabs = st.tabs(["🔑 API Gemini", "🏢 Empresa", "📄 Templates", "🔧 Avanzado", "💾 Backup"])
    with tabs[0]:
        st.markdown("#### API Key Gemini")
        st.info("💡 Obtenga su API Key gratuita en [Google AI Studio](https://aistudio.google.com/)")
        api_key = st.text_input("API Key:", value=config.get("gemini_api_key", ""), type="password")
        
        # NUEVO: Botón de diagnóstico de API
        if st.button("🔌 Probar Conexión API", type="secondary", use_container_width=True):
            if not api_key:
                st.error("⚠️ Ingrese una API Key primero.")
            else:
                with st.spinner("Probando conexión con Google Gemini..."):
                    try:
                        import requests
                        url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={api_key}"
                        payload = {"contents": [{"parts": [{"text": "Responde solo con la palabra 'OK'"}]}]}
                        resp = requests.post(url, json=payload, timeout=10)
                        if resp.status_code == 200:
                            st.success("✅ ¡Conexión exitosa! Tu API Key es válida y la API está habilitada.")
                        elif resp.status_code == 404:
                            st.error("❌ Error 404: La 'Generative Language API' NO está habilitada en tu proyecto.")
                            st.info("👉 **Solución:** Ve a [Google Cloud Console](https://console.cloud.google.com/apis/library/generativelanguage.googleapis.com), selecciona tu proyecto y haz clic en **HABILITAR**.")
                            st.code(resp.text, language="json")
                        elif resp.status_code == 403:
                            st.error("❌ Error 403: Acceso denegado. Tu API Key puede tener restricciones de IP/Referer o no tener permisos.")
                            st.code(resp.text, language="json")
                        else:
                            st.error(f"❌ Error {resp.status_code}: {resp.text}")
                    except Exception as e:
                        st.error(f"❌ Error de conexión: {e}")
        
        st.markdown("#### Selección de Modelo")
        current_model = config.get("gemini_model", "gemini-1.5-pro")
        col1, col2 = st.columns(2)
        models = [
            ("gemini-1.5-pro", "🧠 Gemini 1.5 Pro", "Máxima calidad y razonamiento", "Recomendado"),
            ("gemini-1.5-flash", "⚡ Gemini 1.5 Flash", "Rápido y eficiente", "Estándar"),
        ]
        for i, (model_id, name, desc, badge) in enumerate(models):
            is_selected = current_model == model_id
            with [col1, col2][i]:
                border_color = "#1a5f7a" if is_selected else "#e2e8f0"
                bg_color = "#eff6ff" if is_selected else "#f8fafc"
                selected_text = '<div style="color: #1a5f7a; font-weight: bold; margin-top: 0.5rem;">✓ Seleccionado</div>' if is_selected else ''
                st.markdown(f"""
<div style="background: {bg_color}; border: 2px solid {border_color}; border-radius: 12px; padding: 1rem; text-align: center;">
<div style="font-size: 2rem;">{name.split()[0]}</div>
<h4 style="margin: 0.5rem 0; color: #1e293b;">{name.split(maxsplit=1)[1]}</h4>
<p style="color: #64748b; font-size: 0.85rem; margin: 0;">{desc}</p>
<span style="background: #dbeafe; color: #1d4ed8; padding: 0.15rem 0.5rem; border-radius: 10px; font-size: 0.75rem;">{badge}</span>
{selected_text}
</div>
""", unsafe_allow_html=True)
                btn_type = "primary" if is_selected else "secondary"
                if st.button(f"Seleccionar", key=f"sel_{model_id}", use_container_width=True, type=btn_type):
                    config["gemini_model"] = model_id
                    st.session_state.config = config
                    LocalStorage.save_config(config)
                    st.success(f"✅ Modelo: {name}")
                    st.rerun()
        if st.button("💾 Guardar API Key", type="primary", use_container_width=True):
            config["gemini_api_key"] = api_key
            st.session_state.config = config
            LocalStorage.save_config(config)
            st.success("✅ API Key guardada")
    with tabs[1]:
        st.markdown("#### Datos de la Empresa")
        company = st.text_input("Empresa:", value=config.get("company_name", ""), placeholder="Ej: Orica Perú")
        dept = st.text_input("Departamento:", value=config.get("department", ""), placeholder="Ej: Mantenimiento")
        author = st.text_input("Autor por defecto:", value=config.get("default_author", ""))
        area = st.text_input("Área por defecto:", value=config.get("default_area", ""), placeholder="Ej: Planta Lurín")
        if st.button("💾 Guardar Datos", type="primary", use_container_width=True):
            config["company_name"] = company
            config["department"] = dept
            config["default_author"] = author
            config["default_area"] = area
            st.session_state.config = config
            LocalStorage.save_config(config)
            st.success("✅ Datos guardados")
    with tabs[2]:
        st.markdown("#### Carga de Templates Oficiales")
        st.warning("⚠️ **Importante:** Suba los archivos oficiales (.pptx para MoC y Kaizen, .docx para A3).")
        moc_ok = st.session_state.get("template_moc_bytes") is not None
        a3_ok = st.session_state.get("template_a3_bytes") is not None
        kzn_ok = st.session_state.get("template_kaizen_bytes") is not None
        col1, col2, col3 = st.columns(3)
        with col1:
            status_moc = "✅ Cargado" if moc_ok else "❌ Pendiente"
            st.markdown(f"**Template MoC**<br>{status_moc}", unsafe_allow_html=True)
            moc_file = st.file_uploader("Subir MoC (.pptx)", type=["pptx"], key="upload_moc")
        with col2:
            status_a3 = "✅ Cargado" if a3_ok else "❌ Pendiente"
            st.markdown(f"**Template A3**<br>{status_a3}", unsafe_allow_html=True)
            a3_file = st.file_uploader("Subir A3 (.docx)", type=["docx"], key="upload_a3")
        with col3:
            status_kzn = "✅ Cargado" if kzn_ok else "❌ Pendiente"
            st.markdown(f"**Template Kaizen**<br>{status_kzn}", unsafe_allow_html=True)
            kzn_file = st.file_uploader("Subir Kaizen (.pptx)", type=["pptx"], key="upload_kzn")
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("💾 Guardar Templates", type="primary", use_container_width=True):
            saved_any = False
            if moc_file is not None:
                bytes_data = moc_file.getvalue()
                st.session_state.template_moc_bytes = bytes_data
                LocalStorage.save_template_bytes("moc", bytes_data)
                st.success("✅ Template MoC guardado")
                saved_any = True
            if a3_file is not None:
                bytes_data = a3_file.getvalue()
                st.session_state.template_a3_bytes = bytes_data
                LocalStorage.save_template_bytes("a3", bytes_data)
                st.success("✅ Template A3 guardado")
                saved_any = True
            if kzn_file is not None:
                bytes_data = kzn_file.getvalue()
                st.session_state.template_kaizen_bytes = bytes_data
                LocalStorage.save_template_bytes("kaizen", bytes_data)
                st.success("✅ Template Kaizen guardado")
                saved_any = True
            if not saved_any:
                st.warning("⚠️ No se seleccionó ningún archivo nuevo.")
            else:
                st.rerun()
        if moc_ok and a3_ok and kzn_ok:
            st.balloons()
            st.success("🎉 ¡Todos los templates están cargados!")
    with tabs[3]:
        st.markdown("#### Configuración Avanzada")
        auto_correct = st.toggle("Corrección automática en campos de entrada", value=config.get("auto_correct", True))
        spell_check = st.toggle("Corrector ortográfico con IA en revisión", value=config.get("spell_check", True))
        thinking = st.select_slider("Profundidad de generación:", ["Básico", "Estándar", "Profundo"],
                                    value=config.get("thinking_level", "Estándar"))
        col1, col2, col3 = st.columns(3)
        with col1:
            last_moc = st.number_input("Último MoC:", value=config.get("last_moc_number", 0), min_value=0)
        with col2:
            last_a3 = st.number_input("Último A3:", value=config.get("last_a3_number", 0), min_value=0)
        with col3:
            last_kzn = st.number_input("Último Kaizen:", value=config.get("last_kaizen_number", 0), min_value=0)
        if st.button("💾 Guardar Configuración Avanzada", type="primary", use_container_width=True):
            config["auto_correct"] = auto_correct
            config["spell_check"] = spell_check
            config["thinking_level"] = thinking
            config["last_moc_number"] = last_moc
            config["last_a3_number"] = last_a3
            config["last_kaizen_number"] = last_kzn
            st.session_state.config = config
            LocalStorage.save_config(config)
            st.success("✅ Configuración avanzada guardada")
    with tabs[4]:
        st.markdown("#### Backup y Restauración")
        export_data = {
            "config": st.session_state.config,
            "history": st.session_state.history,
            "export_date": datetime.now().isoformat(),
            "version": "7.5.0"
        }
        export_json = json.dumps(export_data, indent=2, ensure_ascii=False)
        st.download_button(
            label="📥 Exportar todo (JSON)",
            data=export_json,
            file_name=f"backup_gestion_documental_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
            mime="application/json",
            use_container_width=True
        )
        st.markdown("#### Restaurar desde archivo")
        restore_file = st.file_uploader("Seleccione archivo de backup:", type=["json"], key="restore_file")
        if restore_file:
            try:
                restore_data = json.loads(restore_file.read())
                if "config" in restore_data:
                    st.session_state.config = restore_data["config"]
                    LocalStorage.save_config(restore_data["config"])
                if "history" in restore_data:
                    st.session_state.history = restore_data["history"]
                    LocalStorage.save_history(restore_data["history"])
                st.success("✅ Restauración completada.")
                st.rerun()
            except Exception as e:
                st.error(f"❌ Error: {e}")
        st.markdown("---")
        st.markdown("#### 🗑️ Zona de Peligro")
        st.warning("⚠️ Las siguientes acciones son irreversibles.")
        col1, col2 = st.columns(2)
        with col1:
            if st.button("🗑️ Borrar Historial Completo", type="secondary", use_container_width=True):
                st.session_state.history = {"documents": []}
                LocalStorage.save_history({"documents": []})
                st.success("✅ Historial borrado")
                st.rerun()
        with col2:
            if st.button("🗑️ Borrar Configuración", type="secondary", use_container_width=True):
                st.session_state.config = {
                    "gemini_api_key": "",
                    "gemini_model": "gemini-1.5-pro",
                    "company_name": "",
                    "department": "",
                    "default_author": "",
                    "default_area": "",
                    "header_text": "",
                    "footer_text": "",
                    "last_moc_number": 0,
                    "last_a3_number": 0,
                    "last_kaizen_number": 0,
                    "spell_check": True,
                    "thinking_level": "Estándar",
                    "auto_correct": True,
                }
                LocalStorage.save_config(st.session_state.config)
                st.success("✅ Configuración restaurada")
                st.rerun()

def main():
    render_sidebar()
    page = st.session_state.page
    if page == "inicio":
        render_welcome()
    elif page == "nueva_moc":
        render_moc_form()
    elif page == "nueva_a3":
        render_a3_form()
    elif page == "nuevo_kaizen":
        render_kaizen_form()
    elif page == "revisar":
        render_review()
    elif page == "historial":
        render_history()
    elif page == "configuracion":
        render_settings()
    else:
        render_welcome()
    st.markdown("""
<div class="app-footer">
<p><strong style="font-size: 1.1rem;">CAVA</strong> - Especialistas en Robótica y Automatización</p>
<p>Diseñado por <strong>Roger Huamani</strong> | Sistema de Gestión Documental v7.5.0</p>
<p style="font-size: 0.75rem; color: #94a3b8;">
Software empresarial para automatización de documentos MoC, A3 y Kaizen.<br>
Formato oficial MDET con Checklist 360° y análisis integral.<br>
Datos persistentes locales. Exportación a PDF integrada.
</p>
</div>
""", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
