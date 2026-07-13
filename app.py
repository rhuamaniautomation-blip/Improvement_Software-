#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SISTEMA DE GESTION DOCUMENTAL - MoC | Mejora A3 | Simple Kaizen
Version 5.0.0 - Completo con Templates Oficiales y Correccion Ortografica
Diseñado por: CAVA - Especialistas en Robotica y Automatizacion
Desarrollador: Roger Huamani
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

REPORTLAB_AVAILABLE = False
try:
    from reportlab.lib.pagesizes import A4, letter
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch, cm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage, PageBreak
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
    REPORTLAB_AVAILABLE = True
except ImportError:
    pass

st.set_page_config(
    page_title="Gestion Documental - MoC | A3 | Kaizen",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

BASE_DIR = Path(__file__).parent if "__file__" in dir() else Path(".")
DATA_DIR = BASE_DIR / "data"
TEMPLATES_DIR = BASE_DIR / "templates"
HISTORY_FILE = DATA_DIR / "history.json"
CONFIG_FILE = DATA_DIR / "config.json"

DATA_DIR.mkdir(exist_ok=True)
TEMPLATES_DIR.mkdir(exist_ok=True)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', 'Segoe UI', sans-serif !important; }
    .main-header { background: linear-gradient(135deg, #1a5f7a 0%, #2e8bc0 100%); padding: 2rem; border-radius: 12px; color: white; margin-bottom: 2rem; box-shadow: 0 4px 20px rgba(26, 95, 122, 0.3); }
    .main-header h1 { color: white !important; font-weight: 700 !important; margin-bottom: 0.5rem !important; }
    .main-header p { color: rgba(255,255,255,0.9) !important; font-size: 1.1rem !important; }
    .doc-card { background: white; border-radius: 16px; padding: 2rem; border: 2px solid #e2e8f0; transition: all 0.3s ease; cursor: pointer; height: 100%; }
    .doc-card:hover { border-color: #1a5f7a; transform: translateY(-4px); box-shadow: 0 12px 40px rgba(26, 95, 122, 0.15); }
    .doc-card-moc { border-left: 5px solid #1a5f7a; }
    .doc-card-a3 { border-left: 5px solid #10b981; }
    .doc-card-kaizen { border-left: 5px solid #f59e0b; }
    .stButton > button { border-radius: 10px !important; font-weight: 600 !important; padding: 0.75rem 2rem !important; transition: all 0.2s ease !important; }
    .stButton > button:hover { transform: translateY(-2px); box-shadow: 0 4px 12px rgba(0,0,0,0.15); }
    .stTextInput > div > div > input, .stTextArea > div > div > textarea, .stSelectbox > div > div > div { border-radius: 8px !important; border: 1px solid #cbd5e1 !important; font-size: 15px !important; }
    .section-header { background: #f1f5f9; padding: 1rem 1.5rem; border-radius: 10px; margin: 1.5rem 0 1rem 0; border-left: 4px solid #1a5f7a; }
    .section-header h3 { margin: 0 !important; color: #1e293b !important; font-weight: 600 !important; }
    .field-card { background: white; border: 1px solid #e2e8f0; border-radius: 10px; padding: 1rem; margin: 0.5rem 0; }
    .field-card:hover { border-color: #1a5f7a; box-shadow: 0 2px 8px rgba(0,0,0,0.05); }
    .gemini-badge { display: inline-block; background: #e0e7ff; color: #4338ca; padding: 0.25rem 0.75rem; border-radius: 20px; font-size: 12px; font-weight: 600; margin-left: 0.5rem; }
    .history-item { background: white; border: 1px solid #e2e8f0; border-radius: 10px; padding: 1rem; margin: 0.5rem 0; transition: all 0.2s; }
    .history-item:hover { border-color: #1a5f7a; box-shadow: 0 2px 8px rgba(0,0,0,0.05); }
    .app-footer { text-align: center; padding: 2rem; margin-top: 3rem; border-top: 1px solid #e2e8f0; color: #64748b; }
    .auto-correct-badge { display: inline-flex; align-items: center; background: #dcfce7; color: #166534; padding: 0.25rem 0.75rem; border-radius: 20px; font-size: 11px; font-weight: 600; margin-bottom: 0.5rem; }
    .image-preview { border: 2px solid #e2e8f0; border-radius: 12px; padding: 0.5rem; background: #f8fafc; text-align: center; }
    .image-preview img { border-radius: 8px; max-width: 100%; }
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    [data-testid="stSidebar"] { background: #1e293b !important; }
    [data-testid="stSidebar"] .stMarkdown { color: #94a3b8 !important; }
    [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 { color: #f8fafc !important; }
    .template-uploader { background: #f8fafc; border: 2px dashed #cbd5e1; border-radius: 12px; padding: 1.5rem; text-align: center; margin: 1rem 0; }
    .template-uploader.ok { background: #f0fdf4; border-color: #10b981; }
</style>
""", unsafe_allow_html=True)

class LocalStorage:
    @staticmethod
    def save_config(config):
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            st.error(f"Error guardando configuracion: {e}")
            return False

    @staticmethod
    def load_config():
        try:
            if CONFIG_FILE.exists():
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    return json.load(f)
        except Exception as e:
            st.warning(f"Error cargando configuracion: {e}")
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

    @staticmethod
    def delete_template(template_name):
        try:
            template_path = TEMPLATES_DIR / f"{template_name}.bin"
            if template_path.exists():
                template_path.unlink()
        except Exception:
            pass

class Utils:
    @staticmethod
    def format_date():
        meses = {1:"ENERO",2:"FEBRERO",3:"MARZO",4:"ABRIL",5:"MAYO",6:"JUNIO",
                 7:"JULIO",8:"AGOSTO",9:"SEPTIEMBRE",10:"OCTUBRE",11:"NOVIEMBRE",12:"DICIEMBRE"}
        now = datetime.now()
        return f"{meses[now.month]} {now.year}"

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
    def correct_spelling_comprehensive(text):
        if not text or not text.strip():
            return text
        result = text
        accent_corrections = {
            "tecnico": "tecnico", "tecnica": "tecnica", "tecnicos": "tecnicos", "tecnicas": "tecnicas",
            "Tecnico": "Tecnico", "Tecnica": "Tecnica", "Tecnicos": "Tecnicos", "Tecnicas": "Tecnicas",
            "TECNICO": "TECNICO", "TECNICA": "TECNICA", "TECNICOS": "TECNICOS", "TECNICAS": "TECNICAS",
            "produccion": "produccion", "Produccion": "Produccion", "PRODUCCION": "PRODUCCION",
            "implementacion": "implementacion", "Implementacion": "Implementacion", "IMPLEMENTACION": "IMPLEMENTACION",
            "evaluacion": "evaluacion", "Evaluacion": "Evaluacion", "EVALUACION": "EVALUACION",
            "operacion": "operacion", "Operacion": "Operacion", "OPERACION": "OPERACION",
            "condicion": "condicion", "Condicion": "Condicion", "CONDICION": "CONDICION",
            "modificacion": "modificacion", "Modificacion": "Modificacion", "MODIFICACION": "MODIFICACION",
            "verificacion": "verificacion", "Verificacion": "Verificacion", "VERIFICACION": "VERIFICACION",
            "capacitacion": "capacitacion", "Capacitacion": "Capacitacion", "CAPACITACION": "CAPACITACION",
            "socializacion": "socializacion", "Socializacion": "Socializacion", "SOCIALIZACION": "SOCIALIZACION",
            "documentacion": "documentacion", "Documentacion": "Documentacion", "DOCUMENTACION": "DOCUMENTACION",
            "estandarizacion": "estandarizacion", "Estandarizacion": "Estandarizacion", "ESTANDARIZACION": "ESTANDARIZACION",
            "optimizacion": "optimizacion", "Optimizacion": "Optimizacion", "OPTIMIZACION": "OPTIMIZACION",
            "identificacion": "identificacion", "Identificacion": "Identificacion", "IDENTIFICACION": "IDENTIFICACION",
            "clasificacion": "clasificacion", "Clasificacion": "Clasificacion", "CLASIFICACION": "CLASIFICACION",
            "notificacion": "notificacion", "Notificacion": "Notificacion", "NOTIFICACION": "NOTIFICACION",
            "coordinacion": "coordinacion", "Coordinacion": "Coordinacion", "COORDINACION": "COORDINACION",
            "aprobacion": "aprobacion", "Aprobacion": "Aprobacion", "APROBACION": "APROBACION",
            "revision": "revision", "Revision": "Revision", "REVISION": "REVISION",
            "revison": "revision", "Revison": "Revision", "REVISON": "REVISION",
            "ejecucion": "ejecucion", "Ejecucion": "Ejecucion", "EJECUCION": "EJECUCION",
            "inspeccion": "inspeccion", "Inspeccion": "Inspeccion", "INSPECCION": "INSPECCION",
            "proteccion": "proteccion", "Proteccion": "Proteccion", "PROTECCION": "PROTECCION",
            "deteccion": "deteccion", "Deteccion": "Deteccion", "DETECCION": "DETECCION",
            "prevencion": "prevencion", "Prevencion": "Prevencion", "PREVENCION": "PREVENCION",
            "intervencion": "intervencion", "Intervencion": "Intervencion", "INTERVENCION": "INTERVENCION",
            "supervision": "supervision", "Supervision": "Supervision", "SUPERVISION": "SUPERVISION",
            "comunicacion": "comunicacion", "Comunicacion": "Comunicacion", "COMUNICACION": "COMUNICACION",
            "organizacion": "organizacion", "Organizacion": "Organizacion", "ORGANIZACION": "ORGANIZACION",
            "planificacion": "planificacion", "Planificacion": "Planificacion", "PLANIFICACION": "PLANIFICACION",
            "calificacion": "calificacion", "Calificacion": "Calificacion", "CALIFICACION": "CALIFICACION",
            "certificacion": "certificacion", "Certificacion": "Certificacion", "CERTIFICACION": "CERTIFICACION",
            "validacion": "validacion", "Validacion": "Validacion", "VALIDACION": "VALIDACION",
            "calibracion": "calibracion", "Calibracion": "Calibracion", "CALIBRACION": "CALIBRACION",
            "configuracion": "configuracion", "Configuracion": "Configuracion", "CONFIGURACION": "CONFIGURACION",
            "programacion": "programacion", "Programacion": "Programacion", "PROGRAMACION": "PROGRAMACION",
            "automatizacion": "automatizacion", "Automatizacion": "Automatizacion", "AUTOMATIZACION": "AUTOMATIZACION",
            "integracion": "integracion", "Integracion": "Integracion", "INTEGRACION": "INTEGRACION",
            "funcion": "funcion", "Funcion": "Funcion", "FUNCION": "FUNCION",
            "relacion": "relacion", "Relacion": "Relacion", "RELACION": "RELACION",
            "conexion": "conexion", "Conexion": "Conexion", "CONECTION": "CONEXION",
            "direccion": "direccion", "Direccion": "Direccion", "DIRECCION": "DIRECCION",
            "seleccion": "seleccion", "Seleccion": "Seleccion", "SELECCION": "SELECCION",
            "proyeccion": "proyeccion", "Proyeccion": "Proyeccion", "PROYECCION": "PROYECCION",
            "restriccion": "restriccion", "Restriccion": "Restriccion", "RESTRICCION": "RESTRICCION",
            "distribucion": "distribucion", "Distribucion": "Distribucion", "DISTRIBUCION": "DISTRIBUCION",
            "construccion": "construccion", "Construccion": "Construccion", "CONSTRUCCION": "CONSTRUCCION",
            "destruccion": "destruccion", "Destruccion": "Destruccion", "DESTRUCCION": "DESTRUCCION",
            "instruccion": "instruccion", "Instruccion": "Instruccion", "INSTRUCCION": "INSTRUCCION",
            "conduccion": "conduccion", "Conduccion": "Conduccion", "CONDUCCION": "CONDUCCION",
            "introduccion": "introduccion", "Introduccion": "Introduccion", "INTRODUCCION": "INTRODUCCION",
            "reduccion": "reduccion", "Reduccion": "Reduccion", "REDUCCION": "REDUCCION",
            "reproduccion": "reproduccion", "Reproduccion": "Reproduccion", "REPRODUCCION": "REPRODUCCION",
            "traduccion": "traduccion", "Traduccion": "Traduccion", "TRADUCCION": "TRADUCCION",
            "deduccion": "deduccion", "Deduccion": "Deduccion", "DEDUCCION": "DEDUCCION",
            "induccion": "induccion", "Induccion": "Induccion", "INDUCCION": "INDUCCION",
            "seduccion": "seduccion", "Seduccion": "Seduccion", "SEDUCCION": "SEDUCCION",
            "seguridad": "seguridad", "Seguridad": "Seguridad", "SEGURIDAD": "SEGURIDAD",
            "maquina": "maquina", "maquinas": "maquinas", "Maquina": "Maquina", "Maquinas": "Maquinas",
            "MAQUINA": "MAQUINA", "MAQUINAS": "MAQUINAS",
            "numero": "numero", "numeros": "numeros", "Numero": "Numero", "Numeros": "Numeros",
            "NUMERO": "NUMERO", "NUMEROS": "NUMEROS",
            "maximo": "maximo", "maxima": "maxima", "Maximo": "Maximo", "Maxima": "Maxima",
            "MAXIMO": "MAXIMO", "MAXIMA": "MAXIMA",
            "minimo": "minimo", "minima": "minima", "Minimo": "Minimo", "Minima": "Minima",
            "MINIMO": "MINIMO", "MINIMA": "MINIMA",
            "rapido": "rapido", "rapida": "rapida", "Rapido": "Rapido", "Rapida": "Rapida",
            "RAPIDO": "RAPIDO", "RAPIDA": "RAPIDA",
            "facil": "facil", "Facil": "Facil", "FACIL": "FACIL",
            "dificil": "dificil", "Dificil": "Dificil", "DIFICIL": "DIFICIL",
            "electronico": "electronico", "electronica": "electronica", "Electronicos": "Electronicos",
            "Electrico": "Electrico", "Electrica": "Electrica", "electrico": "electrico", "electrica": "electrica",
            "ELECTRICO": "ELECTRICO", "ELECTRICA": "ELECTRICA",
            "mecanico": "mecanico", "mecanica": "mecanica", "Mecanico": "Mecanico", "Mecanica": "Mecanica",
            "MECANICO": "MECANICO", "MECANICA": "MECANICA",
            "hidraulico": "hidraulico", "hidraulica": "hidraulica", "Hidraulico": "Hidraulico", "Hidraulica": "Hidraulica",
            "HIDRAULICO": "HIDRAULICO", "HIDRAULICA": "HIDRAULICA",
            "neumatico": "neumatico", "neumatica": "neumatica", "Neumatico": "Neumatico", "Neumatica": "Neumatica",
            "NEUMATICO": "NEUMATICO", "NEUMATICA": "NEUMATICA",
            "termico": "termico", "termica": "termica", "Termico": "Termico", "Termica": "Termica",
            "TERMICO": "TERMICO", "TERMICA": "TERMICA",
            "critico": "critico", "critica": "critica", "Critico": "Critico", "Critica": "Critica",
            "CRITICO": "CRITICO", "CRITICA": "CRITICA",
            "periodico": "periodico", "periodica": "periodica", "Periodico": "Periodico", "Periodica": "Periodica",
            "PERIODICO": "PERIODICO", "PERIODICA": "PERIODICA",
            "logico": "logico", "logica": "logica", "Logico": "Logico", "Logica": "Logica",
            "LOGICO": "LOGICO", "LOGICA": "LOGICA",
            "publico": "publico", "publica": "publica", "Publico": "Publico", "Publica": "Publica",
            "PUBLICO": "PUBLICO", "PUBLICA": "PUBLICA",
            "unico": "unico", "unica": "unica", "Unico": "Unico", "Unica": "Unica",
            "UNICO": "UNICO", "UNICA": "UNICA",
            "fisico": "fisico", "fisica": "fisica", "Fisico": "Fisico", "Fisica": "Fisica",
            "FISICO": "FISICO", "FISICA": "FISICA",
            "quimico": "quimico", "quimica": "quimica", "Quimico": "Quimico", "Quimica": "Quimica",
            "QUIMICO": "QUIMICO", "QUIMICA": "QUIMICA",
            "medico": "medico", "medica": "medica", "Medico": "Medico", "Medica": "Medica",
            "MEDICO": "MEDICO", "MEDICA": "MEDICA",
            "politico": "politico", "politica": "politica", "Politico": "Politico", "Politica": "Politica",
            "POLITICO": "POLITICO", "POLITICA": "POLITICA",
            "economico": "economico", "economica": "economica", "Economico": "Economico", "Economica": "Economica",
            "ECONOMICO": "ECONOMICO", "ECONOMICA": "ECONOMICA",
            "didactico": "didactico", "didactica": "didactica", "Didactico": "Didactico", "Didactica": "Didactica",
            "DIDACTICO": "DIDACTICO", "DIDACTICA": "DIDACTICA",
            "tactico": "tactico", "tactica": "tactica", "Tactico": "Tactico", "Tactica": "Tactica",
            "TACTICO": "TACTICO", "TACTICA": "TACTICA",
            "practico": "practico", "practica": "practica", "Practico": "Practico", "Practica": "Practica",
            "PRACTICO": "PRACTICO", "PRACTICA": "PRACTICA",
            "sistematico": "sistematico", "sistematica": "sistematica", "Sistematico": "Sistematico", "Sistematica": "Sistematica",
            "SISTEMATICO": "SISTEMATICO", "SISTEMATICA": "SISTEMATICA",
            "automatico": "automatico", "automatica": "automatica", "Automatico": "Automatico", "Automatica": "Automatica",
            "AUTOMATICO": "AUTOMATICO", "AUTOMATICA": "AUTOMATICA",
            "caracteristico": "caracteristico", "caracteristica": "caracteristica",
            "Caracteristico": "Caracteristico", "Caracteristica": "Caracteristica",
            "CARACTERISTICO": "CARACTERISTICO", "CARACTERISTICA": "CARACTERISTICA",
            "analisis": "analisis", "Analisis": "Analisis", "ANALISIS": "ANALISIS",
            "crisis": "crisis", "Crisis": "Crisis", "CRISIS": "CRISIS",
            "tesis": "tesis", "Tesis": "Tesis", "TESIS": "TESIS",
            "hipotesis": "hipotesis", "Hipotesis": "Hipotesis", "HIPOTESIS": "HIPOTESIS",
            "sintesis": "sintesis", "Sintesis": "Sintesis", "SINTESIS": "SINTESIS",
            "parentesis": "parentesis", "Parentesis": "Parentesis", "PARENTESIS": "PARENTESIS",
            "enfasis": "enfasis", "Enfasis": "Enfasis", "ENFASIS": "ENFASIS",
            "diasis": "diasis", "Diasis": "Diasis", "DIASIS": "DIASIS",
            "apocalipsis": "apocalipsis", "Apocalipsis": "Apocalipsis", "APOCALIPSIS": "APOCALIPSIS",
            "paralisis": "paralisis", "Paralisis": "Paralisis", "PARALISIS": "PARALISIS",
            "linguistica": "linguistica", "Linguistica": "Linguistica", "LINGUISTICA": "LINGUISTICA",
            "bilingue": "bilingue", "Bilingue": "Bilingue", "BILINGUE": "BILINGUE",
            "ambiguo": "ambiguo", "Ambiguo": "Ambiguo", "AMBIGUO": "AMBIGUO",
            "antiguo": "antiguo", "antigua": "antigua", "Antiguo": "Antiguo", "Antigua": "Antigua",
            "ANTIGUO": "ANTIGUO", "ANTIGUA": "ANTIGUA",
            "arguir": "arguir", "Arguir": "Arguir", "ARGUIR": "ARGUIR",
            "averiguo": "averiguo", "Averiguo": "Averiguo", "AVERIGUO": "AVERIGUO",
            "bilinguo": "bilinguo", "Bilinguo": "Bilinguo", "BILINGUO": "BILINGUO",
            "ciguena": "ciguena", "Ciguena": "Ciguena", "CIGUENA": "CIGUENA",
            "contiguo": "contiguo", "contigua": "contigua", "Contiguo": "Contiguo", "Contigua": "Contigua",
            "CONTIGUO": "CONTIGUO", "CONTIGUA": "CONTIGUA",
            "desaguisado": "desaguisado", "Desaguisado": "Desaguisado", "DESAGUISADO": "DESAGUISADO",
            "exiguo": "exiguo", "exigua": "exigua", "Exiguo": "Exiguo", "Exigua": "Exigua",
            "EXIGUO": "EXIGUO", "EXIGUA": "EXIGUA",
            "pinguino": "pinguino", "Pinguino": "Pinguino", "PINGUINO": "PINGUINO",
            "prosiguio": "prosiguio", "Prosiguio": "Prosiguio", "PROSIGUIO": "PROSIGUIO",
            "santiguo": "santiguo", "Santiguo": "Santiguo", "SANTIGUO": "SANTIGUO",
            "trilingue": "trilingue", "Trilingue": "Trilingue", "TRILINGUE": "TRILINGUE",
            "veriguo": "veriguo", "Veriguo": "Veriguo", "VERIGUO": "VERIGUO",
            "sere": "sere", "Sere": "Sere", "SERE": "SERE",
            "tene": "tene", "Tene": "Tene", "TENE": "TENE",
            "vene": "vene", "Vene": "Vene", "VENE": "VENE",
            "pode": "pode", "Pode": "Pode", "PODE": "PODE",
            "debe": "debe", "Debe": "Debe", "DEBE": "DEBE",
            "sabe": "sabe", "Sabe": "Sabe", "SABE": "SABE",
            "habe": "habe", "Habe": "Habe", "HABE": "HABE",
            "hace": "hace", "Hace": "Hace", "HACE": "HACE",
            "dece": "dece", "Dece": "Dece", "DECE": "DECE",
            "pede": "pede", "Pede": "Pede", "PEDE": "PEDE",
            "mede": "mede", "Mede": "Mede", "MEDE": "MEDE",
            "segue": "segue", "Segue": "Segue", "SEGUE": "SEGUE",
            "consegue": "consegue", "Consegue": "Consegue", "CONSEGUE": "CONSEGUE",
            "persegue": "persegue", "Persegue": "Persegue", "PERSEGUE": "PERSEGUE",
            "prosegue": "prosegue", "Prosegue": "Prosegue", "PROSEGUE": "PROSEGUE",
            "rei": "rei", "Rei": "Rei", "REI": "REI",
            "sonrei": "sonrei", "Sonrei": "Sonrei", "SONREI": "SONREI",
            "frei": "frei", "Frei": "Frei", "FREI": "FREI",
            "reune": "reune", "Reune": "Reune", "REUNE": "REUNE",
            "reuni": "reuni", "Reuni": "Reuni", "REUNI": "REUNI",
            "reunion": "reunion", "Reunion": "Reunion", "REUNION": "REUNION",
            "diversion": "diversion", "Diversion": "Diversion", "DIVERSION": "DIVERSION",
            "version": "version", "Version": "Version", "VERSION": "VERSION",
            "conversion": "conversion", "Conversion": "Conversion", "CONVERSION": "CONVERSION",
            "inversion": "inversion", "Inversion": "Inversion", "INVERSION": "INVERSION",
            "aversion": "aversion", "Aversion": "Aversion", "AVERSION": "AVERSION",
            "excursion": "excursion", "Excursion": "Excursion", "EXCURSION": "EXCURSION",
            "pasion": "pasion", "Pasion": "Pasion", "PASION": "PASION",
            "profesion": "profesion", "Profesion": "Profesion", "PROFESION": "PROFESION",
            "presion": "presion", "Presion": "Presion", "PRESION": "PRESION",
            "expresion": "expresion", "Expresion": "Expresion", "EXPRESION": "EXPRESION",
            "compresion": "compresion", "Compresion": "Compresion", "COMPRESION": "COMPRESION",
            "opresion": "opresion", "Opresion": "Opresion", "OPRESION": "OPRESION",
            "supresion": "supresion", "Supresion": "Supresion", "SUPRESION": "SUPRESION",
            "transgresion": "transgresion", "Transgresion": "Transgresion", "TRANSGRESION": "TRANSGRESION",
            "sesion": "sesion", "Sesion": "Sesion", "SESION": "SESION",
            "ascension": "ascension", "Ascension": "Ascension", "ASCENSION": "ASCENSION",
            "descension": "descension", "Descension": "Descension", "DESCENSION": "DESCENSION",
            "extension": "extension", "Extension": "Extension", "EXTENSION": "EXTENSION",
            "intension": "intension", "Intension": "Intension", "INTENSION": "INTENSION",
            "pension": "pension", "Pension": "Pension", "PENSION": "PENSION",
            "tension": "tension", "Tension": "Tension", "TENSION": "TENSION",
            "atencion": "atencion", "Atencion": "Atencion", "ATENCION": "ATENCION",
            "intencion": "intencion", "Intencion": "Intencion", "INTENCION": "INTENCION",
            "contencion": "contencion", "Contencion": "Contencion", "CONTENCION": "CONTENCION",
            "detencion": "detencion", "Detencion": "Detencion", "DETENCION": "DETENCION",
            "retencion": "retencion", "Retencion": "Retencion", "RETENCION": "RETENCION",
            "sustencion": "sustencion", "Sustencion": "Sustencion", "SUSTENCION": "SUSTENCION",
            "prevencion": "prevencion", "Prevencion": "Prevencion", "PREVENCION": "PREVENCION",
            "conveniencia": "conveniencia", "Conveniencia": "Conveniencia", "CONVENIENCIA": "CONVENIENCIA",
            "experiencia": "experiencia", "Experiencia": "Experiencia", "EXPERIENCIA": "EXPERIENCIA",
            "ciencia": "ciencia", "Ciencia": "Ciencia", "CIENCIA": "CIENCIA",
            "conciencia": "conciencia", "Conciencia": "Conciencia", "CONCIENCIA": "CONCIENCIA",
            "incienso": "incienso", "Incienso": "Incienso", "INCIENSO": "INCIENSO",
            "eficiencia": "eficiencia", "Eficiencia": "Eficiencia", "EFICIENCIA": "EFICIENCIA",
            "suficiencia": "suficiencia", "Suficiencia": "Suficiencia", "SUFICIENCIA": "SUFICIENCIA",
            "insuficiencia": "insuficiencia", "Insuficiencia": "Insuficiencia", "INSUFICIENCIA": "INSUFICIENCIA",
            "deficiencia": "deficiencia", "Deficiencia": "Deficiencia", "DEFICIENCIA": "DEFICIENCIA",
            "proficiencia": "proficiencia", "Proficiencia": "Proficiencia", "PROFICIENCIA": "PROFICIENCIA",
            "tendencia": "tendencia", "Tendencia": "Tendencia", "TENDENCIA": "TENDENCIA",
            "contendencia": "contendencia", "Contendencia": "Contendencia", "CONTENDENCIA": "CONTENDENCIA",
            "distendencia": "distendencia", "Distendencia": "Distendencia", "DISTENDENCIA": "DISTENDENCIA",
            "extravagancia": "extravagancia", "Extravagancia": "Extravagancia", "EXTRAVAGANCIA": "EXTRAVAGANCIA",
            "vagancia": "vagancia", "Vagancia": "Vagancia", "VAGANCIA": "VAGANCIA",
            "paciencia": "paciencia", "Paciencia": "Paciencia", "PACIENCIA": "PACIENCIA",
            "impaciencia": "impaciencia", "Impaciencia": "Impaciencia", "IMPACIENCIA": "IMPACIENCIA",
            "magnificencia": "magnificencia", "Magnificencia": "Magnificencia", "MAGNIFICENCIA": "MAGNIFICENCIA",
            "terrificencia": "terrificencia", "Terrificencia": "Terrificencia", "TERRIFICENCIA": "TERRIFICENCIA",
            "munificencia": "munificencia", "Munificencia": "Munificencia", "MUNIFICENCIA": "MUNIFICENCIA",
            "omnipotencia": "omnipotencia", "Omnipotencia": "Omnipotencia", "OMNIPOTENCIA": "OMNIPOTENCIA",
            "potencia": "potencia", "Potencia": "Potencia", "POTENCIA": "POTENCIA",
            "impotencia": "impotencia", "Impotencia": "Impotencia", "IMPOTENCIA": "IMPOTENCIA",
            "competencia": "competencia", "Competencia": "Competencia", "COMPETENCIA": "COMPETENCIA",
            "incompetencia": "incompetencia", "Incompetencia": "Incompetencia", "INCOMPETENCIA": "INCOMPETENCIA",
            "emergencia": "emergencia", "Emergencia": "Emergencia", "EMERGENCIA": "EMERGENCIA",
            "urgencia": "urgencia", "Urgencia": "Urgencia", "URGENCIA": "URGENCIA",
            "diligencia": "diligencia", "Diligencia": "Diligencia", "DILIGENCIA": "DILIGENCIA",
            "indiligencia": "indiligencia", "Indiligencia": "Indiligencia", "INDILIGENCIA": "INDILIGENCIA",
            "inteligencia": "inteligencia", "Inteligencia": "Inteligencia", "INTELIGENCIA": "INTELIGENCIA",
            "negligencia": "negligencia", "Negligencia": "Negligencia", "NEGLIGENCIA": "NEGLIGENCIA",
            "insignificancia": "insignificancia", "Insignificancia": "Insignificancia", "INSIGNIFICANCIA": "INSIGNIFICANCIA",
            "significancia": "significancia", "Significancia": "Significancia", "SIGNIFICANCIA": "SIGNIFICANCIA",
            "innocencia": "innocencia", "Innocencia": "Innocencia", "INNOCENCIA": "INNOCENCIA",
            "consecuencia": "consecuencia", "Consecuencia": "Consecuencia", "CONSECUENCIA": "CONSECUENCIA",
            "inconsecuencia": "inconsecuencia", "Inconsecuencia": "Inconsecuencia", "INCONSECUENCIA": "INCONSECUENCIA",
            "frecuencia": "frecuencia", "Frecuencia": "Frecuencia", "FRECUENCIA": "FRECUENCIA",
            "secuencia": "secuencia", "Secuencia": "Secuencia", "SECUENCIA": "SECUENCIA",
            "pertenencia": "pertenencia", "Pertenencia": "Pertenencia", "PERTENENCIA": "PERTENENCIA",
            "exigencia": "exigencia", "Exigencia": "Exigencia", "EXIGENCIA": "EXIGENCIA",
            "indulgencia": "indulgencia", "Indulgencia": "Indulgencia", "INDULGENCIA": "INDULGENCIA",
            "correspondencia": "correspondencia", "Correspondencia": "Correspondencia", "CORRESPONDENCIA": "CORRESPONDENCIA",
            "descendencia": "descendencia", "Descendencia": "Descendencia", "DESCENDENCIA": "DESCENDENCIA",
            "ascendencia": "ascendencia", "Ascendencia": "Ascendencia", "ASCENDENCIA": "ASCENDENCIA",
            "incidencia": "incidencia", "Incidencia": "Incidencia", "INCIDENCIA": "INCIDENCIA",
            "coincidencia": "coincidencia", "Coincidencia": "Coincidencia", "COINCIDENCIA": "COINCIDENCIA",
            "decadencia": "decadencia", "Decadencia": "Decadencia", "DECADENCIA": "DECADENCIA",
            "incandescencia": "incandescencia", "Incandescencia": "Incandescencia", "INCANDESCENCIA": "INCANDESCENCIA",
            "condescendencia": "condescendencia", "Condescendencia": "Condescendencia", "CONDESCENDENCIA": "CONDESCENDENCIA",
            "transcendencia": "transcendencia", "Transcendencia": "Transcendencia", "TRANSCENDENCIA": "TRANSCENDENCIA",
            "influencia": "influencia", "Influencia": "Influencia", "INFLUENCIA": "INFLUENCIA",
            "congruencia": "congruencia", "Congruencia": "Congruencia", "CONGRUENCIA": "CONGRUENCIA",
            "incongruencia": "incongruencia", "Incongruencia": "Incongruencia", "INCONGRUENCIA": "INCONGRUENCIA",
            "continuencia": "continuencia", "Continuencia": "Continuencia", "CONTINUENCIA": "CONTINUENCIA",
            "incontinuencia": "incontinuencia", "Incontinuencia": "Incontinuencia", "INCONTINUENCIA": "INCONTINUENCIA",
            "permanencia": "permanencia", "Permanencia": "Permanencia", "PERMANENCIA": "PERMANENCIA",
            "immanencia": "immanencia", "Immanencia": "Immanencia", "IMMANENCIA": "IMMANENCIA",
            "emanencia": "emanencia", "Emanencia": "Emanencia", "EMANENCIA": "EMANENCIA",
            "prominencia": "prominencia", "Prominencia": "Prominencia", "PROMINENCIA": "PROMINENCIA",
            "eminencia": "eminencia", "Eminencia": "Eminencia", "EMINENCIA": "EMINENCIA",
            "imminencia": "imminencia", "Imminencia": "Imminencia", "IMMINENCIA": "IMMINENCIA",
            "preeminencia": "preeminencia", "Preeminencia": "Preeminencia", "PREEMINENCIA": "PREEMINENCIA",
            "abstinencia": "abstinencia", "Abstinencia": "Abstinencia", "ABSTINENCIA": "ABSTINENCIA",
            "continencia": "continencia", "Continencia": "Continencia", "CONTINENCIA": "CONTINENCIA",
            "incontinencia": "incontinencia", "Incontinencia": "Incontinencia", "INCONTINENCIA": "INCONTINENCIA",
            "pertinencia": "pertinencia", "Pertinencia": "Pertinencia", "PERTINENCIA": "PERTINENCIA",
            "impertinencia": "impertinencia", "Impertinencia": "Impertinencia", "IMPERTINENCIA": "IMPERTINENCIA",
            "reticencia": "reticencia", "Reticencia": "Reticencia", "RETICENCIA": "RETICENCIA",
            "licencia": "licencia", "Licencia": "Licencia", "LICENCIA": "LICENCIA",
            "delicencia": "delicencia", "Delicencia": "Delicencia", "DELICENCIA": "DELICENCIA",
        }
        for wrong, correct in accent_corrections.items():
            result = result.replace(wrong, correct)
        result = re.sub(r'  +', ' ', result)
        result = re.sub(r' ([.,;:!?])', r'\1', result)
        return result


class GeminiService:
    MODELS = {
        "gemini-1.5-flash-lite": {"name": "3.1 Flash-Lite", "desc": "Respuestas rapidas"},
        "gemini-1.5-flash": {"name": "3.5 Flash", "desc": "Ayuda completa"},
        "gemini-1.5-pro": {"name": "3.1 Pro", "desc": "Advanced math and code"},
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
        response = requests.post(url, json=payload, timeout=60)
        response.raise_for_status()
        result = response.json()
        if "candidates" in result and len(result["candidates"]) > 0:
            return result["candidates"][0]["content"]["parts"][0]["text"]
        return ""

    def _extract_json(self, text):
        import json
        json_match = re.search(r'```json\s*(.*?)\s*```', text, re.DOTALL)
        if json_match:
            return json.loads(json_match.group(1))
        json_match = re.search(r'\{.*\}', text, re.DOTALL)
        if json_match:
            try:
                return json.loads(json_match.group())
            except:
                pass
        return {"generated_text": text}

    def generate_moc(self, problem, context="", equipo=""):
        if not self.api_key:
            return self._generate_local_moc(problem, context, equipo)

        prompt = f"""Eres un experto senior en gestion de cambios industriales (MoC) con 20 anos de experiencia en mineria, manufactura y operaciones industriales. Redactas documentos profesionales, impecables, con redaccion humanizada, natural y tecnica, sin errores ortograficos.

INSTRUCCIONES DE REDACCION:
- Usa lenguaje profesional pero natural, como lo haria un ingeniero senior experimentado.
- Evita frases genericas o robotizadas como "se identifico la siguiente situacion".
- Se especifico, detallado y tecnico. Incluye referencias a normas, procedimientos y mejores practicas.
- Usa conectores logicos, parrafos bien estructurados y vocabulario tecnico apropiado.
- Incluye datos cuantitativos cuando sea posible (tiempos, porcentajes, metricas).
- La redaccion debe parecer escrita por un profesional humano, no por una IA.

PROBLEMA REPORTADO: {problem}
CONTEXTO ADICIONAL: {context}
EQUIPO INVOLUCRADO: {equipo}

Genera en ESPANOL formato JSON con los siguientes campos detallados y humanizados:

1. descripcion_problema: Descripcion tecnica detallada del problema, con causas identificadas, impacto operacional, riesgos actuales y consecuencias si no se actua. Minimo 300 palabras. Redaccion fluida y profesional.

2. condicion_actual: Descripcion tecnica exhaustiva del estado actual del equipo/proceso, incluyendo especificaciones tecnicas, parametros operativos, limitaciones documentadas y referencias a normativas aplicables.

3. condicion_propuesta: Descripcion detallada de la solucion propuesta, incluyendo especificaciones tecnicas de la modificacion, beneficios esperados cuantificados, alineacion con estandares corporativos y mejores practicas de la industria.

4. razones_cambio: Lista numerada y justificada de las razones tecnicas, de seguridad, regulatorias y economicas que sustentan el cambio. Cada razon debe tener una explicacion de al menos 2-3 lineas.

5. alternativas_retorno: Analisis de alternativas evaluadas con pros y contras de cada una. Incluir plan de retorno detallado con pasos especificos, responsables y tiempos estimados.

6. recursos: Listado exhaustivo de recursos humanos (con roles y responsabilidades especificas), materiales (con especificaciones tecnicas), tecnicos (herramientas, software, documentacion) y financieros requeridos.

7. plan_implementacion: Plan detallado por fases con actividades especificas, responsables asignados, duracion estimada, hitos de control y criterios de aceptacion para cada fase. Minimo 4 fases bien definidas.

8. tiempo_duracion: Estimacion detallada del tiempo total con desglose por fase, consideraciones de ventanas de mantenimiento, contingencias y factores que pueden afectar la duracion.

9. riesgos_controles: Array de minimo 5 objetos {{"riesgo":"descripcion detallada del riesgo","control":"medida de control especifica con responsable y frecuencia"}}. Incluir riesgos tecnicos, operacionales y de calidad.

10. riesgos_shes: Array de minimo 5 objetos {{"riesgo":"descripcion detallada del riesgo SHES","control":"plan de accion especifico con medidas preventivas","plazo":"plazo de implementacion del control"}}. Cubrir seguridad, salud, medio ambiente y comunidad.

Responde SOLO JSON valido sin comentarios adicionales."""

        try:
            text = self._call_api(prompt, temperature=0.4, max_tokens=8192)
            result = self._extract_json(text)
            for key in result:
                if isinstance(result[key], str):
                    result[key] = Utils.correct_spelling_comprehensive(result[key])
                elif isinstance(result[key], list):
                    for item in result[key]:
                        if isinstance(item, dict):
                            for k in item:
                                if isinstance(item[k], str):
                                    item[k] = Utils.correct_spelling_comprehensive(item[k])
            return result
        except Exception as e:
            st.error(f"Error API: {e}. Usando generacion local.")
            return self._generate_local_moc(problem, context, equipo)

    def generate_a3(self, problem, context=""):
        if not self.api_key:
            return self._generate_local_a3(problem, context)

        prompt = f"""Eres un experto senior en metodologia A3 Lean con 15 anos de experiencia en mejora continua industrial. Redactas documentos A3 con redaccion humanizada, tecnica y profesional.

INSTRUCCIONES DE REDACCION:
- Usa lenguaje profesional, directo y tecnico como lo haria un Black Belt en Lean Six Sigma.
- Evita frases genericas. Se especifico con datos, metricas y analisis cuantitativos.
- Incluye referencias a herramientas Lean (5S, SMED, TPM, VSM, etc.) cuando aplique.
- La redaccion debe ser fluida, con parrafos bien estructurados y conectores logicos.
- Incluye datos hipoteticos pero realistas cuando el usuario no proporcione numeros especificos.

PROBLEMA REPORTADO: {problem}
CONTEXTO ADICIONAL: {context}

Genera en ESPANOL formato JSON con los siguientes campos detallados:

1. titulo: Titulo conciso y descriptivo de la mejora (maximo 10 palabras).

2. antecedentes: Contexto historico del problema, datos de linea base, tendencias observadas y por que es relevante abordarlo ahora. Minimo 200 palabras.

3. problema_actual: Descripcion detallada del problema con datos cuantitativos, frecuencia de ocurrencia, impacto en KPIs criticos y consecuencias operacionales. Minimo 250 palabras.

4. analisis_situacion: Analisis de la situacion actual usando datos, graficos conceptuales descritos en texto, comparativas con benchmarks de la industria y analisis de variabilidad del proceso.

5. objetivos: Objetivo general SMART y 3-5 objetivos especificos con metricas cuantificables, plazos y lineas base.

6. analisis_causa_raiz: Analisis de causa raiz detallado usando la metodologia de los 5 Porques, diagrama de Ishikawa conceptual descrito en texto, y validacion de hipotesis. Identificar la causa raiz fundamental.

7. contramedidas: Lista de 5-8 contramedidas especificas, priorizadas, con responsable asignado, fecha de implementacion y criterio de exito medible para cada una.

8. resultados_esperados: Resultados cuantificados esperados con proyecciones de ahorro, mejora en indicadores clave, retorno de inversion estimado y beneficios intangibles.

9. plan_seguimiento: Plan de seguimiento detallado con frecuencia de revision, indicadores a monitorear, responsables de seguimiento y criterios de exito a corto, mediano y largo plazo.

10. lecciones_aprendidas: Reflexiones sobre el proceso de analisis, desafios encontrados, aprendizajes clave y recomendaciones para futuras iniciativas similares.

11. estandarizacion: Plan de estandarizacion detallado con documentos a actualizar, capacitaciones requeridas, integracion al SGC y mecanismos de sostenibilidad.

Responde SOLO JSON valido sin comentarios adicionales."""

        try:
            text = self._call_api(prompt, temperature=0.4, max_tokens=8192)
            result = self._extract_json(text)
            for key in result:
                if isinstance(result[key], str):
                    result[key] = Utils.correct_spelling_comprehensive(result[key])
            return result
        except:
            return self._generate_local_a3(problem, context)

    def generate_kaizen(self, activity, context=""):
        if not self.api_key:
            return self._generate_local_kaizen(activity, context)

        prompt = f"""Eres un experto en Kaizen y Lean Manufacturing con amplia experiencia en gemba walks y mejora continua en operaciones industriales. Redactas registros Kaizen con redaccion humanizada, practica y motivadora.

INSTRUCCIONES DE REDACCION:
- Usa lenguaje practico, directo y motivador como lo haria un lider de mejora continua en el gemba.
- Incluye datos cuantitativos especificos: tiempos antes/despues, cantidades, porcentajes de mejora.
- Describe la situacion actual con detalle visual para que el lector pueda imaginar el antes y despues.
- La redaccion debe ser natural, con frases cortas y claras, evitando tecnicismos innecesarios.
- Incluye el impacto humano: como beneficia al operario, al equipo y a la organizacion.

ACTIVIDAD DE MEJORA: {activity}
CONTEXTO ADICIONAL: {context}

Genera en ESPANOL formato JSON con los siguientes campos detallados:

1. titulo: Titulo atractivo y descriptivo de la mejora Kaizen (maximo 8 palabras).

2. area: Area especifica donde se implemento la mejora.

3. descripcion_problema: Descripcion vivida del problema o desperdicio identificado, con datos cuantitativos del antes (tiempos, movimientos, distancias, cantidad de material). Minimo 200 palabras.

4. solucion: Descripcion detallada de la solucion implementada paso a paso, materiales utilizados, tiempo de implementacion, participantes y desafios superados. Minimo 200 palabras.

5. beneficios: Lista de beneficios cuantificados y cualitativos con datos antes/despues, ahorros estimados, mejoras en seguridad, calidad, productividad y ambiente de trabajo.

6. tipo_desperdicio: Tipo(s) de desperdicio Lean eliminado(s) de la lista: Motion, Skills, Inventory, Transportation, Over Production, Over Processing, Waiting, Defects.

7. impacto_bto: Categoria BTO impactada: Safe and Sustainable, People & Culture, Network Optimisation, Supply Chain and Manufacturing Excellence.

8. proximos_pasos: Plan de accion concretos para sostener la mejora, replicarla en otras areas, reconocer al equipo y establecer el nuevo estandar.

Responde SOLO JSON valido sin comentarios adicionales."""

        try:
            text = self._call_api(prompt, temperature=0.4, max_tokens=4096)
            result = self._extract_json(text)
            for key in result:
                if isinstance(result[key], str):
                    result[key] = Utils.correct_spelling_comprehensive(result[key])
            return result
        except:
            return self._generate_local_kaizen(activity, context)

    def translate_document(self, data):
        if not self.api_key:
            return data
        prompt = f"""Traduce del espanol al ingles profesional industrial, manteniendo la terminologia tecnica apropiada (OSHA, ISO, ANSI, etc.):
{json.dumps(data, ensure_ascii=False, indent=2)}
Responde SOLO el JSON traducido, misma estructura exacta."""
        try:
            text = self._call_api(prompt, temperature=0.2, max_tokens=8192)
            return self._extract_json(text)
        except:
            return data

    def correct_spelling(self, text):
        if not self.api_key or not text.strip():
            return Utils.correct_spelling_comprehensive(text)
        prompt = f"""Corrige ortografia, gramatica, puntuacion y mejora la redaccion del siguiente texto en espanol. Manten el significado tecnico exacto. Mejora la fluidez y naturalidad sin hacerlo robotico. Devuelve SOLO el texto corregido, sin explicaciones.

TEXTO:
{text}"""
        try:
            corrected = self._call_api(prompt, temperature=0.2, max_tokens=4096).strip()
            return Utils.correct_spelling_comprehensive(corrected)
        except:
            return Utils.correct_spelling_comprehensive(text)

    def _generate_local_moc(self, problem, context, equipo):
        return {
            "descripcion_problema": "Se ha identificado una condicion critica que requiere gestion formal mediante el proceso de Management of Change (MoC). El problema reportado es: " + problem + ".\n\nEsta situacion presenta riesgos significativos para la seguridad operacional, la integridad del proceso y la continuidad de la produccion. Durante las evaluaciones preliminares se ha determinado que la condicion actual no cumple con los estandares corporativos de seguridad y calidad establecidos, generando una exposicion potencial al personal operativo y a los equipos criticos.\n\nEs imperativo implementar un cambio controlado y documentado que mitigue los riesgos identificados, garantice el cumplimiento normativo y restablezca las condiciones operativas seguras y eficientes del proceso.",
            "condicion_actual": "El equipo o proceso actual opera bajo condiciones que presentan las siguientes limitaciones tecnicas documentadas: " + context + ".\n\nSe han identificado deficiencias en la configuracion actual que afectan directamente el rendimiento operativo y la seguridad del personal. Los parametros criticos del proceso se encuentran fuera de los rangos optimos establecidos en los procedimientos operativos estandar (SOP).\n\nSe requiere una evaluacion tecnica exhaustiva para establecer una linea base de referencia completa antes de proceder con cualquier modificacion, asegurando que todos los cambios sean trazables y verificables.",
            "condicion_propuesta": "Se propone implementar modificaciones tecnicas estructuradas que optimicen el rendimiento operativo del equipo critico, mejoren significativamente las condiciones de seguridad del proceso y alineen las operaciones con los estandares corporativos y regulatorios vigentes.\n\nLa propuesta incluye la actualizacion de componentes criticos, la implementacion de controles adicionales de seguridad, la estandarizacion de procedimientos operativos y la capacitacion del personal involucrado. Todas las modificaciones seran disenadas siguiendo las mejores practicas de la industria y los requisitos normativos aplicables.",
            "razones_cambio": "1. SEGURIDAD OPERACIONAL: La condicion actual presenta riesgos identificados que pueden comprometer la integridad fisica del personal. La implementacion del cambio reducira significativamente la probabilidad de incidentes y accidentes laborales, alineandose con la politica de cero accidentes de la organizacion.\n\n2. OPTIMIZACION DE RENDIMIENTO: El equipo critico opera por debajo de su capacidad optima debido a las limitaciones tecnicas identificadas. El cambio propuesto mejorara la confiabilidad, disponibilidad y eficiencia del equipo, reduciendo tiempos de parada no planificados.\n\n3. CUMPLIMIENTO NORMATIVO: La modificacion asegura el cumplimiento de estandares corporativos, regulaciones nacionales e internacionales aplicables al sector industrial, evitando sanciones y manteniendo la licencia operativa.\n\n4. REDUCCION DE RIESGOS SHES: Las evaluaciones previas han identificado riesgos en seguridad, salud y medio ambiente que seran mitigados proactivamente con las medidas de control propuestas en este documento.\n\n5. MEJORA CONTINUA: El cambio esta alineado con los objetivos estrategicos de la organizacion en materia de excelencia operacional, sostenibilidad y mejora continua.",
            "alternativas_retorno": "ALTERNATIVAS EVALUADAS:\n\n1. MANTENIMIENTO CORRECTIVO TRADICIONAL (DESCARTADO): Aunque de menor costo inicial, presenta un alcance limitado que no aborda las causas raiz del problema. La recurrencia de fallas seria alta, generando costos operacionales mayores a largo plazo.\n\n2. REEMPLAZO TOTAL DEL SISTEMA (DESCARTADO): Ofrece la solucion mas completa pero con un costo de inversion elevado que excede el presupuesto aprobado para este periodo. Ademas, el tiempo de implementacion seria excesivo para las necesidades operativas actuales.\n\n3. MODIFICACION CONTROLADA (SELECCIONADA): Representa la mejor relacion costo-beneficio, abordando las causas raiz identificadas con un alcance definido, tiempos de implementacion razonables y un retorno de inversion favorable dentro del primer ano.\n\nPLAN DE RETORNO:\nEn caso de que el cambio no produzca los resultados esperados o se presenten complicaciones durante la implementacion, se ejecutara el siguiente plan de retorno: Restauracion inmediata de la configuracion original del equipo, activacion del protocolo de contingencia establecido, notificacion oportuna a supervision directa y areas de apoyo, documentacion detallada de las lecciones aprendidas y analisis de causa raiz de la falla para prevenir recurrencias.",
            "recursos": "RECURSOS HUMANOS REQUERIDOS:\n- Tecnico especializado de mantenimiento mecanico/electronico (1 persona, tiempo completo durante implementacion)\n- Supervisor de area operativa (1 persona, supervision continua)\n- Especialista SHES (1 persona, verificacion de controles y permisos)\n- Operador de area certificado (1-2 personas, apoyo operativo y pruebas)\n- Ingeniero de procesos (1 persona, validacion tecnica y ajustes de parametros)\n\nRECURSOS MATERIALES:\n- Herramientas especializadas certificadas y calibradas\n- Repuestos de calidad certificada con trazabilidad\n- EPP completo: casco de seguridad, gafas de proteccion, guantes anticorte, botas dielectricas, arnes cuando aplique\n- Materiales de senalizacion, demarcacion y etiquetado\n- Materiales de limpieza y preparacion de area\n\nRECURSOS TECNICOS:\n- Documentacion tecnica actualizada del equipo (manuales, diagramas, especificaciones)\n- SOP vigentes y procedimientos de trabajo seguro\n- Permisos de trabajo segun tipo (trabajo en caliente, espacio confinado, trabajo en altura, etc.)\n- Checklist de verificacion pre y post implementacion\n- Equipos de prueba y medicion calibrados",
            "plan_implementacion": "FASE 1: PREPARACION Y PLANIFICACION (Dias 1-2)\n- Reunion de coordinacion multidisciplinaria con produccion, mantenimiento y SHES\n- Verificacion exhaustiva de disponibilidad de todos los recursos materiales y humanos\n- Preparacion del area de trabajo: limpieza profunda, senalizacion de perimetro, aplicacion de LOTO (Lock Out Tag Out)\n- Briefing de seguridad con todo el equipo involucrado, revision de riesgos y controles\n- Verificacion final de permisos de trabajo y autorizaciones requeridas\n\nFASE 2: EJECUCION DE MODIFICACIONES (Dias 3-5)\n- Implementacion progresiva de las modificaciones tecnicas segun plan detallado\n- Pruebas funcionales iniciales despues de cada sub-etapa critica\n- Registro fotografico detallado del antes, durante y despues de cada modificacion\n- Verificacion intermedia SHES al finalizar cada dia de trabajo\n- Comunicacion continua con supervision de produccion sobre avances\n\nFASE 3: VALIDACION Y PRUEBAS (Dias 6-7)\n- Pruebas funcionales bajo condiciones normales de operacion\n- Verificacion de todos los parametros criticos del proceso contra especificaciones\n- Validacion conjunta por supervisor de area, produccion y especialista tecnico\n- Pruebas de estres y verificacion de limites operativos\n- Documentacion de resultados de pruebas y ajustes finales\n\nFASE 4: CIERRE Y ESTANDARIZACION (Dia 8)\n- Actualizacion completa de toda la documentacion tecnica y operativa\n- Capacitacion formal al personal operativo sobre nuevos procedimientos\n- Socializacion de lecciones aprendidas con todas las areas involucradas\n- Cierre formal del MoC con firmas de aprobacion de todas las partes\n- Archivo del documento completo en el sistema de gestion documental",
            "tiempo_duracion": "ESTIMACION TOTAL DEL CAMBIO: 8 dias habiles distribuidos en 4 fases bien definidas.\n\nDESGLOSE POR FASE:\n- Fase 1 (Preparacion): 2 dias\n- Fase 2 (Ejecucion): 3 dias\n- Fase 3 (Validacion): 2 dias\n- Fase 4 (Cierre): 1 dia\n\nCONSIDERACIONES:\nLa duracion puede ajustarse segun condiciones operativas, disponibilidad de recursos y resultados de las verificaciones intermedias. Se ha incluido un margen de contingencia del 20% para imprevistos. Las ventanas de mantenimiento seran coordinadas con produccion con al menos 48 horas de anticipacion.",
            "riesgos_controles": [
                {"riesgo": "Interrupcion del proceso productivo durante la implementacion de modificaciones, generando perdidas de produccion estimadas", "control": "Coordinacion previa detallada con produccion para definir ventana de mantenimiento planificado. Comunicacion oportuna a todas las areas involucradas con 24 horas de anticipacion. Monitoreo continuo del plan de produccion durante la ejecucion."},
                {"riesgo": "Falla tecnica durante la modificacion que pueda afectar equipos adyacentes o sistemas interconectados", "control": "Verificacion previa exhaustiva de todas las interconexiones. Disponibilidad inmediata de repuestos de emergencia. Supervision tecnica continua por ingeniero senior. Protocolo de parada de emergencia activo durante toda la ejecucion."},
                {"riesgo": "Exposicion a riesgos de seguridad durante la ejecucion de trabajos en campo (golpes, cortes, caidas)", "control": "Permisos de trabajo especificos segun tipo de riesgo identificado. Uso obligatorio y verificado de EPP completo. Supervision continua por especialista SHES. Aplicacion estricta de LOTO en todos los puntos de energia."},
                {"riesgo": "Error humano durante la implementacion que genere dano al equipo o configuracion incorrecta", "control": "Checklist de verificacion paso a paso firmado por tecnico y supervisor. Doble verificacion critica (dos personas) en puntos de control clave. Registro fotografico de cada etapa para trazabilidad."},
                {"riesgo": "Demora en la entrega de repuestos o materiales criticos que retrase la implementacion", "control": "Verificacion de disponibilidad de materiales 48 horas antes del inicio. Identificacion de proveedores alternativos. Stock de seguridad de componentes criticos. Plan de contingencia con materiales sustitutos pre-aprobados."}
            ],
            "riesgos_shes": [
                {"riesgo": "Lesiones por manipulacion manual de equipos, componentes pesados o herramientas durante la ejecucion", "control": "Capacitacion especifica en tecnicas de levantamiento seguro antes del inicio. Uso obligatorio de EPP completo incluyendo guantes anticorte y calzado de seguridad. Senalizacion clara del area de trabajo. Supervisor SHES presente durante trabajos de alto riesgo.", "plazo": "Antes del inicio de actividades"},
                {"riesgo": "Generacion de residuos solidos, liquidos o peligrosos durante el proceso de modificacion", "control": "Manejo seguro segun procedimiento ambiental corporativo. Clasificacion en origen de todos los residuos generados. Disposicion unicamente en areas autorizadas y con registro de trazabilidad. Contenedores identificados y segregados.", "plazo": "Durante toda la ejecucion"},
                {"riesgo": "Exposicion a ruido excesivo, vibraciones o agentes quimicos durante trabajos de modificacion", "control": "Monitoreo continuo de niveles de ruido y agentes quimicos. Uso obligatorio de protectores auditivos cuando los niveles excedan 85 dB. Ventilacion adecuada en areas cerradas. Limitacion de horario de exposicion segun limites permisibles legales.", "plazo": "Durante toda la ejecucion"},
                {"riesgo": "Incendio o explosion por trabajo en caliente, chispas o acumulacion de vapores inflamables", "control": "Permiso de trabajo en caliente con analisis de atmosfera previo. Vigia de fuego designado y capacitado. Extintores portatiles disponibles y verificados. Limpieza del area de materiales combustibles antes del inicio. Monitoreo continuo de gases inflamables.", "plazo": "Durante trabajos de soldadura/corte"},
                {"riesgo": "Contaminacion del suelo o cuerpos de agua por derrames accidentales de lubricantes, solventes o productos quimicos", "control": "Kit de contencion de derrames disponible en el area. Uso de bandejas de retencion bajo equipos que manipulen liquidos. Prohibicion de drenaje directo a sistemas de alcantarillado. Limpieza inmediata de cualquier derrame con materiales absorbentes aprobados.", "plazo": "Durante toda la ejecucion"}
            ]
        }

    def _generate_local_a3(self, problem, context):
        return {
            "titulo": "Optimizacion del proceso: " + problem[:50],
            "antecedentes": "Durante los ultimos seis meses, el area operativa ha experimentado una degradacion progresiva en sus indicadores clave de desempeno. Se han registrado incrementos en tiempos de ciclo, aumento en la tasa de defectos y una reduccion en la productividad general del proceso.\n\nEl analisis preliminar de datos historicos revela una tendencia creciente que, si no se aborda de manera estructurada, comprometera los objetivos anuales de la organizacion. La metodologia A3 fue seleccionada como herramienta principal para el analisis estructurado de esta situacion, permitiendo una vision integral del problema y facilitando la identificacion de soluciones sostenibles.",
            "problema_actual": problem,
            "analisis_situacion": "La situacion actual presenta multiples indicadores de desempeno con oportunidades significativas de mejora. Se requiere una recopilacion sistematica y rigurosa de datos para establecer una linea base solida que permita cuantificar el impacto de las contramedidas propuestas.\n\nEl analisis de variabilidad del proceso muestra fluctuaciones que exceden los limites de control establecidos, indicando la presencia de causas especiales que deben ser identificadas y eliminadas. La comparativa con benchmarks de la industria revela una brecha de desempeno del 15-25% respecto a los mejores en clase.",
            "objetivos": "OBJETIVO GENERAL:\nOptimizar integralmente el proceso eliminando los desperdicios identificados y estableciendo un nuevo estandar de desempeno sostenible.\n\nOBJETIVOS ESPECIFICOS (SMART):\n1. Reducir el tiempo de ciclo en un 15% dentro de los proximos 3 meses, pasando de 45 minutos a 38 minutos por unidad.\n2. Disminuir la tasa de defectos en un 20% durante el proximo trimestre, reduciendo de 8% a 6.4%.\n3. Mejorar la productividad general del area en un 10% dentro de 6 meses.\n4. Incrementar la satisfaccion interna del cliente (area siguiente) en un 25% segun encuesta trimestral.\n5. Reducir el costo operativo unitario en un 8% dentro del primer ano de implementacion.",
            "analisis_causa_raiz": "ANALISIS DE LOS 5 PORQUES:\n1. POR QUE ocurre el problema? -> Porque el proceso opera con una configuracion inadecuada que genera variabilidad excesiva.\n2. POR QUE la configuracion es inadecuada? -> Porque no existe una estandarizacion formal de los parametros operativos criticos.\n3. POR QUE no hay estandarizacion? -> Porque los procedimientos operativos estandar (SOP) no han sido actualizados en los ultimos 18 meses.\n4. POR QUE no estan actualizados? -> Porque no existe un sistema de gestion documental efectivo que asegure la revision periodica.\n5. POR QUE no hay sistema? -> Porque falta una politica clara de gestion del conocimiento y mejora continua con responsables asignados.\n\nCAUSA RAIZ IDENTIFICADA:\nAusencia de un sistema integral de gestion, actualizacion y control de SOP, combinado con la falta de responsables claros y metricas de seguimiento del desempeno del proceso.",
            "contramedidas": "1. ACTUALIZAR SOP DEL PROCESO: Revisar y actualizar todos los procedimientos operativos con instrucciones claras, paso a paso, con fotos de referencia y puntos de control critico. Responsable: Ingeniero de Procesos. Plazo: 2 semanas.\n\n2. IMPLEMENTAR CHECKLISTS DIARIOS: Disenar y desplegar checklists de verificacion diaria en cada puesto de trabajo para asegurar el cumplimiento de estandares. Responsable: Supervisor de Area. Plazo: 1 semana.\n\n3. CAPACITAR AL PERSONAL: Programar y ejecutar capacitaciones formales sobre los nuevos estandares, con evaluacion de competencias y certificacion. Responsable: Especialista de Capacitacion. Plazo: 3 semanas.\n\n4. ESTABLECER KPIs VISUALES: Implementar tableros visuales en el area con indicadores clave de desempeno actualizados diariamente. Responsable: Lider de Mejora Continua. Plazo: 2 semanas.\n\n5. PROGRAMAR AUDITORIAS MENSUALES: Establecer auditorias formales mensuales de cumplimiento con criterios de evaluacion definidos y plan de accion para desviaciones. Responsable: Auditor Interno. Plazo: Inicio inmediato, recurrente.\n\n6. IMPLEMENTAR SISTEMA DE GESTION DOCUMENTAL: Desarrollar o adquirir una solucion digital para control de versiones, aprobaciones y distribucion de documentos. Responsable: IT + Calidad. Plazo: 2 meses.\n\n7. DEFINIR RESPONSABLES DE ESTANDARES: Asignar responsables claros por area para la gestion, actualizacion y seguimiento de estandares operativos. Responsable: Gerente de Operaciones. Plazo: 1 semana.",
            "resultados_esperados": "- Reduccion medible y sostenida de desperdicios identificados (Motion, Waiting, Defects)\n- Mejora sostenida en calidad del producto y consistencia del proceso\n- Estandarizacion efectiva que reduzca la variabilidad operativa en al menos 30%\n- Reduccion del tiempo de ciclo en 15% con impacto directo en capacidad productiva\n- Incremento en satisfaccion del cliente interno medido mediante encuestas\n- Fortalecimiento de la cultura de mejora continua y trabajo en equipo\n- Retorno de inversion estimado del 180% dentro del primer ano\n- Reduccion de costos operativos unitarios en 8%\n- Mejora en el ambiente de trabajo y motivacion del personal",
            "plan_seguimiento": "SEMANA 1-2: Implementacion de contramedidas 1 y 2 (SOP y checklists). Monitoreo diario de cumplimiento.\n\nSEMANA 3-4: Ejecucion de capacitaciones (contramedida 3). Evaluacion de competencias. Monitoreo inicial de indicadores. Ajustes segun resultados.\n\nMES 2: Primera auditoria formal (contramedida 5). Evaluacion de avance vs. objetivos iniciales. Implementacion de KPIs visuales.\n\nMES 3: Evaluacion integral vs. objetivos SMART establecidos. Analisis de tendencias. Ajustes a contramedidas si es necesario.\n\nMES 6: Revision de sostenibilidad de mejoras. Analisis de replicabilidad en otras areas. Reconocimiento al equipo.\n\nTRIMESTRAL: Revisiones formales con gerencia. Actualizacion de objetivos segun evolucion del proceso.",
            "lecciones_aprendidas": "La aplicacion de la metodologia A3 permitio visualizar de manera integral la complejidad del problema y las interconexiones entre sus multiples causas. La participacion activa y multidisciplinaria del equipo fue fundamental para identificar la causa raiz real, que inicialmente no era evidente.\n\nSe aprendio que los problemas aparentemente tecnicos frecuentemente tienen raices en sistemas de gestion deficientes. La inversion en capacitacion y estandarizacion genera retornos significativos a mediano plazo. La visualizacion de datos y el seguimiento constante son criticos para mantener las mejoras.",
            "estandarizacion": "Los procedimientos actualizados seran documentados formalmente con control de versiones, aprobados por gerencia de operaciones y calidad, socializados mediante capacitaciones estructuradas con evaluacion de competencias, integrados al Sistema de Gestion de Calidad (SGC) existente y sujetos a revision periodica anual como minimo. Se estableceran metricas de cumplimiento y auditorias programadas para asegurar la sostenibilidad de los estandares implementados."
        }

    def _generate_local_kaizen(self, activity, context):
        return {
            "titulo": "Kaizen: " + activity[:50],
            "area": "Area de Mantenimiento / Produccion / Calidad (especificar segun contexto)",
            "descripcion_problema": activity + "\n\nDurante las actividades diarias de gemba walk, el equipo identifico esta oportunidad de mejora que representa un desperdicio significativo en el proceso. La situacion actual genera movimientos innecesarios, tiempos de espera o riesgos de calidad que impactan directamente en la productividad del area y en la satisfaccion del personal.\n\nSe realizo un analisis rapido de la situacion que confirmo la viabilidad de implementar una mejora inmediata con recursos disponibles en el area, siguiendo el principio fundamental del Kaizen: mejorar un poco cada dia.",
            "solucion": "Se implemento una mejora estructurada orientada a eliminar el desperdicio identificado y optimizar el flujo del proceso, aplicando principios fundamentales de Lean Manufacturing y el pensamiento Kaizen de mejora continua.\n\nLa solucion fue disenada y ejecutada por el equipo de trabajo del area con apoyo del lider de mejora continua, utilizando materiales disponibles y aplicando el concepto de low cost, high impact. Se realizaron pruebas piloto antes de la implementacion definitiva para validar la efectividad de la propuesta.\n\nEl equipo documento el antes y despues con fotografias y mediciones de tiempo para cuantificar el impacto de la mejora implementada.",
            "beneficios": "- Reduccion del tiempo de ejecucion en aproximadamente 20-30%\n- Mejora significativa en calidad y consistencia del proceso\n- Mayor seguridad para el personal al eliminar movimientos riesgosos\n- Reduccion de costos operativos derivados de la eliminacion de desperdicios\n- Mejora en el ambiente de trabajo y orden del area\n- Eliminacion de movimientos innecesarios y tiempos de busqueda\n- Incremento en la motivacion del equipo al ver resultados inmediatos\n- Facil replicabilidad en otras areas similares",
            "tipo_desperdicio": "Motion / Waiting / Skills (seleccionar segun analisis especifico del desperdicio identificado)",
            "impacto_bto": "Supply Chain and Manufacturing Excellence / Safe and Sustainable (seleccionar segun el impacto principal de la mejora)",
            "proximos_pasos": "1. Documentar formalmente la mejora con fotografias, descripcion detallada y datos de impacto\n2. Socializar la mejora con otras areas relacionadas mediante presentacion breve en reunion de coordinacion\n3. Replicar la mejora en procesos similares identificados durante el gemba walk\n4. Establecer monitoreo mensual para asegurar que la mejora se mantiene en el tiempo\n5. Reconocer formalmente al equipo participante en la mejora\n6. Integrar el nuevo estandar al SOP del area\n7. Programar revision de sostenibilidad a los 3 meses"
        }


def replace_text_in_shape(shape, old_text, new_text):
    """Reemplaza texto en un shape preservando el formato de los runs"""
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
    """Reemplaza multiples textos en toda la presentacion"""
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
    """Llena una celda de tabla preservando formato"""
    if cell.text_frame.paragraphs:
        first_para = cell.text_frame.paragraphs[0]
        if first_para.runs:
            first_run = first_para.runs[0]
            first_run.text = str(text)
        else:
            first_para.text = str(text)
    else:
        cell.text = str(text)


class DocumentGenerator:
    """Genera documentos usando templates cargados en memoria"""

    def generate_moc(self, data, images=None, template_bytes=None):
        """Genera MoC desde template en memoria"""
        if template_bytes is None:
            st.error("Template MoC no cargado. Vaya a Configuracion > Templates.")
            return None

        prs = Presentation(BytesIO(template_bytes))

        replacements = {
            "MOC: Control de Acceso a panel de control (HMI) – carga de detonadores (219)": f"MOC: {data.get('moc_title', '')}",
            "MAYO 2026": data.get('fecha', Utils.format_date()),
            "nÚmero DE LA MOC: XXXXXXXXXXXXX": f"nÚmero DE LA MOC: {data.get('moc_number', '')}",
            "Naturaleza de la moc: permanente": f"Naturaleza de la moc: {data.get('naturaleza', 'permanente')}",
            "ORIGINADOR DE LA MOC: ROGER HUAMANI": f"ORIGINADOR DE LA MOC: {data.get('originador', '')}",
            "Producción: \nSpecialist SHES: \nMantenimiento:": f"Producción: {data.get('produccion', '')}\nSpecialist SHES: {data.get('specialist_shes', '')}\nMantenimiento: {data.get('mantenimiento', '')}",
            "Revisor 1:\nRevisor 2:\nRevisor 3:\nRevisor 4:\nAprobador Final:": f"Revisor 1: {data.get('revisores', '')}\nRevisor 2:\nRevisor 3:\nRevisor 4:\nAprobador Final:",
            "Experto aprobador:": f"Experto aprobador: {data.get('experto_aprobador', '')}",
        }

        replace_all_text_in_presentation(prs, replacements)

        # SLIDE 3: Descripcion del Problema
        slide3 = prs.slides[2]
        for shape in slide3.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if run.text.strip() == "3":
                            run.text = data.get('descripcion_problema', '')

        # SLIDE 4: Tabla Condicion Actual / Condicion Propuesta
        slide4 = prs.slides[3]
        for shape in slide4.shapes:
            if shape.has_table:
                table = shape.table
                if len(table.rows) >= 2 and len(table.columns) >= 2:
                    fill_table_cell(table.cell(1, 0), data.get('condicion_actual', ''))
                    fill_table_cell(table.cell(1, 1), data.get('condicion_propuesta', ''))

        # SLIDE 5: Razones del Cambio
        slide5 = prs.slides[4]
        for shape in slide5.shapes:
            if shape.has_text_frame:
                text = shape.text_frame.text
                if text.strip() == "5":
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            if run.text.strip() == "5":
                                run.text = data.get('razones_cambio', '')
                if "Alternativas" in text:
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            if "Alternativas" in run.text:
                                run.text = data.get('alternativas_retorno', '')

        # SLIDE 6: Recursos y Plan
        slide6 = prs.slides[5]
        for shape in slide6.shapes:
            if shape.has_text_frame:
                text = shape.text_frame.text
                if "Recursos:" in text:
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            if "Recursos:" in run.text:
                                run.text = f"Recursos:\n{data.get('recursos', '')}"

        # SLIDE 7: Tiempo
        slide7 = prs.slides[6]
        for shape in slide7.shapes:
            if shape.has_text_frame:
                text = shape.text_frame.text
                if text.strip() == "7" or "Tiempo" in text:
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            if run.text.strip() == "7":
                                run.text = data.get('tiempo_duracion', '')

        # SLIDE 8: Riesgos (Tabla)
        slide8 = prs.slides[7]
        for shape in slide8.shapes:
            if shape.has_table:
                table = shape.table
                risks = data.get('riesgos_controles', [])
                for i, risk in enumerate(risks):
                    row_idx = i + 1
                    if row_idx < len(table.rows):
                        fill_table_cell(table.cell(row_idx, 0), str(i + 1))
                        fill_table_cell(table.cell(row_idx, 1), risk.get('riesgo', ''))
                        fill_table_cell(table.cell(row_idx, 2), risk.get('control', ''))

        # SLIDE 9: Riesgos SHES
        slide9 = prs.slides[8]
        for shape in slide9.shapes:
            if shape.has_table:
                table = shape.table
                risks = data.get('riesgos_shes', [])
                for i, risk in enumerate(risks):
                    row_idx = i + 1
                    if row_idx < len(table.rows):
                        fill_table_cell(table.cell(row_idx, 0), str(i + 1))
                        fill_table_cell(table.cell(row_idx, 1), risk.get('riesgo', ''))
                        fill_table_cell(table.cell(row_idx, 2), risk.get('control', ''))
                        fill_table_cell(table.cell(row_idx, 3), risk.get('plazo', ''))

        # Agregar imagenes de soporte
        if images:
            for idx, img_info in enumerate(images, 1):
                blank_layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]
                new_slide = prs.slides.add_slide(blank_layout)

                title_box = new_slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.5))
                tf = title_box.text_frame
                tf.text = f"Figura {idx} - Imagen de Soporte"
                for p in tf.paragraphs:
                    p.font.size = Pt(18)
                    p.font.bold = True
                    p.font.color.rgb = RGBColor(0x1a, 0x5f, 0x7a)

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

                    desc = img_info.get("desc", f"Figura {idx}") if isinstance(img_info, dict) else f"Figura {idx}"
                    desc_box = new_slide.shapes.add_textbox(Inches(0.5), Inches(7.3), Inches(9), Inches(0.5))
                    dtf = desc_box.text_frame
                    dtf.text = desc
                    for p in dtf.paragraphs:
                        p.font.size = Pt(11)
                        p.font.italic = True
                        p.font.color.rgb = RGBColor(0x47, 0x55, 0x69)
                        p.alignment = PP_ALIGN.CENTER
                except Exception as e:
                    st.warning(f"Error con imagen {idx}: {e}")

        output_buffer = BytesIO()
        prs.save(output_buffer)
        output_buffer.seek(0)
        return output_buffer

    def generate_a3(self, data, images=None, template_bytes=None):
        """Genera A3 desde template en memoria"""
        if template_bytes is None:
            st.error("Template A3 no cargado. Vaya a Configuracion > Templates.")
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
            ("ANTECEDENTES", "antecedentes"),
            ("PROBLEMA ACTUAL", "problema_actual"),
            ("ANALISIS DE LA SITUACION", "analisis_situacion"),
            ("OBJETIVOS", "objetivos"),
            ("ANALISIS DE CAUSA RAIZ", "analisis_causa_raiz"),
            ("CONTRAMEDIDAS", "contramedidas"),
            ("RESULTADOS ESPERADOS", "resultados_esperados"),
            ("PLAN DE SEGUIMIENTO", "plan_seguimiento"),
            ("LECCIONES APRENDIDAS", "lecciones_aprendidas"),
            ("ESTANDARIZACION", "estandarizacion"),
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
            doc.add_heading('IMAGENES DE SOPORTE', level=1)
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
        """Genera Kaizen desde template en memoria"""
        if template_bytes is None:
            st.error("Template Kaizen no cargado. Vaya a Configuracion > Templates.")
            return None

        prs = Presentation(BytesIO(template_bytes))
        slide = prs.slides[0]

        for shape in slide.shapes:
            if shape.has_table:
                table = shape.table
                tipo_desp = data.get('tipo_desperdicio', '')
                waste_map = {
                    'motion': 1, 'skills': 2, 'inventory': 3, 'transportation': 4,
                    'over production': 5, 'over processing': 6, 'waiting': 7, 'defects': 8
                }
                for row_idx in range(1, len(table.rows)):
                    waste_cell = table.cell(row_idx, 1).text.strip().lower()
                    for waste_key, mapped_row in waste_map.items():
                        if waste_key in tipo_desp.lower() and row_idx == mapped_row:
                            table.cell(row_idx, 2).text = "X"
                            for para in table.cell(row_idx, 2).text_frame.paragraphs:
                                para.font.bold = True
                                para.font.size = Pt(14)
                                para.alignment = PP_ALIGN.CENTER

            if shape.has_text_frame:
                text = shape.text_frame.text
                if "Leader" in text and "Picture" in text:
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            if "Leader" in run.text:
                                run.text = f"Leader: {data.get('autor', '')}\nArea: {data.get('area', '')}\nFecha: {data.get('fecha', '')}"

        blank_layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]
        detail_slide = prs.slides.add_slide(blank_layout)

        title_box = detail_slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.5))
        tf = title_box.text_frame
        tf.text = data.get('titulo', 'Kaizen')
        for p in tf.paragraphs:
            p.font.size = Pt(24)
            p.font.bold = True
            p.font.color.rgb = RGBColor(0x1a, 0x5f, 0x7a)

        sections = [
            ("Descripcion del Problema", "descripcion_problema"),
            ("Solucion Implementada", "solucion"),
            ("Beneficios", "beneficios"),
            ("Proximos Pasos", "proximos_pasos"),
        ]

        y_pos = 1.0
        for section_title, key in sections:
            sec_box = detail_slide.shapes.add_textbox(Inches(0.5), Inches(y_pos), Inches(9), Inches(0.3))
            stf = sec_box.text_frame
            stf.text = section_title
            for p in stf.paragraphs:
                p.font.size = Pt(14)
                p.font.bold = True
                p.font.color.rgb = RGBColor(0x1a, 0x5f, 0x7a)

            y_pos += 0.3
            content = data.get(key, '')
            if content:
                content_box = detail_slide.shapes.add_textbox(Inches(0.5), Inches(y_pos), Inches(9), Inches(1.5))
                ctf = content_box.text_frame
                ctf.text = content
                for p in ctf.paragraphs:
                    p.font.size = Pt(12)
                    p.word_wrap = True
                y_pos += 1.5

        if images:
            for idx, img_info in enumerate(images, 1):
                blank_layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]
                img_slide = prs.slides.add_slide(blank_layout)

                title_box = img_slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.5))
                tf = title_box.text_frame
                tf.text = f"Figura {idx}"
                for p in tf.paragraphs:
                    p.font.size = Pt(18)
                    p.font.bold = True
                    p.font.color.rgb = RGBColor(0x1a, 0x5f, 0x7a)

                try:
                    img_path = img_info["path"] if isinstance(img_info, dict) else img_info
                    img = Image.open(img_path)
                    img_w, img_h = img.size
                    aspect = img_w / img_h

                    if aspect > (9/6):
                        w = Inches(9)
                        h = w / aspect
                    else:
                        h = Inches(6)
                        w = h * aspect

                    left = (Inches(10) - w) / 2
                    img_slide.shapes.add_picture(img_path, left, Inches(1), w, h)
                except Exception as e:
                    st.warning(f"Error imagen: {e}")

        output_buffer = BytesIO()
        prs.save(output_buffer)
        output_buffer.seek(0)
        return output_buffer


class PDFExporter:
    """Exporta documentos a PDF usando multiples metodos"""

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
            st.warning(f"Conversion LibreOffice fallo: {e}")
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
            st.warning(f"Conversion LibreOffice fallo: {e}")
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
                'CustomTitle',
                parent=styles['Heading1'],
                fontSize=20,
                textColor=colors.HexColor('#1a5f7a'),
                spaceAfter=30,
                alignment=TA_CENTER,
                fontName='Helvetica-Bold'
            )

            heading_style = ParagraphStyle(
                'CustomHeading',
                parent=styles['Heading2'],
                fontSize=14,
                textColor=colors.HexColor('#1a5f7a'),
                spaceAfter=12,
                spaceBefore=12,
                fontName='Helvetica-Bold'
            )

            body_style = ParagraphStyle(
                'CustomBody',
                parent=styles['BodyText'],
                fontSize=10,
                leading=14,
                alignment=TA_JUSTIFY,
                fontName='Helvetica'
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
                    ("Descripcion del Problema", "descripcion_problema"),
                    ("Condicion Actual", "condicion_actual"),
                    ("Condicion Propuesta", "condicion_propuesta"),
                    ("Razones del Cambio", "razones_cambio"),
                    ("Alternativas y Plan de Retorno", "alternativas_retorno"),
                    ("Recursos", "recursos"),
                    ("Plan de Implementacion", "plan_implementacion"),
                    ("Tiempo de Duracion", "tiempo_duracion"),
                ]
                for title, key in sections:
                    story.append(Paragraph(f"<b>{title}</b>", heading_style))
                    content = data.get(key, '').replace('\n', '<br/>')
                    story.append(Paragraph(content, body_style))
                    story.append(Spacer(1, 10))

                story.append(Paragraph("<b>Riesgos y Controles</b>", heading_style))
                risks = data.get('riesgos_controles', [])
                if risks:
                    risk_data = [["N°", "Riesgo Identificado", "Control Recomendado"]]
                    for i, risk in enumerate(risks, 1):
                        risk_data.append([str(i), risk.get('riesgo', ''), risk.get('control', '')])
                    risk_table = Table(risk_data, colWidths=[30, 200, 200])
                    risk_table.setStyle(TableStyle([
                        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a5f7a')),
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                        ('FONTSIZE', (0, 0), (-1, 0), 10),
                        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                        ('GRID', (0, 0), (-1, -1), 1, colors.black),
                        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                        ('FONTSIZE', (0, 1), (-1, -1), 8),
                    ]))
                    story.append(risk_table)
                    story.append(Spacer(1, 10))

                story.append(Paragraph("<b>Riesgos SHES</b>", heading_style))
                risks_shes = data.get('riesgos_shes', [])
                if risks_shes:
                    risk_data = [["N°", "Riesgo", "Control", "Plazo"]]
                    for i, risk in enumerate(risks_shes, 1):
                        risk_data.append([str(i), risk.get('riesgo', ''), risk.get('control', ''), risk.get('plazo', '')])
                    risk_table = Table(risk_data, colWidths=[30, 150, 150, 100])
                    risk_table.setStyle(TableStyle([
                        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a5f7a')),
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                        ('FONTSIZE', (0, 0), (-1, 0), 10),
                        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                        ('GRID', (0, 0), (-1, -1), 1, colors.black),
                        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                        ('FONTSIZE', (0, 1), (-1, -1), 8),
                    ]))
                    story.append(risk_table)

            elif doc_type == "a3":
                sections = [
                    ("Antecedentes", "antecedentes"), ("Problema Actual", "problema_actual"),
                    ("Analisis de la Situacion", "analisis_situacion"), ("Objetivos", "objetivos"),
                    ("Analisis de Causa Raiz", "analisis_causa_raiz"), ("Contramedidas", "contramedidas"),
                    ("Resultados Esperados", "resultados_esperados"), ("Plan de Seguimiento", "plan_seguimiento"),
                    ("Lecciones Aprendidas", "lecciones_aprendidas"), ("Estandarizacion", "estandarizacion"),
                ]
                for title, key in sections:
                    story.append(Paragraph(f"<b>{title}</b>", heading_style))
                    content = data.get(key, '').replace('\n', '<br/>')
                    story.append(Paragraph(content, body_style))
                    story.append(Spacer(1, 10))

            elif doc_type == "kaizen":
                sections = [
                    ("Descripcion del Problema", "descripcion_problema"),
                    ("Solucion Implementada", "solucion"),
                    ("Beneficios", "beneficios"),
                    ("Proximos Pasos", "proximos_pasos"),
                ]
                for title, key in sections:
                    story.append(Paragraph(f"<b>{title}</b>", heading_style))
                    content = data.get(key, '').replace('\n', '<br/>')
                    story.append(Paragraph(content, body_style))
                    story.append(Spacer(1, 10))

                story.append(Paragraph(f"<b>Tipo de Desperdicio:</b> {data.get('tipo_desperdicio', '')}", body_style))
                story.append(Paragraph(f"<b>Impacto BTO:</b> {data.get('impacto_bto', '')}", body_style))

            doc.build(story)
            buffer.seek(0)
            return buffer.getvalue()

        except Exception as e:
            st.error(f"Error generando PDF con ReportLab: {e}")
            return None
