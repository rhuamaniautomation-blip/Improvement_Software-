#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
================================================================================
SISTEMA DE GESTION DOCUMENTAL - MoC | Mejora A3 | Simple Kaizen
Version 9.0.0 - Ingeniero Senior + Imágenes por Slide + Modelos Actualizados
================================================================================
Diseñado por: CAVA - Especialistas en Robotica y Automatizacion
Desarrollador: Roger Huamani
Version: 9.0.0
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

SPELLCHECKER_AVAILABLE = False
try:
    from spellchecker import SpellChecker
    SPELLCHECKER_AVAILABLE = True
except ImportError:
    pass

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

st.set_page_config(
    page_title="Gestión Documental - MoC | A3 | Kaizen",
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
.gemini-badge {
    display: inline-block; background: #e0e7ff; color: #4338ca;
    padding: 0.25rem 0.75rem; border-radius: 20px;
    font-size: 12px; font-weight: 600; margin-left: 0.5rem;
}
.history-item {
    background: white; border: 1px solid #e2e8f0;
    border-radius: 10px; padding: 1rem; margin: 0.5rem 0;
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
.slide-image-card {
    background: #f8fafc; border: 1px solid #e2e8f0;
    border-radius: 8px; padding: 0.75rem; margin: 0.5rem 0;
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
        if not text or not text.strip():
            return text
        corrections = {
            "tecnico": "técnico", "Tecnico": "Técnico", "TECNICO": "TÉCNICO",
            "tecnica": "técnica", "Tecnica": "Técnica",
            "tecnologia": "tecnología", "Tecnologia": "Tecnología",
            "produccion": "producción", "Produccion": "Producción",
            "implementacion": "implementación", "Implementacion": "Implementación",
            "evaluacion": "evaluación", "Evaluacion": "Evaluación",
            "operacion": "operación", "Operacion": "Operación",
            "condicion": "condición", "Condicion": "Condición",
            "modificacion": "modificación", "Modificacion": "Modificación",
            "verificacion": "verificación", "Verificacion": "Verificación",
            "capacitacion": "capacitación", "Capacitacion": "Capacitación",
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
            "distribucion": "distribución", "Distribucion": "Distribución",
            "construccion": "construcción", "Construccion": "Construcción",
            "instruccion": "instrucción", "Instruccion": "Instrucción",
            "reduccion": "reducción", "Reduccion": "Reducción",
            "traduccion": "traducción", "Traduccion": "Traducción",
            "maquina": "máquina", "Maquina": "Máquina",
            "maquinas": "máquinas", "Maquinas": "Máquinas",
            "podria": "podría", "Podria": "Podría",
            "habria": "habría", "Habria": "Habría",
            "seria": "sería", "Seria": "Sería",
            "tendria": "tendría", "Tendria": "Tendría",
            "haria": "haría", "Haria": "Haría",
            "daria": "daría", "Daria": "Daría",
            "estaria": "estaría", "Estaria": "Estaría",
            "deberia": "debería", "Deberia": "Debería",
            "mas": "más", "Mas": "Más",
            "aun": "aún", "Aun": "Aún",
            "tambien": "también", "Tambien": "También",
            "asi": "así", "Asi": "Así",
            "aqui": "aquí", "Aqui": "Aquí",
            "alli": "allí", "Alli": "Allí",
            "alla": "allá", "Alla": "Allá",
            "despues": "después", "Despues": "Después",
            "ademas": "además", "Ademas": "Además",
            "segun": "según", "Segun": "Según",
            "numero": "número", "Numero": "Número",
            "maximo": "máximo", "Maximo": "Máximo",
            "minimo": "mínimo", "Minimo": "Mínimo",
            "optimo": "óptimo", "Optimo": "Óptimo",
            "ultimo": "último", "Ultimo": "Último",
            "periodo": "período", "Periodo": "Período",
            "area": "área", "Area": "Área",
            "dia": "día", "Dia": "Día",
            "manana": "mañana", "Manana": "Mañana",
            "proximo": "próximo", "Proximo": "Próximo",
            "analisis": "análisis", "Analisis": "Análisis",
            "sintesis": "síntesis", "Sintesis": "Síntesis",
            "hipotesis": "hipótesis", "Hipotesis": "Hipótesis",
            "metodo": "método", "Metodo": "Método",
            "parametro": "parámetro", "Parametro": "Parámetro",
            "parametros": "parámetros", "Parametros": "Parámetros",
            "caracteristica": "característica", "Caracteristica": "Característica",
            "caracteristicas": "características", "Caracteristicas": "Características",
            "especifico": "específico", "Especifico": "Específico",
            "especifica": "específica", "Especifica": "Específica",
            "generico": "genérico", "Generico": "Genérico",
            "electronico": "electrónico", "Electronico": "Electrónico",
            "electrico": "eléctrico", "Electrico": "Eléctrico",
            "hidraulico": "hidráulico", "Hidraulico": "Hidráulico",
            "neumatico": "neumático", "Neumatico": "Neumático",
            "termico": "térmico", "Termico": "Térmico",
            "optico": "óptico", "Optico": "Óptico",
            "quimico": "químico", "Quimico": "Químico",
            "fisico": "físico", "Fisico": "Físico",
            "biologico": "biológico", "Biologico": "Biológico",
            "version": "versión", "Version": "Versión",
            "conversion": "conversión", "Conversion": "Conversión",
            "descripcion": "descripción", "Descripcion": "Descripción",
            "solucion": "solución", "Solucion": "Solución",
            "situacion": "situación", "Situacion": "Situación",
            "presentacion": "presentación", "Presentacion": "Presentación",
            "revision": "revisión", "Revision": "Revisión",
            "habilitacion": "habilitación", "Habilitacion": "Habilitación",
            "limites": "límites", "Limites": "Límites",
            "limite": "límite", "Limite": "Límite",
            "linea": "línea", "Linea": "Línea",
            "lineas": "líneas", "Lineas": "Líneas",
            "unico": "único", "Unico": "Único",
            "unica": "única", "Unica": "Única",
            "facil": "fácil", "Facil": "Fácil",
            "dificil": "difícil", "Dificil": "Difícil",
            "rapido": "rápido", "Rapido": "Rápido",
            "rapida": "rápida", "Rapida": "Rápida",
            "trabagar": "trabajar",
            "podra": "podrá",
            "esta": "está",
        }
        result = text
        for wrong, correct in corrections.items():
            result = result.replace(wrong, correct)
        result = re.sub(r'  +', ' ', result)
        result = re.sub(r' ([.,;:!?])', r'\1', result)
        return result

# =============================================================================
# SERVICIO GEMINI API - MODELOS ACTUALIZADOS Y PROMPT DE INGENIERO SENIOR
# =============================================================================
class GeminiService:
    MODELS = {
        "gemini-2.5-flash": {"name": "Gemini 2.5 Flash", "desc": "Rápido, eficiente y recomendado"},
        "gemini-2.5-pro": {"name": "Gemini 2.5 Pro", "desc": "Máxima calidad y razonamiento"},
    }
    DEPRECATED_MODELS = {
        "gemini-1.5-flash-lite", "gemini-1.5-flash", "gemini-1.5-pro",
        "gemini-1.0-pro", "gemini-2.0-flash", "gemini-2.0-flash-lite",
        "gemini-2.0-pro", "gemini-2.5-flash-lite"
    }
    DEFAULT_MODEL = "gemini-2.5-flash"

    def __init__(self, api_key="", model=None):
        self.api_key = api_key
        if model is None or model in self.DEPRECATED_MODELS or model not in self.MODELS:
            self.model = self.DEFAULT_MODEL
        else:
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
            response = requests.post(url, json=payload, timeout=180)
            response.raise_for_status()
            result = response.json()
            if "candidates" in result and len(result["candidates"]) > 0:
                return result["candidates"][0]["content"]["parts"][0]["text"]
            return ""
        except requests.exceptions.HTTPError as e:
            if e.response.status_code == 404:
                raise Exception(f"Error 404: Modelo '{self.model}' no disponible. Verifique que la 'Generative Language API' esté habilitada.")
            elif e.response.status_code == 403:
                raise Exception("Error 403: Acceso denegado. Verifique API Key.")
            else:
                raise Exception(f"Error HTTP {e.response.status_code}: {e.response.text}")

    def _extract_json(self, text):
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

    def generate_moc(self, problem, context="", equipo="", alternativas=""):
        """Genera MoC con rol de Ingeniero Senior - 14 slides detallados"""
        if not self.api_key:
            st.error("❌ API Key no configurada. Configure en Configuración > API Gemini")
            return None

        prompt = f"""Actúa como un Ingeniero Senior de Ingeniería, Automatización, Mantenimiento, Seguridad de Procesos y Gestión del Cambio (MoC) en una planta industrial de manufactura. Tienes 20 años de experiencia en minería, manufactura y operaciones industriales.

CONTEXTO DEL PROBLEMA IDENTIFICADO:
{problem}

INFORMACIÓN ADICIONAL:
{context}

EQUIPO INVOLUCRADO:
{equipo}

ALTERNATIVAS CONSIDERADAS (si existen):
{alternativas if alternativas else 'No proporcionadas por el usuario.'}

INSTRUCCIONES CRÍTICAS DE REDACCIÓN:
1. Redacta SIEMPRE en español técnico y corporativo, con lenguaje profesional de ingeniería senior.
2. Utiliza enfoque de ingeniería industrial, automatización, calidad, seguridad y confiabilidad.
3. NO pidas información adicional. Si falta algún dato, asume la mejor alternativa técnicamente razonable.
4. Redacta en párrafos bien estructurados de 4-6 oraciones con conectores lógicos.
5. Evita listas, excepto en "alternativas_consideradas".
6. Considera riesgos operativos, de calidad, seguridad, productividad, mantenimiento y cumplimiento normativo.
7. Cuando aplique, considera PLC, HMI, SCADA, sensores, instrumentación, equipos industriales, seguridad de máquinas, ISO 13849, enclavamientos, sistemas de visión, control de procesos, validaciones, calidad y gestión de activos.
8. Si el cambio es de software, automatización o parámetros, indica que no modifica la integridad mecánica del equipo salvo que el contexto indique lo contrario.
9. Si el proveedor es extranjero, considera tiempos de coordinación, ingeniería y validación.
10. ORTOGRAFÍA IMPECABLE: Tildes correctas en todas las palabras.
11. PROHIBIDO usar frases genéricas como "degradación progresiva de componentes" o "parámetros fuera de rango" a menos que el usuario las haya escrito explícitamente.
12. TODO el contenido debe basarse EXCLUSIVAMENTE en el problema específico del usuario.

Genera en ESPAÑOL formato JSON con esta estructura EXACTA (14 campos para 14 slides):

{{
  "moc_title": "Título técnico conciso del cambio (máximo 12 palabras)",
  
  "condicion_actual": "SLIDE 3 (columna izquierda). Describe detalladamente cómo opera actualmente el sistema o proceso, incluyendo limitaciones, desviaciones, riesgos, ineficiencias y situación existente. Mínimo 200 palabras en párrafos bien estructurados.",
  
  "condicion_propuesta": "SLIDE 3 (columna derecha). Describe detalladamente la solución propuesta, indicando cómo funcionará el sistema después del cambio y cuáles serán las mejoras obtenidas. Mínimo 200 palabras.",
  
  "justificacion_moc": "SLIDE 4. Explica la oportunidad de mejora, observación u obligación que motiva el cambio y por qué es necesario implementarlo. Mínimo 150 palabras.",
  
  "descripcion_problema": "SLIDE 5. Desarrolla técnicamente el problema identificado, incluyendo causas, efectos y consecuencias para la operación. Mínimo 250 palabras.",
  
  "razones_cambio": "SLIDE 6. Explica las razones técnicas, operativas, de calidad, productividad, confiabilidad, mantenimiento o seguridad que justifican la modificación. Mínimo 200 palabras.",
  
  "alternativas_consideradas": "SLIDE 4 (parte inferior). Genera TRES alternativas en formato de lista con viñetas '❖':
❖ Alternativa 1: Mantener condición actual. Explica ventajas, desventajas y motivo de descarte.
❖ Alternativa 2: Solución parcial o administrativa. Explica ventajas, desventajas y motivo de descarte.
❖ Alternativa 3: Solución seleccionada. Explica por qué se selecciona.",
  
  "plan_retorno": "SLIDE 4 (parte inferior). Redacta un único párrafo de aproximadamente tres líneas indicando cómo retornar a la condición original en caso de falla de la implementación.",
  
  "recursos": "SLIDE 8 (parte superior). Describe de forma breve los recursos necesarios según apliquen: Ingeniería, Automatización, Mantenimiento, Calidad, Producción, Seguridad, Proveedor, Materiales, Software, Validación. Mínimo 150 palabras.",
  
  "plan_implementacion": "SLIDE 8 (parte media). Redacta un párrafo corto y resumido describiendo las principales etapas: evaluación, ingeniería, programación, instalación, pruebas, validación y liberación. Mínimo 150 palabras.",
  
  "tiempo_duracion": "SLIDE 8 (parte inferior). Estima razonablemente la duración total del cambio considerando: Ingeniería, Aprobaciones, Compras, Soporte del proveedor, Instalación, Programación, Validación. Indica un rango de tiempo realista.",
  
  "riesgos_controles": "SLIDE 12 (parte superior). Identifica riesgos residuales posteriores a la implementación y sus controles para asegurar la sostenibilidad del cambio. Array de 3-5 objetos con estructura: {{'riesgo': 'descripción del riesgo', 'control': 'medida de control específica', 'plazo': 'plazo de implementación'}}",
  
  "riesgos_shes": "SLIDE 12 (parte inferior). Genera riesgos SHES posteriores a la implementación. Si el cambio no genera riesgos adicionales de SHES, indícalo y propone controles de monitoreo, validación o mantenimiento. Array de 3-5 objetos con estructura: {{'riesgo': 'descripción del riesgo SHES', 'control': 'plan de acción específico', 'plazo': 'plazo'}}",
  
  "impacto_esperado": "SLIDE adicional. Agrega un resumen ejecutivo de uno o dos párrafos indicando el beneficio esperado en: Seguridad, Calidad, Productividad, Confiabilidad, Mantenimiento, Costos, Cumplimiento normativo.",
  
  "resumen_ejecutivo": "SLIDE adicional. Redacta un párrafo corto dirigido a los aprobadores de la MoC, explicando por qué el cambio debe ser aprobado y cuáles son los beneficios principales.",
  
  "checklist_360": [
    {{"numero": 1, "factor": "Interacción o impacto con otras áreas/procesos", "aplica": "SI/NO", "descripcion": "Descripción del impacto o dejar vacío si NO"}},
    {{"numero": 2, "factor": "Cambios en los procedimientos operativos, arranque y parada", "aplica": "SI/NO", "descripcion": "..."}},
    {{"numero": 3, "factor": "Parámetros operativos y límites de control", "aplica": "SI/NO", "descripcion": "..."}},
    {{"numero": 4, "factor": "Cambios en interfaces hombre-máquina y gestión de alarmas", "aplica": "SI/NO", "descripcion": "..."}},
    {{"numero": 5, "factor": "Compatibilidad de materiales, sustancias y equipos", "aplica": "SI/NO", "descripcion": "..."}},
    {{"numero": 6, "factor": "Exposición ocupacional (ruido, polvo, ergonomía, etc.)", "aplica": "SI/NO", "descripcion": "..."}},
    {{"numero": 7, "factor": "Requerimientos de EPP y su compatibilidad", "aplica": "SI/NO", "descripcion": "..."}},
    {{"numero": 8, "factor": "Escenarios de emergencia y capacidad de respuesta", "aplica": "SI/NO", "descripcion": "..."}},
    {{"numero": 9, "factor": "Impacto en el almacenamiento y tránsito interno/externo", "aplica": "SI/NO", "descripcion": "..."}},
    {{"numero": 10, "factor": "Impactos ambientales y generación de residuos", "aplica": "SI/NO", "descripcion": "..."}},
    {{"numero": 11, "factor": "Impacto en la calidad del producto o servicio", "aplica": "SI/NO", "descripcion": "..."}},
    {{"numero": 12, "factor": "Cambios en roles, competencias y carga de trabajo", "aplica": "SI/NO", "descripcion": "..."}},
    {{"numero": 13, "factor": "Integridad de equipos, protecciones y sistemas de control", "aplica": "SI/NO", "descripcion": "..."}},
    {{"numero": 14, "factor": "Cumplimiento legal, normativo y permisos aplicables", "aplica": "SI/NO", "descripcion": "..."}},
    {{"numero": 15, "factor": "Cambios en las condiciones para trabajos especiales", "aplica": "SI/NO", "descripcion": "..."}},
    {{"numero": 16, "factor": "Cambios sucesivos que incrementan el riesgo global", "aplica": "SI/NO", "descripcion": "..."}}
  ],
  
  "documentos_impactados": [
    {{"numero": 1, "documento": "JSERA - IPERC", "aplica": "SI/NO", "modificacion": "Describir modificación o vacío si NO"}},
    {{"numero": 2, "documento": "Procedimiento de Trabajo, Instructivo/PO", "aplica": "SI/NO", "modificacion": "..."}},
    {{"numero": 3, "documento": "Formato/Checklist operativos", "aplica": "SI/NO", "modificacion": "..."}},
    {{"numero": 4, "documento": "Matriz de EPP", "aplica": "SI/NO", "modificacion": "..."}},
    {{"numero": 5, "documento": "MSDS de sustancias involucradas", "aplica": "SI/NO", "modificacion": "..."}},
    {{"numero": 6, "documento": "Mapa de Riesgos", "aplica": "SI/NO", "modificacion": "..."}},
    {{"numero": 7, "documento": "Plan de emergencias", "aplica": "SI/NO", "modificacion": "..."}},
    {{"numero": 8, "documento": "Plan de Mantenimiento", "aplica": "SI/NO", "modificacion": "..."}},
    {{"numero": 9, "documento": "Matriz de impactos ambientales", "aplica": "SI/NO", "modificacion": "..."}},
    {{"numero": 10, "documento": "Plan Monitoreos SSO requeridos", "aplica": "SI/NO", "modificacion": "..."}},
    {{"numero": 11, "documento": "Plan de tráfico", "aplica": "SI/NO", "modificacion": "..."}},
    {{"numero": 12, "documento": "Matriz de competencias, plan de entrenamiento", "aplica": "SI/NO", "modificacion": "..."}},
    {{"numero": 13, "documento": "Plan de calidad", "aplica": "SI/NO", "modificacion": "..."}},
    {{"numero": 14, "documento": "Planos y diagramas (layout, P&ID)", "aplica": "SI/NO", "modificacion": "..."}},
    {{"numero": 15, "documento": "Licencias y permisos aplicables", "aplica": "SI/NO", "modificacion": "..."}}
  ]
}}

IMPORTANTE:
- Responde SOLO con JSON válido, sin comentarios ni texto adicional.
- Todos los textos en ESPAÑOL.
- Ortografía impecable con todas las tildes correctas.
- Redacción profesional, técnica y humanizada.
- Párrafos extensos y bien estructurados.
- TODO EL CONTENIDO DEBE BASARSE EXCLUSIVAMENTE EN EL PROBLEMA ESPECÍFICO DEL USUARIO."""
        try:
            text = self._call_api(prompt, temperature=0.4, max_tokens=16000)
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
        prompt = f"""Eres un experto senior en metodología A3 Lean con 15 años de experiencia. Redactas documentos A3 con redacción humanizada, técnica y profesional.
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
        prompt = f"""Eres un experto en Kaizen y Lean Manufacturing. Redactas registros Kaizen con redacción humanizada y práctica.
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
        prompt = f"""Corrige ortografía, gramática y puntuación. Mantén el significado técnico. Asegura tildes correctas. Devuelve SOLO el texto corregido.
TEXTO:
{text}"""
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

# =============================================================================
# GENERADOR DE DOCUMENTOS
# =============================================================================
class DocumentGenerator:
    def generate_moc(self, data, images_by_slide=None, template_bytes=None):
        """Genera MoC desde template con imágenes por slide"""
        if template_bytes is None:
            st.error("❌ Template MoC no cargado. Vaya a Configuración > Templates.")
            return None
        prs = Presentation(BytesIO(template_bytes))
        
        # Reemplazos globales
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

        # SLIDE 2: Equipo
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

        # SLIDE 3: Tabla Condición Actual / Condición Propuesta
        if len(prs.slides) > 2:
            slide3 = prs.slides[2]
            for shape in slide3.shapes:
                if shape.has_table:
                    table = shape.table
                    if len(table.rows) >= 2 and len(table.columns) >= 2:
                        fill_table_cell(table.cell(1, 0), data.get('condicion_actual', ''))
                        fill_table_cell(table.cell(1, 1), data.get('condicion_propuesta', ''))

        # SLIDE 4: Razones + Alternativas + Plan de retorno
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
                                first_para.runs[0].text = data.get('alternativas_consideradas', '')
                            else:
                                first_para.text = data.get('alternativas_consideradas', '')
                    if "Plan de retorno" in text or "Reinstalación del sistema actual" in text:
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                run.text = ""
                        if shape.text_frame.paragraphs:
                            first_para = shape.text_frame.paragraphs[0]
                            if first_para.runs:
                                first_para.runs[0].text = data.get('plan_retorno', '')
                            else:
                                first_para.text = data.get('plan_retorno', '')

        # SLIDE 5: Descripción del Problema
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

        # SLIDE 6 y 7: Imágenes con numeración correlativa y texto explicativo
        if images_by_slide and len(prs.slides) > 5:
            # Recopilar todas las imágenes de los slides 6 y 7
            all_slide_images = []
            for slide_num in [6, 7]:
                slide_key = f"slide_{slide_num}"
                if slide_key in images_by_slide:
                    for img_info in images_by_slide[slide_key]:
                        all_slide_images.append((slide_num, img_info))
            
            # Insertar imágenes en los slides 6 y 7 existentes
            for idx, (slide_num, img_info) in enumerate(all_slide_images[:2]):
                target_slide = prs.slides[5 + idx] if (5 + idx) < len(prs.slides) else None
                if target_slide:
                    try:
                        img_path = img_info["path"]
                        img = Image.open(img_path)
                        img_w, img_h = img.size
                        aspect = img_w / img_h
                        max_w = Inches(7.5)
                        max_h = Inches(4.5)
                        if aspect > (7.5/4.5):
                            w = max_w
                            h = w / aspect
                        else:
                            h = max_h
                            w = h * aspect
                        img_left = (Inches(10) - w) / 2
                        target_slide.shapes.add_picture(img_path, img_left, Inches(1.5), w, h)
                        
                        # Agregar texto explicativo con numeración correlativa
                        desc = img_info.get("desc", f"Figura {img_info.get('number', idx+1)}")
                        desc_box = target_slide.shapes.add_textbox(Inches(0.5), Inches(6.5), Inches(9), Inches(0.8))
                        dtf = desc_box.text_frame
                        dtf.text = desc
                        for p in dtf.paragraphs:
                            p.font.size = Pt(11)
                            p.font.italic = True
                            p.alignment = PP_ALIGN.CENTER
                    except Exception as e:
                        st.warning(f"Error con imagen {idx+1}: {e}")
            
            # Si hay más imágenes, agregar slides adicionales
            for idx, (slide_num, img_info) in enumerate(all_slide_images[2:]):
                blank_layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]
                new_slide = prs.slides.add_slide(blank_layout)
                try:
                    img_path = img_info["path"]
                    img = Image.open(img_path)
                    img_w, img_h = img.size
                    aspect = img_w / img_h
                    max_w = Inches(8.5)
                    max_h = Inches(5.5)
                    if aspect > (8.5/5.5):
                        w = max_w
                        h = w / aspect
                    else:
                        h = max_h
                        w = h * aspect
                    img_left = (Inches(10) - w) / 2
                    new_slide.shapes.add_picture(img_path, img_left, Inches(1), w, h)
                    
                    desc = img_info.get("desc", f"Figura {img_info.get('number', idx+3)}")
                    desc_box = new_slide.shapes.add_textbox(Inches(0.5), Inches(6.8), Inches(9), Inches(0.8))
                    dtf = desc_box.text_frame
                    dtf.text = desc
                    for p in dtf.paragraphs:
                        p.font.size = Pt(11)
                        p.font.italic = True
                        p.alignment = PP_ALIGN.CENTER
                except Exception as e:
                    st.warning(f"Error con imagen: {e}")

        # SLIDE 8: Recursos + Plan + Tiempo
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

        # SLIDE 9: Checklist 360°
        if len(prs.slides) > 8:
            slide9 = prs.slides[8]
            checklist = data.get('checklist_360', [])
            for shape in slide9.shapes:
                if shape.has_table:
                    table = shape.table
                    for i, item in enumerate(checklist):
                        row_idx = i + 1
                        if row_idx < len(table.rows):
                            if len(table.columns) > 0:
                                fill_table_cell(table.cell(row_idx, 0), str(item.get('numero', i+1)))
                            if len(table.columns) > 1:
                                fill_table_cell(table.cell(row_idx, 1), item.get('factor', ''))
                            if len(table.columns) > 2:
                                fill_table_cell(table.cell(row_idx, 2), item.get('aplica', 'NO'))
                            if len(table.columns) > 3:
                                fill_table_cell(table.cell(row_idx, 3), item.get('descripcion', ''))

        # SLIDE 10: Documentos impactados
        if len(prs.slides) > 9:
            slide10 = prs.slides[9]
            docs_impactados = data.get('documentos_impactados', [])
            for shape in slide10.shapes:
                if shape.has_table:
                    table = shape.table
                    for i, item in enumerate(docs_impactados):
                        row_idx = i + 1
                        if row_idx < len(table.rows):
                            if len(table.columns) > 0:
                                fill_table_cell(table.cell(row_idx, 0), str(item.get('numero', i+1)))
                            if len(table.columns) > 1:
                                fill_table_cell(table.cell(row_idx, 1), item.get('documento', ''))
                            if len(table.columns) > 2:
                                fill_table_cell(table.cell(row_idx, 2), item.get('aplica', 'NO'))
                            if len(table.columns) > 3:
                                fill_table_cell(table.cell(row_idx, 3), item.get('modificacion', ''))

        # SLIDE 12: Riesgos SHES
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
            st.error("❌ Template A3 no cargado.")
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
        doc.add_page_break()
        heading = doc.add_heading(data.get('titulo', 'Mejora A3'), level=1)
        sections = [
            ("ANTECEDENTES", "antecedentes"), ("PROBLEMA ACTUAL", "problema_actual"),
            ("ANÁLISIS DE LA SITUACIÓN", "analisis_situacion"), ("OBJETIVOS", "objetivos"),
            ("ANÁLISIS DE CAUSA RAÍZ", "analisis_causa_raiz"), ("CONTRAMEDIDAS", "contramedidas"),
            ("RESULTADOS ESPERADOS", "resultados_esperados"), ("PLAN DE SEGUIMIENTO", "plan_seguimiento"),
            ("LECCIONES APRENDIDAS", "lecciones_aprendidas"), ("ESTANDARIZACIÓN", "estandarizacion"),
        ]
        for section_title, key in sections:
            h = doc.add_heading(section_title, level=2)
            content = data.get(key, '')
            if content:
                doc.add_paragraph(content)
        output_buffer = BytesIO()
        doc.save(output_buffer)
        output_buffer.seek(0)
        return output_buffer

    def generate_kaizen(self, data, images=None, template_bytes=None):
        if template_bytes is None:
            st.error("❌ Template Kaizen no cargado.")
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
        except Exception as e:
            st.warning(f"Conversión LibreOffice falló: {e}")
        return None

# =============================================================================
# INICIALIZACIÓN CON MIGRACIÓN AUTOMÁTICA
# =============================================================================
def init_session_state():
    saved_config = LocalStorage.load_config()
    saved_history = LocalStorage.load_history()
    
    if saved_config and saved_config.get("gemini_model") in GeminiService.DEPRECATED_MODELS:
        old_model = saved_config.get("gemini_model")
        saved_config["gemini_model"] = GeminiService.DEFAULT_MODEL
        LocalStorage.save_config(saved_config)
    
    defaults = {
        "page": "inicio",
        "config": saved_config or {
            "gemini_api_key": "",
            "gemini_model": GeminiService.DEFAULT_MODEL,
            "company_name": "",
            "department": "",
            "default_author": "",
            "default_area": "",
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
        "doc_images_by_slide": {},  # NUEVO: imágenes por slide
        "doc_images": [],
        "doc_type": None,
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
# COMPONENTE: CARGA DE IMÁGENES POR SLIDE CON NUMERACIÓN CORRELATIVA
# =============================================================================
def render_slide_image_uploader(slide_number, slide_title, images_by_slide):
    """Renderiza un uploader de imágenes para un slide específico con numeración correlativa"""
    slide_key = f"slide_{slide_number}"
    
    st.markdown(f"**📷 Slide {slide_number}: {slide_title}**")
    
    # Calcular el siguiente número correlativo global
    all_images = []
    for key, imgs in images_by_slide.items():
        all_images.extend(imgs)
    next_number = len(all_images) + 1
    
    # Uploader de imágenes
    uploaded_files = st.file_uploader(
        f"Cargar imagen(es) para el Slide {slide_number}:",
        type=["png", "jpg", "jpeg"],
        accept_multiple_files=True,
        key=f"uploader_slide_{slide_number}"
    )
    
    if uploaded_files:
        if slide_key not in images_by_slide:
            images_by_slide[slide_key] = []
        
        for img_file in uploaded_files:
            # Verificar si ya está cargada
            already_loaded = any(
                img.get("filename") == img_file.name 
                for img in images_by_slide[slide_key]
            )
            if not already_loaded:
                img_path = f"/tmp/temp_moc_slide{slide_number}_{next_number}_{img_file.name}"
                with open(img_path, "wb") as f:
                    f.write(img_file.getbuffer())
                
                images_by_slide[slide_key].append({
                    "path": img_path,
                    "filename": img_file.name,
                    "number": next_number,
                    "desc": f"Figura {next_number}. [Ingrese texto explicativo]",
                    "slide": slide_number
                })
                next_number += 1
        
        st.success(f"✅ {len(uploaded_files)} imagen(es) cargada(s) para el Slide {slide_number}")
    
    # Mostrar imágenes cargadas con texto explicativo editable
    if slide_key in images_by_slide and images_by_slide[slide_key]:
        for idx, img_info in enumerate(images_by_slide[slide_key]):
            st.markdown("---")
            col1, col2 = st.columns([2, 3])
            with col1:
                st.image(img_info["path"], width=200)
                st.caption(f"**Figura {img_info['number']}**")
            with col2:
                new_desc = st.text_area(
                    f"Texto explicativo (Obligatorio) - Figura {img_info['number']}:",
                    value=img_info["desc"],
                    height=80,
                    key=f"desc_slide_{slide_number}_{idx}",
                    help="Describa el contenido de la imagen. Este texto aparecerá en el documento."
                )
                images_by_slide[slide_key][idx]["desc"] = new_desc
                
                if st.button(f"🗑️ Eliminar Figura {img_info['number']}", key=f"del_img_{slide_number}_{idx}"):
                    images_by_slide[slide_key].pop(idx)
                    # Re-numerar las imágenes restantes
                    for i, img in enumerate(images_by_slide[slide_key]):
                        img["number"] = i + 1
                        if img["desc"].startswith("Figura "):
                            img["desc"] = f"Figura {i+1}. [Ingrese texto explicativo]"
                    st.rerun()

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
    model_name = GeminiService.MODELS.get(config.get("gemini_model", GeminiService.DEFAULT_MODEL), {}).get("name", "Gemini 2.5 Flash")
    st.sidebar.markdown(f"""
<div style="text-align: center; color: #64748b; font-size: 0.75rem;">
<p>Modelo IA: <span class="gemini-badge">{model_name}</span></p>
<p>v9.0.0 · Agosto 2026</p>
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
        st.warning("⚠️ **Templates no cargados.** Vaya a Configuración > Templates.")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("""
<div class="doc-card doc-card-moc">
<h3 style="color: #1a5f7a; margin-top: 0;">📋 Management of Change</h3>
<p style="color: #64748b; font-size: 0.9rem;">Formato oficial MDET de 12+ slides con Checklist 360° y análisis integral.</p>
<ul style="color: #475569; font-size: 0.85rem; padding-left: 1.2rem;">
<li>14 slides estandarizados</li><li>Imágenes por slide</li><li>Checklist 360° automático</li><li>Riesgos SHES detallados</li>
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
    st.info("💡 Complete la información y describa el problema con detalle. La IA (como Ingeniero Senior) generará automáticamente los 14 slides del formato oficial MDET con redacción profesional.")
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
    st.markdown("#### 3. Contexto del Problema Identificado (DETALLADO)")
    st.warning("⚠️ **IMPORTANTE:** Describa el problema con el mayor detalle posible. La IA actuará como Ingeniero Senior y generará contenido basado EXCLUSIVAMENTE en su contexto.")
    problem_desc = auto_correct_text_input(
        "Contexto del problema identificado:",
        "",
        "moc_problem_desc",
        height=300,
        help_text="Describa: qué está pasando, desde cuándo, impacto, equipos involucrados, riesgos observados, dimensiones, parámetros técnicos, normas aplicables (ISO 13849, etc.), solución propuesta."
    )
    st.markdown("#### 4. Contexto Adicional (Opcional)")
    context = auto_correct_text_input(
        "Información adicional:",
        "",
        "moc_context",
        height=120,
        help_text="TAG del equipo, área específica, normativas aplicables, fechas relevantes, datos numéricos, planos de referencia, etc."
    )
    st.markdown("#### 5. Alternativas Consideradas (Opcional)")
    alternativas = auto_correct_text_input(
        "Alternativas ya evaluadas (si existen):",
        "",
        "moc_alternativas",
        height=100,
        help_text="Si ya tiene alternativas evaluadas, descríbalas aquí. Si no, la IA generará 3 alternativas automáticamente."
    )
    
    st.markdown("#### 6. Imágenes de Soporte por Slide")
    st.info("💡 Cargue imágenes específicas para cada slide. La numeración será correlativa automática y el texto explicativo es obligatorio.")
    
    # Inicializar imágenes por slide si no existe
    if "doc_images_by_slide" not in st.session_state:
        st.session_state.doc_images_by_slide = {}
    images_by_slide = st.session_state.doc_images_by_slide
    
    # Slides donde se pueden cargar imágenes
    slide_definitions = [
        (6, "Vista general del equipo/proceso"),
        (7, "Planos de referencia / Diagramas"),
    ]
    
    for slide_num, slide_title in slide_definitions:
        render_slide_image_uploader(slide_num, slide_title, images_by_slide)
    
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🤖 Generar Documento MoC con IA (Ingeniero Senior)", type="primary", use_container_width=True):
        if not problem_desc.strip():
            st.error("❌ Describa el problema antes de generar.")
            return
        with st.spinner("🧠 El Ingeniero Senior IA está generando los 14 slides del formato oficial MDET..."):
            gemini = GeminiService(config.get("gemini_api_key", ""), config.get("gemini_model"))
            equipo_data = {
                "produccion": produccion, "specialist_shes": specialist_shes,
                "mantenimiento": mantenimiento, "revisores": revisores,
                "experto_aprobador": experto_aprobador
            }
            result = gemini.generate_moc(problem_desc, context, json.dumps(equipo_data), alternativas)
            if result is None:
                st.error("❌ No se pudo generar el documento. Verifique su API Key en Configuración.")
                return
            st.session_state.generated_data = result
            st.session_state.doc_meta = {
                "moc_title": moc_title, "moc_number": moc_number, "naturaleza": naturaleza,
                "originador": originador, "fecha": fecha, **equipo_data
            }
            st.session_state.doc_images_by_slide = images_by_slide
            st.session_state.doc_type = "moc"
            st.session_state.page = "revisar"
            st.rerun()

def render_a3_form():
    config = st.session_state.config
    st.markdown('<div class="section-header"><h3>📊 Nueva Mejora A3</h3></div>', unsafe_allow_html=True)
    st.info("💡 Describa el problema con detalle y la IA generará el documento A3 completo.")
    if not st.session_state.get("template_a3_bytes"):
        st.error("❌ **Template A3 no cargado.**")
        return
    col1, col2 = st.columns(2)
    with col1:
        a3_title = st.text_input("Título:")
        area = st.text_input("Área:", value=config.get("default_area", ""))
    with col2:
        autor = st.text_input("Autor:", value=config.get("default_author", ""))
        doc_number = st.text_input("Número:", value=Utils.generate_doc_number("a3"), disabled=True)
        fecha = st.text_input("Fecha:", value=Utils.format_date(), disabled=True)
    problem_desc = auto_correct_text_input("Describa el problema actual:", "", "a3_problem_desc", height=250)
    context = auto_correct_text_input("Contexto adicional:", "", "a3_context", height=100)
    if st.button("🤖 Generar Documento A3 con IA", type="primary", use_container_width=True):
        if not problem_desc.strip():
            st.error("❌ Describa el problema antes de generar.")
            return
        with st.spinner("🧠 Generando documento A3..."):
            gemini = GeminiService(config.get("gemini_api_key", ""), config.get("gemini_model"))
            result = gemini.generate_a3(problem_desc, context)
            if result is None:
                return
            st.session_state.generated_data = result
            st.session_state.doc_meta = {"titulo": a3_title, "area": area, "autor": autor, "doc_number": doc_number, "fecha": fecha}
            st.session_state.doc_type = "a3"
            st.session_state.page = "revisar"
            st.rerun()

def render_kaizen_form():
    config = st.session_state.config
    st.markdown('<div class="section-header"><h3>⚡ Nuevo Simple Kaizen</h3></div>', unsafe_allow_html=True)
    if not st.session_state.get("template_kaizen_bytes"):
        st.error("❌ **Template Kaizen no cargado.**")
        return
    col1, col2 = st.columns(2)
    with col1:
        kaizen_title = st.text_input("Título (Name):")
        area = st.text_input("Plant/Area:", value=config.get("default_area", ""))
    with col2:
        leader = st.text_input("Leader:", value=config.get("default_author", ""))
        doc_number = st.text_input("Número:", value=Utils.generate_doc_number("kaizen"), disabled=True)
        fecha = st.text_input("Date:", value=Utils.format_date(), disabled=True)
        team_members = st.text_input("Team Members:")
    activity_desc = auto_correct_text_input("Describa la mejora realizada:", "", "kzn_activity_desc", height=250)
    tipo_desp = st.multiselect("Tipo de Desperdicio:", ["Motion", "Skills", "Inventory", "Transportation", "Over Production", "Over Processing", "Waiting", "Defects"])
    impacto_bto = st.selectbox("Impacto BTO:", ["Safe and Sustainable", "People & Culture", "Network Optimisation", "Supply Chain and Manufacturing Excellence"])
    beneficios = auto_correct_text_input("Beneficios:", "", "kzn_beneficios", height=120)
    if st.button("🤖 Generar Documento Kaizen con IA", type="primary", use_container_width=True):
        if not activity_desc.strip():
            st.error("❌ Describa la actividad antes de generar.")
            return
        with st.spinner("🧠 Generando documento Kaizen..."):
            gemini = GeminiService(config.get("gemini_api_key", ""), config.get("gemini_model"))
            result = gemini.generate_kaizen(activity_desc, "")
            if result is None:
                return
            result["tipo_desperdicio"] = ", ".join(tipo_desp) if tipo_desp else result.get("tipo_desperdicio", "")
            result["impacto_bto"] = impacto_bto
            result["leader"] = leader
            result["team_members"] = team_members
            result["beneficios"] = beneficios if beneficios else result.get("beneficios", "")
            st.session_state.generated_data = result
            st.session_state.doc_meta = {"titulo": kaizen_title, "area": area, "leader": leader, "team_members": team_members, "doc_number": doc_number, "fecha": fecha}
            st.session_state.doc_type = "kaizen"
            st.session_state.page = "revisar"
            st.rerun()

def render_review():
    doc_type = st.session_state.doc_type
    data = st.session_state.get("generated_data", {})
    meta = st.session_state.get("doc_meta", {})
    images_by_slide = st.session_state.get("doc_images_by_slide", {})
    config = st.session_state.config
    type_names = {"moc": "MoC", "a3": "Mejora A3", "kaizen": "Simple Kaizen"}
    type_name = type_names.get(doc_type, "Documento")
    st.markdown(f'<div class="section-header"><h3>👁️ Revisar y Editar {type_name}</h3></div>', unsafe_allow_html=True)
    if doc_type == "moc":
        _render_moc_review(data, meta, images_by_slide, config)
    elif doc_type == "a3":
        _render_a3_review(data, meta, config)
    elif doc_type == "kaizen":
        _render_kaizen_review(data, meta, config)

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

def _render_moc_review(data, meta, images_by_slide, config):
    gemini = GeminiService(config.get("gemini_api_key", ""), config.get("gemini_model"))
    tabs = st.tabs(["📋 General", "📝 Contenido", "📷 Imágenes", "📊 Checklist", "📄 Documentos", "⚠️ Riesgos", "⚙️ Generar"])
    
    with tabs[0]:
        st.markdown("#### Información del Documento")
        meta["moc_title"] = st.text_input("Título:", value=meta.get("moc_title", ""))
        meta["moc_number"] = st.text_input("Número:", value=meta.get("moc_number", ""), disabled=True)
        meta["naturaleza"] = st.selectbox("Naturaleza:", ["permanente", "temporal", "emergencia"],
                                          index=["permanente", "temporal", "emergencia"].index(meta.get("naturaleza", "permanente")))
        meta["originador"] = st.text_input("Originador:", value=meta.get("originador", ""))
        st.markdown("#### Equipo de Revisión")
        meta["produccion"] = st.text_input("Producción:", value=meta.get("produccion", ""))
        meta["specialist_shes"] = st.text_input("Specialist SHES:", value=meta.get("specialist_shes", ""))
        meta["mantenimiento"] = st.text_input("Mantenimiento:", value=meta.get("mantenimiento", ""))
        meta["revisores"] = st.text_input("Revisores:", value=meta.get("revisores", ""))
        meta["experto_aprobador"] = st.text_input("Experto Aprobador:", value=meta.get("experto_aprobador", ""))
    
    with tabs[1]:
        st.markdown("#### Condición Actual (Slide 3)")
        data["condicion_actual"] = _spell_check_field("", data.get("condicion_actual", ""), "moc_actual", gemini)
        st.markdown("#### Condición Propuesta (Slide 3)")
        data["condicion_propuesta"] = _spell_check_field("", data.get("condicion_propuesta", ""), "moc_prop", gemini)
        st.markdown("#### Justificación de la MoC (Slide 4)")
        data["justificacion_moc"] = _spell_check_field("", data.get("justificacion_moc", ""), "moc_just", gemini)
        st.markdown("#### Descripción del Problema (Slide 5)")
        data["descripcion_problema"] = _spell_check_field("", data.get("descripcion_problema", ""), "moc_desc", gemini)
        st.markdown("#### Razones del Cambio (Slide 6)")
        data["razones_cambio"] = _spell_check_field("", data.get("razones_cambio", ""), "moc_raz", gemini)
        st.markdown("#### Alternativas Consideradas (Slide 4)")
        data["alternativas_consideradas"] = _spell_check_field("", data.get("alternativas_consideradas", ""), "moc_alt", gemini)
        st.markdown("#### Plan de Retorno (Slide 4)")
        data["plan_retorno"] = _spell_check_field("", data.get("plan_retorno", ""), "moc_ret", gemini)
        st.markdown("#### Recursos (Slide 8)")
        data["recursos"] = _spell_check_field("", data.get("recursos", ""), "moc_rec", gemini)
        st.markdown("#### Plan de Implementación (Slide 8)")
        data["plan_implementacion"] = _spell_check_field("", data.get("plan_implementacion", ""), "moc_plan", gemini)
        st.markdown("#### Tiempo de Duración (Slide 8)")
        data["tiempo_duracion"] = _spell_check_field("", data.get("tiempo_duracion", ""), "moc_tiempo", gemini)
        st.markdown("#### Impacto Esperado")
        data["impacto_esperado"] = _spell_check_field("", data.get("impacto_esperado", ""), "moc_impacto", gemini)
        st.markdown("#### Resumen Ejecutivo para Aprobación")
        data["resumen_ejecutivo"] = _spell_check_field("", data.get("resumen_ejecutivo", ""), "moc_resumen", gemini)
    
    with tabs[2]:
        st.markdown("#### 📷 Imágenes Cargadas por Slide")
        st.info("💡 Cada imagen tiene numeración correlativa automática y texto explicativo obligatorio.")
        
        if images_by_slide:
            total_images = sum(len(imgs) for imgs in images_by_slide.values())
            st.success(f"✅ Total de imágenes cargadas: {total_images}")
            
            for slide_key, imgs in images_by_slide.items():
                slide_num = slide_key.replace("slide_", "")
                st.markdown(f"### Slide {slide_num}")
                for idx, img_info in enumerate(imgs):
                    st.markdown("---")
                    col1, col2 = st.columns([2, 3])
                    with col1:
                        st.image(img_info["path"], width=250)
                        st.caption(f"**Figura {img_info['number']}**")
                    with col2:
                        new_desc = st.text_area(
                            f"Texto explicativo - Figura {img_info['number']} (Obligatorio):",
                            value=img_info["desc"],
                            height=80,
                            key=f"rev_desc_{slide_key}_{idx}"
                        )
                        images_by_slide[slide_key][idx]["desc"] = new_desc
        else:
            st.info("No se cargaron imágenes. Puede agregarlas en el formulario de creación.")
    
    with tabs[3]:
        st.markdown("#### Checklist 360° - 16 Factores")
        checklist = data.get("checklist_360", [])
        updated_checklist = []
        for item in checklist:
            st.markdown(f"**{item.get('numero')}. {item.get('factor')}**")
            col1, col2 = st.columns([1, 3])
            with col1:
                aplica = st.selectbox("Aplica:", ["SI", "NO"],
                                      index=0 if item.get("aplica", "NO") == "SI" else 1,
                                      key=f"chk_{item.get('numero')}")
            with col2:
                desc = st.text_input("Descripción:", value=item.get("descripcion", ""),
                                     key=f"chk_desc_{item.get('numero')}")
            updated_checklist.append({"numero": item.get("numero"), "factor": item.get("factor"),
                                      "aplica": aplica, "descripcion": desc if aplica == "SI" else ""})
        data["checklist_360"] = updated_checklist
    
    with tabs[4]:
        st.markdown("#### Documentos Impactados - 15 Documentos")
        docs_imp = data.get("documentos_impactados", [])
        updated_docs = []
        for item in docs_imp:
            st.markdown(f"**{item.get('numero')}. {item.get('documento')}**")
            col1, col2 = st.columns([1, 3])
            with col1:
                aplica = st.selectbox("Aplica:", ["SI", "NO"],
                                      index=0 if item.get("aplica", "NO") == "SI" else 1,
                                      key=f"doc_{item.get('numero')}")
            with col2:
                modif = st.text_input("Modificación:", value=item.get("modificacion", ""),
                                      key=f"doc_mod_{item.get('numero')}")
            updated_docs.append({"numero": item.get("numero"), "documento": item.get("documento"),
                                 "aplica": aplica, "modificacion": modif if aplica == "SI" else ""})
        data["documentos_impactados"] = updated_docs
    
    with tabs[5]:
        st.markdown("#### Riesgos de Calidad")
        risks_cal = data.get("riesgos_controles", [])
        updated_cal = []
        for i, risk in enumerate(risks_cal):
            st.markdown(f"**Riesgo de Calidad {i+1}**")
            col1, col2, col3 = st.columns([2, 2, 1])
            with col1:
                r_riesgo = st.text_input(f"Riesgo:", value=risk.get("riesgo", ""), key=f"rcal_r_{i}")
            with col2:
                r_control = st.text_input(f"Control:", value=risk.get("control", ""), key=f"rcal_c_{i}")
            with col3:
                r_plazo = st.text_input(f"Plazo:", value=risk.get("plazo", ""), key=f"rcal_p_{i}")
            updated_cal.append({"riesgo": r_riesgo, "control": r_control, "plazo": r_plazo})
        data["riesgos_controles"] = updated_cal
        
        st.markdown("#### Riesgos SHES")
        risks_shes = data.get("riesgos_shes", [])
        updated_shes = []
        for i, risk in enumerate(risks_shes):
            st.markdown(f"**Riesgo SHES {i+1}**")
            col1, col2, col3 = st.columns([2, 2, 1])
            with col1:
                s_riesgo = st.text_input(f"Riesgo:", value=risk.get("riesgo", ""), key=f"rshes_r_{i}")
            with col2:
                s_control = st.text_input(f"Control:", value=risk.get("control", ""), key=f"rshes_c_{i}")
            with col3:
                s_plazo = st.text_input(f"Plazo:", value=risk.get("plazo", ""), key=f"rshes_p_{i}")
            updated_shes.append({"riesgo": s_riesgo, "control": s_control, "plazo": s_plazo})
        data["riesgos_shes"] = updated_shes
    
    with tabs[6]:
        st.markdown("#### Generar Documento Final")
        st.success("✅ Documento listo para generar (14 slides formato oficial MDET)")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            if st.button("🇪🇸 PPTX Español", type="primary", use_container_width=True):
                _finalize_document(data, meta, images_by_slide, "es", "moc", "pptx")
        with col2:
            if st.button("🇺🇸 PPTX Inglés", type="primary", use_container_width=True):
                _finalize_document(data, meta, images_by_slide, "en", "moc", "pptx")
        with col3:
            if st.button("📄 PDF Español", type="secondary", use_container_width=True):
                _finalize_document(data, meta, images_by_slide, "es", "moc", "pdf")
        with col4:
            if st.button("🔄 Regenerar", use_container_width=True):
                st.session_state.page = "nueva_moc"
                st.rerun()
    
    st.session_state.generated_data = data
    st.session_state.doc_meta = meta
    st.session_state.doc_images_by_slide = images_by_slide

def _render_a3_review(data, meta, config):
    gemini = GeminiService(config.get("gemini_api_key", ""), config.get("gemini_model"))
    tabs = st.tabs(["📋 General", "📝 Contenido", "⚙️ Generar"])
    with tabs[0]:
        meta["titulo"] = st.text_input("Título:", value=meta.get("titulo", ""))
        meta["area"] = st.text_input("Área:", value=meta.get("area", ""))
        meta["autor"] = st.text_input("Autor:", value=meta.get("autor", ""))
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
        col1, col2 = st.columns(2)
        with col1:
            if st.button("✅ DOCX Español", type="primary", use_container_width=True):
                _finalize_document(data, meta, {}, "es", "a3", "docx")
        with col2:
            if st.button("📄 PDF Español", type="secondary", use_container_width=True):
                _finalize_document(data, meta, {}, "es", "a3", "pdf")
    st.session_state.generated_data = data
    st.session_state.doc_meta = meta

def _render_kaizen_review(data, meta, config):
    gemini = GeminiService(config.get("gemini_api_key", ""), config.get("gemini_model"))
    tabs = st.tabs(["📋 General", "📝 Contenido", "⚙️ Generar"])
    with tabs[0]:
        meta["titulo"] = st.text_input("Título:", value=meta.get("titulo", ""))
        meta["area"] = st.text_input("Plant/Area:", value=meta.get("area", ""))
        meta["leader"] = st.text_input("Leader:", value=meta.get("leader", ""))
        meta["team_members"] = st.text_input("Team Members:", value=meta.get("team_members", ""))
    with tabs[1]:
        st.markdown("**Descripción del Problema**")
        data["descripcion_problema"] = _spell_check_field("", data.get("descripcion_problema", ""), "kzn_desc", gemini)
        st.markdown("**Solución**")
        data["solucion"] = _spell_check_field("", data.get("solucion", ""), "kzn_sol", gemini)
        st.markdown("**Beneficios**")
        data["beneficios"] = _spell_check_field("", data.get("beneficios", ""), "kzn_ben", gemini)
    with tabs[2]:
        col1, col2 = st.columns(2)
        with col1:
            if st.button("✅ PPTX Español", type="primary", use_container_width=True):
                _finalize_document(data, meta, {}, "es", "kaizen", "pptx")
        with col2:
            if st.button("📄 PDF Español", type="secondary", use_container_width=True):
                _finalize_document(data, meta, {}, "es", "kaizen", "pdf")
    st.session_state.generated_data = data
    st.session_state.doc_meta = meta

def _finalize_document(data, meta, images_by_slide, language, doc_type, output_format="pptx"):
    config = st.session_state.config
    gemini = GeminiService(config.get("gemini_api_key", ""), config.get("gemini_model"))
    with st.spinner(f"📄 Generando documento..."):
        final_data = {**meta, **data}
        if language == "en" and doc_type == "moc":
            final_data = gemini.translate_document(final_data)
        generator = DocumentGenerator()
        pdf_exporter = PDFExporter()
        
        if doc_type == "moc":
            buffer = generator.generate_moc(final_data, images_by_slide, st.session_state.get("template_moc_bytes"))
            ext = "pptx"
            mime = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            
            if output_format == "pdf" and buffer:
                pdf_bytes = pdf_exporter.pptx_to_pdf_libreoffice(buffer.getvalue(), "moc.pdf")
                if pdf_bytes:
                    buffer = BytesIO(pdf_bytes)
                    ext = "pdf"
                    mime = "application/pdf"
        elif doc_type == "a3":
            buffer = generator.generate_a3(final_data, None, st.session_state.get("template_a3_bytes"))
            ext = "docx"
            mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        else:
            buffer = generator.generate_kaizen(final_data, None, st.session_state.get("template_kaizen_bytes"))
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
    st.markdown('<div class="section-header"><h3>📁 Historial</h3></div>', unsafe_allow_html=True)
    docs = st.session_state.history.get("documents", [])
    if not docs:
        st.info("📭 No hay documentos generados aún.")
        return
    for doc in docs:
        type_emoji = {"moc": "📋", "a3": "📊", "kaizen": "⚡"}.get(doc.get("type"), "📄")
        st.markdown(f"""
<div class="history-item">
<h4 style="margin: 0;">{type_emoji} {doc.get('title', 'Sin título')}</h4>
<p style="margin: 0.25rem 0; color: #64748b; font-size: 0.9rem;">
{doc.get('number', '')} · {doc.get('timestamp', '')[:10]}
</p>
</div>
""", unsafe_allow_html=True)

def render_settings():
    st.markdown('<div class="section-header"><h3>⚙️ Configuración</h3></div>', unsafe_allow_html=True)
    config = st.session_state.config
    tabs = st.tabs(["🔑 API Gemini", "📄 Templates", "💾 Backup"])
    
    with tabs[0]:
        st.markdown("#### API Key Gemini")
        api_key = st.text_input("API Key:", value=config.get("gemini_api_key", ""), type="password")
        
        if st.button("🔌 Probar Conexión API", type="secondary", use_container_width=True):
            if not api_key:
                st.error("⚠️ Ingrese una API Key primero.")
            else:
                with st.spinner("Probando conexión..."):
                    try:
                        import requests
                        url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={api_key}"
                        payload = {"contents": [{"parts": [{"text": "Responde solo 'OK'"}]}]}
                        resp = requests.post(url, json=payload, timeout=15)
                        if resp.status_code == 200:
                            st.success("✅ ¡Conexión exitosa!")
                        else:
                            st.error(f"❌ Error {resp.status_code}: {resp.text}")
                    except Exception as e:
                        st.error(f"❌ Error: {e}")
        
        st.markdown("#### Selección de Modelo")
        current_model = config.get("gemini_model", GeminiService.DEFAULT_MODEL)
        col1, col2 = st.columns(2)
        for i, (model_id, model_info) in enumerate(GeminiService.MODELS.items()):
            is_selected = current_model == model_id
            with [col1, col2][i]:
                st.markdown(f"**{model_info['name']}** - {model_info['desc']}")
                if st.button(f"{'✓ ' if is_selected else ''}Seleccionar", key=f"sel_{model_id}", use_container_width=True):
                    config["gemini_model"] = model_id
                    st.session_state.config = config
                    LocalStorage.save_config(config)
                    st.success(f"✅ Modelo: {model_info['name']}")
                    st.rerun()
        
        if st.button("💾 Guardar API Key", type="primary", use_container_width=True):
            config["gemini_api_key"] = api_key
            st.session_state.config = config
            LocalStorage.save_config(config)
            st.success("✅ API Key guardada")
    
    with tabs[1]:
        st.markdown("#### Carga de Templates Oficiales")
        moc_ok = st.session_state.get("template_moc_bytes") is not None
        a3_ok = st.session_state.get("template_a3_bytes") is not None
        kzn_ok = st.session_state.get("template_kaizen_bytes") is not None
        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown(f"**Template MoC**: {'✅' if moc_ok else '❌'}")
            moc_file = st.file_uploader("Subir MoC (.pptx)", type=["pptx"], key="upload_moc")
        with col2:
            st.markdown(f"**Template A3**: {'✅' if a3_ok else '❌'}")
            a3_file = st.file_uploader("Subir A3 (.docx)", type=["docx"], key="upload_a3")
        with col3:
            st.markdown(f"**Template Kaizen**: {'✅' if kzn_ok else '❌'}")
            kzn_file = st.file_uploader("Subir Kaizen (.pptx)", type=["pptx"], key="upload_kzn")
        
        if st.button("💾 Guardar Templates", type="primary", use_container_width=True):
            if moc_file:
                st.session_state.template_moc_bytes = moc_file.getvalue()
                LocalStorage.save_template_bytes("moc", moc_file.getvalue())
            if a3_file:
                st.session_state.template_a3_bytes = a3_file.getvalue()
                LocalStorage.save_template_bytes("a3", a3_file.getvalue())
            if kzn_file:
                st.session_state.template_kaizen_bytes = kzn_file.getvalue()
                LocalStorage.save_template_bytes("kaizen", kzn_file.getvalue())
            st.success("✅ Templates guardados")
            st.rerun()
    
    with tabs[2]:
        export_data = {
            "config": st.session_state.config,
            "history": st.session_state.history,
            "export_date": datetime.now().isoformat(),
            "version": "9.0.0"
        }
        st.download_button(
            label="📥 Exportar backup (JSON)",
            data=json.dumps(export_data, indent=2, ensure_ascii=False),
            file_name=f"backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
            mime="application/json",
            use_container_width=True
        )

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
<p><strong>CAVA</strong> - Especialistas en Robótica y Automatización</p>
<p>Diseñado por <strong>Roger Huamani</strong> | v9.0.0</p>
</div>
""", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
