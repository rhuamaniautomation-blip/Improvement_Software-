#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
================================================================================
SISTEMA DE GESTION DOCUMENTAL - MoC | Mejora A3 | Simple Kaizen
Version 6.0.0 - Formato Oficial MDET con 12 Slides Estandarizados
================================================================================
Diseñado por: CAVA - Especialistas en Robotica y Automatizacion
Desarrollador: Roger Huamani
Version: 6.0.0
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
            "epoca": "época", "Epoca": "Época",
            "decada": "década", "Decada": "Década",
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
            "lento": "lento",
            "lenta": "lenta",
            "ademas": "además", "Ademas": "Además",
            "tambien": "también", "Tambien": "También",
            "esta": "está",
            "este": "éste",
            "esta": "está",
            "trabagar": "trabajar",
            "podra": "podrá",
        }
        result = text
        for wrong, correct in corrections.items():
            result = result.replace(wrong, correct)
        result = re.sub(r'  +', ' ', result)
        result = re.sub(r' ([.,;:!?])', r'\1', result)
        return result

# =============================================================================
# SERVICIO GEMINI API - PROMPTS MEJORADOS
# =============================================================================
class GeminiService:
    MODELS = {
        "gemini-1.5-flash-lite": {"name": "3.1 Flash-Lite", "desc": "Respuestas rápidas"},
        "gemini-1.5-flash": {"name": "3.5 Flash", "desc": "Ayuda completa"},
        "gemini-1.5-pro": {"name": "3.1 Pro", "desc": "Máxima calidad"},
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
        response = requests.post(url, json=payload, timeout=120)
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
        """Genera MoC con formato oficial MDET de 12 slides"""
        if not self.api_key:
            return self._generate_local_moc(problem, context, equipo)

        prompt = f"""Eres un ingeniero senior con 20 años de experiencia en la industria minera y manufacturera, especializado en gestión de cambios (Management of Change - MoC) bajo estándares internacionales (PSM, ISO 45001, ISO 9001). Redactas documentos técnicos impecables, con lenguaje profesional, humanizado y natural.

CONTEXTO DEL USUARIO:
PROBLEMA/CAMBIO REPORTADO: {problem}
INFORMACIÓN ADICIONAL: {context}
EQUIPO INVOLUCRADO: {equipo}

INSTRUCCIONES CRÍTICAS DE REDACCIÓN:
1. REDACCIÓN HUMANIZADA: Escribe como lo haría un ingeniero senior experimentado. Evita frases robóticas como "se identificó", "se determinó". Usa voz activa y construcciones naturales.
2. PÁRRAFOS DETALLADOS: Cada sección debe tener párrafos completos de 4-6 oraciones bien conectadas con conectores lógicos (por lo tanto, en consecuencia, asimismo, adicionalmente, cabe destacar).
3. VIÑETAS TÉCNICAS: Cuando listes elementos, usa viñetas con "❖" al inicio de cada una, manteniendo coherencia gramatical.
4. DATOS CUANTITATIVOS: Incluye valores numéricos realistas (dimensiones, tiempos, porcentajes, temperaturas, presiones) cuando aplique.
5. REFERENCIAS NORMATIVAS: Cita normas, procedimientos internos y mejores prácticas de la industria.
6. ORTOGRAFÍA IMPECABLE: Tildes correctas en todas las palabras (producción, operación, condición, modificación, verificación, implementación, evaluación, capacitación, documentación, estandarización, optimización, identificación, clasificación, notificación, coordinación, aprobación, revisión, ejecución, inspección, protección, detección, prevención, intervención, supervisión, comunicación, organización, planificación, calificación, certificación, validación, calibración, configuración, programación, automatización, integración, función, relación, conexión, dirección, selección, distribución, construcción, instrucción, reducción, traducción, máquina, podría, habría, sería, tendría, haría, daría, estaría, más, también, así, aquí, allí, allá, después, además, según, número, máximo, mínimo, óptimo, último, período, área, día, próximo, análisis, método, parámetro, característica, específico, genérico, electrónico, eléctrico, hidráulico, neumático, térmico, químico, físico, versión, descripción, solución, situación, límites, línea, único, fácil, rápido).
7. SIN ERRORES: Revisa mentalmente cada palabra antes de escribirla.

GENERAR JSON CON LA SIGUIENTE ESTRUCTURA EXACTA (12 SECCIONES PARA 12 SLIDES):

1. "moc_title": Título técnico conciso del cambio (máximo 12 palabras, estilo "OPTIMIZACIÓN DEL SISTEMA DE...")
2. "descripcion_problema": SLIDE 5. Párrafo técnico extenso (mínimo 250 palabras) describiendo el problema con causas técnicas, mecanismo de falla, consecuencias operativas y riesgos. Usar viñetas con "❖" para listar impactos numerados.
3. "condicion_actual": SLIDE 3 (columna izquierda). Descripción técnica detallada del estado actual (mínimo 150 palabras), incluyendo especificaciones técnicas, dimensiones, parámetros operativos y limitaciones documentadas.
4. "condicion_propuesta": SLIDE 3 (columna derecha). Descripción técnica de la solución propuesta (mínimo 150 palabras), con especificaciones de la modificación, beneficios esperados y método de gestión operativa.
5. "razones_cambio": SLIDE 4 (parte superior). Lista de 4-6 razones técnicas usando viñetas "❖" que justifiquen el cambio desde perspectivas de seguridad, productividad, calidad y cumplimiento normativo.
6. "alternativas_consideradas": SLIDE 4 (parte media). Análisis de al menos 2 alternativas evaluadas con sus pros/contras, explicando por qué se selecciona la propuesta.
7. "plan_retorno": SLIDE 4 (parte inferior). Procedimiento detallado de retorno a condiciones originales en caso de falla, con pasos específicos.
8. "recursos": SLIDE 8 (parte superior). Lista detallada de recursos humanos (con roles), herramientas, equipos y materiales requeridos usando viñetas.
9. "plan_implementacion": SLIDE 8 (parte media). Secuencia de actividades de implementación con pasos numerados o viñetas, incluyendo pruebas y validaciones.
10. "tiempo_duracion": SLIDE 8 (parte inferior). Estimación realista del tiempo total con desglose de actividades y consideraciones de ventanas de mantenimiento.
11. "checklist_360": Array de 16 objetos con esta estructura exacta:
    [
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
    ]
12. "documentos_impactados": Array de 15 objetos con esta estructura exacta:
    [
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
13. "riesgos_calidad": Array de 3-5 objetos {{"riesgo": "...", "control": "...", "plazo": "..."}}
14. "riesgos_shes": Array de 3-5 objetos {{"riesgo": "...", "control": "...", "plazo": "..."}}

IMPORTANTE:
- Responde SOLO con JSON válido, sin comentarios ni texto adicional.
- Todos los textos en ESPAÑOL.
- Ortografía impecable con todas las tildes correctas.
- Redacción profesional, técnica y humanizada.
- Párrafos extensos y bien estructurados.
- Viñetas con "❖" donde corresponda."""

        try:
            text = self._call_api(prompt, temperature=0.4, max_tokens=12000)
            result = self._extract_json(text)
            # Aplicar corrección ortográfica a todos los campos de texto
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
            st.error(f"Error API: {e}. Usando generación local.")
            return self._generate_local_moc(problem, context, equipo)

    def generate_a3(self, problem, context=""):
        if not self.api_key:
            return self._generate_local_a3(problem, context)
        prompt = f"""Eres un experto senior en metodología A3 Lean con 15 años de experiencia en mejora continua industrial. Redactas documentos A3 con redacción humanizada, técnica y profesional.

INSTRUCCIONES DE REDACCIÓN:
- Usa lenguaje profesional, directo y técnico como lo haría un Black Belt en Lean Six Sigma.
- Evita frases genéricas. Sé específico con datos, métricas y análisis cuantitativos.
- Incluye referencias a herramientas Lean (5S, SMED, TPM, VSM, etc.) cuando aplique.
- La redacción debe ser fluida, con párrafos bien estructurados y conectores lógicos.
- Incluye datos hipotéticos pero realistas cuando el usuario no proporcione números específicos.
- CORRIGE TODAS LAS FALTAS DE ORTOGRAFÍA: tildes en producción, operación, condición, modificación, verificación, capacitación, documentación, estandarización, optimización, identificación, clasificación, análisis, método, parámetro, característica, específico, genérico, electrónico, eléctrico, hidráulico, neumático, térmico, químico, físico, versión, descripción, solución, situación, límites, línea, único, fácil, rápido.

PROBLEMA REPORTADO: {problem}
CONTEXTO ADICIONAL: {context}

Genera en ESPAÑOL formato JSON con los siguientes campos:
1. titulo: Título conciso y descriptivo (máximo 10 palabras)
2. antecedentes: Contexto histórico del problema (mínimo 200 palabras)
3. problema_actual: Descripción detallada con datos cuantitativos (mínimo 250 palabras)
4. analisis_situacion: Análisis con datos y comparativas
5. objetivos: Objetivo general SMART y 3-5 objetivos específicos
6. analisis_causa_raiz: Análisis 5 Porqués y diagrama de Ishikawa conceptual
7. contramedidas: Lista de 5-8 contramedidas priorizadas
8. resultados_esperados: Resultados cuantificados esperados
9. plan_seguimiento: Plan de seguimiento detallado
10. lecciones_aprendidas: Reflexiones sobre el proceso
11. estandarizacion: Plan de estandarización

Responde SOLO JSON válido."""
        try:
            text = self._call_api(prompt, temperature=0.4, max_tokens=8192)
            result = self._extract_json(text)
            for key in result:
                if isinstance(result[key], str):
                    result[key] = Utils.correct_spelling_basic(result[key])
            return result
        except:
            return self._generate_local_a3(problem, context)

    def generate_kaizen(self, activity, context=""):
        if not self.api_key:
            return self._generate_local_kaizen(activity, context)
        prompt = f"""Eres un experto en Kaizen y Lean Manufacturing. Redactas registros Kaizen con redacción humanizada, práctica y motivadora.

INSTRUCCIONES:
- Lenguaje práctico, directo y motivador
- Datos cuantitativos específicos: tiempos antes/después, porcentajes
- Descripción visual detallada del antes y después
- Redacción natural, frases cortas y claras
- Impacto humano: beneficios al operario, equipo y organización
- ORTOGRAFÍA IMPECABLE: tildes correctas en producción, operación, condición, modificación, verificación, análisis, método, parámetro, máquina, podría, habría, sería, más, también, así, después, además, según, número, máximo, mínimo, área, día, próximo, descripción, solución, situación, límites.

ACTIVIDAD DE MEJORA: {activity}
CONTEXTO ADICIONAL: {context}

Genera en ESPAÑOL formato JSON con:
1. titulo: Título atractivo (máximo 8 palabras)
2. area: Área específica
3. descripcion_problema: Descripción vívida del problema (mínimo 200 palabras)
4. solucion: Descripción detallada de la solución (mínimo 200 palabras)
5. beneficios: Lista de beneficios cuantificados
6. tipo_desperdicio: Tipo(s) de desperdicio Lean eliminado(s)
7. impacto_bto: Categoría BTO impactada
8. proximos_pasos: Plan de acción concretos
9. leader: Nombre del líder
10. team_members: Lista de miembros

Responde SOLO JSON válido."""
        try:
            text = self._call_api(prompt, temperature=0.4, max_tokens=4096)
            result = self._extract_json(text)
            for key in result:
                if isinstance(result[key], str):
                    result[key] = Utils.correct_spelling_basic(result[key])
            return result
        except:
            return self._generate_local_kaizen(activity, context)

    def translate_document(self, data):
        if not self.api_key:
            return data
        prompt = f"""Traduce del español al inglés profesional industrial, manteniendo la terminología técnica apropiada:
{json.dumps(data, ensure_ascii=False, indent=2)}
Responde SOLO el JSON traducido, misma estructura exacta."""
        try:
            text = self._call_api(prompt, temperature=0.2, max_tokens=8192)
            return self._extract_json(text)
        except:
            return data

    def correct_spelling(self, text):
        if not self.api_key or not text.strip():
            return Utils.correct_spelling_basic(text)
        prompt = f"""Corrige ortografía, gramática, puntuación y mejora la redacción del siguiente texto en español. Mantén el significado técnico exacto. Mejora la fluidez y naturalidad. Asegúrate de poner todas las tildes correctas. Devuelve SOLO el texto corregido.

TEXTO:
{text}"""
        try:
            corrected = self._call_api(prompt, temperature=0.2, max_tokens=4096).strip()
            return Utils.correct_spelling_basic(corrected)
        except:
            return Utils.correct_spelling_basic(text)

    def _generate_local_moc(self, problem, context, equipo):
        """Generación local mejorada con formato MDET de 12 slides"""
        return {
            "moc_title": f"OPTIMIZACIÓN DEL SISTEMA - {problem[:60].upper()}",
            "descripcion_problema": f"Durante las operaciones rutinarias en el área de producción se ha identificado una condición técnica que afecta directamente la continuidad operativa y la confiabilidad del proceso productivo. El problema reportado consiste en: {problem}. Esta situación ha sido documentada mediante inspecciones de campo y análisis de datos operativos, revelando un patrón de recurrencia que compromete los indicadores clave de desempeño del área.\n\nLa condición actual presenta las siguientes características técnicas:\n❖ Operación fuera de los parámetros óptimos establecidos en los procedimientos operativos estándar.\n❖ Degradación progresiva de componentes críticos que afecta la disponibilidad del equipo.\n❖ Incremento en la frecuencia de intervenciones correctivas no programadas.\n❖ Exposición potencial del personal operativo a condiciones subóptimas de trabajo.\n\nEl impacto operacional se manifiesta en:\n1. Reducción de la disponibilidad mecánica del equipo crítico.\n2. Incremento de los costos de mantenimiento correctivo.\n3. Riesgo de paradas no programadas que afectan la programación productiva.\n4. Posible afectación a la calidad del producto final.\n\nDe no implementarse una solución técnica estructurada, se proyecta una deterioración progresiva de las condiciones operativas, con el consecuente incremento en los riesgos de seguridad y la disminución de la productividad global del área.",
            "condicion_actual": f"El sistema opera actualmente bajo las siguientes condiciones técnicas documentadas: {context if context else 'Configuración original de diseño con limitaciones operativas identificadas durante la operación rutinaria.'}\n\nLos parámetros críticos del proceso presentan desviaciones respecto a los valores nominales de diseño, lo cual ha sido validado mediante mediciones instrumentadas y análisis de tendencias históricas. Los componentes principales muestran signos de desgaste acelerado debido a las condiciones operativas actuales, generando la necesidad de intervenciones correctivas frecuentes que impactan la disponibilidad del equipo.\n\nLas limitaciones técnicas identificadas incluyen restricciones en la capacidad de procesamiento, incompatibilidad con nuevos requerimientos operativos y vulnerabilidad ante variaciones en las condiciones de alimentación del proceso.",
            "condicion_propuesta": f"Se propone la implementación de una modificación técnica estructurada que optimice el rendimiento operativo del sistema, mejorando significativamente las condiciones de seguridad y alineando las operaciones con los estándares corporativos y regulatorios vigentes.\n\nLa propuesta contempla:\n❖ Ingeniería detallada de la modificación con validación técnica por especialistas.\n❖ Fabricación e implementación de componentes con especificaciones mejoradas.\n❖ Actualización de los procedimientos operativos y de mantenimiento.\n❖ Capacitación del personal involucrado en los nuevos procedimientos.\n❖ Establecimiento de un plan de monitoreo post-implementación.\n\nEsta modificación se gestionará operativamente como un cambio controlado bajo el proceso de Management of Change (MoC), asegurando la trazabilidad completa del proceso y el cumplimiento de todos los requisitos normativos aplicables.",
            "razones_cambio": "❖ ELIMINACIÓN DE PARADAS NO PROGRAMADAS: La condición actual genera fallas recurrentes que obligan a intervenciones correctivas no planificadas, afectando la continuidad productiva y los compromisos de entrega al cliente.\n❖ ADAPTACIÓN A REQUERIMIENTOS OPERATIVOS ACTUALES: El sistema debe evolucionar para atender las demandas operativas actuales que exceden las capacidades del diseño original.\n❖ REDUCCIÓN DE INTERVENCIONES CORRECTIVAS: La modificación disminuirá significativamente la frecuencia de mantenimientos correctivos, optimizando el uso de recursos de mantenimiento.\n❖ MEJORA EN LA DISPONIBILIDAD DE EQUIPO (OEE): Se proyecta un incremento en el Overall Equipment Effectiveness mediante la eliminación de las causas raíz de las fallas recurrentes.\n❖ FORTALECIMIENTO DE LA SEGURIDAD OPERACIONAL: La modificación reduce los riesgos asociados a las intervenciones frecuentes y a la operación en condiciones subóptimas.\n❖ CUMPLIMIENTO NORMATIVO: El cambio asegura el alineamiento con los estándares corporativos y las mejores prácticas de la industria.",
            "alternativas_consideradas": "ALTERNATIVAS EVALUADAS:\n\n❖ ALTERNATIVA 1 - MANTENER SISTEMA ACTUAL (DESCARTADA): Esta opción implica continuar con intervenciones correctivas frecuentes, generando costos operativos crecientes, disminución de la disponibilidad del equipo y exposición prolongada a condiciones operativas subóptimas. El análisis costo-beneficio muestra que esta alternativa es insostenible a mediano plazo.\n\n❖ ALTERNATIVA 2 - REEMPLAZO TOTAL DEL SISTEMA (DESCARTADA): Aunque ofrece la solución más completa, el costo de inversión y el tiempo de implementación exceden los recursos disponibles en el período presupuestario actual. Adicionalmente, requeriría paradas prolongadas que afectarían significativamente la programación productiva.\n\n❖ ALTERNATIVA 3 - MODIFICACIÓN CONTROLADA (SELECCIONADA): Representa la mejor relación costo-beneficio, abordando las causas raíz identificadas con un alcance definido, tiempos de implementación razonables y un retorno de inversión favorable dentro del primer año de operación.\n\nPLAN DE RETORNO:\nEn caso de que la modificación no produzca los resultados esperados, se ejecutará el siguiente procedimiento:\n❖ Restauración inmediata de la configuración original del equipo.\n❖ Activación del protocolo de contingencia establecido.\n❖ Notificación oportuna a supervisión directa y áreas de apoyo.\n❖ Documentación detallada de lecciones aprendidas.\n❖ Análisis de causa raíz de la falla para prevenir recurrencias.",
            "plan_retorno": "PLAN DE RETORNO A CONDICIONES ORIGINALES:\n\nEn caso de presentarse complicaciones durante o después de la implementación que comprometan la operación segura del equipo, se ejecutará el siguiente procedimiento de retorno:\n\n❖ PASO 1: Detener la operación del equipo siguiendo el procedimiento de parada segura establecido.\n❖ PASO 2: Aislar energéticamente el equipo aplicando el procedimiento LOTO (Lock Out - Tag Out).\n❖ PASO 3: Desmontar los componentes modificados y almacenarlos de forma identificada para análisis posterior.\n❖ PASO 4: Reinstalar la configuración original utilizando los componentes de respaldo disponibles.\n❖ PASO 5: Realizar verificaciones pre-operativas según checklist establecido.\n❖ PASO 6: Ejecutar pruebas funcionales con carga reducida antes de retornar a operación normal.\n❖ PASO 7: Notificar a todas las áreas involucradas sobre la activación del plan de retorno.\n❖ PASO 8: Documentar el evento y convocar reunión de análisis con el equipo multidisciplinario.\n\nEl tiempo estimado de retorno a condiciones originales es de 2 horas, considerando la disponibilidad de componentes de respaldo y personal calificado.",
            "recursos": "RECURSOS HUMANOS:\n❖ Supervisor de área operativa (supervisión continua durante la implementación).\n❖ Técnico especialista de mantenimiento mecánico (2 personas, tiempo completo).\n❖ Especialista SHES (verificación de controles y permisos).\n❖ Operador de área certificado (apoyo operativo y pruebas).\n❖ Ingeniero de procesos (validación técnica y ajustes de parámetros).\n\nRECURSOS MATERIALES:\n❖ Herramientas especializadas certificadas y calibradas.\n❖ Repuestos de calidad certificada con trazabilidad documentada.\n❖ EPP completo: casco de seguridad, gafas de protección, guantes anticorte, botas dieléctricas, protección auditiva.\n❖ Materiales de señalización, demarcación y etiquetado del área de trabajo.\n❖ Materiales de limpieza y preparación de área.\n❖ Kit de contención de derrames (si aplica).\n\nRECURSOS TÉCNICOS:\n❖ Documentación técnica actualizada del equipo (manuales, diagramas, especificaciones).\n❖ Procedimientos operativos estándar (SOP) vigentes.\n❖ Permisos de trabajo según tipo de actividad (trabajo en caliente, espacio confinado, trabajo en altura, etc.).\n❖ Checklist de verificación pre y post implementación.\n❖ Equipos de prueba y medición calibrados.",
            "plan_implementacion": "FASE 1: PREPARACIÓN Y PLANIFICACIÓN\n❖ Reunión de coordinación multidisciplinaria con producción, mantenimiento y SHES.\n❖ Verificación exhaustiva de disponibilidad de todos los recursos materiales y humanos.\n❖ Preparación del área de trabajo: limpieza profunda, señalización de perímetro, aplicación de LOTO.\n❖ Briefing de seguridad con todo el equipo involucrado, revisión de riesgos y controles.\n❖ Verificación final de permisos de trabajo y autorizaciones requeridas.\n\nFASE 2: EJECUCIÓN DE MODIFICACIONES\n❖ Desmontaje controlado de los componentes existentes.\n❖ Implementación progresiva de las modificaciones técnicas según plan detallado.\n❖ Verificación dimensional y de alineamiento durante la instalación.\n❖ Registro fotográfico detallado del antes, durante y después de cada modificación.\n❖ Verificación intermedia SHES al finalizar cada etapa crítica.\n\nFASE 3: VALIDACIÓN Y PRUEBAS\n❖ Pruebas funcionales iniciales sin carga para verificar el correcto ensamblaje.\n❖ Pruebas funcionales bajo condiciones normales de operación.\n❖ Verificación de todos los parámetros críticos contra especificaciones de diseño.\n❖ Validación conjunta por supervisor de área, producción y especialista técnico.\n❖ Pruebas de estrés y verificación de límites operativos.\n\nFASE 4: CIERRE Y ESTANDARIZACIÓN\n❖ Actualización completa de toda la documentación técnica y operativa.\n❖ Capacitación formal al personal operativo sobre nuevos procedimientos.\n❖ Socialización de lecciones aprendidas con todas las áreas involucradas.\n❖ Cierre formal del MoC con firmas de aprobación de todas las partes.\n❖ Archivo del documento completo en el sistema de gestión documental.",
            "tiempo_duracion": "ESTIMACIÓN TOTAL DEL CAMBIO: 8 días hábiles distribuidos en 4 fases bien definidas.\n\nDESGLOSE POR FASE:\n❖ Fase 1 (Preparación y Planificación): 2 días hábiles.\n❖ Fase 2 (Ejecución de Modificaciones): 3 días hábiles.\n❖ Fase 3 (Validación y Pruebas): 2 días hábiles.\n❖ Fase 4 (Cierre y Estandarización): 1 día hábil.\n\nCONSIDERACIONES ESPECIALES:\n❖ La ejecución física de la modificación (desmontaje, montaje y alineación) se estima en 1.5 horas, aprovechando una ventana programada de mantenimiento o cambio de lote productivo.\n❖ Se ha incluido un margen de contingencia del 20% para imprevistos.\n❖ Las ventanas de mantenimiento serán coordinadas con producción con al menos 48 horas de anticipación.\n❖ La duración total puede ajustarse según condiciones operativas, disponibilidad de recursos y resultados de las verificaciones intermedias.\n❖ Se establecerán puntos de control diarios para monitorear el avance contra el cronograma establecido.",
            "checklist_360": [
                {"numero": 1, "factor": "Interacción o impacto con otras áreas/procesos", "aplica": "NO", "descripcion": ""},
                {"numero": 2, "factor": "Cambios en los procedimientos operativos, arranque y parada", "aplica": "SI", "descripcion": "Se actualizarán los procedimientos operativos estándar (SOP) para reflejar la nueva configuración del equipo, incluyendo nuevas secuencias de arranque, operación normal y parada segura. Se establecerán nuevos puntos de control operativo."},
                {"numero": 3, "factor": "Parámetros operativos y límites de control", "aplica": "SI", "descripcion": "Se redefinirán los parámetros operativos críticos y sus límites de control conforme a las nuevas especificaciones técnicas del sistema modificado. Se establecerán nuevas variables de monitoreo."},
                {"numero": 4, "factor": "Cambios en interfaces hombre-máquina y gestión de alarmas", "aplica": "NO", "descripcion": ""},
                {"numero": 5, "factor": "Compatibilidad de materiales, sustancias y equipos", "aplica": "SI", "descripcion": "Se verificará la compatibilidad de los nuevos componentes con los materiales existentes, asegurando la interoperabilidad del sistema modificado con el resto de equipos del proceso."},
                {"numero": 6, "factor": "Exposición ocupacional (ruido, polvo, ergonomía, etc.)", "aplica": "NO", "descripcion": ""},
                {"numero": 7, "factor": "Requerimientos de EPP y su compatibilidad", "aplica": "NO", "descripcion": ""},
                {"numero": 8, "factor": "Escenarios de emergencia y capacidad de respuesta", "aplica": "NO", "descripcion": ""},
                {"numero": 9, "factor": "Impacto en el almacenamiento y tránsito interno/externo", "aplica": "NO", "descripcion": ""},
                {"numero": 10, "factor": "Impactos ambientales y generación de residuos", "aplica": "NO", "descripcion": ""},
                {"numero": 11, "factor": "Impacto en la calidad del producto o servicio", "aplica": "SI", "descripcion": "La modificación mejorará la consistencia del proceso, reduciendo la variabilidad y mejorando los indicadores de calidad del producto final. Se establecerán nuevos controles de calidad."},
                {"numero": 12, "factor": "Cambios en roles, competencias y carga de trabajo", "aplica": "SI", "descripcion": "Se requerirá capacitación específica del personal operativo y de mantenimiento en los nuevos procedimientos. Se actualizará la matriz de competencias del área."},
                {"numero": 13, "factor": "Integridad de equipos, protecciones y sistemas de control", "aplica": "SI", "descripcion": "Se verificará la integridad mecánica de los nuevos componentes y se establecerán nuevos puntos de inspección en el plan de mantenimiento preventivo."},
                {"numero": 14, "factor": "Cumplimiento legal, normativo y permisos aplicables", "aplica": "NO", "descripcion": ""},
                {"numero": 15, "factor": "Cambios en las condiciones para trabajos especiales", "aplica": "NO", "descripcion": ""},
                {"numero": 16, "factor": "Cambios sucesivos que incrementan el riesgo global", "aplica": "NO", "descripcion": ""}
            ],
            "documentos_impactados": [
                {"numero": 1, "documento": "JSERA - IPERC", "aplica": "NO", "modificacion": ""},
                {"numero": 2, "documento": "Procedimiento de Trabajo, Instructivo/PO", "aplica": "SI", "modificacion": "Actualización del procedimiento operativo estándar (SOP) para reflejar la nueva configuración del equipo y los nuevos procedimientos de operación, arranque y parada."},
                {"numero": 3, "documento": "Formato/Checklist operativos", "aplica": "SI", "modificacion": "Actualización del checklist pre-operacional para incluir las nuevas verificaciones específicas del sistema modificado."},
                {"numero": 4, "documento": "Matriz de EPP", "aplica": "NO", "modificacion": ""},
                {"numero": 5, "documento": "MSDS de sustancias involucradas", "aplica": "NO", "modificacion": ""},
                {"numero": 6, "documento": "Mapa de Riesgos", "aplica": "NO", "modificacion": ""},
                {"numero": 7, "documento": "Plan de emergencias", "aplica": "NO", "modificacion": ""},
                {"numero": 8, "documento": "Plan de Mantenimiento", "aplica": "SI", "modificacion": "Actualización del plan de mantenimiento preventivo para incluir las nuevas tareas de inspección y mantenimiento de los componentes modificados."},
                {"numero": 9, "documento": "Matriz de impactos ambientales", "aplica": "NO", "modificacion": ""},
                {"numero": 10, "documento": "Plan Monitoreos SSO requeridos", "aplica": "NO", "modificacion": ""},
                {"numero": 11, "documento": "Plan de tráfico", "aplica": "NO", "modificacion": ""},
                {"numero": 12, "documento": "Matriz de competencias, plan de entrenamiento", "aplica": "SI", "modificacion": "Actualización de la matriz de competencias y diseño de plan de capacitación para el personal operativo y de mantenimiento en los nuevos procedimientos."},
                {"numero": 13, "documento": "Plan de calidad", "aplica": "NO", "modificacion": ""},
                {"numero": 14, "documento": "Planos y diagramas (layout, P&ID)", "aplica": "SI", "modificacion": "Actualización de los planos técnicos y diagramas del equipo para reflejar la nueva configuración del sistema modificado."},
                {"numero": 15, "documento": "Licencias y permisos aplicables", "aplica": "NO", "modificacion": ""}
            ],
            "riesgos_calidad": [
                {"riesgo": "Variabilidad en los parámetros operativos durante la puesta en marcha inicial del sistema modificado", "control": "Establecer protocolo de arranque gradual con monitoreo intensivo de parámetros críticos durante las primeras 72 horas de operación. Designar ingeniero de procesos para acompañamiento continuo.", "plazo": "Durante puesta en marcha"},
                {"riesgo": "Desviaciones en las especificaciones de los componentes fabricados que afecten el desempeño del sistema", "control": "Implementar inspección dimensional 100% de los componentes críticos antes de su instalación. Contar con planos de verificación y equipos de medición calibrados.", "plazo": "Antes de instalación"},
                {"riesgo": "Falta de estandarización en los nuevos procedimientos operativos que genere inconsistencias", "control": "Desarrollar procedimientos detallados con apoyo visual (fotografías, diagramas). Ejecutar capacitaciones prácticas con evaluación de competencias antes de autorizar la operación.", "plazo": "Antes de operación"},
                {"riesgo": "Degradación prematura de componentes por condiciones operativas no contempladas en el diseño", "control": "Establecer programa de monitoreo de condición con inspecciones semanales durante el primer mes y mensuales posteriormente. Definir indicadores de desgaste y umbrales de acción.", "plazo": "Post-implementación continuo"}
            ],
            "riesgos_shes": [
                {"riesgo": "Lesiones por manipulación manual de equipos y componentes durante el desmontaje e instalación", "control": "Capacitación específica en técnicas de levantamiento seguro. Uso obligatorio de EPP completo. Asignación de ayudantes para cargas superiores a 25 kg. Señalización del área de trabajo.", "plazo": "Antes del inicio de actividades"},
                {"riesgo": "Exposición a energías peligrosas durante las actividades de modificación (eléctrica, mecánica, neumática, hidráulica)", "control": "Aplicación estricta del procedimiento LOTO (Lock Out - Tag Out) en todos los puntos de energía. Verificación de ausencia de energía antes de iniciar trabajos. Supervisión continua por personal autorizado.", "plazo": "Durante toda la ejecución"},
                {"riesgo": "Generación de residuos sólidos, líquidos o peligrosos durante el proceso de modificación", "control": "Manejo seguro según procedimiento ambiental corporativo. Clasificación en origen de todos los residuos. Disposición en áreas autorizadas con registro de trazabilidad. Contenedores identificados y segregados.", "plazo": "Durante toda la ejecución"},
                {"riesgo": "Caídas a distinto nivel o mismo nivel durante las actividades de instalación", "control": "Uso obligatorio de arnés de seguridad para trabajos sobre 1.80m. Verificación de condiciones de orden y limpieza. Instalación de líneas de vida cuando aplique. Iluminación adecuada del área de trabajo.", "plazo": "Durante trabajos en altura"}
            ]
        }

    def _generate_local_a3(self, problem, context):
        return {
            "titulo": "Optimización del proceso: " + problem[:50],
            "antecedentes": "Durante los últimos seis meses, el área operativa ha experimentado una degradación progresiva en sus indicadores clave de desempeño. Se han registrado incrementos en tiempos de ciclo, aumento en la tasa de defectos y una reducción en la productividad general del proceso.\n\nEl análisis preliminar de datos históricos revela una tendencia creciente que, si no se aborda de manera estructurada, comprometerá los objetivos anuales de la organización. La metodología A3 fue seleccionada como herramienta principal para el análisis estructurado de esta situación.",
            "problema_actual": problem,
            "analisis_situacion": "La situación actual presenta múltiples indicadores de desempeño con oportunidades significativas de mejora. Se requiere una recopilación sistemática y rigurosa de datos para establecer una línea base sólida que permita cuantificar el impacto de las contramedidas propuestas.",
            "objetivos": "OBJETIVO GENERAL:\nOptimizar integralmente el proceso eliminando los desperdicios identificados y estableciendo un nuevo estándar de desempeño sostenible.\n\nOBJETIVOS ESPECÍFICOS (SMART):\n❖ Reducir el tiempo de ciclo en un 15% dentro de los próximos 3 meses.\n❖ Disminuir la tasa de defectos en un 20% durante el próximo trimestre.\n❖ Mejorar la productividad general del área en un 10% dentro de 6 meses.\n❖ Incrementar la satisfacción interna del cliente en un 25% según encuesta trimestral.\n❖ Reducir el costo operativo unitario en un 8% dentro del primer año.",
            "analisis_causa_raiz": "ANÁLISIS DE LOS 5 PORQUÉS:\n1. ¿POR QUÉ ocurre el problema? → Porque el proceso opera con una configuración inadecuada que genera variabilidad excesiva.\n2. ¿POR QUÉ la configuración es inadecuada? → Porque no existe una estandarización formal de los parámetros operativos críticos.\n3. ¿POR QUÉ no hay estandarización? → Porque los procedimientos operativos estándar (SOP) no han sido actualizados en los últimos 18 meses.\n4. ¿POR QUÉ no están actualizados? → Porque no existe un sistema de gestión documental efectivo.\n5. ¿POR QUÉ no hay sistema? → Porque falta una política clara de gestión del conocimiento.\n\nCAUSA RAÍZ IDENTIFICADA:\nAusencia de un sistema integral de gestión, actualización y control de SOP.",
            "contramedidas": "❖ ACTUALIZAR SOP DEL PROCESO: Revisar y actualizar todos los procedimientos operativos. Responsable: Ingeniero de Procesos. Plazo: 2 semanas.\n❖ IMPLEMENTAR CHECKLISTS DIARIOS: Diseñar y desplegar checklists de verificación diaria. Responsable: Supervisor de Área. Plazo: 1 semana.\n❖ CAPACITAR AL PERSONAL: Programar y ejecutar capacitaciones formales. Responsable: Especialista de Capacitación. Plazo: 3 semanas.\n❖ ESTABLECER KPIs VISUALES: Implementar tableros visuales con indicadores clave. Responsable: Líder de Mejora Continua. Plazo: 2 semanas.\n❖ PROGRAMAR AUDITORÍAS MENSUALES: Establecer auditorías formales mensuales. Responsable: Auditor Interno. Plazo: Inicio inmediato.",
            "resultados_esperados": "❖ Reducción medible y sostenida de desperdicios identificados.\n❖ Mejora sostenida en calidad del producto y consistencia del proceso.\n❖ Estandarización efectiva que reduzca la variabilidad operativa en al menos 30%.\n❖ Reducción del tiempo de ciclo en 15% con impacto directo en capacidad productiva.\n❖ Retorno de inversión estimado del 180% dentro del primer año.\n❖ Reducción de costos operativos unitarios en 8%.",
            "plan_seguimiento": "❖ SEMANA 1-2: Implementación de contramedidas iniciales. Monitoreo diario de cumplimiento.\n❖ SEMANA 3-4: Ejecución de capacitaciones. Evaluación de competencias.\n❖ MES 2: Primera auditoría formal. Evaluación de avance vs. objetivos iniciales.\n❖ MES 3: Evaluación integral vs. objetivos SMART establecidos.\n❖ MES 6: Revisión de sostenibilidad de mejoras.\n❖ TRIMESTRAL: Revisiones formales con gerencia.",
            "lecciones_aprendidas": "La aplicación de la metodología A3 permitió visualizar de manera integral la complejidad del problema y las interconexiones entre sus múltiples causas. La participación activa y multidisciplinaria del equipo fue fundamental para identificar la causa raíz real.\n\nSe aprendió que los problemas aparentemente técnicos frecuentemente tienen raíces en sistemas de gestión deficientes. La inversión en capacitación y estandarización genera retornos significativos a mediano plazo.",
            "estandarizacion": "Los procedimientos actualizados serán documentados formalmente con control de versiones, aprobados por gerencia de operaciones y calidad, socializados mediante capacitaciones estructuradas con evaluación de competencias, integrados al Sistema de Gestión de Calidad (SGC) existente y sujetos a revisión periódica anual como mínimo."
        }

    def _generate_local_kaizen(self, activity, context):
        return {
            "titulo": "Kaizen: " + activity[:50],
            "area": "Área de Mantenimiento / Producción / Calidad",
            "descripcion_problema": activity + "\n\nDurante las actividades diarias de gemba walk, el equipo identificó esta oportunidad de mejora que representa un desperdicio significativo en el proceso. La situación actual genera movimientos innecesarios, tiempos de espera o riesgos de calidad que impactan directamente en la productividad del área y en la satisfacción del personal.",
            "solucion": "Se implementó una mejora estructurada orientada a eliminar el desperdicio identificado y optimizar el flujo del proceso, aplicando principios fundamentales de Lean Manufacturing y el pensamiento Kaizen de mejora continua.\n\nLa solución fue diseñada y ejecutada por el equipo de trabajo del área con apoyo del líder de mejora continua, utilizando materiales disponibles y aplicando el concepto de low cost, high impact.",
            "beneficios": "❖ Reducción del tiempo de ejecución en aproximadamente 20-30%\n❖ Mejora significativa en calidad y consistencia del proceso\n❖ Mayor seguridad para el personal al eliminar movimientos riesgosos\n❖ Reducción de costos operativos derivados de la eliminación de desperdicios\n❖ Mejora en el ambiente de trabajo y orden del área\n❖ Fácil replicabilidad en otras áreas similares",
            "tipo_desperdicio": "Motion / Waiting / Skills",
            "impacto_bto": "Supply Chain and Manufacturing Excellence",
            "proximos_pasos": "❖ Documentar formalmente la mejora con fotografías y datos de impacto\n❖ Socializar la mejora con otras áreas relacionadas\n❖ Replicar la mejora en procesos similares\n❖ Establecer monitoreo mensual para asegurar sostenibilidad\n❖ Reconocer formalmente al equipo participante\n❖ Integrar el nuevo estándar al SOP del área",
            "leader": "",
            "team_members": ""
        }

# =============================================================================
# REEMPLAZO INTELIGENTE DE TEXTO EN POWERPOINT Y WORD
# =============================================================================
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
    """Reemplaza múltiples textos en toda la presentación"""
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

def replace_text_in_docx_preserve_runs(doc, old_text, new_text):
    """Reemplaza texto en documento Word preservando formato de runs"""
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
# GENERADOR DE DOCUMENTOS - FORMATO OFICIAL MDET 12 SLIDES
# =============================================================================
class DocumentGenerator:
    """Genera documentos usando templates cargados en memoria"""

    def generate_moc(self, data, images=None, template_bytes=None):
        """Genera MoC desde template en memoria - Formato oficial MDET 12 slides"""
        if template_bytes is None:
            st.error("❌ Template MoC no cargado. Vaya a Configuración > Templates.")
            return None

        prs = Presentation(BytesIO(template_bytes))

        # Mapeo de reemplazos globales (portada y datos generales)
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

        # SLIDE 2: Equipo que participó en la revisión
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

        # SLIDE 4: Razones del cambio + Alternativas + Plan de retorno
        if len(prs.slides) > 3:
            slide4 = prs.slides[3]
            for shape in slide4.shapes:
                if shape.has_text_frame:
                    text = shape.text_frame.text
                    # Reemplazar razones del cambio
                    if "Razones del cambio" in text or "Eliminación de paradas" in text:
                        for paragraph in shape.text_frame.paragraphs:
                            para_text = paragraph.text
                            if "Eliminación" in para_text or "Adaptación" in para_text or "Reducción" in para_text or "Mejora" in para_text:
                                for run in paragraph.runs:
                                    run.text = ""
                        # Insertar nuevas razones
                        if shape.text_frame.paragraphs:
                            first_para = shape.text_frame.paragraphs[0]
                            if first_para.runs:
                                first_para.runs[0].text = data.get('razones_cambio', '')
                            else:
                                first_para.text = data.get('razones_cambio', '')

                    # Reemplazar alternativas
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

                    # Reemplazar plan de retorno
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
                        # Limpiar contenido existente
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                run.text = ""
                        # Insertar nueva descripción
                        if shape.text_frame.paragraphs:
                            first_para = shape.text_frame.paragraphs[0]
                            if first_para.runs:
                                first_para.runs[0].text = data.get('descripcion_problema', '')
                            else:
                                first_para.text = data.get('descripcion_problema', '')

        # SLIDE 6 y 7: Imágenes (Vista general y Planos)
        if images and len(prs.slides) > 5:
            # Insertar imágenes en slides existentes o agregar nuevos
            for idx, img_info in enumerate(images):
                if idx < 2 and len(prs.slides) > 5 + idx:
                    # Usar slides existentes 6 y 7
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
                    # Agregar slides adicionales para más imágenes
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

        # SLIDE 8: Recursos y Plan de implementación + Tiempo
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

        # SLIDE 9: Checklist 360° (Tabla de 16 factores)
        if len(prs.slides) > 8:
            slide9 = prs.slides[8]
            checklist = data.get('checklist_360', [])
            for shape in slide9.shapes:
                if shape.has_table:
                    table = shape.table
                    for i, item in enumerate(checklist):
                        row_idx = i + 1  # +1 para saltar encabezado
                        if row_idx < len(table.rows):
                            # Columna N°
                            if len(table.columns) > 0:
                                fill_table_cell(table.cell(row_idx, 0), str(item.get('numero', i+1)))
                            # Columna Factor
                            if len(table.columns) > 1:
                                fill_table_cell(table.cell(row_idx, 1), item.get('factor', ''))
                            # Columna Aplica
                            if len(table.columns) > 2:
                                fill_table_cell(table.cell(row_idx, 2), item.get('aplica', 'NO'))
                            # Columna Descripción
                            if len(table.columns) > 3:
                                fill_table_cell(table.cell(row_idx, 3), item.get('descripcion', ''))

        # SLIDE 10: Documentos impactados (Tabla de 15 documentos)
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

        # SLIDE 11: Evaluación o Estudio de Riesgos
        # Esta slide generalmente tiene contenido de referencia, se mantiene

        # SLIDE 12: Riesgos SHES (Tabla)
        if len(prs.slides) > 11:
            slide12 = prs.slides[11]
            riesgos_shes = data.get('riesgos_shes', [])
            riesgos_calidad = data.get('riesgos_calidad', [])
            # Combinar riesgos de calidad y SHES
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
        """Genera A3 desde template en memoria"""
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
            ("ANTECEDENTES", "antecedentes"),
            ("PROBLEMA ACTUAL", "problema_actual"),
            ("ANÁLISIS DE LA SITUACIÓN", "analisis_situacion"),
            ("OBJETIVOS", "objetivos"),
            ("ANÁLISIS DE CAUSA RAÍZ", "analisis_causa_raiz"),
            ("CONTRAMEDIDAS", "contramedidas"),
            ("RESULTADOS ESPERADOS", "resultados_esperados"),
            ("PLAN DE SEGUIMIENTO", "plan_seguimiento"),
            ("LECCIONES APRENDIDAS", "lecciones_aprendidas"),
            ("ESTANDARIZACIÓN", "estandarizacion"),
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
        """Genera Kaizen desde template en memoria"""
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
                    ("Alternativas y Plan de Retorno", "alternativas_consideradas"),
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
    model_name = GeminiService.MODELS.get(config.get("gemini_model", "gemini-1.5-pro"), {}).get("name", "3.1 Pro")
    st.sidebar.markdown(f"""
<div style="text-align: center; color: #64748b; font-size: 0.75rem;">
<p>Modelo IA: <span class="gemini-badge">{model_name}</span></p>
<p>v6.0.0 · Agosto 2026</p>
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
<p style="color: #64748b; font-size: 0.9rem;">Formato oficial MDET de 12 slides con Checklist 360° y análisis integral.</p>
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
    st.info("💡 Complete la información y describa el problema con detalle. La IA generará automáticamente los 12 slides del formato oficial MDET con redacción profesional, Checklist 360° y análisis integral.")
    if not st.session_state.get("template_moc_bytes"):
        st.error("❌ **Template MoC no cargado.** Vaya a Configuración > Templates.")
        if st.button("Ir a Configuración", key="go_config_moc"):
            st.session_state.page = "configuracion"
            st.rerun()
        return
    st.markdown("#### 1. Información General")
    col1, col2 = st.columns(2)
    with col1:
        moc_title = st.text_input("Título de la MoC:", placeholder="Ej: OPTIMIZACIÓN DEL SISTEMA DE ALIMENTACIÓN DE RETARDOS EN ST08")
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
    st.markdown("#### 3. Descripción del Problema/Cambio (Sea lo más detallado posible)")
    problem_desc = auto_correct_text_input(
        "Describa el problema o cambio con sus palabras:",
        "",
        "moc_problem_desc",
        height=250,
        help_text="Cuanto más detalle proporcione, mejor será la generación. Incluya: qué está pasando, desde cuándo, impacto, equipos involucrados, riesgos observados, dimensiones, parámetros técnicos."
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
    if st.button("🤖 Generar Documento MoC con IA (12 Slides)", type="primary", use_container_width=True):
        if not problem_desc.strip():
            st.error("❌ Describa el problema antes de generar.")
            return
        with st.spinner("🧠 La IA está generando los 12 slides del formato oficial MDET..."):
            gemini = GeminiService(config.get("gemini_api_key", ""), config.get("gemini_model", "gemini-1.5-pro"))
            equipo_data = {
                "produccion": produccion, "specialist_shes": specialist_shes,
                "mantenimiento": mantenimiento, "revisores": revisores,
                "experto_aprobador": experto_aprobador
            }
            result = gemini.generate_moc(problem_desc, context, json.dumps(equipo_data))
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
    st.info("💡 Describa el problema con detalle y la IA generará el documento A3 completo.")
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

# =============================================================================
# PANTALLA DE REVISIÓN
# =============================================================================
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
    tabs = st.tabs(["📋 General", "📝 Contenido", "📊 Checklist 360°", "📄 Documentos", "⚠️ Riesgos", "📷 Imágenes", "⚙️ Generar"])
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
        st.markdown("#### Alternativas Consideradas y Plan de Retorno")
        data["alternativas_consideradas"] = _spell_check_field("", data.get("alternativas_consideradas", ""), "moc_alt", gemini)
        st.markdown("#### Recursos")
        data["recursos"] = _spell_check_field("", data.get("recursos", ""), "moc_rec", gemini)
        st.markdown("#### Plan de Implementación")
        data["plan_implementacion"] = _spell_check_field("", data.get("plan_implementacion", ""), "moc_plan", gemini)
        st.markdown("#### Tiempo de Duración")
        data["tiempo_duracion"] = _spell_check_field("", data.get("tiempo_duracion", ""), "moc_tiempo", gemini)
    with tabs[2]:
        st.markdown("#### Checklist 360° - 16 Factores")
        st.info("Análisis integral del cambio. Marque SI/NO y describa el impacto cuando corresponda.")
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
    with tabs[3]:
        st.markdown("#### Documentos Impactados - 15 Documentos")
        st.info("Identifique los documentos que deben actualizarse como consecuencia del cambio.")
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
    with tabs[4]:
        st.markdown("#### Riesgos de Calidad")
        risks_cal = data.get("riesgos_calidad", [])
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
        data["riesgos_calidad"] = updated_cal
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
    with tabs[5]:
        st.markdown("#### Imágenes Cargadas")
        if images:
            for idx, img_info in enumerate(images, 1):
                st.image(img_info["path"], caption=f"Figura {idx}: {img_info['desc']}", width=400)
        else:
            st.info("No se cargaron imágenes")
    with tabs[6]:
        st.markdown("#### Generar Documento Final")
        st.success("✅ Documento listo para generar (12 slides formato oficial MDET)")
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

# =============================================================================
# FINALIZACIÓN DE DOCUMENTO
# =============================================================================
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
                        pdf_bytes = pdf_exporter.generate_pdf_from_data(final_data, doc_type, meta, images)
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
                        pdf_bytes = pdf_exporter.generate_pdf_from_data(final_data, doc_type, meta, images)
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
                        pdf_bytes = pdf_exporter.generate_pdf_from_data(final_data, doc_type, meta, images)
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

# =============================================================================
# HISTORIAL
# =============================================================================
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
            "version": "6.0.0"
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

# =============================================================================
# CONFIGURACIÓN
# =============================================================================
def render_settings():
    st.markdown('<div class="section-header"><h3>⚙️ Configuración del Sistema</h3></div>', unsafe_allow_html=True)
    config = st.session_state.config
    tabs = st.tabs(["🔑 API Gemini", "🏢 Empresa", "📄 Templates", "🔧 Avanzado", "💾 Backup"])
    with tabs[0]:
        st.markdown("#### API Key Gemini")
        st.info("💡 Obtenga su API Key gratuita en [Google AI Studio](https://aistudio.google.com/)")
        api_key = st.text_input("API Key:", value=config.get("gemini_api_key", ""), type="password")
        st.markdown("#### Selección de Modelo")
        current_model = config.get("gemini_model", "gemini-1.5-pro")
        col1, col2, col3 = st.columns(3)
        models = [
            ("gemini-1.5-flash-lite", "⚡ 3.1 Flash-Lite", "Respuestas rápidas", "Económico"),
            ("gemini-1.5-flash", "🔥 3.5 Flash", "Ayuda completa", "Balance"),
            ("gemini-1.5-pro", "🧠 3.1 Pro", "Máxima calidad", "Recomendado"),
        ]
        for i, (model_id, name, desc, badge) in enumerate(models):
            is_selected = current_model == model_id
            with [col1, col2, col3][i]:
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
            "version": "6.0.0"
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

# =============================================================================
# FUNCIÓN PRINCIPAL
# =============================================================================
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
<p>Diseñado por <strong>Roger Huamani</strong> | Sistema de Gestión Documental v6.0.0</p>
<p style="font-size: 0.75rem; color: #94a3b8;">
Software empresarial para automatización de documentos MoC, A3 y Kaizen.<br>
Formato oficial MDET de 12 slides con Checklist 360° y análisis integral.<br>
Datos persistentes locales. Exportación a PDF integrada.
</p>
</div>
""", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
