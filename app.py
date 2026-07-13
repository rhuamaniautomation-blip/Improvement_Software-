#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
================================================================================
SISTEMA DE GESTION DOCUMENTAL - MoC | Mejora A3 | Simple Kaizen
Version 4.0.0 - Completo con Exportación a PDF y Persistencia Local
================================================================================
Diseñado por: CAVA - Especialistas en Robotica y Automatizacion
Desarrollador: Roger Huamani
Version: 4.0.0
Fecha: Julio 2026
================================================================================
MEJORAS v4.0.0:
- Generación automática de formatos MoC, A3 y Kaizen con IA detallada
- Corrector ortográfico automático en todos los campos
- Redacción humanizada y profesional
- Exportación directa a PDF (PowerPoint y Word -> PDF)
- Persistencia local de datos mediante archivos JSON
- Historial completo con búsqueda y filtros
- Templates oficiales integrados (no requieren carga manual)
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

# Intentar importar reportlab para PDF
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

# Crear directorios si no existen
DATA_DIR.mkdir(exist_ok=True)
TEMPLATES_DIR.mkdir(exist_ok=True)

# =============================================================================
# CSS PERSONALIZADO EMPRESARIAL
# =============================================================================
CUSTOM_CSS = """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', 'Segoe UI', sans-serif !important;
    }

    .main-header {
        background: linear-gradient(135deg, #1a5f7a 0%, #2e8bc0 100%);
        padding: 2rem;
        border-radius: 12px;
        color: white;
        margin-bottom: 2rem;
        box-shadow: 0 4px 20px rgba(26, 95, 122, 0.3);
    }

    .main-header h1 { color: white !important; font-weight: 700 !important; margin-bottom: 0.5rem !important; }
    .main-header p { color: rgba(255,255,255,0.9) !important; font-size: 1.1rem !important; }

    .doc-card {
        background: white;
        border-radius: 16px;
        padding: 2rem;
        border: 2px solid #e2e8f0;
        transition: all 0.3s ease;
        cursor: pointer;
        height: 100%;
    }
    .doc-card:hover {
        border-color: #1a5f7a;
        transform: translateY(-4px);
        box-shadow: 0 12px 40px rgba(26, 95, 122, 0.15);
    }
    .doc-card-moc { border-left: 5px solid #1a5f7a; }
    .doc-card-a3 { border-left: 5px solid #10b981; }
    .doc-card-kaizen { border-left: 5px solid #f59e0b; }

    .stButton > button {
        border-radius: 10px !important;
        font-weight: 600 !important;
        padding: 0.75rem 2rem !important;
        transition: all 0.2s ease !important;
    }
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
    }

    .stTextInput > div > div > input,
    .stTextArea > div > div > textarea,
    .stSelectbox > div > div > div {
        border-radius: 8px !important;
        border: 1px solid #cbd5e1 !important;
        font-size: 15px !important;
    }

    .section-header {
        background: #f1f5f9;
        padding: 1rem 1.5rem;
        border-radius: 10px;
        margin: 1.5rem 0 1rem 0;
        border-left: 4px solid #1a5f7a;
    }
    .section-header h3 { margin: 0 !important; color: #1e293b !important; font-weight: 600 !important; }

    .field-card {
        background: white;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 1rem;
        margin: 0.5rem 0;
    }
    .field-card:hover {
        border-color: #1a5f7a;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
    }

    .gemini-badge {
        display: inline-block;
        background: #e0e7ff;
        color: #4338ca;
        padding: 0.25rem 0.75rem;
        border-radius: 20px;
        font-size: 12px;
        font-weight: 600;
        margin-left: 0.5rem;
    }

    .history-item {
        background: white;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 1rem;
        margin: 0.5rem 0;
        transition: all 0.2s;
    }
    .history-item:hover {
        border-color: #1a5f7a;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
    }

    .app-footer {
        text-align: center;
        padding: 2rem;
        margin-top: 3rem;
        border-top: 1px solid #e2e8f0;
        color: #64748b;
    }

    .auto-correct-badge {
        display: inline-flex;
        align-items: center;
        background: #dcfce7;
        color: #166534;
        padding: 0.25rem 0.75rem;
        border-radius: 20px;
        font-size: 11px;
        font-weight: 600;
        margin-bottom: 0.5rem;
    }

    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}

    [data-testid="stSidebar"] { background: #1e293b !important; }
    [data-testid="stSidebar"] .stMarkdown { color: #94a3b8 !important; }
    [data-testid="stSidebar"] h1, 
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3 { color: #f8fafc !important; }

    .template-uploader {
        background: #f8fafc;
        border: 2px dashed #cbd5e1;
        border-radius: 12px;
        padding: 1.5rem;
        text-align: center;
        margin: 1rem 0;
    }
    .template-uploader.ok {
        background: #f0fdf4;
        border-color: #10b981;
    }
</style>
"""

st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# =============================================================================
# PERSISTENCIA LOCAL - GUARDAR/CARGAR DATOS
# =============================================================================
class LocalStorage:
    """Gestiona la persistencia local de configuración e historial"""

    @staticmethod
    def save_config(config):
        """Guarda configuración en archivo JSON local"""
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            st.error(f"Error guardando configuración: {e}")
            return False

    @staticmethod
    def load_config():
        """Carga configuración desde archivo JSON local"""
        try:
            if CONFIG_FILE.exists():
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    return json.load(f)
        except Exception as e:
            st.warning(f"Error cargando configuración: {e}")
        return None

    @staticmethod
    def save_history(history):
        """Guarda historial en archivo JSON local"""
        try:
            with open(HISTORY_FILE, 'w', encoding='utf-8') as f:
                json.dump(history, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            st.error(f"Error guardando historial: {e}")
            return False

    @staticmethod
    def load_history():
        """Carga historial desde archivo JSON local"""
        try:
            if HISTORY_FILE.exists():
                with open(HISTORY_FILE, 'r', encoding='utf-8') as f:
                    return json.load(f)
        except Exception as e:
            st.warning(f"Error cargando historial: {e}")
        return {"documents": []}

    @staticmethod
    def save_template_bytes(template_name, file_bytes):
        """Guarda bytes de template en archivo"""
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
        """Carga bytes de template desde archivo"""
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
        """Elimina archivo de template"""
        try:
            template_path = TEMPLATES_DIR / f"{template_name}.bin"
            if template_path.exists():
                template_path.unlink()
        except Exception:
            pass


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
        """Agrega documento al historial persistente"""
        history = st.session_state.history
        doc_info["id"] = str(uuid.uuid4())
        doc_info["timestamp"] = datetime.now().isoformat()
        history["documents"].insert(0, doc_info)
        st.session_state.history = history
        LocalStorage.save_history(history)

    @staticmethod
    def get_history(doc_type=None):
        """Obtiene historial filtrado"""
        docs = st.session_state.history.get("documents", [])
        if doc_type:
            docs = [d for d in docs if d.get("type") == doc_type]
        return docs

    @staticmethod
    def delete_from_history(doc_id):
        """Elimina documento del historial"""
        history = st.session_state.history
        history["documents"] = [d for d in history["documents"] if d.get("id") != doc_id]
        st.session_state.history = history
        LocalStorage.save_history(history)

    @staticmethod
    def correct_spelling_basic(text):
        """Corrector ortográfico básico local (sin API)"""
        if not text or not text.strip():
            return text

        # Correcciones comunes en español técnico industrial
        corrections = {
            "tecnico": "técnico", "Tecnico": "Técnico", "TECNICO": "TÉCNICO",
            "produccion": "producción", "Produccion": "Producción",
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
        }

        result = text
        for wrong, correct in corrections.items():
            result = result.replace(wrong, correct)

        # Espacios dobles
        result = re.sub(r'  +', ' ', result)
        # Espacio antes de punto/coma
        result = re.sub(r' ([.,;:!?])', r'', result)

        return result

# =============================================================================
# SERVICIO GEMINI API
# =============================================================================
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

        prompt = f"""Eres un experto senior en gestion de cambios industriales (MoC) con 20 años de experiencia en minería, manufactura y operaciones industriales. Redactas documentos profesionales, impecables, con redacción humanizada, natural y técnica, sin errores ortográficos.

INSTRUCCIONES DE REDACCIÓN:
- Usa lenguaje profesional pero natural, como lo haría un ingeniero senior experimentado.
- Evita frases genéricas o robotizadas como "se identificó la siguiente situación".
- Sé específico, detallado y técnico. Incluye referencias a normas, procedimientos y mejores prácticas.
- Usa conectores lógicos, párrafos bien estructurados y vocabulario técnico apropiado.
- Incluye datos cuantitativos cuando sea posible (tiempos, porcentajes, métricas).
- La redacción debe parecer escrita por un profesional humano, no por una IA.

PROBLEMA REPORTADO: {problem}
CONTEXTO ADICIONAL: {context}
EQUIPO INVOLUCRADO: {equipo}

Genera en ESPAÑOL formato JSON con los siguientes campos detallados y humanizados:

1. descripcion_problema: Descripción técnica detallada del problema, con causas identificadas, impacto operacional, riesgos actuales y consecuencias si no se actúa. Mínimo 300 palabras. Redacción fluida y profesional.

2. condicion_actual: Descripción técnica exhaustiva del estado actual del equipo/proceso, incluyendo especificaciones técnicas, parámetros operativos, limitaciones documentadas y referencias a normativas aplicables.

3. condicion_propuesta: Descripción detallada de la solución propuesta, incluyendo especificaciones técnicas de la modificación, beneficios esperados cuantificados, alineación con estándares corporativos y mejores prácticas de la industria.

4. razones_cambio: Lista numerada y justificada de las razones técnicas, de seguridad, regulatorias y económicas que sustentan el cambio. Cada razón debe tener una explicación de al menos 2-3 líneas.

5. alternativas_retorno: Análisis de alternativas evaluadas con pros y contras de cada una. Incluir plan de retorno detallado con pasos específicos, responsables y tiempos estimados.

6. recursos: Listado exhaustivo de recursos humanos (con roles y responsabilidades específicas), materiales (con especificaciones técnicas), técnicos (herramientas, software, documentación) y financieros requeridos.

7. plan_implementacion: Plan detallado por fases con actividades específicas, responsables asignados, duración estimada, hitos de control y criterios de aceptación para cada fase. Mínimo 4 fases bien definidas.

8. tiempo_duracion: Estimación detallada del tiempo total con desglose por fase, consideraciones de ventanas de mantenimiento, contingencias y factores que pueden afectar la duración.

9. riesgos_controles: Array de mínimo 5 objetos {{"riesgo":"descripción detallada del riesgo","control":"medida de control específica con responsable y frecuencia"}}. Incluir riesgos técnicos, operacionales y de calidad.

10. riesgos_shes: Array de mínimo 5 objetos {{"riesgo":"descripción detallada del riesgo SHES","control":"plan de acción específico con medidas preventivas","plazo":"plazo de implementación del control"}}. Cubrir seguridad, salud, medio ambiente y comunidad.

Responde SOLO JSON válido sin comentarios adicionales."""

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

PROBLEMA REPORTADO: {problem}
CONTEXTO ADICIONAL: {context}

Genera en ESPAÑOL formato JSON con los siguientes campos detallados:

1. titulo: Título conciso y descriptivo de la mejora (máximo 10 palabras).

2. antecedentes: Contexto histórico del problema, datos de línea base, tendencias observadas y por qué es relevante abordarlo ahora. Mínimo 200 palabras.

3. problema_actual: Descripción detallada del problema con datos cuantitativos, frecuencia de ocurrencia, impacto en KPIs críticos y consecuencias operacionales. Mínimo 250 palabras.

4. analisis_situacion: Análisis de la situación actual usando datos, gráficos conceptuales descritos en texto, comparativas con benchmarks de la industria y análisis de variabilidad del proceso.

5. objetivos: Objetivo general SMART y 3-5 objetivos específicos con métricas cuantificables, plazos y líneas base.

6. analisis_causa_raiz: Análisis de causa raíz detallado usando la metodología de los 5 Porqués, diagrama de Ishikawa conceptual descrito en texto, y validación de hipótesis. Identificar la causa raíz fundamental.

7. contramedidas: Lista de 5-8 contramedidas específicas, priorizadas, con responsable asignado, fecha de implementación y criterio de éxito medible para cada una.

8. resultados_esperados: Resultados cuantificados esperados con proyecciones de ahorro, mejora en indicadores clave, retorno de inversión estimado y beneficios intangibles.

9. plan_seguimiento: Plan de seguimiento detallado con frecuencia de revisión, indicadores a monitorear, responsables de seguimiento y criterios de éxito a corto, mediano y largo plazo.

10. lecciones_aprendidas: Reflexiones sobre el proceso de análisis, desafíos encontrados, aprendizajes clave y recomendaciones para futuras iniciativas similares.

11. estandarizacion: Plan de estandarización detallado con documentos a actualizar, capacitaciones requeridas, integración al SGC y mecanismos de sostenibilidad.

Responde SOLO JSON válido sin comentarios adicionales."""

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

        prompt = f"""Eres un experto en Kaizen y Lean Manufacturing con amplia experiencia en gemba walks y mejora continua en operaciones industriales. Redactas registros Kaizen con redacción humanizada, práctica y motivadora.

INSTRUCCIONES DE REDACCIÓN:
- Usa lenguaje práctico, directo y motivador como lo haría un líder de mejora continua en el gemba.
- Incluye datos cuantitativos específicos: tiempos antes/después, cantidades, porcentajes de mejora.
- Describe la situación actual con detalle visual para que el lector pueda imaginar el antes y después.
- La redacción debe ser natural, con frases cortas y claras, evitando tecnicismos innecesarios.
- Incluye el impacto humano: cómo beneficia al operario, al equipo y a la organización.

ACTIVIDAD DE MEJORA: {activity}
CONTEXTO ADICIONAL: {context}

Genera en ESPAÑOL formato JSON con los siguientes campos detallados:

1. titulo: Título atractivo y descriptivo de la mejora Kaizen (máximo 8 palabras).

2. area: Área específica donde se implementó la mejora.

3. descripcion_problema: Descripción vívida del problema o desperdicio identificado, con datos cuantitativos del antes (tiempos, movimientos, distancias, cantidad de material). Mínimo 200 palabras.

4. solucion: Descripción detallada de la solución implementada paso a paso, materiales utilizados, tiempo de implementación, participantes y desafíos superados. Mínimo 200 palabras.

5. beneficios: Lista de beneficios cuantificados y cualitativos con datos antes/después, ahorros estimados, mejoras en seguridad, calidad, productividad y ambiente de trabajo.

6. tipo_desperdicio: Tipo(s) de desperdicio Lean eliminado(s) de la lista: Motion, Skills, Inventory, Transportation, Over Production, Over Processing, Waiting, Defects.

7. impacto_bto: Categoría BTO impactada: Safe and Sustainable, People & Culture, Network Optimisation, Supply Chain and Manufacturing Excellence.

8. proximos_pasos: Plan de acción concretos para sostener la mejora, replicarla en otras áreas, reconocer al equipo y establecer el nuevo estándar.

Responde SOLO JSON válido sin comentarios adicionales."""

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
        prompt = f"""Traduce del español al inglés profesional industrial, manteniendo la terminología técnica apropiada (OSHA, ISO, ANSI, etc.):
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
        prompt = f"""Corrige ortografía, gramática, puntuación y mejora la redacción del siguiente texto en español. Mantén el significado técnico exacto. Mejora la fluidez y naturalidad sin hacerlo robótico. Devuelve SOLO el texto corregido, sin explicaciones.

TEXTO:
{text}"""
        try:
            corrected = self._call_api(prompt, temperature=0.2, max_tokens=4096).strip()
            return Utils.correct_spelling_basic(corrected)
        except:
            return Utils.correct_spelling_basic(text)

    def _generate_local_moc(self, problem, context, equipo):
        return {
            "descripcion_problema": "Se ha identificado una condición crítica que requiere gestión formal mediante el proceso de Management of Change (MoC). El problema reportado es: " + problem + ".\n\nEsta situación presenta riesgos significativos para la seguridad operacional, la integridad del proceso y la continuidad de la producción. Durante las evaluaciones preliminares se ha determinado que la condición actual no cumple con los estándares corporativos de seguridad y calidad establecidos, generando una exposición potencial al personal operativo y a los equipos críticos.\n\nEs imperativo implementar un cambio controlado y documentado que mitigue los riesgos identificados, garantice el cumplimiento normativo y restablezca las condiciones operativas seguras y eficientes del proceso.",
            "condicion_actual": "El equipo o proceso actual opera bajo condiciones que presentan las siguientes limitaciones técnicas documentadas: " + context + ".\n\nSe han identificado deficiencias en la configuración actual que afectan directamente el rendimiento operativo y la seguridad del personal. Los parámetros críticos del proceso se encuentran fuera de los rangos óptimos establecidos en los procedimientos operativos estándar (SOP).\n\nSe requiere una evaluación técnica exhaustiva para establecer una línea base de referencia completa antes de proceder con cualquier modificación, asegurando que todos los cambios sean trazables y verificables.",
            "condicion_propuesta": "Se propone implementar modificaciones técnicas estructuradas que optimicen el rendimiento operativo del equipo crítico, mejoren significativamente las condiciones de seguridad del proceso y alineen las operaciones con los estándares corporativos y regulatorios vigentes.\n\nLa propuesta incluye la actualización de componentes críticos, la implementación de controles adicionales de seguridad, la estandarización de procedimientos operativos y la capacitación del personal involucrado. Todas las modificaciones serán diseñadas siguiendo las mejores prácticas de la industria y los requisitos normativos aplicables.",
            "razones_cambio": "1. SEGURIDAD OPERACIONAL: La condición actual presenta riesgos identificados que pueden comprometer la integridad física del personal. La implementación del cambio reducirá significativamente la probabilidad de incidentes y accidentes laborales, alineándose con la política de cero accidentes de la organización.\n\n2. OPTIMIZACIÓN DE RENDIMIENTO: El equipo crítico opera por debajo de su capacidad óptima debido a las limitaciones técnicas identificadas. El cambio propuesto mejorará la confiabilidad, disponibilidad y eficiencia del equipo, reduciendo tiempos de parada no planificados.\n\n3. CUMPLIMIENTO NORMATIVO: La modificación asegura el cumplimiento de estándares corporativos, regulaciones nacionales e internacionales aplicables al sector industrial, evitando sanciones y manteniendo la licencia operativa.\n\n4. REDUCCIÓN DE RIESGOS SHES: Las evaluaciones previas han identificado riesgos en seguridad, salud y medio ambiente que serán mitigados proactivamente con las medidas de control propuestas en este documento.\n\n5. MEJORA CONTINUA: El cambio está alineado con los objetivos estratégicos de la organización en materia de excelencia operacional, sostenibilidad y mejora continua.",
            "alternativas_retorno": "ALTERNATIVAS EVALUADAS:\n\n1. MANTENIMIENTO CORRECTIVO TRADICIONAL (DESCARTADO): Aunque de menor costo inicial, presenta un alcance limitado que no aborda las causas raíz del problema. La recurrencia de fallas sería alta, generando costos operacionales mayores a largo plazo.\n\n2. REEMPLAZO TOTAL DEL SISTEMA (DESCARTADO): Ofrece la solución más completa pero con un costo de inversión elevado que excede el presupuesto aprobado para este período. Además, el tiempo de implementación sería excesivo para las necesidades operativas actuales.\n\n3. MODIFICACIÓN CONTROLADA (SELECCIONADA): Representa la mejor relación costo-beneficio, abordando las causas raíz identificadas con un alcance definido, tiempos de implementación razonables y un retorno de inversión favorable dentro del primer año.\n\nPLAN DE RETORNO:\nEn caso de que el cambio no produzca los resultados esperados o se presenten complicaciones durante la implementación, se ejecutará el siguiente plan de retorno: Restauración inmediata de la configuración original del equipo, activación del protocolo de contingencia establecido, notificación oportuna a supervisión directa y áreas de apoyo, documentación detallada de las lecciones aprendidas y análisis de causa raíz de la falla para prevenir recurrencias.",
            "recursos": "RECURSOS HUMANOS REQUERIDOS:\n- Técnico especializado de mantenimiento mecánico/electrónico (1 persona, tiempo completo durante implementación)\n- Supervisor de área operativa (1 persona, supervisión continua)\n- Especialista SHES (1 persona, verificación de controles y permisos)\n- Operador de área certificado (1-2 personas, apoyo operativo y pruebas)\n- Ingeniero de procesos (1 persona, validación técnica y ajustes de parámetros)\n\nRECURSOS MATERIALES:\n- Herramientas especializadas certificadas y calibradas\n- Repuestos de calidad certificada con trazabilidad\n- EPP completo: casco de seguridad, gafas de protección, guantes anticorte, botas dieléctricas, arnés cuando aplique\n- Materiales de señalización, demarcación y etiquetado\n- Materiales de limpieza y preparación de área\n\nRECURSOS TÉCNICOS:\n- Documentación técnica actualizada del equipo (manuales, diagramas, especificaciones)\n- SOP vigentes y procedimientos de trabajo seguro\n- Permisos de trabajo según tipo (trabajo en caliente, espacio confinado, trabajo en altura, etc.)\n- Checklist de verificación pre y post implementación\n- Equipos de prueba y medición calibrados",
            "plan_implementacion": "FASE 1: PREPARACIÓN Y PLANIFICACIÓN (Días 1-2)\n- Reunión de coordinación multidisciplinaria con producción, mantenimiento y SHES\n- Verificación exhaustiva de disponibilidad de todos los recursos materiales y humanos\n- Preparación del área de trabajo: limpieza profunda, señalización de perímetro, aplicación de LOTO (Lock Out Tag Out)\n- Briefing de seguridad con todo el equipo involucrado, revisión de riesgos y controles\n- Verificación final de permisos de trabajo y autorizaciones requeridas\n\nFASE 2: EJECUCIÓN DE MODIFICACIONES (Días 3-5)\n- Implementación progresiva de las modificaciones técnicas según plan detallado\n- Pruebas funcionales iniciales después de cada sub-etapa crítica\n- Registro fotográfico detallado del antes, durante y después de cada modificación\n- Verificación intermedia SHES al finalizar cada día de trabajo\n- Comunicación continua con supervisión de producción sobre avances\n\nFASE 3: VALIDACIÓN Y PRUEBAS (Días 6-7)\n- Pruebas funcionales bajo condiciones normales de operación\n- Verificación de todos los parámetros críticos del proceso contra especificaciones\n- Validación conjunta por supervisor de área, producción y especialista técnico\n- Pruebas de estrés y verificación de límites operativos\n- Documentación de resultados de pruebas y ajustes finales\n\nFASE 4: CIERRE Y ESTANDARIZACIÓN (Día 8)\n- Actualización completa de toda la documentación técnica y operativa\n- Capacitación formal al personal operativo sobre nuevos procedimientos\n- Socialización de lecciones aprendidas con todas las áreas involucradas\n- Cierre formal del MoC con firmas de aprobación de todas las partes\n- Archivo del documento completo en el sistema de gestión documental",
            "tiempo_duracion": "ESTIMACIÓN TOTAL DEL CAMBIO: 8 días hábiles distribuidos en 4 fases bien definidas.\n\nDESGLOSE POR FASE:\n- Fase 1 (Preparación): 2 días\n- Fase 2 (Ejecución): 3 días\n- Fase 3 (Validación): 2 días\n- Fase 4 (Cierre): 1 día\n\nCONSIDERACIONES:\nLa duración puede ajustarse según condiciones operativas, disponibilidad de recursos y resultados de las verificaciones intermedias. Se ha incluido un margen de contingencia del 20% para imprevistos. Las ventanas de mantenimiento serán coordinadas con producción con al menos 48 horas de anticipación.",
            "riesgos_controles": [
                {"riesgo": "Interrupción del proceso productivo durante la implementación de modificaciones, generando pérdidas de producción estimadas", "control": "Coordinación previa detallada con producción para definir ventana de mantenimiento planificado. Comunicación oportuna a todas las áreas involucradas con 24 horas de anticipación. Monitoreo continuo del plan de producción durante la ejecución."},
                {"riesgo": "Falla técnica durante la modificación que pueda afectar equipos adyacentes o sistemas interconectados", "control": "Verificación previa exhaustiva de todas las interconexiones. Disponibilidad inmediata de repuestos de emergencia. Supervisión técnica continua por ingeniero senior. Protocolo de parada de emergencia activo durante toda la ejecución."},
                {"riesgo": "Exposición a riesgos de seguridad durante la ejecución de trabajos en campo (golpes, cortes, caídas)", "control": "Permisos de trabajo específicos según tipo de riesgo identificado. Uso obligatorio y verificado de EPP completo. Supervisión continua por especialista SHES. Aplicación estricta de LOTO en todos los puntos de energía."},
                {"riesgo": "Error humano durante la implementación que genere daño al equipo o configuración incorrecta", "control": "Checklist de verificación paso a paso firmado por técnico y supervisor. Doble verificación crítica (dos personas) en puntos de control clave. Registro fotográfico de cada etapa para trazabilidad."},
                {"riesgo": "Demora en la entrega de repuestos o materiales críticos que retrase la implementación", "control": "Verificación de disponibilidad de materiales 48 horas antes del inicio. Identificación de proveedores alternativos. Stock de seguridad de componentes críticos. Plan de contingencia con materiales sustitutos pre-aprobados."}
            ],
            "riesgos_shes": [
                {"riesgo": "Lesiones por manipulación manual de equipos, componentes pesados o herramientas durante la ejecución", "control": "Capacitación específica en técnicas de levantamiento seguro antes del inicio. Uso obligatorio de EPP completo incluyendo guantes anticorte y calzado de seguridad. Señalización clara del área de trabajo. Supervisor SHES presente durante trabajos de alto riesgo.", "plazo": "Antes del inicio de actividades"},
                {"riesgo": "Generación de residuos sólidos, líquidos o peligrosos durante el proceso de modificación", "control": "Manejo seguro según procedimiento ambiental corporativo. Clasificación en origen de todos los residuos generados. Disposición únicamente en áreas autorizadas y con registro de trazabilidad. Contenedores identificados y segregados.", "plazo": "Durante toda la ejecución"},
                {"riesgo": "Exposición a ruido excesivo, vibraciones o agentes químicos durante trabajos de modificación", "control": "Monitoreo continuo de niveles de ruido y agentes químicos. Uso obligatorio de protectores auditivos cuando los niveles excedan 85 dB. Ventilación adecuada en áreas cerradas. Limitación de horario de exposición según límites permisibles legales.", "plazo": "Durante toda la ejecución"},
                {"riesgo": "Incendio o explosión por trabajo en caliente, chispas o acumulación de vapores inflamables", "control": "Permiso de trabajo en caliente con análisis de atmósfera previo. Vigía de fuego designado y capacitado. Extintores portátiles disponibles y verificados. Limpieza del área de materiales combustibles antes del inicio. Monitoreo continuo de gases inflamables.", "plazo": "Durante trabajos de soldadura/corte"},
                {"riesgo": "Contaminación del suelo o cuerpos de agua por derrames accidentales de lubricantes, solventes o productos químicos", "control": "Kit de contención de derrames disponible en el área. Uso de bandejas de retención bajo equipos que manipulen líquidos. Prohibición de drenaje directo a sistemas de alcantarillado. Limpieza inmediata de cualquier derrame con materiales absorbentes aprobados.", "plazo": "Durante toda la ejecución"}
            ]
        }

    def _generate_local_a3(self, problem, context):
        return {
            "titulo": "Optimización del proceso: " + problem[:50],
            "antecedentes": "Durante los últimos seis meses, el área operativa ha experimentado una degradación progresiva en sus indicadores clave de desempeño. Se han registrado incrementos en tiempos de ciclo, aumento en la tasa de defectos y una reducción en la productividad general del proceso.\n\nEl análisis preliminar de datos históricos revela una tendencia creciente que, si no se aborda de manera estructurada, comprometerá los objetivos anuales de la organización. La metodología A3 fue seleccionada como herramienta principal para el análisis estructurado de esta situación, permitiendo una visión integral del problema y facilitando la identificación de soluciones sostenibles.",
            "problema_actual": problem,
            "analisis_situacion": "La situación actual presenta múltiples indicadores de desempeño con oportunidades significativas de mejora. Se requiere una recopilación sistemática y rigurosa de datos para establecer una línea base sólida que permita cuantificar el impacto de las contramedidas propuestas.\n\nEl análisis de variabilidad del proceso muestra fluctuaciones que exceden los límites de control establecidos, indicando la presencia de causas especiales que deben ser identificadas y eliminadas. La comparativa con benchmarks de la industria revela una brecha de desempeño del 15-25% respecto a los mejores en clase.",
            "objetivos": "OBJETIVO GENERAL:\nOptimizar integralmente el proceso eliminando los desperdicios identificados y estableciendo un nuevo estándar de desempeño sostenible.\n\nOBJETIVOS ESPECÍFICOS (SMART):\n1. Reducir el tiempo de ciclo en un 15% dentro de los próximos 3 meses, pasando de 45 minutos a 38 minutos por unidad.\n2. Disminuir la tasa de defectos en un 20% durante el próximo trimestre, reduciendo de 8% a 6.4%.\n3. Mejorar la productividad general del área en un 10% dentro de 6 meses.\n4. Incrementar la satisfacción interna del cliente (área siguiente) en un 25% según encuesta trimestral.\n5. Reducir el costo operativo unitario en un 8% dentro del primer año de implementación.",
            "analisis_causa_raiz": "ANÁLISIS DE LOS 5 PORQUÉS:\n1. ¿POR QUÉ ocurre el problema? → Porque el proceso opera con una configuración inadecuada que genera variabilidad excesiva.\n2. ¿POR QUÉ la configuración es inadecuada? → Porque no existe una estandarización formal de los parámetros operativos críticos.\n3. ¿POR QUÉ no hay estandarización? → Porque los procedimientos operativos estándar (SOP) no han sido actualizados en los últimos 18 meses.\n4. ¿POR QUÉ no están actualizados? → Porque no existe un sistema de gestión documental efectivo que asegure la revisión periódica.\n5. ¿POR QUÉ no hay sistema? → Porque falta una política clara de gestión del conocimiento y mejora continua con responsables asignados.\n\nCAUSA RAÍZ IDENTIFICADA:\nAusencia de un sistema integral de gestión, actualización y control de SOP, combinado con la falta de responsables claros y métricas de seguimiento del desempeño del proceso.",
            "contramedidas": "1. ACTUALIZAR SOP DEL PROCESO: Revisar y actualizar todos los procedimientos operativos con instrucciones claras, paso a paso, con fotos de referencia y puntos de control crítico. Responsable: Ingeniero de Procesos. Plazo: 2 semanas.\n\n2. IMPLEMENTAR CHECKLISTS DIARIOS: Diseñar y desplegar checklists de verificación diaria en cada puesto de trabajo para asegurar el cumplimiento de estándares. Responsable: Supervisor de Área. Plazo: 1 semana.\n\n3. CAPACITAR AL PERSONAL: Programar y ejecutar capacitaciones formales sobre los nuevos estándares, con evaluación de competencias y certificación. Responsable: Especialista de Capacitación. Plazo: 3 semanas.\n\n4. ESTABLECER KPIs VISUALES: Implementar tableros visuales en el área con indicadores clave de desempeño actualizados diariamente. Responsable: Líder de Mejora Continua. Plazo: 2 semanas.\n\n5. PROGRAMAR AUDITORÍAS MENSUALES: Establecer auditorías formales mensuales de cumplimiento con criterios de evaluación definidos y plan de acción para desviaciones. Responsable: Auditor Interno. Plazo: Inicio inmediato, recurrente.\n\n6. IMPLEMENTAR SISTEMA DE GESTIÓN DOCUMENTAL: Desarrollar o adquirir una solución digital para control de versiones, aprobaciones y distribución de documentos. Responsable: IT + Calidad. Plazo: 2 meses.\n\n7. DEFINIR RESPONSABLES DE ESTÁNDARES: Asignar responsables claros por área para la gestión, actualización y seguimiento de estándares operativos. Responsable: Gerente de Operaciones. Plazo: 1 semana.",
            "resultados_esperados": "- Reducción medible y sostenida de desperdicios identificados (Motion, Waiting, Defects)\n- Mejora sostenida en calidad del producto y consistencia del proceso\n- Estandarización efectiva que reduzca la variabilidad operativa en al menos 30%\n- Reducción del tiempo de ciclo en 15% con impacto directo en capacidad productiva\n- Incremento en satisfacción del cliente interno medido mediante encuestas\n- Fortalecimiento de la cultura de mejora continua y trabajo en equipo\n- Retorno de inversión estimado del 180% dentro del primer año\n- Reducción de costos operativos unitarios en 8%\n- Mejora en el ambiente de trabajo y motivación del personal",
            "plan_seguimiento": "SEMANA 1-2: Implementación de contramedidas 1 y 2 (SOP y checklists). Monitoreo diario de cumplimiento.\n\nSEMANA 3-4: Ejecución de capacitaciones (contramedida 3). Evaluación de competencias. Monitoreo inicial de indicadores. Ajustes según resultados.\n\nMES 2: Primera auditoría formal (contramedida 5). Evaluación de avance vs. objetivos iniciales. Implementación de KPIs visuales.\n\nMES 3: Evaluación integral vs. objetivos SMART establecidos. Análisis de tendencias. Ajustes a contramedidas si es necesario.\n\nMES 6: Revisión de sostenibilidad de mejoras. Análisis de replicabilidad en otras áreas. Reconocimiento al equipo.\n\nTRIMESTRAL: Revisiones formales con gerencia. Actualización de objetivos según evolución del proceso.",
            "lecciones_aprendidas": "La aplicación de la metodología A3 permitió visualizar de manera integral la complejidad del problema y las interconexiones entre sus múltiples causas. La participación activa y multidisciplinaria del equipo fue fundamental para identificar la causa raíz real, que inicialmente no era evidente.\n\nSe aprendió que los problemas aparentemente técnicos frecuentemente tienen raíces en sistemas de gestión deficientes. La inversión en capacitación y estandarización genera retornos significativos a mediano plazo. La visualización de datos y el seguimiento constante son críticos para mantener las mejoras.",
            "estandarizacion": "Los procedimientos actualizados serán documentados formalmente con control de versiones, aprobados por gerencia de operaciones y calidad, socializados mediante capacitaciones estructuradas con evaluación de competencias, integrados al Sistema de Gestión de Calidad (SGC) existente y sujetos a revisión periódica anual como mínimo. Se establecerán métricas de cumplimiento y auditorías programadas para asegurar la sostenibilidad de los estándares implementados."
        }

    def _generate_local_kaizen(self, activity, context):
        return {
            "titulo": "Kaizen: " + activity[:50],
            "area": "Área de Mantenimiento / Producción / Calidad (especificar según contexto)",
            "descripcion_problema": activity + "\n\nDurante las actividades diarias de gemba walk, el equipo identificó esta oportunidad de mejora que representa un desperdicio significativo en el proceso. La situación actual genera movimientos innecesarios, tiempos de espera o riesgos de calidad que impactan directamente en la productividad del área y en la satisfacción del personal.\n\nSe realizó un análisis rápido de la situación que confirmó la viabilidad de implementar una mejora inmediata con recursos disponibles en el área, siguiendo el principio fundamental del Kaizen: mejorar un poco cada día.",
            "solucion": "Se implementó una mejora estructurada orientada a eliminar el desperdicio identificado y optimizar el flujo del proceso, aplicando principios fundamentales de Lean Manufacturing y el pensamiento Kaizen de mejora continua.\n\nLa solución fue diseñada y ejecutada por el equipo de trabajo del área con apoyo del líder de mejora continua, utilizando materiales disponibles y aplicando el concepto de low cost, high impact. Se realizaron pruebas piloto antes de la implementación definitiva para validar la efectividad de la propuesta.\n\nEl equipo documentó el antes y después con fotografías y mediciones de tiempo para cuantificar el impacto de la mejora implementada.",
            "beneficios": "- Reducción del tiempo de ejecución en aproximadamente 20-30%\n- Mejora significativa en calidad y consistencia del proceso\n- Mayor seguridad para el personal al eliminar movimientos riesgosos\n- Reducción de costos operativos derivados de la eliminación de desperdicios\n- Mejora en el ambiente de trabajo y orden del área\n- Eliminación de movimientos innecesarios y tiempos de búsqueda\n- Incremento en la motivación del equipo al ver resultados inmediatos\n- Fácil replicabilidad en otras áreas similares",
            "tipo_desperdicio": "Motion / Waiting / Skills (seleccionar según análisis específico del desperdicio identificado)",
            "impacto_bto": "Supply Chain and Manufacturing Excellence / Safe and Sustainable (seleccionar según el impacto principal de la mejora)",
            "proximos_pasos": "1. Documentar formalmente la mejora con fotografías, descripción detallada y datos de impacto\n2. Socializar la mejora con otras áreas relacionadas mediante presentación breve en reunión de coordinación\n3. Replicar la mejora en procesos similares identificados durante el gemba walk\n4. Establecer monitoreo mensual para asegurar que la mejora se mantiene en el tiempo\n5. Reconocer formalmente al equipo participante en la mejora\n6. Integrar el nuevo estándar al SOP del área\n7. Programar revisión de sostenibilidad a los 3 meses"
        }

# =============================================================================
# REEMPLAZO INTELIGENTE DE TEXTO EN POWERPOINT (preserva formato)
# =============================================================================
def replace_text_in_shape(shape, old_text, new_text):
    """Reemplaza texto en un shape preservando el formato de los runs"""
    if not shape.has_text_frame:
        return False

    text_frame = shape.text_frame
    text = text_frame.text

    if old_text not in text:
        return False

    # Reemplazar en cada párrafo, preservando runs
    for paragraph in text_frame.paragraphs:
        paragraph_text = paragraph.text
        if old_text in paragraph_text:
            # Encontrar el run que contiene el texto
            for run in paragraph.runs:
                if old_text in run.text:
                    run.text = run.text.replace(old_text, new_text)
                    return True
            # Si el texto está dividido entre runs, reemplazar en el primer run
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
            # Reemplazar en text frames
            if shape.has_text_frame:
                for old_text, new_text in replacements.items():
                    replace_text_in_shape(shape, old_text, new_text)

            # Reemplazar en tablas
            if shape.has_table:
                table = shape.table
                for row in table.rows:
                    for cell in row.cells:
                        for old_text, new_text in replacements.items():
                            if old_text in cell.text:
                                # Preservar formato de la celda
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


# =============================================================================
# GENERADOR DE DOCUMENTOS
# =============================================================================
class DocumentGenerator:
    """Genera documentos usando templates cargados en memoria"""

    def generate_moc(self, data, images=None, template_bytes=None):
        """Genera MoC desde template en memoria"""
        if template_bytes is None:
            st.error("❌ Template MoC no cargado. Vaya a Configuración > Templates.")
            return None

        prs = Presentation(BytesIO(template_bytes))

        # Reemplazos globales en toda la presentación
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

        # SLIDE 3: Descripción del Problema
        slide3 = prs.slides[2]
        for shape in slide3.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if run.text.strip() == "3":
                            run.text = data.get('descripcion_problema', '')

        # SLIDE 4: Tabla Condición Actual / Condición Propuesta
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

        # Agregar imágenes de soporte
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
                                run.text = f"Leader: {data.get('autor', '')}\nÁrea: {data.get('area', '')}\nFecha: {data.get('fecha', '')}"

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
            ("Descripción del Problema", "descripcion_problema"),
            ("Solución Implementada", "solucion"),
            ("Beneficios", "beneficios"),
            ("Próximos Pasos", "proximos_pasos"),
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

# =============================================================================
# EXPORTADOR A PDF
# =============================================================================
class PDFExporter:
    """Exporta documentos a PDF usando múltiples métodos"""

    @staticmethod
    def pptx_to_pdf_libreoffice(pptx_bytes, output_filename):
        """Convierte PPTX a PDF usando LibreOffice (método preferido)"""
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
        """Convierte DOCX a PDF usando LibreOffice"""
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
        """Genera PDF desde datos usando ReportLab (fallback)"""
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

                story.append(Paragraph(f"<b>Tipo de Desperdicio:</b> {data.get('tipo_desperdicio', '')}", body_style))
                story.append(Paragraph(f"<b>Impacto BTO:</b> {data.get('impacto_bto', '')}", body_style))

            doc.build(story)
            buffer.seek(0)
            return buffer.getvalue()

        except Exception as e:
            st.error(f"Error generando PDF con ReportLab: {e}")
            return None

# =============================================================================
# INICIALIZACION DE SESSION STATE CON PERSISTENCIA LOCAL
# =============================================================================
def init_session_state():
    """Inicializa todas las variables de session_state con persistencia local"""

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
# INTERFAZ DE USUARIO - SIDEBAR, HEADER, WELCOME
# =============================================================================

def render_sidebar():
    """Renderiza sidebar de navegacion"""
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
        <p>v4.0.0 · Julio 2026</p>
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
        st.warning("⚠️ **Templates no cargados.** Vaya a Configuración > Templates para subir los formatos oficiales de la empresa antes de generar documentos.")

    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown("""
        <div class="doc-card doc-card-moc">
            <h3 style="color: #1a5f7a; margin-top: 0;">📋 Management of Change</h3>
            <p style="color: #64748b; font-size: 0.9rem;">Documento base para cualquier cambio en proceso o maquinaria. Incluye evaluación de riesgos, plan de implementación y controles SHES.</p>
            <ul style="color: #475569; font-size: 0.85rem; padding-left: 1.2rem;">
                <li>Evaluación técnica completa</li><li>Riesgos y controles detallados</li><li>Plan de implementación por fases</li><li>Exportación a PDF</li>
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
            <p style="color: #64748b; font-size: 0.9rem;">Formato estructurado para registrar mejoras. Metodología Lean con análisis de causa raíz 5 Porqués y contramedidas SMART.</p>
            <ul style="color: #475569; font-size: 0.85rem; padding-left: 1.2rem;">
                <li>Análisis 5 Porqués detallado</li><li>Contramedidas priorizadas</li><li>Plan de seguimiento</li><li>Estandarización completa</li>
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
            <p style="color: #64748b; font-size: 0.9rem;">Registro rápido de mejoras del día a día. Captura ideas de mejora continua con clasificación de desperdicios Lean.</p>
            <ul style="color: #475569; font-size: 0.85rem; padding-left: 1.2rem;">
                <li>8 Desperdicios (Wastes)</li><li>Impacto BTO cuantificado</li><li>Beneficios medibles</li><li>Replicabilidad inmediata</li>
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

# =============================================================================
# FORMULARIOS DE CREACIÓN CON CORRECTOR AUTOMÁTICO
# =============================================================================

def auto_correct_text_input(label, value, key, height=100, help_text=""):
    """Campo de texto con corrección ortográfica automática"""
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
    st.info("💡 Complete la información y describa el problema con el mayor detalle posible. La IA generará automáticamente todos los campos técnicos con redacción profesional y humanizada.")

    if not st.session_state.get("template_moc_bytes"):
        st.error("❌ **Template MoC no cargado.** Vaya a Configuración > Templates y suba el archivo .pptx oficial.")
        if st.button("Ir a Configuración", key="go_config_moc"):
            st.session_state.page = "configuracion"
            st.rerun()
        return

    st.markdown("#### 1. Información General")
    col1, col2 = st.columns(2)
    with col1:
        moc_title = st.text_input("Título de la MoC:", placeholder="Ej: Control de Acceso a panel HMI - Carga de detonadores")
        moc_number = st.text_input("Número:", value=Utils.generate_doc_number("moc"), disabled=True)
    with col2:
        naturaleza = st.selectbox("Naturaleza:", ["permanente", "temporal", "emergencia"])
        originador = st.text_input("Originador:", value=config.get("default_author", ""))
    fecha = st.text_input("Fecha:", value=Utils.format_date(), disabled=True)

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

    st.markdown("#### 3. Descripción del Problema (Sea lo más detallado posible)")
    problem_desc = auto_correct_text_input(
        "Describa el problema con sus palabras:", 
        "", 
        "moc_problem_desc",
        height=250,
        help_text="Cuanto más detalle proporcione, mejor será la generación automática. Incluya: qué está pasando, desde cuándo, impacto, equipos involucrados, riesgos observados."
    )

    st.markdown("#### 4. Contexto Adicional (Opcional pero recomendado)")
    context = auto_correct_text_input(
        "Información adicional:", 
        "", 
        "moc_context",
        height=120,
        help_text="TAG del equipo, área específica, normativas aplicables, fechas relevantes, datos numéricos, etc."
    )

    st.markdown("#### 5. Imágenes de Soporte (Opcional)")
    uploaded_images = st.file_uploader("Seleccione imágenes:", type=["png", "jpg", "jpeg"], accept_multiple_files=True)

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

        with st.spinner("🧠 La IA está generando el documento completo con redacción profesional..."):
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
    st.info("💡 Describa el problema con detalle y la IA generará el documento A3 completo con análisis de causa raíz y contramedidas.")

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
    st.info("💡 Describa la actividad de mejora realizada con detalle. La IA generará un registro completo con beneficios cuantificados.")

    if not st.session_state.get("template_kaizen_bytes"):
        st.error("❌ **Template Kaizen no cargado.** Vaya a Configuración > Templates.")
        if st.button("Ir a Configuración", key="go_config_kzn"):
            st.session_state.page = "configuracion"
            st.rerun()
        return

    st.markdown("#### 1. Información General")
    col1, col2 = st.columns(2)
    with col1:
        kaizen_title = st.text_input("Título:", placeholder="Ej: Organización de área de herramientas con Shadow Board")
        area = st.text_input("Área:", value=config.get("default_area", ""))
    with col2:
        autor = st.text_input("Autor:", value=config.get("default_author", ""))
        doc_number = st.text_input("Número:", value=Utils.generate_doc_number("kaizen"), disabled=True)
    fecha = st.text_input("Fecha:", value=Utils.format_date(), disabled=True)

    st.markdown("#### 2. Descripción de la Actividad (Detallada)")
    activity_desc = auto_correct_text_input(
        "Describa la mejora realizada:", 
        "", 
        "kzn_activity_desc",
        height=250,
        help_text="Describa el antes, durante y después. Incluya datos de tiempo, movimientos, cantidades. ¿Qué hizo? ¿Cómo lo hizo? ¿Quién participó?"
    )

    st.markdown("#### 3. Clasificación")
    tipo_desp = st.multiselect("Tipo de Desperdicio eliminado:",
                               ["Motion", "Skills", "Inventory", "Transportation",
                                "Over Production", "Over Processing", "Waiting", "Defects"])
    impacto_bto = st.selectbox("Impacto BTO:",
                               ["Safe and Sustainable", "People & Culture",
                                "Network Optimisation", "Supply Chain and Manufacturing Excellence"])

    context = auto_correct_text_input(
        "Contexto adicional:", 
        "", 
        "kzn_context",
        height=80,
        help_text="Costos, tiempos medidos, materiales utilizados, etc."
    )

    st.markdown("#### 4. Imágenes de Soporte")
    uploaded_images = st.file_uploader("Seleccione imágenes:", type=["png", "jpg", "jpeg"], accept_multiple_files=True)

    image_paths = []
    if uploaded_images:
        for idx, img_file in enumerate(uploaded_images, 1):
            img_path = f"/tmp/temp_kzn_img_{doc_number}_{idx}.png"
            with open(img_path, "wb") as f:
                f.write(img_file.getbuffer())
            image_paths.append({"path": img_path, "desc": f"Figura {idx} - {img_file.name}"})
        st.success(f"📷 {len(image_paths)} imagen(es) cargada(s)")

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🤖 Generar Documento Kaizen con IA", type="primary", use_container_width=True):
        if not activity_desc.strip():
            st.error("❌ Describa la actividad antes de generar.")
            return

        with st.spinner("🧠 Generando documento Kaizen con beneficios cuantificados..."):
            gemini = GeminiService(config.get("gemini_api_key", ""), config.get("gemini_model", "gemini-1.5-pro"))
            result = gemini.generate_kaizen(activity_desc, context)
            result["tipo_desperdicio"] = ", ".join(tipo_desp) if tipo_desp else result.get("tipo_desperdicio", "")
            result["impacto_bto"] = impacto_bto

            st.session_state.generated_data = result
            st.session_state.doc_meta = {"titulo": kaizen_title, "area": area, "autor": autor, "doc_number": doc_number, "fecha": fecha}
            st.session_state.doc_images = image_paths
            st.session_state.doc_type = "kaizen"
            st.session_state.page = "revisar"
            st.rerun()


# =============================================================================
# PANTALLA DE REVISIÓN UNIFICADA
# =============================================================================

def render_review():
    """Renderiza pantalla de revision unificada para cualquier tipo de documento"""
    doc_type = st.session_state.doc_type
    data = st.session_state.get("generated_data", {})
    meta = st.session_state.get("doc_meta", {})
    images = st.session_state.get("doc_images", [])
    config = st.session_state.config

    type_names = {"moc": "MoC", "a3": "Mejora A3", "kaizen": "Simple Kaizen"}
    type_name = type_names.get(doc_type, "Documento")

    st.markdown(f'<div class="section-header"><h3>👁️ Revisar y Editar {type_name}</h3></div>', unsafe_allow_html=True)
    st.info("💡 Revise cada campo, realice correcciones manuales si es necesario y genere el documento final. Todos los cambios se guardan automáticamente.")

    if doc_type == "moc":
        _render_moc_review(data, meta, images, config)
    elif doc_type == "a3":
        _render_a3_review(data, meta, images, config)
    elif doc_type == "kaizen":
        _render_kaizen_review(data, meta, images, config)


def _spell_check_field(label, value, key_prefix, gemini):
    """Campo de texto con boton de correccion ortografica"""
    col1, col2 = st.columns([6, 1])
    with col1:
        text = st.text_area(label, value=value, height=120, key=f"{key_prefix}_field")
    with col2:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("✨ Corregir", key=f"{key_prefix}_spell", help="Corregir ortografía y mejorar redacción con IA"):
            with st.spinner("Corrigiendo y puliendo redacción..."):
                corrected = gemini.correct_spelling(text)
                st.session_state[f"{key_prefix}_corrected"] = corrected
                st.rerun()

    if st.session_state.get(f"{key_prefix}_corrected"):
        text = st.session_state[f"{key_prefix}_corrected"]
        del st.session_state[f"{key_prefix}_corrected"]

    return text


def _render_moc_review(data, meta, images, config):
    """Revision especifica MoC"""
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
            st.rerun()
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
            st.rerun()
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
    """Revision especifica A3"""
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
    """Revision especifica Kaizen"""
    gemini = GeminiService(config.get("gemini_api_key", ""), config.get("gemini_model", "gemini-1.5-pro"))

    tabs = st.tabs(["📋 General", "📝 Contenido", "📷 Imágenes", "⚙️ Generar"])

    with tabs[0]:
        meta["titulo"] = st.text_input("Título:", value=meta.get("titulo", ""), key="kzn_rev_title")
        meta["area"] = st.text_input("Área:", value=meta.get("area", ""), key="kzn_rev_area")
        meta["autor"] = st.text_input("Autor:", value=meta.get("autor", ""), key="kzn_rev_autor")
        meta["doc_number"] = st.text_input("Número:", value=meta.get("doc_number", ""), disabled=True)
        meta["fecha"] = st.text_input("Fecha:", value=meta.get("fecha", ""), disabled=True)

    with tabs[1]:
        sections = [
            ("Descripción del Problema", "descripcion_problema"), ("Solución Implementada", "solucion"),
            ("Beneficios", "beneficios"), ("Tipo de Desperdicio", "tipo_desperdicio"),
            ("Impacto BTO", "impacto_bto"), ("Próximos Pasos", "proximos_pasos"),
        ]
        for label, key in sections:
            st.markdown(f"**{label}**")
            data[key] = _spell_check_field("", data.get(key, ""), f"kzn_{key}", gemini)

    with tabs[2]:
        if images:
            for idx, img_info in enumerate(images, 1):
                st.image(img_info["path"], caption=f"Figura {idx}: {img_info['desc']}", width=400)
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
# FINALIZACIÓN DE DOCUMENTO - GENERACIÓN Y EXPORTACIÓN
# =============================================================================

def _finalize_document(data, meta, images, language, doc_type, output_format="pptx"):
    """Finaliza y genera documento en el formato solicitado"""
    config = st.session_state.config
    gemini = GeminiService(config.get("gemini_api_key", ""), config.get("gemini_model", "gemini-1.5-pro"))

    with st.spinner(f"📄 Generando documento {'en Inglés' if language == 'en' else 'en Español'}..."):
        final_data = {**meta, **data}

        if language == "en" and doc_type == "moc":
            st.info("🌐 Traduciendo documento al inglés profesional...")
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
                            st.warning("No se pudo generar PDF. Descargando PPTX en su lugar.")
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
                            st.warning("No se pudo generar PDF. Descargando DOCX en su lugar.")
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
                            st.warning("No se pudo generar PDF. Descargando PPTX en su lugar.")
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

        if ext != "pdf":
            st.info("💡 Para obtener PDF: use el botón PDF directamente arriba, o abra el archivo en Office/LibreOffice y exporte a PDF.")


# =============================================================================
# HISTORIAL CON PERSISTENCIA LOCAL
# =============================================================================

def render_history():
    """Renderiza historial de documentos con exportacion/importacion"""
    st.markdown('<div class="section-header"><h3>📁 Historial de Documentos Generados</h3></div>', unsafe_allow_html=True)

    docs = st.session_state.history.get("documents", [])

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("#### 💾 Exportar Configuración")
        export_data = {
            "config": st.session_state.config,
            "history": st.session_state.history,
            "export_date": datetime.now().isoformat(),
            "version": "4.0.0"
        }
        export_json = json.dumps(export_data, indent=2, ensure_ascii=False)
        st.download_button(
            label="📥 Descargar backup completo (JSON)",
            data=export_json,
            file_name=f"gestion_documental_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
            mime="application/json",
            use_container_width=True
        )
        st.caption("Guarde este archivo para restaurar su configuración e historial en cualquier momento.")

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
        st.info("📭 No hay documentos generados aún. Cree su primer documento desde el menú principal.")
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
# CONFIGURACIÓN COMPLETA
# =============================================================================

def render_settings():
    """Renderiza configuracion completa"""
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
        st.warning("⚠️ **Importante:** Los templates son los formatos oficiales de su empresa. Suba los archivos .pptx (MoC, Kaizen) y .docx (A3).")

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
            st.success("🎉 ¡Todos los templates están cargados! Puede comenzar a generar documentos.")

    with tabs[3]:
        st.markdown("#### Configuración Avanzada")

        st.markdown("##### Corrección Ortográfica")
        auto_correct = st.toggle("Corrección automática en campos de entrada", value=config.get("auto_correct", True))
        spell_check = st.toggle("Corrector ortográfico con IA en revisión", value=config.get("spell_check", True))

        st.markdown("##### Nivel de Pensamiento IA")
        thinking = st.select_slider("Profundidad de generación:", ["Básico", "Estándar", "Profundo"],
                                    value=config.get("thinking_level", "Estándar"))

        st.markdown("##### Numeración de Documentos")
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
        st.info("💡 Exporte su configuración e historial para no perderlos. Los datos se guardan automáticamente, pero este backup le permite migrar a otro equipo.")

        export_data = {
            "config": st.session_state.config,
            "history": st.session_state.history,
            "export_date": datetime.now().isoformat(),
            "version": "4.0.0"
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
                st.success("✅ Restauración completada. Recargando...")
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
                st.success("✅ Configuración restaurada a valores por defecto")
                st.rerun()


# =============================================================================
# FUNCIÓN PRINCIPAL
# =============================================================================

def main():
    """Punto de entrada principal"""
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

    # Footer
    st.markdown("""
    <div class="app-footer">
        <p><strong style="font-size: 1.1rem;">CAVA</strong> - Especialistas en Robótica y Automatización</p>
        <p>Diseñado por <strong>Roger Huamani</strong> | Sistema de Gestión Documental v4.0.0</p>
        <p style="font-size: 0.75rem; color: #94a3b8;">
            Software empresarial para automatización de documentos MoC, A3 y Kaizen.<br>
            Mantiene los formatos oficiales de la empresa sin modificaciones.<br>
            Datos persistentes locales. Exportación a PDF integrada.
        </p>
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
