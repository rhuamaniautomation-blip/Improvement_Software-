#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
================================================================================
SISTEMA DE GESTION DOCUMENTAL - MoC | Mejora A3 | Simple Kaizen
Version Streamlit Web - Acceso online para todo el equipo
================================================================================
Diseñado por: CAVA - Especialistas en Robotica y Automatizacion
Desarrollador: Roger Huamani
Version: 2.0.0
Fecha: Mayo 2026
================================================================================
"""

import streamlit as st
import os
import json
import re
import uuid
from datetime import datetime
from pathlib import Path
from PIL import Image
import io
import base64

# Configuracion de pagina Streamlit
st.set_page_config(
    page_title="Gestion Documental - MoC | A3 | Kaizen",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =============================================================================
# CONFIGURACION DE DIRECTORIOS
# =============================================================================
BASE_DIR = Path(__file__).parent
TEMPLATES_DIR = BASE_DIR / "templates"
OUTPUT_DIR = BASE_DIR / "output"
DATA_DIR = BASE_DIR / "data"
CONFIG_FILE = DATA_DIR / "config.json"
HISTORY_FILE = DATA_DIR / "history.json"

for d in [OUTPUT_DIR, DATA_DIR]:
    d.mkdir(parents=True, exist_ok=True)

# =============================================================================
# CSS PERSONALIZADO - TEMA EMPRESARIAL PROFESIONAL
# =============================================================================
CUSTOM_CSS = """
<style>
    /* Fuentes y colores base */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', 'Segoe UI', sans-serif !important;
    }

    /* Header principal */
    .main-header {
        background: linear-gradient(135deg, #1a5f7a 0%, #2e8bc0 100%);
        padding: 2rem;
        border-radius: 12px;
        color: white;
        margin-bottom: 2rem;
        box-shadow: 0 4px 20px rgba(26, 95, 122, 0.3);
    }

    .main-header h1 {
        color: white !important;
        font-weight: 700 !important;
        margin-bottom: 0.5rem !important;
    }

    .main-header p {
        color: rgba(255,255,255,0.9) !important;
        font-size: 1.1rem !important;
    }

    /* Tarjetas de seleccion */
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

    /* Indicadores de color */
    .indicator-moc { background: #1a5f7a; }
    .indicator-a3 { background: #10b981; }
    .indicator-kaizen { background: #f59e0b; }

    /* Botones principales */
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

    /* Campos de texto */
    .stTextInput > div > div > input,
    .stTextArea > div > div > textarea,
    .stSelectbox > div > div > div {
        border-radius: 8px !important;
        border: 1px solid #cbd5e1 !important;
        font-size: 15px !important;
    }

    .stTextInput > div > div > input:focus,
    .stTextArea > div > div > textarea:focus {
        border-color: #1a5f7a !important;
        box-shadow: 0 0 0 3px rgba(26, 95, 122, 0.1) !important;
    }

    /* Secciones */
    .section-header {
        background: #f1f5f9;
        padding: 1rem 1.5rem;
        border-radius: 10px;
        margin: 1.5rem 0 1rem 0;
        border-left: 4px solid #1a5f7a;
    }

    .section-header h3 {
        margin: 0 !important;
        color: #1e293b !important;
        font-weight: 600 !important;
    }

    /* Tablas de riesgos */
    .risk-card {
        background: white;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 1rem;
        margin: 0.5rem 0;
    }

    /* Footer */
    .app-footer {
        text-align: center;
        padding: 2rem;
        margin-top: 3rem;
        border-top: 1px solid #e2e8f0;
        color: #64748b;
    }

    /* Corrector ortografico hint */
    .spell-hint {
        background: #fef3c7;
        border: 1px solid #f59e0b;
        border-radius: 8px;
        padding: 0.75rem 1rem;
        margin: 0.5rem 0;
        font-size: 13px;
        color: #92400e;
    }

    /* Badge de version gemini */
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

    /* Historial */
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

    /* Animacion de carga */
    @keyframes pulse {
        0%, 100% { opacity: 1; }
        50% { opacity: 0.5; }
    }

    .generating {
        animation: pulse 2s cubic-bezier(0.4, 0, 0.6, 1) infinite;
    }

    /* Ocultar elementos de Streamlit por defecto */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}

    /* Sidebar */
    [data-testid="stSidebar"] {
        background: #1e293b !important;
    }

    [data-testid="stSidebar"] .stMarkdown {
        color: #94a3b8 !important;
    }

    [data-testid="stSidebar"] h1, 
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3 {
        color: #f8fafc !important;
    }
</style>
"""

st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# =============================================================================
# UTILIDADES Y CONFIGURACION
# =============================================================================
class ConfigManager:
    """Gestiona la configuracion persistente del sistema"""

    @staticmethod
    def load():
        if CONFIG_FILE.exists():
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        return {
            "gemini_api_key": "",
            "gemini_model": "gemini-1.5-pro",
            "company_name": "Empresa",
            "department": "Mantenimiento",
            "default_author": "",
            "default_area": "",
            "header_text": "",
            "footer_text": "",
            "last_moc_number": 0,
            "last_a3_number": 0,
            "last_kaizen_number": 0
        }

    @staticmethod
    def save(config):
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=2, ensure_ascii=False)

    @staticmethod
    def get_next_number(doc_type):
        config = ConfigManager.load()
        key = f"last_{doc_type}_number"
        config[key] = config.get(key, 0) + 1
        ConfigManager.save(config)
        return config[key]


class HistoryManager:
    """Gestiona el historial de documentos generados"""

    @staticmethod
    def load():
        if HISTORY_FILE.exists():
            with open(HISTORY_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        return {"documents": []}

    @staticmethod
    def save(history):
        with open(HISTORY_FILE, 'w', encoding='utf-8') as f:
            json.dump(history, f, indent=2, ensure_ascii=False)

    @staticmethod
    def add_document(doc_info):
        history = HistoryManager.load()
        doc_info["id"] = str(uuid.uuid4())
        doc_info["timestamp"] = datetime.now().isoformat()
        history["documents"].insert(0, doc_info)
        HistoryManager.save(history)
        return doc_info["id"]

    @staticmethod
    def get_documents(doc_type=None, limit=50):
        history = HistoryManager.load()
        docs = history["documents"]
        if doc_type:
            docs = [d for d in docs if d.get("type") == doc_type]
        return docs[:limit]


class Utils:
    @staticmethod
    def format_date():
        meses = {1:"ENERO",2:"FEBRERO",3:"MARZO",4:"ABRIL",5:"MAYO",6:"JUNIO",
                 7:"JULIO",8:"AGOSTO",9:"SEPTIEMBRE",10:"OCTUBRE",11:"NOVIEMBRE",12:"DICIEMBRE"}
        now = datetime.now()
        return f"{meses[now.month]} {now.year}"

    @staticmethod
    def generate_doc_number(doc_type):
        prefix = {"moc": "MOC", "a3": "A3", "kaizen": "KZN"}
        num = ConfigManager.get_next_number(doc_type)
        now = datetime.now()
        return f"{prefix.get(doc_type, 'DOC')}-{now.year}{now.month:02d}{now.day:02d}-{num:04d}"

    @staticmethod
    def sanitize_filename(filename):
        return re.sub(r'[<>:"/\\|?*]', '_', filename)[:50]


# =============================================================================
# SERVICIO GEMINI API
# =============================================================================
class GeminiService:
    """Servicio de integracion con Gemini API"""

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
        """Llama a la API de Gemini"""
        if not self.api_key:
            raise ValueError("API Key no configurada")

        import requests

        url = f"{self.base_url}/{self.model}:generateContent?key={self.api_key}"

        payload = {
            "contents": [{"parts": [{"text": prompt}]}],
            "generationConfig": {
                "temperature": temperature,
                "maxOutputTokens": max_tokens
            }
        }

        response = requests.post(url, json=payload, timeout=60)
        response.raise_for_status()

        result = response.json()
        if "candidates" in result and len(result["candidates"]) > 0:
            text = result["candidates"][0]["content"]["parts"][0]["text"]
            return text
        return ""

    def _extract_json(self, text):
        """Extrae JSON de la respuesta de la API"""
        import json
        # Buscar JSON entre ```json y ```
        json_match = re.search(r'```json\s*(.*?)\s*```', text, re.DOTALL)
        if json_match:
            return json.loads(json_match.group(1))

        # Buscar JSON entre { y }
        json_match = re.search(r'\{.*\}', text, re.DOTALL)
        if json_match:
            try:
                return json.loads(json_match.group())
            except:
                pass

        # Si no hay JSON, devolver el texto como campo unico
        return {"generated_text": text}

    def generate_moc(self, problem, context="", equipo=""):
        """Genera contenido completo de MoC"""
        if not self.api_key:
            return self._generate_local_moc(problem, context, equipo)

        prompt = f"""Eres un experto senior en gestion de cambios industriales (MoC) con 20 años de experiencia en mantenimiento de plantas mineras y manufactura. Redactas documentos MoC profesionales, impecables, sin errores ortograficos ni gramaticales.

PROBLEMA DESCRITO POR EL TECNICO:
{problem}

CONTEXTO ADICIONAL:
{context}

EQUIPO DE REVISION:
{equipo}

Genera el siguiente contenido en ESPANOL, con redaccion profesional, clara, estructurada y SIN ERRORES ORTOGRAFICOS:

1. descripcion_problema: Redacta como parrafo bien estructurado, claro y profesional.
2. condicion_actual: Describe la situacion actual.
3. condicion_propuesta: Describe la solucion.
4. razones_cambio: Justificacion tecnica.
5. alternativas_retorno: Alternativas evaluadas y plan de contingencia.
6. recursos: Lista detallada de recursos.
7. plan_implementacion: Pasos detallados por fases.
8. tiempo_duracion: Estimacion de duracion.
9. riesgos_controles: Array de objetos {{"riesgo": "...", "control": "..."}} (minimo 3).
10. riesgos_shes: Array de objetos {{"riesgo": "...", "control": "...", "plazo": "..."}} (minimo 3).

Responde UNICAMENTE en formato JSON valido.
"""

        try:
            text = self._call_api(prompt, temperature=0.2, max_tokens=8192)
            return self._extract_json(text)
        except Exception as e:
            st.error(f"Error API Gemini: {e}. Usando generacion local.")
            return self._generate_local_moc(problem, context, equipo)

    def generate_a3(self, problem, context=""):
        """Genera contenido de Mejora A3"""
        if not self.api_key:
            return self._generate_local_a3(problem, context)

        prompt = f"""Eres un experto en metodologia A3 para mejora continua (Lean Manufacturing). Redactas documentos A3 impecables, sin errores ortograficos.

PROBLEMA: {problem}
CONTEXTO: {context}

Genera en ESPANOL formato JSON con:
titulo, antecedentes, problema_actual, analisis_situacion, objetivos, analisis_causa_raiz, contramedidas, resultados_esperados, plan_seguimiento, lecciones_aprendidas, estandarizacion

Responde JSON valido."""

        try:
            text = self._call_api(prompt, temperature=0.2, max_tokens=8192)
            return self._extract_json(text)
        except:
            return self._generate_local_a3(problem, context)

    def generate_kaizen(self, activity, context=""):
        """Genera contenido de Kaizen"""
        if not self.api_key:
            return self._generate_local_kaizen(activity, context)

        prompt = f"""Eres un experto en Kaizen y mejora continua. Redactas documentos Simple Kaizen impecables.

ACTIVIDAD: {activity}
CONTEXTO: {context}

Genera en ESPANOL formato JSON con:
titulo, area, descripcion_problema, solucion, beneficios, tipo_desperdicio, impacto_bto, proximos_pasos

Responde JSON valido."""

        try:
            text = self._call_api(prompt, temperature=0.2, max_tokens=4096)
            return self._extract_json(text)
        except:
            return self._generate_local_kaizen(activity, context)

    def translate_document(self, data):
        """Traduce documento completo al ingles"""
        if not self.api_key:
            return data

        prompt = f"""Traduce el siguiente documento tecnico de gestion de cambios del espanol al ingles profesional.
Mantén terminologia tecnica de industria minera/manufacturera.

DOCUMENTO:
{json.dumps(data, ensure_ascii=False, indent=2)}

Responde SOLO con el JSON traducido, misma estructura."""

        try:
            text = self._call_api(prompt, temperature=0.1, max_tokens=8192)
            return self._extract_json(text)
        except:
            return data

    def correct_spelling(self, text):
        """Corrige ortografia y gramatica"""
        if not self.api_key or not text.strip():
            return text

        prompt = f"""Corrige la ortografia, gramatica y puntuacion del siguiente texto en espanol.
Mantén el significado original y el tono profesional.
No agregues comentarios, devuelve SOLO el texto corregido.

TEXTO:
{text}"""

        try:
            corrected = self._call_api(prompt, temperature=0.1, max_tokens=4096)
            return corrected.strip()
        except:
            return text

    def _generate_local_moc(self, problem, context, equipo):
        """Generacion local sin API"""
        return {
            "descripcion_problema": f"Se identifico la siguiente situacion que requiere gestion mediante MoC: {problem}. Esta condicion presenta riesgos operacionales que deben ser mitigados mediante un cambio controlado y documentado, garantizando la integridad del proceso y la seguridad del personal.",
            "condicion_actual": "El equipo o proceso actual opera bajo condiciones que presentan limitaciones tecnicas documentadas. Se requiere evaluacion detallada para establecer la linea base de referencia antes de la modificacion, identificando todos los parametros criticos y condiciones operativas actuales.",
            "condicion_propuesta": "Implementar las modificaciones tecnicas necesarias para optimizar el rendimiento operativo, mejorar la seguridad del proceso y alinear las condiciones con los estandares corporativos vigentes y mejores practicas de la industria minera.",
            "razones_cambio": "1. Mejora de la seguridad operacional del proceso y reduccion de riesgos para el personal.\n2. Optimizacion del rendimiento y confiabilidad del equipo critico.\n3. Cumplimiento de estandares corporativos, regulatorios y normativos aplicables.\n4. Reduccion de riesgos identificados en evaluaciones previas de SHES.\n5. Mejora continua alineada con objetivos estrategicos de la organizacion.",
            "alternativas_retorno": "Alternativas evaluadas durante el analisis tecnico:\n\n1. Mantenimiento correctivo tradicional (DESCARTADO): Alcance limitado y temporalidad insuficiente para resolver la causa raiz.\n2. Reemplazo total del sistema (DESCARTADO): Costo elevado, tiempo de implementacion prolongado y disponibilidad limitada de equipos de reemplazo.\n3. Modificacion controlada con plan de implementacion estructurado (SELECCIONADA): Mejor relacion costo-beneficio, menor impacto operativo y alineacion con estandares de gestion de cambios.\n\nPlan de retorno:\nEn caso de que los resultados de la implementacion no sean satisfactorios o se presenten efectos adversos no previstos, se ejecutara el plan de retorno que incluye: restauracion de la configuracion original documentada, activacion del protocolo de contingencia establecido, notificacion inmediata a supervision y produccion, y registro del evento para analisis de lecciones aprendidas.",
            "recursos": "RECURSOS HUMANOS:\n- Tecnico especializado de mantenimiento electrico/mecanico con certificacion vigente.\n- Supervisor de area de mantenimiento con experiencia en gestion de cambios.\n- Especialista SHES (Seguridad, Salud, Ambiente y Sociedad) para supervision continua.\n- Operador de area para pruebas funcionales y validacion operativa.\n\nRECURSOS MATERIALES:\n- Herramientas especializadas certificadas y calibradas.\n- Repuestos y materiales de calidad certificada conforme a especificaciones tecnicas del fabricante.\n- Equipos de proteccion personal (EPP) completos: casco, gafas de seguridad, guantes anticorte, botas punta de acero, arnes de seguridad.\n- Materiales de senalizacion, demarcacion de area y cinta de precaucion.\n\nRECURSOS TECNICOS:\n- Documentacion tecnica del fabricante actualizada y procedimientos operativos estandar (SOP) vigentes.\n- Permisos de trabajo segun aplique: trabajo en caliente, en altura, espacios confinados, trabajo electrico.\n- Checklist de verificacion pre-operacional y formatos de registro.",
            "plan_implementacion": "FASE 1: PREPARACION (Dias 1-2)\n- Reunion de coordinacion con produccion para programar ventana de mantenimiento.\n- Verificacion de disponibilidad de todos los recursos (humanos, materiales, tecnicos).\n- Preparacion del area de trabajo: limpieza, senalizacion, bloqueo y etiquetado (LOTO).\n- Briefing de seguridad con todo el personal involucrado.\n\nFASE 2: EJECUCION (Dias 3-5)\n- Implementacion de modificaciones segun especificaciones tecnicas aprobadas.\n- Pruebas funcionales iniciales en condiciones controladas.\n- Registro fotografico detallado de todo el proceso (antes, durante y despues).\n- Verificacion intermedia por especialista SHES.\n\nFASE 3: VALIDACION (Dias 6-7)\n- Pruebas de operacion bajo condiciones normales de proceso.\n- Verificacion de parametros criticos y comparacion con linea base.\n- Validacion por supervisor de area y representante de produccion.\n- Documentacion de resultados de pruebas.\n\nFASE 4: CIERRE Y ESTANDARIZACION (Dia 8)\n- Actualizacion de documentacion tecnica y procedimientos operativos.\n- Entrenamiento al personal operativo sobre nuevas condiciones y parametros.\n- Socializacion de lecciones aprendidas al equipo de trabajo.\n- Cierre formal de la MoC con firmas de aprobacion correspondientes.",
            "tiempo_duracion": "Estimacion total del cambio: 8 dias habiles, distribuidos en 4 fases (Preparacion, Ejecucion, Validacion y Cierre). Esta estimacion considera la complejidad tecnica identificada, la disponibilidad de recursos y las ventanas de mantenimiento acordadas con produccion. Puede ajustarse segun condiciones operativas.",
            "riesgos_controles": [
                {"riesgo": "Interrupcion del proceso productivo durante la implementacion del cambio", "control": "Coordinacion previa con produccion para programar la intervencion en ventana de mantenimiento planificado. Comunicacion oportuna a todas las areas involucradas mediante oficio formal."},
                {"riesgo": "Falla tecnica durante la modificacion que afecte equipos adyacentes o sistemas conexos", "control": "Verificacion previa exhaustiva de todos los componentes y materiales. Disponibilidad de repuestos de emergencia. Supervision tecnica continua durante la ejecucion por personal especializado."},
                {"riesgo": "Exposicion a riesgos SHES durante la ejecucion de trabajos en campo", "control": "Aplicacion estricta de permisos de trabajo segun el tipo de intervencion. Uso obligatorio de EPP completo. Supervision continua del especialista SHES. Aplicacion de principio LOTO (Bloqueo y Etiquetado)."}
            ],
            "riesgos_shes": [
                {"riesgo": "Riesgo de lesiones por manipulacion manual de equipos, componentes pesados o herramientas", "control": "Capacitacion especifica al personal sobre tecnicas de levantamiento seguro. Uso de EPP adecuado incluyendo guantes anticorte y botas punta de acero. Senalizacion del area de trabajo. Supervisor designado.", "plazo": "Antes del inicio de trabajos"},
                {"riesgo": "Generacion de residuos solidos, liquidos o materiales de desecho durante la intervencion", "control": "Manejo adecuado de residuos segun procedimiento ambiental corporativo. Clasificacion en origen. Disposicion en areas autorizadas. Registro de disposicion final.", "plazo": "Durante toda la ejecucion"},
                {"riesgo": "Exposicion a ruido, vibraciones, agentes quimicos o condiciones ambientales adversas", "control": "Monitoreo continuo de condiciones ambientales. Uso de protectores auditivos cuando aplique. Ventilacion adecuada del area. Limitacion de horario de trabajo si se exceden limites permisibles.", "plazo": "Durante toda la ejecucion"}
            ]
        }

    def _generate_local_a3(self, problem, context):
        return {
            "titulo": "Mejora del proceso: " + problem[:50],
            "antecedentes": "Se identifico una oportunidad de mejora significativa en el proceso actual que requiere analisis estructurado mediante la metodologia A3 de resolucion de problemas, promoviendo el pensamiento lean y la mejora continua en la organizacion.",
            "problema_actual": problem,
            "analisis_situacion": "La situacion actual presenta indicadores de desempeno que reflejan oportunidades de mejora significativas. Se requiere recopilacion sistematica de datos cuantitativos para establecer una linea base solida que permita medir el impacto de las contramedidas implementadas.",
            "objetivos": "OBJETIVO GENERAL:\nOptimizar el proceso para eliminar el desperdicio identificado y mejorar los indicadores clave de desempeno del area de manera sostenible.\n\nOBJETIVOS ESPECIFICOS (SMART):\n- Reducir el tiempo de ciclo del proceso en un 15% durante los proximos 3 meses.\n- Disminuir la tasa de defectos o retrabajos en un 20% en el siguiente trimestre.\n- Mejorar la productividad del area en un 10% dentro de 6 meses.\n- Incrementar la satisfaccion del cliente interno en un 25%.",
            "analisis_causa_raiz": "ANALISIS DE CAUSA RAIZ - Metodo de los 5 Porques:\n\n1. POR QUE ocurre el problema identificado? → Porque existe una condicion operativa inadecuada que no cumple con los estandares establecidos.\n2. POR QUE la condicion es inadecuada? → Porque falta estandarizacion formal del proceso y los procedimientos no estan actualizados.\n3. POR QUE falta estandarizacion? → Porque los procedimientos operativos no han sido revisados ni actualizados periodicamente.\n4. POR QUE no estan actualizados? → Porque no existe un sistema de gestion documental que establezca frecuencias de revision y responsables claros.\n5. POR QUE no existe ese sistema? → Porque no hay una politica clara de gestion del conocimiento ni compromiso visible de la gerencia.\n\nCAUSA RAIZ IDENTIFICADA: Ausencia de sistema de gestion, actualizacion y control de procedimientos operativos estandar.",
            "contramedidas": "1. ACTUALIZAR procedimientos operativos estandar (SOP) del proceso, incluyendo instrucciones claras, parametros de control y criterios de aceptacion.\n2. IMPLEMENTAR checklists de verificacion diaria al inicio y fin de turno para asegurar cumplimiento.\n3. CAPACITAR al personal operativo y de supervision en los nuevos estandares, documentando asistencia y evaluacion de competencias.\n4. ESTABLECER indicadores de control (KPIs) visuales en el area: tablero de control con metricas diarias/semanales/mensuales.\n5. PROGRAMAR auditorias mensuales de cumplimiento lideradas por supervision con checklist estandar.",
            "resultados_esperados": "- Reduccion medible de los desperdicios identificados (Motion, Waiting, Defects).\n- Mejora sostenida en indicadores de calidad, productividad y seguridad.\n- Estandarizacion efectiva del proceso con documentacion actualizada y accesible.\n- Reduccion significativa de la variabilidad operativa entre turnos y operadores.\n- Incremento en la satisfaccion del cliente interno y externo.\n- Cultura de mejora continua fortalecida en el equipo de trabajo.",
            "plan_seguimiento": "SEMANA 1-2: Implementacion de contramedidas identificadas, capacitacion inicial del personal y puesta en marcha de KPIs.\nSEMANA 3-4: Monitoreo inicial de indicadores, ajustes rapidos a contramedidas si es necesario.\nMES 2: Primera auditoria formal de cumplimiento, revision de resultados parciales vs. objetivos.\nMES 3: Evaluacion integral de resultados, comparacion con linea base, decision de estandarizacion o ajuste.\nMES 6: Revision de sostenibilidad, analisis de tendencias, documentacion de lecciones aprendidas y replicacion en procesos similares.",
            "lecciones_aprendidas": "La aplicacion de la metodologia A3 permitio visualizar de manera integral el problema, sus causas raiz y las soluciones correspondientes en un solo documento estructurado. La participacion activa del equipo multidisciplinario (operaciones, mantenimiento, calidad, SHES) fue fundamental para identificar la causa raiz real y no quedarse en causas superficiales. Se recomienda replicar esta metodologia en otros procesos criticos de la planta.",
            "estandarizacion": "Los nuevos procedimientos y estandares desarrollados seran documentados formalmente, revisados y aprobados por las gerencias correspondientes. Se socializaran a todo el personal involucrado mediante capacitaciones programadas. Se integraran al Sistema de Gestion de Calidad de la empresa y se establecera un calendario de revision periodica (minimo anual) para mantener la vigencia y mejora continua."
        }

    def _generate_local_kaizen(self, activity, context):
        return {
            "titulo": "Kaizen: " + activity[:50],
            "area": "Area de Mantenimiento / Produccion / Calidad",
            "descripcion_problema": activity,
            "solucion": "Se implemento una mejora orientada a eliminar el desperdicio identificado y optimizar el flujo del proceso, aplicando principios fundamentales de Lean Manufacturing y el pensamiento de mejora continua (Kaizen). La solucion fue desarrollada con la participacion activa del equipo de trabajo.",
            "beneficios": "- Reduccion del tiempo requerido para ejecutar la actividad en un porcentaje significativo.\n- Mejora inmediata en la calidad y consistencia del resultado obtenido.\n- Mayor seguridad para el personal durante la ejecucion de la tarea.\n- Reduccion de costos operativos asociados al proceso.\n- Mejora en el ambiente de trabajo y moral del equipo.\n- Eliminacion de movimientos innecesarios y tiempos de espera.",
            "tipo_desperdicio": "Motion / Waiting / Skills (seleccionar segun el desperdicio principal identificado en el analisis)",
            "impacto_bto": "Supply Chain and Manufacturing Excellence / Safe and Sustainable (seleccionar segun el impacto principal identificado)",
            "proximos_pasos": "1. Documentar formalmente la mejora en el sistema de gestion de la empresa.\n2. Socializar con otras areas relacionadas que puedan beneficiarse de esta mejora.\n3. Replicar la mejora en procesos o areas similares identificadas durante el analisis.\n4. Establecer monitoreo mensual de sostenibilidad de la mejora implementada.\n5. Reconocer al equipo participante como parte de la cultura de mejora continua."
        }


# =============================================================================
# GENERADOR DE DOCUMENTOS - LLENADO PRECISO DE TEMPLATES
# =============================================================================
class DocumentGenerator:
    """Genera documentos finales manteniendo los formatos oficiales exactos"""

    def __init__(self):
        self.output_dir = OUTPUT_DIR

    def generate_moc(self, data, images=None, header_text="", footer_text=""):
        """Genera MoC en PowerPoint manteniendo el formato original"""
        from pptx import Presentation
        from pptx.util import Inches, Pt, Emu
        from pptx.dml.color import RGBColor
        from pptx.enum.text import PP_ALIGN
        from copy import deepcopy

        template_path = TEMPLATES_DIR / "moc_template.pptx"
        prs = Presentation(str(template_path))

        # Funcion helper para limpiar y llenar celdas de tabla
        def fill_table_cell(cell, text):
            cell.text = ""
            p = cell.text_frame.paragraphs[0]
            p.text = str(text)
            p.font.size = Pt(11)
            p.font.name = "Calibri"
            p.alignment = PP_ALIGN.LEFT
            p.word_wrap = True

        # SLIDE 1: Portada
        slide1 = prs.slides[0]
        for shape in slide1.shapes:
            if not shape.has_text_frame:
                continue
            text = shape.text_frame.text

            if "MOC:" in text and shape.shape_type == 14:  # Placeholder
                # Es el titulo principal
                for para in shape.text_frame.paragraphs:
                    para.clear()
                p = shape.text_frame.paragraphs[0]
                p.text = f"MOC: {data.get('moc_title', '')}"
                p.font.size = Pt(24)
                p.font.bold = True
                p.font.color.rgb = RGBColor(0x1a, 0x5f, 0x7a)

            elif "MAYO" in text or "ENERO" in text or "FEBRERO" in text:
                for para in shape.text_frame.paragraphs:
                    para.clear()
                p = shape.text_frame.paragraphs[0]
                p.text = data.get('fecha', Utils.format_date())
                p.font.size = Pt(18)

            elif "nÚmero" in text or "NÚMERO" in text or "numero" in text.lower():
                for para in shape.text_frame.paragraphs:
                    para.clear()
                p = shape.text_frame.paragraphs[0]
                p.text = f"NÚMERO DE LA MOC: {data.get('moc_number', '')}"
                p.font.size = Pt(16)
                p.font.color.rgb = RGBColor(0x00, 0x70, 0xC0)

            elif "Naturaleza" in text:
                for para in shape.text_frame.paragraphs:
                    para.clear()
                p = shape.text_frame.paragraphs[0]
                p.text = f"Naturaleza de la moc: {data.get('naturaleza', 'permanente')}"
                p.font.size = Pt(14)

            elif "ORIGINADOR" in text:
                for para in shape.text_frame.paragraphs:
                    para.clear()
                p = shape.text_frame.paragraphs[0]
                p.text = f"ORIGINADOR DE LA MOC: {data.get('originador', '')}"
                p.font.size = Pt(14)

        # SLIDE 2: Equipo
        slide2 = prs.slides[1]
        for shape in slide2.shapes:
            if not shape.has_text_frame:
                continue
            text = shape.text_frame.text

            if "Producción" in text or "Produccion" in text:
                for para in shape.text_frame.paragraphs:
                    para.clear()
                p = shape.text_frame.paragraphs[0]
                p.text = f"Producción: {data.get('produccion', '')}\nSpecialist SHES: {data.get('specialist_shes', '')}\nMantenimiento: {data.get('mantenimiento', '')}"
                p.font.size = Pt(12)

            elif "Revisor 1" in text:
                for para in shape.text_frame.paragraphs:
                    para.clear()
                p = shape.text_frame.paragraphs[0]
                revisores = data.get('revisores', '')
                p.text = f"Revisor 1: {revisores}\nRevisor 2:\nRevisor 3:\nRevisor 4:\nAprobador Final:"
                p.font.size = Pt(12)

            elif "Experto" in text and "aprobador" in text:
                for para in shape.text_frame.paragraphs:
                    para.clear()
                p = shape.text_frame.paragraphs[0]
                p.text = f"Experto aprobador: {data.get('experto_aprobador', '')}"
                p.font.size = Pt(12)

        # SLIDE 3: Descripcion del Problema
        slide3 = prs.slides[2]
        for shape in slide3.shapes:
            if shape.has_text_frame and "3" in shape.text_frame.text:
                # Es el placeholder del numero, lo reemplazamos con el contenido
                for para in shape.text_frame.paragraphs:
                    para.clear()
                p = shape.text_frame.paragraphs[0]
                desc = data.get('descripcion_problema', '')
                p.text = desc
                p.font.size = Pt(14)
                p.alignment = PP_ALIGN.LEFT
                p.word_wrap = True
                break

        # SLIDE 4: Descripcion del Cambio (Tabla)
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
            if not shape.has_text_frame:
                continue
            text = shape.text_frame.text
            if "5" in text and shape.shape_type == 14:
                for para in shape.text_frame.paragraphs:
                    para.clear()
                p = shape.text_frame.paragraphs[0]
                p.text = data.get('razones_cambio', '')
                p.font.size = Pt(13)
                p.alignment = PP_ALIGN.LEFT
                p.word_wrap = True
            elif "Alternativas" in text:
                for para in shape.text_frame.paragraphs:
                    para.clear()
                p = shape.text_frame.paragraphs[0]
                p.text = data.get('alternativas_retorno', '')
                p.font.size = Pt(12)
                p.alignment = PP_ALIGN.LEFT
                p.word_wrap = True

        # SLIDE 6: Recursos y Plan
        slide6 = prs.slides[5]
        for shape in slide6.shapes:
            if not shape.has_text_frame:
                continue
            text = shape.text_frame.text
            if "Recursos:" in text:
                for para in shape.text_frame.paragraphs:
                    para.clear()
                p = shape.text_frame.paragraphs[0]
                p.text = f"Recursos:\n{data.get('recursos', '')}"
                p.font.size = Pt(12)
                p.alignment = PP_ALIGN.LEFT
                p.word_wrap = True
            elif "Plan de Implementación" in text:
                # Ya esta cubierto arriba o en otro shape
                pass

        # SLIDE 7: Tiempo
        slide7 = prs.slides[6]
        for shape in slide7.shapes:
            if shape.has_text_frame and "7" in shape.text_frame.text:
                for para in shape.text_frame.paragraphs:
                    para.clear()
                p = shape.text_frame.paragraphs[0]
                p.text = data.get('tiempo_duracion', '')
                p.font.size = Pt(14)
                p.alignment = PP_ALIGN.LEFT
                p.word_wrap = True
                break

        # SLIDE 8: Riesgos (Tabla 3x3)
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

        # SLIDE 9: Riesgos SHES (Tabla 4x4)
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

        # Agregar imagenes como slides adicionales
        if images:
            for idx, img_info in enumerate(images, 1):
                blank_layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]
                new_slide = prs.slides.add_slide(blank_layout)

                # Titulo
                left = Inches(0.5)
                top = Inches(0.3)
                width = Inches(9)
                height = Inches(0.5)
                title_box = new_slide.shapes.add_textbox(left, top, width, height)
                tf = title_box.text_frame
                tf.text = f"Figura {idx} - Imagen de Soporte"
                for p in tf.paragraphs:
                    p.font.size = Pt(18)
                    p.font.bold = True
                    p.font.color.rgb = RGBColor(0x1a, 0x5f, 0x7a)

                # Imagen
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
                    img_top = Inches(1)

                    new_slide.shapes.add_picture(img_path, img_left, img_top, w, h)

                    # Descripcion
                    desc = img_info.get("desc", f"Descripcion de la Figura {idx}") if isinstance(img_info, dict) else f"Figura {idx}"
                    desc_box = new_slide.shapes.add_textbox(Inches(0.5), Inches(7.3), Inches(9), Inches(0.5))
                    dtf = desc_box.text_frame
                    dtf.text = desc
                    for p in dtf.paragraphs:
                        p.font.size = Pt(11)
                        p.font.italic = True
                        p.font.color.rgb = RGBColor(0x47, 0x55, 0x69)
                        p.alignment = PP_ALIGN.CENTER
                except Exception as e:
                    print(f"Error imagen {idx}: {e}")

        # Guardar
        output_path = self.output_dir / f"{data.get('moc_number', 'MOC')}.pptx"
        prs.save(str(output_path))
        return output_path

    def generate_a3(self, data, images=None):
        """Genera A3 en Word manteniendo el formato original con textboxes"""
        from docx import Document
        from docx.shared import Inches, Pt, RGBColor
        from docx.enum.text import WD_ALIGN_PARAGRAPH

        template_path = TEMPLATES_DIR / "a3_template.docx"
        doc = Document(str(template_path))

        # El A3 tiene textboxes incrustados. Necesitamos buscar en el XML
        # Por ahora, reemplazamos lo que podemos en parrafos normales
        # y agregamos contenido al final si no hay placeholders claros

        # Reemplazar en parrafos
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
                        run.font.size = Pt(11)

        # Agregar contenido estructurado al final del documento
        # ya que el template A3 tiene textboxes que son dificiles de llenar via python-docx
        doc.add_page_break()

        # Titulo
        heading = doc.add_heading(data.get('titulo', 'Mejora A3'), level=1)
        for run in heading.runs:
            run.font.color.rgb = RGBColor(0x1a, 0x5f, 0x7a)
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
                run.font.color.rgb = RGBColor(0x1a, 0x5f, 0x7a)
                run.font.name = 'Calibri'

            content = data.get(key, '')
            if content:
                p = doc.add_paragraph(content)
                for run in p.runs:
                    run.font.name = 'Calibri'
                    run.font.size = Pt(11)

        # Agregar imagenes
        if images:
            doc.add_page_break()
            doc.add_heading('IMAGENES DE SOPORTE', level=1)

            for idx, img_info in enumerate(images, 1):
                doc.add_paragraph()
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run(f"Figura {idx}")
                run.bold = True
                run.font.size = Pt(12)
                run.font.color.rgb = RGBColor(0x1a, 0x5f, 0x7a)

                try:
                    img_path = img_info["path"] if isinstance(img_info, dict) else img_info
                    doc.add_picture(img_path, width=Inches(5.5))
                    last_paragraph = doc.paragraphs[-1]
                    last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

                    desc = img_info.get("desc", f"Descripcion de la Figura {idx}") if isinstance(img_info, dict) else f"Figura {idx}"
                    caption = doc.add_paragraph()
                    caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    cap_run = caption.add_run(desc)
                    cap_run.italic = True
                    cap_run.font.size = Pt(10)
                    cap_run.font.color.rgb = RGBColor(0x47, 0x55, 0x69)
                except Exception as e:
                    doc.add_paragraph(f"[Error al cargar imagen {idx}: {e}]")

        output_path = self.output_dir / f"{data.get('doc_number', 'A3')}.docx"
        doc.save(str(output_path))
        return output_path

    def generate_kaizen(self, data, images=None):
        """Genera Kaizen en PowerPoint manteniendo el formato original"""
        from pptx import Presentation
        from pptx.util import Inches, Pt
        from pptx.dml.color import RGBColor
        from pptx.enum.text import PP_ALIGN

        template_path = TEMPLATES_DIR / "kaizen_template.pptx"
        prs = Presentation(str(template_path))

        slide = prs.slides[0]

        # Llenar tablas de Kaizen
        for shape in slide.shapes:
            if shape.has_table:
                table = shape.table
                # La tabla tiene BTO, 8 Wastes, X
                # Llenar columna X segun seleccion
                tipo_desp = data.get('tipo_desperdicio', '')
                impacto = data.get('impacto_bto', '')

                # Mapear tipos de desperdicio a filas
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
                    # Reemplazar con datos del lider
                    for para in shape.text_frame.paragraphs:
                        para.clear()
                    p = shape.text_frame.paragraphs[0]
                    p.text = f"Leader: {data.get('autor', '')}\nArea: {data.get('area', '')}\nFecha: {data.get('fecha', '')}"
                    p.font.size = Pt(12)

        # Agregar slide con descripcion detallada
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
            # Titulo de seccion
            sec_box = detail_slide.shapes.add_textbox(Inches(0.5), Inches(y_pos), Inches(9), Inches(0.3))
            stf = sec_box.text_frame
            stf.text = section_title
            for p in stf.paragraphs:
                p.font.size = Pt(14)
                p.font.bold = True
                p.font.color.rgb = RGBColor(0x1a, 0x5f, 0x7a)

            y_pos += 0.3

            # Contenido
            content = data.get(key, '')
            if content:
                content_box = detail_slide.shapes.add_textbox(Inches(0.5), Inches(y_pos), Inches(9), Inches(1.5))
                ctf = content_box.text_frame
                ctf.text = content
                for p in ctf.paragraphs:
                    p.font.size = Pt(12)
                    p.word_wrap = True
                y_pos += 1.5

        # Agregar imagenes
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
                    print(f"Error imagen: {e}")

        output_path = self.output_dir / f"{data.get('doc_number', 'KZN')}.pptx"
        prs.save(str(output_path))
        return output_path

    def convert_to_pdf(self, doc_path):
        """Convierte documento a PDF"""
        try:
            import subprocess
            pdf_path = str(doc_path).replace('.pptx', '.pdf').replace('.docx', '.pdf')

            # Intentar usar LibreOffice
            result = subprocess.run([
                'soffice', '--headless', '--convert-to', 'pdf',
                '--outdir', str(self.output_dir), str(doc_path)
            ], capture_output=True, text=True, timeout=60)

            if os.path.exists(pdf_path):
                return Path(pdf_path)

            # Fallback: usar reportlab para crear PDF basico
            return self._create_basic_pdf(doc_path)
        except Exception as e:
            st.warning(f"No se pudo convertir a PDF automaticamente: {e}. El documento Office esta disponible.")
            return None

    def _create_basic_pdf(self, doc_path):
        """Crea un PDF basico como fallback"""
        try:
            from reportlab.lib.pagesizes import letter, A4
            from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage
            from reportlab.lib.styles import getSampleStyleSheet
            from reportlab.lib.units import inch

            pdf_path = str(doc_path).replace('.pptx', '.pdf').replace('.docx', '.pdf')
            doc = SimpleDocTemplate(pdf_path, pagesize=A4)
            styles = getSampleStyleSheet()
            story = []

            story.append(Paragraph("Documento Generado", styles['Title']))
            story.append(Spacer(1, 0.2*inch))
            story.append(Paragraph(f"Archivo original: {doc_path.name}", styles['Normal']))
            story.append(Paragraph("Por favor abra el archivo original en Microsoft Office o LibreOffice.", styles['Normal']))

            doc.build(story)
            return Path(pdf_path)
        except:
            return None


# =============================================================================
# FUNCIONES DE INTERFAZ STREAMLIT - SIDEBAR Y NAVEGACION
# =============================================================================

def render_sidebar():
    """Renderiza la barra lateral de navegacion"""
    config = ConfigManager.load()

    st.sidebar.markdown("""
    <div style="text-align: center; padding: 1rem 0;">
        <h2 style="color: #f8fafc; margin: 0; font-size: 1.3rem;">📋 GESTIÓN</h2>
        <h2 style="color: #f8fafc; margin: 0; font-size: 1.3rem;">DOCUMENTAL</h2>
        <p style="color: #94a3b8; margin-top: 0.5rem; font-size: 0.85rem;">MoC · A3 · Kaizen</p>
    </div>
    <hr style="border-color: #334155; margin: 1rem 0;">
    """, unsafe_allow_html=True)

    # Navegacion
    st.sidebar.markdown("### 🧭 Navegación")

    nav_options = {
        "inicio": "🏠 Inicio",
        "nueva_moc": "📋 Nueva MoC",
        "nueva_a3": "📊 Nueva Mejora A3",
        "nuevo_kaizen": "⚡ Nuevo Kaizen",
        "historial": "📁 Historial",
        "configuracion": "⚙️ Configuración"
    }

    for key, label in nav_options.items():
        if st.sidebar.button(label, key=f"nav_{key}", use_container_width=True):
            st.session_state.page = key
            st.rerun()

    st.sidebar.markdown("<hr style='border-color: #334155; margin: 1rem 0;'>", unsafe_allow_html=True)

    # Info de version
    model_name = GeminiService.MODELS.get(config.get("gemini_model", "gemini-1.5-pro"), {}).get("name", "3.1 Pro")
    st.sidebar.markdown(f"""
    <div style="text-align: center; color: #64748b; font-size: 0.75rem;">
        <p>Modelo IA: <span class="gemini-badge">{model_name}</span></p>
        <p>v2.0.0 · Mayo 2026</p>
    </div>
    """, unsafe_allow_html=True)

    # Footer empresa
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
    """Renderiza el header principal"""
    st.markdown("""
    <div class="main-header">
        <h1>🎯 Sistema de Gestión Documental</h1>
        <p>Automatización inteligente de documentos MoC, Mejora A3 y Simple Kaizen</p>
    </div>
    """, unsafe_allow_html=True)


def render_welcome():
    """Renderiza la pantalla de bienvenida"""
    render_header()

    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown("""
        <div class="doc-card doc-card-moc" onclick="">
            <h3 style="color: #1a5f7a; margin-top: 0;">📋 Management of Change</h3>
            <p style="color: #64748b; font-size: 0.9rem;">Documento base para cualquier cambio en proceso o maquinaria. Incluye evaluación de riesgos, plan de implementación y controles SHES.</p>
            <ul style="color: #475569; font-size: 0.85rem; padding-left: 1.2rem;">
                <li>Evaluación técnica completa</li>
                <li>Riesgos y controles</li>
                <li>Plan de implementación</li>
                <li>Traducción al inglés</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
        if st.button("➕ Crear MoC", key="btn_moc", use_container_width=True):
            st.session_state.page = "nueva_moc"
            st.rerun()

    with col2:
        st.markdown("""
        <div class="doc-card doc-card-a3" onclick="">
            <h3 style="color: #10b981; margin-top: 0;">📊 Mejora A3</h3>
            <p style="color: #64748b; font-size: 0.9rem;">Formato estructurado para registrar mejoras. Metodología Lean con análisis de causa raíz y contramedidas.</p>
            <ul style="color: #475569; font-size: 0.85rem; padding-left: 1.2rem;">
                <li>Análisis 5 Porqués</li>
                <li>Contramedidas SMART</li>
                <li>Plan de seguimiento</li>
                <li>Estandarización</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
        if st.button("➕ Crear A3", key="btn_a3", use_container_width=True):
            st.session_state.page = "nueva_a3"
            st.rerun()

    with col3:
        st.markdown("""
        <div class="doc-card doc-card-kaizen" onclick="">
            <h3 style="color: #f59e0b; margin-top: 0;">⚡ Simple Kaizen</h3>
            <p style="color: #64748b; font-size: 0.9rem;">Registro rápido de mejoras del día a día. Captura ideas de mejora continua del equipo.</p>
            <ul style="color: #475569; font-size: 0.85rem; padding-left: 1.2rem;">
                <li>8 Desperdicios (Wastes)</li>
                <li>Impacto BTO</li>
                <li>Beneficios cuantificados</li>
                <li>Replicabilidad</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
        if st.button("➕ Crear Kaizen", key="btn_kaizen", use_container_width=True):
            st.session_state.page = "nuevo_kaizen"
            st.rerun()

    # Estadisticas rapidas
    st.markdown("<br>", unsafe_allow_html=True)
    history = HistoryManager.load()
    docs = history.get("documents", [])

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        moc_count = len([d for d in docs if d.get("type") == "moc"])
        st.metric("📋 MoC Generadas", moc_count)
    with col2:
        a3_count = len([d for d in docs if d.get("type") == "a3"])
        st.metric("📊 A3 Generadas", a3_count)
    with col3:
        kzn_count = len([d for d in docs if d.get("type") == "kaizen"])
        st.metric("⚡ Kaizen Generados", kzn_count)
    with col4:
        total = len(docs)
        st.metric("📁 Total Documentos", total)


# =============================================================================
# FORMULARIOS DE CREACION
# =============================================================================

def render_moc_form():
    """Renderiza formulario de creacion MoC"""
    config = ConfigManager.load()

    st.markdown("""
    <div class="section-header">
        <h3>📋 Nueva Management of Change (MoC)</h3>
    </div>
    """, unsafe_allow_html=True)

    st.info("💡 **Instrucciones:** Complete la información general y describa el problema con sus propias palabras. La IA generará automáticamente todos los campos técnicos del documento.")

    # Informacion General
    st.markdown("#### 1. Información General")

    col1, col2 = st.columns(2)
    with col1:
        moc_title = st.text_input("Título de la MoC:", 
                                  placeholder="Ej: Control de Acceso a panel HMI - carga detonadores",
                                  help="Describa brevemente el alcance del cambio")
        moc_number = st.text_input("Número de MoC:", 
                                   value=Utils.generate_doc_number("moc"),
                                   disabled=True,
                                   help="Generado automáticamente")
    with col2:
        naturaleza = st.selectbox("Naturaleza:", 
                                  ["permanente", "temporal", "emergencia"],
                                  help="Tipo de cambio según duración")
        originador = st.text_input("Originador:", 
                                   value=config.get("default_author", ""),
                                   help="Nombre del responsable que origina la MoC")

    fecha = st.text_input("Fecha:", value=Utils.format_date(), disabled=True)

    # Equipo de Revision
    st.markdown("#### 2. Equipo de Revisión")

    col1, col2, col3 = st.columns(3)
    with col1:
        produccion = st.text_input("Producción:", placeholder="Nombre del representante")
        specialist_shes = st.text_input("Specialist SHES:", placeholder="Nombre del especialista")
    with col2:
        mantenimiento = st.text_input("Mantenimiento:", placeholder="Nombre del supervisor")
        revisores = st.text_input("Revisores Enablon:", placeholder="Nombres separados por coma")
    with col3:
        experto_aprobador = st.text_input("Experto Aprobador:", placeholder="Nombre del experto")

    # Descripcion del Problema
    st.markdown("#### 3. Descripción del Problema")

    st.markdown("""
    <div class="spell-hint">
    ✍️ <strong>Corrector ortográfico activo:</strong> La IA corregirá automáticamente errores de ortografía, 
    gramática y puntuación al generar el documento. Escriba con sus palabras, no se preocupe por errores menores.
    </div>
    """, unsafe_allow_html=True)

    problem_desc = st.text_area("Describa el problema con sus propias palabras:",
                                height=200,
                                placeholder="Ejemplo: El panel HMI de la línea de carga de detonadores presenta fallas intermitentes en la visualización de parámetros críticos. Los operadores no pueden ver correctamente los valores de presión y temperatura durante el proceso de carga, lo que genera riesgos de operación y potenciales paradas no planificadas...",
                                help="Sea lo más detallado posible. Incluya: qué pasa, dónde, cuándo, impacto")

    # Contexto adicional
    st.markdown("#### 4. Contexto Adicional (Opcional)")
    context = st.text_area("Información adicional relevante:",
                           height=100,
                           placeholder="Equipo TAG, área específica, fechas relevantes, personal involucrado...",
                           help="Cualquier información que ayude a la IA a contextualizar mejor")

    # Imagenes
    st.markdown("#### 5. Imágenes de Soporte (Opcional)")
    uploaded_images = st.file_uploader("Seleccione imágenes:", 
                                       type=["png", "jpg", "jpeg"],
                                       accept_multiple_files=True,
                                       help="Se numerarán automáticamente como Figura 1, 2, 3...")

    # Guardar imagenes temporalmente
    image_paths = []
    if uploaded_images:
        for idx, img_file in enumerate(uploaded_images, 1):
            img_path = OUTPUT_DIR / f"temp_img_{moc_number}_{idx}.png"
            with open(img_path, "wb") as f:
                f.write(img_file.getbuffer())
            image_paths.append({"path": str(img_path), "desc": f"Figura {idx} - {img_file.name}"})
        st.success(f"📷 {len(image_paths)} imagen(es) cargada(s) correctamente")

    # Boton Generar
    st.markdown("<br>", unsafe_allow_html=True)

    col1, col2 = st.columns([1, 1])
    with col1:
        generate_clicked = st.button("🤖 Generar Documento MoC con IA", 
                                     type="primary",
                                     use_container_width=True)
    with col2:
        st.markdown("""
        <div style="padding-top: 0.5rem; color: #64748b; font-size: 0.9rem;">
            ⚡ La IA completará: descripción, condiciones, razones, recursos, plan, riesgos y controles
        </div>
        """, unsafe_allow_html=True)

    if generate_clicked:
        if not problem_desc.strip():
            st.error("❌ Por favor describa el problema antes de generar.")
            return

        with st.spinner("🧠 La IA está analizando y generando el documento completo..."):
            gemini = GeminiService(config.get("gemini_api_key", ""), config.get("gemini_model", "gemini-1.5-pro"))

            equipo_data = {
                "produccion": produccion,
                "specialist_shes": specialist_shes,
                "mantenimiento": mantenimiento,
                "revisores": revisores,
                "experto_aprobador": experto_aprobador
            }

            result = gemini.generate_moc(problem_desc, context, json.dumps(equipo_data))

            # Guardar en session state para revision
            st.session_state.generated_data = result
            st.session_state.moc_meta = {
                "moc_title": moc_title,
                "moc_number": moc_number,
                "naturaleza": naturaleza,
                "originador": originador,
                "fecha": fecha,
                **equipo_data
            }
            st.session_state.moc_images = image_paths
            st.session_state.page = "revisar_moc"
            st.rerun()


def render_a3_form():
    """Renderiza formulario de creacion A3"""
    config = ConfigManager.load()

    st.markdown("""
    <div class="section-header">
        <h3>📊 Nueva Mejora A3</h3>
    </div>
    """, unsafe_allow_html=True)

    st.info("💡 **Instrucciones:** Describa el problema actual y la IA generará el documento A3 completo con análisis de causa raíz, contramedidas y plan de acción.")

    # Informacion General
    st.markdown("#### 1. Información General")

    col1, col2 = st.columns(2)
    with col1:
        a3_title = st.text_input("Título de la Mejora:", 
                                 placeholder="Ej: Reducción de tiempo de cambio de formato")
        area = st.text_input("Área/Proceso:", 
                             value=config.get("default_area", ""),
                             placeholder="Ej: Línea de empaque")
    with col2:
        autor = st.text_input("Autor:", 
                              value=config.get("default_author", ""))
        doc_number = st.text_input("Número:", 
                                   value=Utils.generate_doc_number("a3"),
                                   disabled=True)

    fecha = st.text_input("Fecha:", value=Utils.format_date(), disabled=True)

    # Problema
    st.markdown("#### 2. Descripción del Problema Actual")

    st.markdown("""
    <div class="spell-hint">
    ✍️ <strong>Corrector ortográfico activo:</strong> Escriba libremente, la IA corregirá ortografía y gramática.
    </div>
    """, unsafe_allow_html=True)

    problem_desc = st.text_area("Describa el problema actual:",
                                height=200,
                                placeholder="¿Qué está pasando? ¿Cuál es el impacto? ¿Desde cuándo ocurre? ¿Tiene datos o métricas?",
                                help="Sea específico y cuantifique el impacto cuando sea posible")

    # Contexto
    st.markdown("#### 3. Contexto Adicional (Opcional)")
    context = st.text_area("Información adicional:",
                           height=80,
                           placeholder="Datos actuales, objetivos, restricciones...")

    # Imagenes
    st.markdown("#### 4. Imágenes de Soporte (Opcional)")
    uploaded_images = st.file_uploader("Seleccione imágenes:", 
                                       type=["png", "jpg", "jpeg"],
                                       accept_multiple_files=True)

    image_paths = []
    if uploaded_images:
        for idx, img_file in enumerate(uploaded_images, 1):
            img_path = OUTPUT_DIR / f"temp_img_a3_{doc_number}_{idx}.png"
            with open(img_path, "wb") as f:
                f.write(img_file.getbuffer())
            image_paths.append({"path": str(img_path), "desc": f"Figura {idx} - {img_file.name}"})
        st.success(f"📷 {len(image_paths)} imagen(es) cargada(s)")

    # Boton
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🤖 Generar Documento A3 con IA", type="primary", use_container_width=True):
        if not problem_desc.strip():
            st.error("❌ Por favor describa el problema antes de generar.")
            return

        with st.spinner("🧠 Generando documento A3 completo..."):
            gemini = GeminiService(config.get("gemini_api_key", ""), config.get("gemini_model", "gemini-1.5-pro"))
            result = gemini.generate_a3(problem_desc, context)

            st.session_state.generated_data = result
            st.session_state.a3_meta = {
                "titulo": a3_title,
                "area": area,
                "autor": autor,
                "doc_number": doc_number,
                "fecha": fecha
            }
            st.session_state.a3_images = image_paths
            st.session_state.page = "revisar_a3"
            st.rerun()


def render_kaizen_form():
    """Renderiza formulario de creacion Kaizen"""
    config = ConfigManager.load()

    st.markdown("""
    <div class="section-header">
        <h3>⚡ Nuevo Simple Kaizen</h3>
    </div>
    """, unsafe_allow_html=True)

    st.info("💡 **Instrucciones:** Describa la actividad de mejora realizada. La IA completará los campos del formato Simple Kaizen.")

    # Informacion
    st.markdown("#### 1. Información General")

    col1, col2 = st.columns(2)
    with col1:
        kaizen_title = st.text_input("Título del Kaizen:", 
                                     placeholder="Ej: Organización de área de herramientas")
        area = st.text_input("Área:", 
                             value=config.get("default_area", ""),
                             placeholder="Ej: Taller de mantenimiento")
    with col2:
        autor = st.text_input("Autor:", 
                              value=config.get("default_author", ""))
        doc_number = st.text_input("Número:", 
                                   value=Utils.generate_doc_number("kaizen"),
                                   disabled=True)

    fecha = st.text_input("Fecha:", value=Utils.format_date(), disabled=True)

    # Actividad
    st.markdown("#### 2. Descripción de la Actividad Realizada")

    st.markdown("""
    <div class="spell-hint">
    ✍️ <strong>Corrector ortográfico activo:</strong> Describa la mejora que realizó, la IA optimizará la redacción.
    </div>
    """, unsafe_allow_html=True)

    activity_desc = st.text_area("Describa la mejora o actividad:",
                                 height=200,
                                 placeholder="Ejemplo: Se reorganizó el área de herramientas implementando un sistema de shadow board que permite identificar visualmente las herramientas faltantes y reducir el tiempo de búsqueda en un 40%...",
                                 help="Incluya: qué hizo, cómo lo hizo, resultado obtenido")

    # Tipo de desperdicio
    st.markdown("#### 3. Tipo de Desperdicio (Opcional)")
    tipo_desp = st.multiselect("Seleccione el tipo de desperdicio eliminado:",
                               ["Motion", "Skills", "Inventory", "Transportation",
                                "Over Production", "Over Processing", "Waiting", "Defects"],
                               help="Puede seleccionar múltiples")

    # Impacto BTO
    impacto_bto = st.selectbox("Impacto en BTO:",
                               ["Safe and Sustainable", "People & Culture",
                                "Network Optimisation", "Supply Chain and Manufacturing Excellence"],
                               help="Seleccione el pilar de BTO impactado")

    # Contexto
    context = st.text_area("Contexto adicional:", height=80, placeholder="Información complementaria...")

    # Imagenes
    st.markdown("#### 4. Imágenes de Soporte (Opcional)")
    uploaded_images = st.file_uploader("Seleccione imágenes:", 
                                       type=["png", "jpg", "jpeg"],
                                       accept_multiple_files=True)

    image_paths = []
    if uploaded_images:
        for idx, img_file in enumerate(uploaded_images, 1):
            img_path = OUTPUT_DIR / f"temp_img_kzn_{doc_number}_{idx}.png"
            with open(img_path, "wb") as f:
                f.write(img_file.getbuffer())
            image_paths.append({"path": str(img_path), "desc": f"Figura {idx} - {img_file.name}"})
        st.success(f"📷 {len(image_paths)} imagen(es) cargada(s)")

    # Boton
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🤖 Generar Documento Kaizen con IA", type="primary", use_container_width=True):
        if not activity_desc.strip():
            st.error("❌ Por favor describa la actividad antes de generar.")
            return

        with st.spinner("🧠 Generando documento Kaizen..."):
            gemini = GeminiService(config.get("gemini_api_key", ""), config.get("gemini_model", "gemini-1.5-pro"))
            result = gemini.generate_kaizen(activity_desc, context)

            # Agregar datos seleccionados manualmente
            result["tipo_desperdicio"] = ", ".join(tipo_desp) if tipo_desp else result.get("tipo_desperdicio", "")
            result["impacto_bto"] = impacto_bto

            st.session_state.generated_data = result
            st.session_state.kaizen_meta = {
                "titulo": kaizen_title,
                "area": area,
                "autor": autor,
                "doc_number": doc_number,
                "fecha": fecha
            }
            st.session_state.kaizen_images = image_paths
            st.session_state.page = "revisar_kaizen"
            st.rerun()


# =============================================================================
# PANTALLAS DE REVISION Y EDICION
# =============================================================================

def render_moc_review():
    """Renderiza pantalla de revision y edicion MoC"""
    st.markdown("""
    <div class="section-header">
        <h3>👁️ Revisar y Editar Documento MoC</h3>
    </div>
    """, unsafe_allow_html=True)

    st.info("💡 Revise cada campo, realice las correcciones necesarias y genere el documento final. Todos los cambios se guardan automáticamente.")

    data = st.session_state.get("generated_data", {})
    meta = st.session_state.get("moc_meta", {})
    images = st.session_state.get("moc_images", [])
    config = ConfigManager.load()

    # Tabs para organizar la revision
    tabs = st.tabs(["📋 Información General", "📝 Descripción y Cambio", "📊 Riesgos y Controles", "📷 Imágenes", "⚙️ Acciones"])

    with tabs[0]:
        st.markdown("#### Información del Documento")

        meta["moc_title"] = st.text_input("Título:", value=meta.get("moc_title", ""))
        meta["moc_number"] = st.text_input("Número:", value=meta.get("moc_number", ""), disabled=True)
        meta["naturaleza"] = st.selectbox("Naturaleza:", ["permanente", "temporal", "emergencia"], 
                                          index=["permanente", "temporal", "emergencia"].index(meta.get("naturaleza", "permanente")))
        meta["originador"] = st.text_input("Originador:", value=meta.get("originador", ""))
        meta["fecha"] = st.text_input("Fecha:", value=meta.get("fecha", ""), disabled=True)

        st.markdown("#### Equipo de Revisión")
        meta["produccion"] = st.text_input("Producción:", value=meta.get("produccion", ""))
        meta["specialist_shes"] = st.text_input("Specialist SHES:", value=meta.get("specialist_shes", ""))
        meta["mantenimiento"] = st.text_input("Mantenimiento:", value=meta.get("mantenimiento", ""))
        meta["revisores"] = st.text_input("Revisores:", value=meta.get("revisores", ""))
        meta["experto_aprobador"] = st.text_input("Experto Aprobador:", value=meta.get("experto_aprobador", ""))

    with tabs[1]:
        st.markdown("#### Descripción del Problema")
        data["descripcion_problema"] = st.text_area("", value=data.get("descripcion_problema", ""), height=150)

        st.markdown("#### Condición Actual")
        data["condicion_actual"] = st.text_area("", value=data.get("condicion_actual", ""), height=120)

        st.markdown("#### Condición Propuesta")
        data["condicion_propuesta"] = st.text_area("", value=data.get("condicion_propuesta", ""), height=120)

        st.markdown("#### Razones del Cambio")
        data["razones_cambio"] = st.text_area("", value=data.get("razones_cambio", ""), height=120)

        st.markdown("#### Alternativas y Plan de Retorno")
        data["alternativas_retorno"] = st.text_area("", value=data.get("alternativas_retorno", ""), height=150)

        st.markdown("#### Recursos")
        data["recursos"] = st.text_area("", value=data.get("recursos", ""), height=150)

        st.markdown("#### Plan de Implementación")
        data["plan_implementacion"] = st.text_area("", value=data.get("plan_implementacion", ""), height=150)

        st.markdown("#### Tiempo de Duración")
        data["tiempo_duracion"] = st.text_area("", value=data.get("tiempo_duracion", ""), height=80)

    with tabs[2]:
        st.markdown("#### Riesgos y Controles")

        risks = data.get("riesgos_controles", [])
        updated_risks = []

        for i, risk in enumerate(risks):
            with st.container():
                st.markdown(f"**Riesgo {i+1}**")
                col1, col2 = st.columns(2)
                with col1:
                    r_riesgo = st.text_input(f"Riesgo {i+1}:", value=risk.get("riesgo", ""), key=f"riesgo_{i}")
                with col2:
                    r_control = st.text_input(f"Control {i+1}:", value=risk.get("control", ""), key=f"control_{i}")
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
            with st.container():
                st.markdown(f"**Riesgo SHES {i+1}**")
                col1, col2, col3 = st.columns(3)
                with col1:
                    s_riesgo = st.text_input(f"Riesgo S{i+1}:", value=risk.get("riesgo", ""), key=f"shes_riesgo_{i}")
                with col2:
                    s_control = st.text_input(f"Control S{i+1}:", value=risk.get("control", ""), key=f"shes_control_{i}")
                with col3:
                    s_plazo = st.text_input(f"Plazo S{i+1}:", value=risk.get("plazo", ""), key=f"shes_plazo_{i}")
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

        st.markdown("""
        <div style="background: #f0fdf4; border: 1px solid #10b981; border-radius: 10px; padding: 1rem; margin: 1rem 0;">
            <h4 style="color: #065f46; margin: 0;">✅ Documento listo para generar</h4>
            <p style="color: #047857; margin: 0.5rem 0 0 0;">
                Seleccione el formato de salida. El documento se guardará automáticamente en el historial.
            </p>
        </div>
        """, unsafe_allow_html=True)

        col1, col2, col3 = st.columns(3)

        with col1:
            if st.button("🇪🇸 Generar en Español", type="primary", use_container_width=True):
                _finalize_moc(data, meta, images, "es")

        with col2:
            if st.button("🇺🇸 Generar en Inglés", type="primary", use_container_width=True):
                _finalize_moc(data, meta, images, "en")

        with col3:
            if st.button("🔄 Regenerar con IA", use_container_width=True):
                st.session_state.page = "nueva_moc"
                st.rerun()

    # Actualizar session state
    st.session_state.generated_data = data
    st.session_state.moc_meta = meta


def _finalize_moc(data, meta, images, language):
    """Finaliza y genera el documento MoC"""
    with st.spinner(f"📄 Generando documento MoC en {'Inglés' if language == 'en' else 'Español'}..."):
        config = ConfigManager.load()
        gemini = GeminiService(config.get("gemini_api_key", ""), config.get("gemini_model", "gemini-1.5-pro"))

        # Combinar datos
        final_data = {**meta, **data}

        # Traducir si es necesario
        if language == "en":
            st.info("🌐 Traduciendo documento al inglés...")
            final_data = gemini.translate_document(final_data)

        # Generar documento
        generator = DocumentGenerator()
        output_path = generator.generate_moc(final_data, images, config.get("header_text", ""), config.get("footer_text", ""))

        # Guardar en historial
        doc_info = {
            "type": "moc",
            "title": meta.get("moc_title", ""),
            "number": meta.get("moc_number", ""),
            "language": language,
            "filename": output_path.name,
            "path": str(output_path)
        }
        HistoryManager.add_document(doc_info)

        # Mostrar resultado
        st.success(f"✅ Documento generado: {output_path.name}")

        # Botones de descarga
        with open(output_path, "rb") as f:
            st.download_button(
                label="📥 Descargar PowerPoint",
                data=f,
                file_name=output_path.name,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True
            )

        # Intentar generar PDF
        try:
            pdf_path = generator.convert_to_pdf(output_path)
            if pdf_path and pdf_path.exists():
                with open(pdf_path, "rb") as f:
                    st.download_button(
                        label="📄 Descargar PDF",
                        data=f,
                        file_name=pdf_path.name,
                        mime="application/pdf",
                        use_container_width=True
                    )
        except:
            st.info("💡 El PDF se puede generar abriendo el archivo en PowerPoint y exportando")


def render_a3_review():
    """Renderiza pantalla de revision A3"""
    st.markdown("""
    <div class="section-header">
        <h3>👁️ Revisar y Editar Documento A3</h3>
    </div>
    """, unsafe_allow_html=True)

    data = st.session_state.get("generated_data", {})
    meta = st.session_state.get("a3_meta", {})
    images = st.session_state.get("a3_images", [])

    tabs = st.tabs(["📋 Información", "📝 Contenido", "📷 Imágenes", "⚙️ Acciones"])

    with tabs[0]:
        meta["titulo"] = st.text_input("Título:", value=meta.get("titulo", ""))
        meta["area"] = st.text_input("Área:", value=meta.get("area", ""))
        meta["autor"] = st.text_input("Autor:", value=meta.get("autor", ""))
        meta["doc_number"] = st.text_input("Número:", value=meta.get("doc_number", ""), disabled=True)
        meta["fecha"] = st.text_input("Fecha:", value=meta.get("fecha", ""), disabled=True)

    with tabs[1]:
        sections = [
            ("Antecedentes", "antecedentes"),
            ("Problema Actual", "problema_actual"),
            ("Análisis de Situación", "analisis_situacion"),
            ("Objetivos", "objetivos"),
            ("Análisis Causa Raíz", "analisis_causa_raiz"),
            ("Contramedidas", "contramedidas"),
            ("Resultados Esperados", "resultados_esperados"),
            ("Plan de Seguimiento", "plan_seguimiento"),
            ("Lecciones Aprendidas", "lecciones_aprendidas"),
            ("Estandarización", "estandarizacion"),
        ]

        for label, key in sections:
            st.markdown(f"**{label}**")
            data[key] = st.text_area("", value=data.get(key, ""), height=100, key=f"a3_{key}")

    with tabs[2]:
        if images:
            for idx, img_info in enumerate(images, 1):
                st.image(img_info["path"], caption=f"Figura {idx}: {img_info['desc']}", width=400)
        else:
            st.info("No se cargaron imágenes")

    with tabs[3]:
        col1, col2 = st.columns(2)
        with col1:
            if st.button("✅ Generar Documento A3", type="primary", use_container_width=True):
                _finalize_a3(data, meta, images)
        with col2:
            if st.button("🔄 Regenerar", use_container_width=True):
                st.session_state.page = "nueva_a3"
                st.rerun()

    st.session_state.generated_data = data
    st.session_state.a3_meta = meta


def _finalize_a3(data, meta, images):
    """Finaliza y genera documento A3"""
    with st.spinner("📄 Generando documento A3..."):
        config = ConfigManager.load()

        final_data = {**meta, **data}

        generator = DocumentGenerator()
        output_path = generator.generate_a3(final_data, images)

        doc_info = {
            "type": "a3",
            "title": meta.get("titulo", ""),
            "number": meta.get("doc_number", ""),
            "language": "es",
            "filename": output_path.name,
            "path": str(output_path)
        }
        HistoryManager.add_document(doc_info)

        st.success(f"✅ Documento generado: {output_path.name}")

        with open(output_path, "rb") as f:
            st.download_button(
                label="📥 Descargar Word",
                data=f,
                file_name=output_path.name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

        # PDF
        try:
            pdf_path = generator.convert_to_pdf(output_path)
            if pdf_path and pdf_path.exists():
                with open(pdf_path, "rb") as f:
                    st.download_button(
                        label="📄 Descargar PDF",
                        data=f,
                        file_name=pdf_path.name,
                        mime="application/pdf",
                        use_container_width=True
                    )
        except:
            st.info("💡 El PDF se puede generar abriendo el archivo en Word y exportando")


def render_kaizen_review():
    """Renderiza pantalla de revision Kaizen"""
    st.markdown("""
    <div class="section-header">
        <h3>👁️ Revisar y Editar Documento Kaizen</h3>
    </div>
    """, unsafe_allow_html=True)

    data = st.session_state.get("generated_data", {})
    meta = st.session_state.get("kaizen_meta", {})
    images = st.session_state.get("kaizen_images", [])

    tabs = st.tabs(["📋 Información", "📝 Contenido", "📷 Imágenes", "⚙️ Acciones"])

    with tabs[0]:
        meta["titulo"] = st.text_input("Título:", value=meta.get("titulo", ""))
        meta["area"] = st.text_input("Área:", value=meta.get("area", ""))
        meta["autor"] = st.text_input("Autor:", value=meta.get("autor", ""))
        meta["doc_number"] = st.text_input("Número:", value=meta.get("doc_number", ""), disabled=True)
        meta["fecha"] = st.text_input("Fecha:", value=meta.get("fecha", ""), disabled=True)

    with tabs[1]:
        sections = [
            ("Descripción del Problema", "descripcion_problema"),
            ("Solución Implementada", "solucion"),
            ("Beneficios", "beneficios"),
            ("Tipo de Desperdicio", "tipo_desperdicio"),
            ("Impacto BTO", "impacto_bto"),
            ("Próximos Pasos", "proximos_pasos"),
        ]

        for label, key in sections:
            st.markdown(f"**{label}**")
            data[key] = st.text_area("", value=data.get(key, ""), height=100, key=f"kzn_{key}")

    with tabs[2]:
        if images:
            for idx, img_info in enumerate(images, 1):
                st.image(img_info["path"], caption=f"Figura {idx}: {img_info['desc']}", width=400)
        else:
            st.info("No se cargaron imágenes")

    with tabs[3]:
        col1, col2 = st.columns(2)
        with col1:
            if st.button("✅ Generar Documento Kaizen", type="primary", use_container_width=True):
                _finalize_kaizen(data, meta, images)
        with col2:
            if st.button("🔄 Regenerar", use_container_width=True):
                st.session_state.page = "nuevo_kaizen"
                st.rerun()

    st.session_state.generated_data = data
    st.session_state.kaizen_meta = meta


def _finalize_kaizen(data, meta, images):
    """Finaliza y genera documento Kaizen"""
    with st.spinner("📄 Generando documento Kaizen..."):
        config = ConfigManager.load()

        final_data = {**meta, **data}

        generator = DocumentGenerator()
        output_path = generator.generate_kaizen(final_data, images)

        doc_info = {
            "type": "kaizen",
            "title": meta.get("titulo", ""),
            "number": meta.get("doc_number", ""),
            "language": "es",
            "filename": output_path.name,
            "path": str(output_path)
        }
        HistoryManager.add_document(doc_info)

        st.success(f"✅ Documento generado: {output_path.name}")

        with open(output_path, "rb") as f:
            st.download_button(
                label="📥 Descargar PowerPoint",
                data=f,
                file_name=output_path.name,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True
            )

        try:
            pdf_path = generator.convert_to_pdf(output_path)
            if pdf_path and pdf_path.exists():
                with open(pdf_path, "rb") as f:
                    st.download_button(
                        label="📄 Descargar PDF",
                        data=f,
                        file_name=pdf_path.name,
                        mime="application/pdf",
                        use_container_width=True
                    )
        except:
            st.info("💡 El PDF se puede generar abriendo el archivo en PowerPoint y exportando")


# =============================================================================
# HISTORIAL Y CONFIGURACION
# =============================================================================

def render_history():
    """Renderiza pantalla de historial de documentos"""
    st.markdown("""
    <div class="section-header">
        <h3>📁 Historial de Documentos Generados</h3>
    </div>
    """, unsafe_allow_html=True)

    history = HistoryManager.load()
    docs = history.get("documents", [])

    if not docs:
        st.info("📭 No hay documentos generados aún. Cree su primer documento desde el menú principal.")
        return

    # Filtros
    col1, col2, col3 = st.columns(3)
    with col1:
        filter_type = st.selectbox("Filtrar por tipo:", 
                                   ["Todos", "MoC", "A3", "Kaizen"],
                                   key="filter_type")
    with col2:
        filter_lang = st.selectbox("Idioma:", 
                                   ["Todos", "Español", "Inglés"],
                                   key="filter_lang")
    with col3:
        search = st.text_input("Buscar:", placeholder="Título o número...", key="search_docs")

    # Filtrar
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

    # Mostrar documentos
    st.markdown(f"**Mostrando {len(filtered)} documento(s)**")

    for doc in filtered:
        type_emoji = {"moc": "📋", "a3": "📊", "kaizen": "⚡"}.get(doc.get("type"), "📄")
        type_label = {"moc": "MoC", "a3": "A3", "kaizen": "Kaizen"}.get(doc.get("type"), "Doc")
        lang_flag = "🇪🇸" if doc.get("language") == "es" else "🇺🇸"

        with st.container():
            st.markdown(f"""
            <div class="history-item">
                <div style="display: flex; justify-content: space-between; align-items: center;">
                    <div>
                        <h4 style="margin: 0; color: #1e293b;">{type_emoji} {doc.get('title', 'Sin título')}</h4>
                        <p style="margin: 0.25rem 0; color: #64748b; font-size: 0.9rem;">
                            {type_label} · {doc.get('number', '')} · {lang_flag} · {doc.get('timestamp', '')[:10]}
                        </p>
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            col1, col2, col3 = st.columns([2, 2, 1])

            doc_path = Path(doc.get("path", ""))
            if doc_path.exists():
                with col1:
                    with open(doc_path, "rb") as f:
                        mime = "application/vnd.openxmlformats-officedocument.presentationml.presentation" if doc_path.suffix == ".pptx" else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        st.download_button(
                            label=f"📥 Descargar {doc_path.suffix.upper()[1:]}",
                            data=f,
                            file_name=doc_path.name,
                            mime=mime,
                            key=f"dl_{doc.get('id', 'x')}",
                            use_container_width=True
                        )

                # Intentar PDF
                pdf_path = doc_path.with_suffix('.pdf')
                if pdf_path.exists():
                    with col2:
                        with open(pdf_path, "rb") as f:
                            st.download_button(
                                label="📄 Descargar PDF",
                                data=f,
                                file_name=pdf_path.name,
                                mime="application/pdf",
                                key=f"dl_pdf_{doc.get('id', 'x')}",
                                use_container_width=True
                            )
                else:
                    with col2:
                        st.button("📄 PDF no disponible", disabled=True, 
                                 key=f"no_pdf_{doc.get('id', 'x')}", use_container_width=True)

            with col3:
                if st.button("🗑️", key=f"del_{doc.get('id', 'x')}", use_container_width=True):
                    # Eliminar del historial
                    history["documents"] = [d for d in history["documents"] if d.get("id") != doc.get("id")]
                    HistoryManager.save(history)
                    # Eliminar archivo
                    if doc_path.exists():
                        doc_path.unlink()
                    st.rerun()


def render_settings():
    """Renderiza pantalla de configuracion"""
    st.markdown("""
    <div class="section-header">
        <h3>⚙️ Configuración del Sistema</h3>
    </div>
    """, unsafe_allow_html=True)

    config = ConfigManager.load()

    tabs = st.tabs(["🔑 API Gemini", "🏢 Empresa", "📄 Templates", "📐 Header/Footer"])

    with tabs[0]:
        st.markdown("#### Configuración de API Gemini")

        st.info("""
        💡 **Cómo obtener API Key:**
        1. Visite [Google AI Studio](https://aistudio.google.com/)
        2. Cree una cuenta o inicie sesión
        3. Genere una API Key gratuita
        4. Péguela aquí

        Sin API Key, el software funciona con generación local básica.
        """)

        api_key = st.text_input("API Key Gemini:", 
                                value=config.get("gemini_api_key", ""),
                                type="password",
                                help="Su clave privada de API de Google AI Studio")

        st.markdown("#### Selección de Modelo")
        st.markdown("**Seleccione la versión de Gemini a utilizar:**")

        model_options = {
            "gemini-1.5-flash-lite": "⚡ 3.1 Flash-Lite - Respuestas rápidas",
            "gemini-1.5-flash": "🔥 3.5 Flash - Ayuda completa",
            "gemini-1.5-pro": "🧠 3.1 Pro - Advanced math and code (Recomendado)",
        }

        current_model = config.get("gemini_model", "gemini-1.5-pro")

        col1, col2, col3 = st.columns(3)

        with col1:
            st.markdown("""
            <div style="background: #f8fafc; border: 2px solid #e2e8f0; border-radius: 12px; padding: 1rem; text-align: center;">
                <div style="font-size: 2rem;">⚡</div>
                <h4 style="margin: 0.5rem 0; color: #1e293b;">3.1 Flash-Lite</h4>
                <p style="color: #64748b; font-size: 0.85rem; margin: 0;">Respuestas rápidas</p>
                <span style="background: #dbeafe; color: #1d4ed8; padding: 0.15rem 0.5rem; border-radius: 10px; font-size: 0.75rem;">Nuevo</span>
            </div>
            """, unsafe_allow_html=True)
            if st.button("Seleccionar Flash-Lite", key="sel_flash_lite", use_container_width=True,
                        type="secondary" if current_model != "gemini-1.5-flash-lite" else "primary"):
                config["gemini_model"] = "gemini-1.5-flash-lite"
                ConfigManager.save(config)
                st.success("✅ Modelo actualizado a Flash-Lite")
                st.rerun()

        with col2:
            st.markdown("""
            <div style="background: #f8fafc; border: 2px solid #e2e8f0; border-radius: 12px; padding: 1rem; text-align: center;">
                <div style="font-size: 2rem;">🔥</div>
                <h4 style="margin: 0.5rem 0; color: #1e293b;">3.5 Flash</h4>
                <p style="color: #64748b; font-size: 0.85rem; margin: 0;">Ayuda completa</p>
                <span style="background: #dbeafe; color: #1d4ed8; padding: 0.15rem 0.5rem; border-radius: 10px; font-size: 0.75rem;">Nuevo</span>
            </div>
            """, unsafe_allow_html=True)
            if st.button("Seleccionar Flash", key="sel_flash", use_container_width=True,
                        type="secondary" if current_model != "gemini-1.5-flash" else "primary"):
                config["gemini_model"] = "gemini-1.5-flash"
                ConfigManager.save(config)
                st.success("✅ Modelo actualizado a Flash")
                st.rerun()

        with col3:
            is_pro = current_model == "gemini-1.5-pro"
            st.markdown(f"""
            <div style="background: {'#eff6ff' if is_pro else '#f8fafc'}; border: 2px solid {'#1a5f7a' if is_pro else '#e2e8f0'}; border-radius: 12px; padding: 1rem; text-align: center;">
                <div style="font-size: 2rem;">🧠</div>
                <h4 style="margin: 0.5rem 0; color: #1e293b;">3.1 Pro</h4>
                <p style="color: #64748b; font-size: 0.85rem; margin: 0;">Advanced math and code</p>
                {'<div style="color: #1a5f7a; font-weight: bold; margin-top: 0.5rem;">✓ Seleccionado</div>' if is_pro else ''}
            </div>
            """, unsafe_allow_html=True)
            if st.button("Seleccionar Pro", key="sel_pro", use_container_width=True,
                        type="secondary" if not is_pro else "primary"):
                config["gemini_model"] = "gemini-1.5-pro"
                ConfigManager.save(config)
                st.success("✅ Modelo actualizado a Pro")
                st.rerun()

        # Nivel de pensamiento
        st.markdown("#### Nivel de Pensamiento")
        thinking_level = st.select_slider(
            "Profundidad de razonamiento:",
            options=["Básico", "Estándar", "Profundo"],
            value=config.get("thinking_level", "Estándar"),
            help="Afecta la calidad y profundidad del contenido generado"
        )

        # Nivel de correccion ortografica
        st.markdown("#### Corrección Ortográfica")
        spell_check = st.toggle("Corrección automática activa", 
                                value=config.get("spell_check", True),
                                help="Corrige automáticamente ortografía, gramática y puntuación")

        if st.button("💾 Guardar Configuración API", type="primary", use_container_width=True):
            config["gemini_api_key"] = api_key
            config["spell_check"] = spell_check
            config["thinking_level"] = thinking_level
            ConfigManager.save(config)
            st.success("✅ Configuración guardada correctamente")

    with tabs[1]:
        st.markdown("#### Datos de la Empresa")

        company_name = st.text_input("Nombre de la Empresa:", 
                                     value=config.get("company_name", ""),
                                     placeholder="Ej: Orica Peru")
        department = st.text_input("Departamento:", 
                                   value=config.get("department", ""),
                                   placeholder="Ej: Mantenimiento")
        default_author = st.text_input("Autor por defecto:", 
                                         value=config.get("default_author", ""),
                                         placeholder="Nombre completo del autor")
        default_area = st.text_input("Área por defecto:", 
                                       value=config.get("default_area", ""),
                                       placeholder="Ej: Planta Lurín")

        if st.button("💾 Guardar Datos Empresa", type="primary", use_container_width=True):
            config["company_name"] = company_name
            config["department"] = department
            config["default_author"] = default_author
            config["default_area"] = default_area
            ConfigManager.save(config)
            st.success("✅ Datos guardados correctamente")

    with tabs[2]:
        st.markdown("#### Rutas de Templates")

        st.info("📄 Los templates son los formatos oficiales de la empresa. No modifique su estructura.")

        template_moc = st.text_input("Template MoC (.pptx):", 
                                     value=str(TEMPLATES_DIR / "moc_template.pptx"),
                                     disabled=True)
        template_a3 = st.text_input("Template A3 (.docx):", 
                                    value=str(TEMPLATES_DIR / "a3_template.docx"),
                                    disabled=True)
        template_kaizen = st.text_input("Template Kaizen (.pptx):", 
                                        value=str(TEMPLATES_DIR / "kaizen_template.pptx"),
                                        disabled=True)

        st.markdown("#### Subir Nuevos Templates")

        col1, col2, col3 = st.columns(3)
        with col1:
            new_moc = st.file_uploader("Nuevo Template MoC", type=["pptx"], key="new_moc_tpl")
            if new_moc:
                with open(TEMPLATES_DIR / "moc_template.pptx", "wb") as f:
                    f.write(new_moc.getbuffer())
                st.success("✅ Template MoC actualizado")

        with col2:
            new_a3 = st.file_uploader("Nuevo Template A3", type=["docx"], key="new_a3_tpl")
            if new_a3:
                with open(TEMPLATES_DIR / "a3_template.docx", "wb") as f:
                    f.write(new_a3.getbuffer())
                st.success("✅ Template A3 actualizado")

        with col3:
            new_kzn = st.file_uploader("Nuevo Template Kaizen", type=["pptx"], key="new_kzn_tpl")
            if new_kzn:
                with open(TEMPLATES_DIR / "kaizen_template.pptx", "wb") as f:
                    f.write(new_kzn.getbuffer())
                st.success("✅ Template Kaizen actualizado")

    with tabs[3]:
        st.markdown("#### Encabezado y Pie de Página")

        header_text = st.text_area("Texto de Encabezado (aparece en cada página):", 
                                   value=config.get("header_text", ""),
                                   height=80,
                                   placeholder="Ej: CONFIDENCIAL - USO INTERNO")
        footer_text = st.text_area("Texto de Pie de Página:", 
                                   value=config.get("footer_text", ""),
                                   height=80,
                                   placeholder="Ej: CAVA - Especialistas en Robótica y Automatización")

        if st.button("💾 Guardar Header/Footer", type="primary", use_container_width=True):
            config["header_text"] = header_text
            config["footer_text"] = footer_text
            ConfigManager.save(config)
            st.success("✅ Configuración guardada")


# =============================================================================
# FUNCION PRINCIPAL DE LA APLICACION
# =============================================================================

def main():
    """Funcion principal de la aplicacion Streamlit"""

    # Inicializar session state
    if "page" not in st.session_state:
        st.session_state.page = "inicio"

    # Renderizar sidebar
    render_sidebar()

    # Renderizar contenido segun pagina
    page = st.session_state.page

    if page == "inicio":
        render_welcome()
    elif page == "nueva_moc":
        render_moc_form()
    elif page == "revisar_moc":
        render_moc_review()
    elif page == "nueva_a3":
        render_a3_form()
    elif page == "revisar_a3":
        render_a3_review()
    elif page == "nuevo_kaizen":
        render_kaizen_form()
    elif page == "revisar_kaizen":
        render_kaizen_review()
    elif page == "historial":
        render_history()
    elif page == "configuracion":
        render_settings()
    else:
        render_welcome()

    # Footer global
    st.markdown("""
    <div class="app-footer">
        <p><strong>CAVA</strong> - Especialistas en Robótica y Automatización</p>
        <p>Diseñado por Roger Huamani | Sistema de Gestión Documental v2.0.0</p>
        <p style="font-size: 0.75rem; color: #94a3b8;">
            Software empresarial para automatización de documentos MoC, A3 y Kaizen.<br>
            Mantiene los formatos oficiales de la empresa sin modificaciones.
        </p>
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
