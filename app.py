#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
================================================================================
SISTEMA DE GESTION DOCUMENTAL - MoC | Mejora A3 | Simple Kaizen
Version Streamlit Cloud - 100% Compatible Web
================================================================================
Diseñado por: CAVA - Especialistas en Robotica y Automatizacion
Desarrollador: Roger Huamani
Version: 3.0.0
Fecha: Mayo 2026
================================================================================
NOTA: Este software esta optimizado para ejecutarse en Streamlit Cloud.
Los templates se cargan via interfaz web la primera vez.
La configuracion se persiste en session_state + archivo descargable.
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
from copy import deepcopy

# =============================================================================
# CONFIGURACION INICIAL DE PAGINA
# =============================================================================
st.set_page_config(
    page_title="Gestion Documental - MoC | A3 | Kaizen",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

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

    .spell-hint {
        background: #fef3c7;
        border: 1px solid #f59e0b;
        border-radius: 8px;
        padding: 0.75rem 1rem;
        margin: 0.5rem 0;
        font-size: 13px;
        color: #92400e;
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
# INICIALIZACION DE SESSION STATE - PERSISTENCIA EN MEMORIA
# =============================================================================
def init_session_state():
    """Inicializa todas las variables de session_state"""
    defaults = {
        "page": "inicio",
        "config": {
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
            "thinking_level": "Estándar"
        },
        "history": {"documents": []},
        "generated_data": {},
        "doc_meta": {},
        "doc_images": [],
        "doc_type": None,
        "templates_uploaded": False,
        "template_moc_bytes": None,
        "template_a3_bytes": None,
        "template_kaizen_bytes": None,
    }

    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

init_session_state()

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
        now = datetime.now()
        prefix = {"moc": "MOC", "a3": "A3", "kaizen": "KZN"}
        return f"{prefix.get(doc_type, 'DOC')}-{now.year}{now.month:02d}{now.day:02d}-{config[key]:04d}"

    @staticmethod
    def sanitize_filename(filename):
        return re.sub(r'[<>:"/\\|?*]', '_', filename)[:50]

    @staticmethod
    def add_to_history(doc_info):
        """Agrega documento al historial en session_state"""
        history = st.session_state.history
        doc_info["id"] = str(uuid.uuid4())
        doc_info["timestamp"] = datetime.now().isoformat()
        history["documents"].insert(0, doc_info)
        st.session_state.history = history

    @staticmethod
    def get_history(doc_type=None):
        """Obtiene historial filtrado"""
        docs = st.session_state.history.get("documents", [])
        if doc_type:
            docs = [d for d in docs if d.get("type") == doc_type]
        return docs


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

        prompt = f"""Eres un experto senior en gestion de cambios industriales (MoC) con 20 años de experiencia. Redactas documentos profesionales, impecables, sin errores ortograficos.

PROBLEMA: {problem}
CONTEXTO: {context}
EQUIPO: {equipo}

Genera en ESPANOL formato JSON con:
1. descripcion_problema (parrafo bien estructurado)
2. condicion_actual
3. condicion_propuesta
4. razones_cambio
5. alternativas_retorno
6. recursos
7. plan_implementacion
8. tiempo_duracion
9. riesgos_controles: array de {{"riesgo":"...","control":"..."}} (min 3)
10. riesgos_shes: array de {{"riesgo":"...","control":"...","plazo":"..."}} (min 3)

Responde SOLO JSON valido."""

        try:
            text = self._call_api(prompt, temperature=0.2, max_tokens=8192)
            return self._extract_json(text)
        except Exception as e:
            st.error(f"Error API: {e}. Usando generacion local.")
            return self._generate_local_moc(problem, context, equipo)

    def generate_a3(self, problem, context=""):
        if not self.api_key:
            return self._generate_local_a3(problem, context)

        prompt = f"""Eres experto en metodologia A3 Lean. PROBLEMA: {problem} CONTEXTO: {context}
Genera JSON con: titulo, antecedentes, problema_actual, analisis_situacion, objetivos, analisis_causa_raiz, contramedidas, resultados_esperados, plan_seguimiento, lecciones_aprendidas, estandarizacion. SOLO JSON."""

        try:
            text = self._call_api(prompt, temperature=0.2, max_tokens=8192)
            return self._extract_json(text)
        except:
            return self._generate_local_a3(problem, context)

    def generate_kaizen(self, activity, context=""):
        if not self.api_key:
            return self._generate_local_kaizen(activity, context)

        prompt = f"""Eres experto en Kaizen. ACTIVIDAD: {activity} CONTEXTO: {context}
Genera JSON con: titulo, area, descripcion_problema, solucion, beneficios, tipo_desperdicio, impacto_bto, proximos_pasos. SOLO JSON."""

        try:
            text = self._call_api(prompt, temperature=0.2, max_tokens=4096)
            return self._extract_json(text)
        except:
            return self._generate_local_kaizen(activity, context)

    def translate_document(self, data):
        if not self.api_key:
            return data
        prompt = f"""Traduce del espanol al ingles profesional industrial:
{json.dumps(data, ensure_ascii=False, indent=2)}
Responde SOLO el JSON traducido, misma estructura."""
        try:
            text = self._call_api(prompt, temperature=0.1, max_tokens=8192)
            return self._extract_json(text)
        except:
            return data

    def correct_spelling(self, text):
        if not self.api_key or not text.strip():
            return text
        prompt = f"""Corrige ortografia, gramatica y puntuacion del siguiente texto en espanol. Manten el significado. Devuelve SOLO el texto corregido.

TEXTO:
{text}"""
        try:
            return self._call_api(prompt, temperature=0.1, max_tokens=4096).strip()
        except:
            return text

    def _generate_local_moc(self, problem, context, equipo):
        return {
            "descripcion_problema": f"Se identifico la siguiente situacion que requiere gestion mediante MoC: {problem}. Esta condicion presenta riesgos operacionales que deben ser mitigados mediante un cambio controlado y documentado, garantizando la integridad del proceso y la seguridad del personal.",
            "condicion_actual": "El equipo o proceso actual opera bajo condiciones que presentan limitaciones tecnicas documentadas. Se requiere evaluacion detallada para establecer la linea base de referencia antes de la modificacion.",
            "condicion_propuesta": "Implementar las modificaciones tecnicas necesarias para optimizar el rendimiento operativo, mejorar la seguridad del proceso y alinear las condiciones con los estandares corporativos vigentes.",
            "razones_cambio": "1. Mejora de la seguridad operacional y reduccion de riesgos para el personal.\n2. Optimizacion del rendimiento y confiabilidad del equipo critico.\n3. Cumplimiento de estandares corporativos y regulatorios.\n4. Reduccion de riesgos identificados en evaluaciones previas de SHES.\n5. Mejora continua alineada con objetivos estrategicos.",
            "alternativas_retorno": "Alternativas evaluadas:\n1. Mantenimiento correctivo tradicional (DESCARTADO): Alcance limitado.\n2. Reemplazo total del sistema (DESCARTADO): Costo elevado.\n3. Modificacion controlada (SELECCIONADA): Mejor relacion costo-beneficio.\n\nPlan de retorno: Restauracion de configuracion original, protocolo de contingencia, notificacion a supervision, registro de lecciones aprendidas.",
            "recursos": "RECURSOS HUMANOS:\n- Tecnico especializado de mantenimiento\n- Supervisor de area\n- Especialista SHES\n- Operador de area\n\nRECURSOS MATERIALES:\n- Herramientas especializadas certificadas\n- Repuestos de calidad certificada\n- EPP completo: casco, gafas, guantes, botas\n- Materiales de senalizacion\n\nRECURSOS TECNICOS:\n- Documentacion tecnica actualizada\n- SOP vigentes\n- Permisos de trabajo segun aplique\n- Checklist de verificacion",
            "plan_implementacion": "FASE 1: PREPARACION (Dias 1-2)\n- Reunion de coordinacion con produccion\n- Verificacion de disponibilidad de recursos\n- Preparacion del area: limpieza, senalizacion, LOTO\n- Briefing de seguridad\n\nFASE 2: EJECUCION (Dias 3-5)\n- Implementacion de modificaciones\n- Pruebas funcionales iniciales\n- Registro fotografico detallado\n- Verificacion intermedia SHES\n\nFASE 3: VALIDACION (Dias 6-7)\n- Pruebas bajo condiciones normales\n- Verificacion de parametros criticos\n- Validacion por supervisor y produccion\n\nFASE 4: CIERRE (Dia 8)\n- Actualizacion de documentacion\n- Entrenamiento al personal\n- Socializacion de lecciones aprendidas\n- Cierre formal con firmas",
            "tiempo_duracion": "Estimacion total: 8 dias habiles distribuidos en 4 fases. Puede ajustarse segun condiciones operativas.",
            "riesgos_controles": [
                {"riesgo": "Interrupcion del proceso productivo durante la implementacion", "control": "Coordinacion previa con produccion para ventana de mantenimiento planificado. Comunicacion oportuna a areas involucradas."},
                {"riesgo": "Falla tecnica durante modificacion que afecte equipos adyacentes", "control": "Verificacion previa exhaustiva. Disponibilidad de repuestos de emergencia. Supervision tecnica continua."},
                {"riesgo": "Exposicion a riesgos SHES durante ejecucion de trabajos", "control": "Permisos de trabajo segun tipo. Uso obligatorio de EPP. Supervision continua SHES. Aplicacion de LOTO."}
            ],
            "riesgos_shes": [
                {"riesgo": "Lesiones por manipulacion manual de equipos y componentes", "control": "Capacitacion en tecnicas de levantamiento seguro. EPP completo. Senalizacion del area.", "plazo": "Antes del inicio"},
                {"riesgo": "Generacion de residuos solidos y liquidos", "control": "Manejo segun procedimiento ambiental. Clasificacion en origen. Disposicion en areas autorizadas.", "plazo": "Durante ejecucion"},
                {"riesgo": "Exposicion a ruido, vibraciones o agentes quimicos", "control": "Monitoreo continuo. Protectores auditivos. Ventilacion adecuada. Limitacion de horario.", "plazo": "Durante ejecucion"}
            ]
        }

    def _generate_local_a3(self, problem, context):
        return {
            "titulo": "Mejora del proceso: " + problem[:50],
            "antecedentes": "Se identifico una oportunidad de mejora significativa que requiere analisis estructurado mediante metodologia A3 de resolucion de problemas.",
            "problema_actual": problem,
            "analisis_situacion": "La situacion actual presenta indicadores de desempeno con oportunidades de mejora. Se requiere recopilacion sistematica de datos para establecer linea base solida.",
            "objetivos": "OBJETIVO GENERAL:\nOptimizar el proceso eliminando desperdicios identificados.\n\nOBJETIVOS ESPECIFICOS:\n- Reducir tiempo de ciclo en 15% (3 meses)\n- Disminuir tasa de defectos en 20% (trimestre)\n- Mejorar productividad en 10% (6 meses)\n- Incrementar satisfaccion del cliente en 25%",
            "analisis_causa_raiz": "ANALISIS 5 PORQUES:\n1. POR QUE ocurre? → Condicion operativa inadecuada.\n2. POR QUE inadecuada? → Falta estandarizacion formal.\n3. POR QUE falta? → Procedimientos no actualizados.\n4. POR QUE no actualizados? → No hay sistema de gestion documental.\n5. POR QUE no hay sistema? → Falta politica clara de gestion del conocimiento.\n\nCAUSA RAIZ: Ausencia de sistema de gestion, actualizacion y control de SOP.",
            "contramedidas": "1. Actualizar SOP del proceso con instrucciones claras.\n2. Implementar checklists de verificacion diaria.\n3. Capacitar personal en nuevos estandares.\n4. Establecer KPIs visuales en el area.\n5. Programar auditorias mensuales de cumplimiento.",
            "resultados_esperados": "- Reduccion medible de desperdicios\n- Mejora sostenida en calidad y productividad\n- Estandarizacion efectiva del proceso\n- Reduccion de variabilidad operativa\n- Incremento en satisfaccion del cliente\n- Cultura de mejora continua fortalecida",
            "plan_seguimiento": "SEMANA 1-2: Implementacion de contramedidas y capacitacion.\nSEMANA 3-4: Monitoreo inicial y ajustes.\nMES 2: Primera auditoria formal.\nMES 3: Evaluacion integral vs. objetivos.\nMES 6: Revision de sostenibilidad y replicacion.",
            "lecciones_aprendidas": "La metodologia A3 permitio visualizar integralmente el problema y soluciones. La participacion multidisciplinaria fue fundamental para identificar la causa raiz real.",
            "estandarizacion": "Los procedimientos seran documentados, aprobados por gerencia, socializados mediante capacitaciones, integrados al SGC y con revision periodica anual."
        }

    def _generate_local_kaizen(self, activity, context):
        return {
            "titulo": "Kaizen: " + activity[:50],
            "area": "Area de Mantenimiento / Produccion / Calidad",
            "descripcion_problema": activity,
            "solucion": "Se implemento mejora orientada a eliminar desperdicio identificado y optimizar flujo del proceso, aplicando principios Lean Manufacturing y pensamiento Kaizen.",
            "beneficios": "- Reduccion del tiempo de ejecucion\n- Mejora en calidad y consistencia\n- Mayor seguridad para el personal\n- Reduccion de costos operativos\n- Mejora en ambiente de trabajo\n- Eliminacion de movimientos innecesarios",
            "tipo_desperdicio": "Motion / Waiting / Skills (seleccionar segun analisis)",
            "impacto_bto": "Supply Chain and Manufacturing Excellence / Safe and Sustainable (seleccionar segun impacto)",
            "proximos_pasos": "1. Documentar formalmente la mejora\n2. Socializar con otras areas relacionadas\n3. Replicar en procesos similares\n4. Establecer monitoreo mensual\n5. Reconocer al equipo participante"
        }


# =============================================================================
# GENERADOR DE DOCUMENTOS - USANDO BYTES DE TEMPLATES EN MEMORIA
# =============================================================================
class DocumentGenerator:
    """Genera documentos usando templates cargados en session_state (bytes)"""

    def generate_moc(self, data, images=None, template_bytes=None):
        """Genera MoC desde template en memoria"""
        from pptx import Presentation
        from pptx.util import Inches, Pt
        from pptx.dml.color import RGBColor
        from pptx.enum.text import PP_ALIGN
        import io

        if template_bytes is None:
            st.error("❌ Template MoC no cargado. Vaya a Configuracion > Templates.")
            return None

        prs = Presentation(io.BytesIO(template_bytes))

        def fill_cell(cell, text):
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

            if "MOC:" in text:
                for para in shape.text_frame.paragraphs:
                    para.clear()
                p = shape.text_frame.paragraphs[0]
                p.text = f"MOC: {data.get('moc_title', '')}"
                p.font.size = Pt(24)
                p.font.bold = True
                p.font.color.rgb = RGBColor(0x1a, 0x5f, 0x7a)
            elif any(m in text for m in ["MAYO", "ENERO", "FEBRERO", "MARZO", "ABRIL", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]):
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
            if shape.has_text_frame:
                text = shape.text_frame.text
                if text.strip() == "3" or "Descripción del Problema" in text or "Descripcion" in text:
                    for para in shape.text_frame.paragraphs:
                        para.clear()
                    p = shape.text_frame.paragraphs[0]
                    p.text = data.get('descripcion_problema', '')
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
                    fill_cell(table.cell(1, 0), data.get('condicion_actual', ''))
                    fill_cell(table.cell(1, 1), data.get('condicion_propuesta', ''))

        # SLIDE 5: Razones del Cambio
        slide5 = prs.slides[4]
        for shape in slide5.shapes:
            if not shape.has_text_frame:
                continue
            text = shape.text_frame.text
            if text.strip() == "5":
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

        # SLIDE 7: Tiempo
        slide7 = prs.slides[6]
        for shape in slide7.shapes:
            if shape.has_text_frame:
                text = shape.text_frame.text
                if text.strip() == "7" or "Tiempo" in text:
                    for para in shape.text_frame.paragraphs:
                        para.clear()
                    p = shape.text_frame.paragraphs[0]
                    p.text = data.get('tiempo_duracion', '')
                    p.font.size = Pt(14)
                    p.alignment = PP_ALIGN.LEFT
                    p.word_wrap = True
                    break

        # SLIDE 8: Riesgos (Tabla)
        slide8 = prs.slides[7]
        for shape in slide8.shapes:
            if shape.has_table:
                table = shape.table
                risks = data.get('riesgos_controles', [])
                for i, risk in enumerate(risks):
                    row_idx = i + 1
                    if row_idx < len(table.rows):
                        fill_cell(table.cell(row_idx, 0), str(i + 1))
                        fill_cell(table.cell(row_idx, 1), risk.get('riesgo', ''))
                        fill_cell(table.cell(row_idx, 2), risk.get('control', ''))

        # SLIDE 9: Riesgos SHES
        slide9 = prs.slides[8]
        for shape in slide9.shapes:
            if shape.has_table:
                table = shape.table
                risks = data.get('riesgos_shes', [])
                for i, risk in enumerate(risks):
                    row_idx = i + 1
                    if row_idx < len(table.rows):
                        fill_cell(table.cell(row_idx, 0), str(i + 1))
                        fill_cell(table.cell(row_idx, 1), risk.get('riesgo', ''))
                        fill_cell(table.cell(row_idx, 2), risk.get('control', ''))
                        fill_cell(table.cell(row_idx, 3), risk.get('plazo', ''))

        # Agregar imagenes
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

        # Guardar en memoria
        output_buffer = io.BytesIO()
        prs.save(output_buffer)
        output_buffer.seek(0)

        return output_buffer

    def generate_a3(self, data, images=None, template_bytes=None):
        """Genera A3 desde template en memoria"""
        from docx import Document
        from docx.shared import Inches, Pt, RGBColor
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        import io

        if template_bytes is None:
            st.error("❌ Template A3 no cargado. Vaya a Configuracion > Templates.")
            return None

        doc = Document(io.BytesIO(template_bytes))

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

        # Agregar contenido estructurado
        doc.add_page_break()

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

        # Imagenes
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

                    desc = img_info.get("desc", f"Figura {idx}") if isinstance(img_info, dict) else f"Figura {idx}"
                    caption = doc.add_paragraph()
                    caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    cap_run = caption.add_run(desc)
                    cap_run.italic = True
                    cap_run.font.size = Pt(10)
                    cap_run.font.color.rgb = RGBColor(0x47, 0x55, 0x69)
                except Exception as e:
                    doc.add_paragraph(f"[Error imagen {idx}: {e}]")

        output_buffer = io.BytesIO()
        doc.save(output_buffer)
        output_buffer.seek(0)

        return output_buffer

    def generate_kaizen(self, data, images=None, template_bytes=None):
        """Genera Kaizen desde template en memoria"""
        from pptx import Presentation
        from pptx.util import Inches, Pt
        from pptx.dml.color import RGBColor
        from pptx.enum.text import PP_ALIGN
        import io

        if template_bytes is None:
            st.error("❌ Template Kaizen no cargado. Vaya a Configuracion > Templates.")
            return None

        prs = Presentation(io.BytesIO(template_bytes))

        slide = prs.slides[0]

        # Llenar tablas
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
                    for para in shape.text_frame.paragraphs:
                        para.clear()
                    p = shape.text_frame.paragraphs[0]
                    p.text = f"Leader: {data.get('autor', '')}\nArea: {data.get('area', '')}\nFecha: {data.get('fecha', '')}"
                    p.font.size = Pt(12)

        # Slide de detalle
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

        # Imagenes
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

        output_buffer = io.BytesIO()
        prs.save(output_buffer)
        output_buffer.seek(0)

        return output_buffer


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
        <p>v3.0.0 · Mayo 2026</p>
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
        <p>Automatización inteligente de documentos MoC, Mejora A3 y Simple Kaizen</p>
    </div>
    """, unsafe_allow_html=True)


def render_welcome():
    render_header()

    # Verificar templates
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
            <p style="color: #64748b; font-size: 0.9rem;">Documento base para cualquier cambio en proceso o maquinaria. Incluye evaluación de riesgos y plan de implementación.</p>
            <ul style="color: #475569; font-size: 0.85rem; padding-left: 1.2rem;">
                <li>Evaluación técnica completa</li><li>Riesgos y controles</li><li>Plan de implementación</li><li>Traducción al inglés</li>
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
            <p style="color: #64748b; font-size: 0.9rem;">Formato estructurado para registrar mejoras. Metodología Lean con análisis de causa raíz.</p>
            <ul style="color: #475569; font-size: 0.85rem; padding-left: 1.2rem;">
                <li>Análisis 5 Porqués</li><li>Contramedidas SMART</li><li>Plan de seguimiento</li><li>Estandarización</li>
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
            <p style="color: #64748b; font-size: 0.9rem;">Registro rápido de mejoras del día a día. Captura ideas de mejora continua.</p>
            <ul style="color: #475569; font-size: 0.85rem; padding-left: 1.2rem;">
                <li>8 Desperdicios (Wastes)</li><li>Impacto BTO</li><li>Beneficios cuantificados</li><li>Replicabilidad</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
        if st.button("➕ Crear Kaizen", key="btn_kaizen", use_container_width=True):
            st.session_state.page = "nuevo_kaizen"
            st.rerun()

    # Estadisticas
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
# FORMULARIOS DE CREACION
# =============================================================================

def render_moc_form():
    config = st.session_state.config

    st.markdown('<div class="section-header"><h3>📋 Nueva Management of Change (MoC)</h3></div>', unsafe_allow_html=True)
    st.info("💡 Complete la información y describa el problema. La IA generará automáticamente todos los campos técnicos.")

    # Verificar template
    if not st.session_state.get("template_moc_bytes"):
        st.error("❌ **Template MoC no cargado.** Vaya a Configuración > Templates y suba el archivo .pptx oficial.")
        if st.button("Ir a Configuración", key="go_config_moc"):
            st.session_state.page = "configuracion"
            st.rerun()
        return

    st.markdown("#### 1. Información General")
    col1, col2 = st.columns(2)
    with col1:
        moc_title = st.text_input("Título de la MoC:", placeholder="Ej: Control de Acceso a panel HMI")
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

    st.markdown("#### 3. Descripción del Problema")
    st.markdown('<div class="spell-hint">✍️ <strong>Corrector activo:</strong> La IA corregirá ortografía y gramática automáticamente al generar.</div>', unsafe_allow_html=True)

    problem_desc = st.text_area("Describa el problema con sus palabras:", height=200,
                                placeholder="Ejemplo: El panel HMI presenta fallas intermitentes en visualización de parámetros críticos...")

    st.markdown("#### 4. Contexto Adicional (Opcional)")
    context = st.text_area("Información adicional:", height=100, placeholder="TAG del equipo, área, fechas relevantes...")

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

        with st.spinner("🧠 La IA está generando el documento completo..."):
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
    st.info("💡 Describa el problema y la IA generará el documento A3 completo.")

    if not st.session_state.get("template_a3_bytes"):
        st.error("❌ **Template A3 no cargado.** Vaya a Configuración > Templates.")
        if st.button("Ir a Configuración", key="go_config_a3"):
            st.session_state.page = "configuracion"
            st.rerun()
        return

    st.markdown("#### 1. Información General")
    col1, col2 = st.columns(2)
    with col1:
        a3_title = st.text_input("Título:", placeholder="Ej: Reducción de tiempo de cambio")
        area = st.text_input("Área:", value=config.get("default_area", ""))
    with col2:
        autor = st.text_input("Autor:", value=config.get("default_author", ""))
        doc_number = st.text_input("Número:", value=Utils.generate_doc_number("a3"), disabled=True)
    fecha = st.text_input("Fecha:", value=Utils.format_date(), disabled=True)

    st.markdown("#### 2. Descripción del Problema")
    st.markdown('<div class="spell-hint">✍️ Corrector ortográfico activo.</div>', unsafe_allow_html=True)
    problem_desc = st.text_area("Describa el problema actual:", height=200,
                                placeholder="¿Qué está pasando? ¿Impacto? ¿Desde cuándo? ¿Datos?")

    context = st.text_area("Contexto adicional:", height=80)

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

        with st.spinner("🧠 Generando documento A3..."):
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
    st.info("💡 Describa la actividad de mejora realizada.")

    if not st.session_state.get("template_kaizen_bytes"):
        st.error("❌ **Template Kaizen no cargado.** Vaya a Configuración > Templates.")
        if st.button("Ir a Configuración", key="go_config_kzn"):
            st.session_state.page = "configuracion"
            st.rerun()
        return

    st.markdown("#### 1. Información General")
    col1, col2 = st.columns(2)
    with col1:
        kaizen_title = st.text_input("Título:", placeholder="Ej: Organización de área de herramientas")
        area = st.text_input("Área:", value=config.get("default_area", ""))
    with col2:
        autor = st.text_input("Autor:", value=config.get("default_author", ""))
        doc_number = st.text_input("Número:", value=Utils.generate_doc_number("kaizen"), disabled=True)
    fecha = st.text_input("Fecha:", value=Utils.format_date(), disabled=True)

    st.markdown("#### 2. Descripción de la Actividad")
    st.markdown('<div class="spell-hint">✍️ Corrector ortográfico activo.</div>', unsafe_allow_html=True)
    activity_desc = st.text_area("Describa la mejora realizada:", height=200,
                                 placeholder="Ejemplo: Se reorganizó el área implementando shadow board...")

    st.markdown("#### 3. Clasificación")
    tipo_desp = st.multiselect("Tipo de Desperdicio eliminado:",
                               ["Motion", "Skills", "Inventory", "Transportation",
                                "Over Production", "Over Processing", "Waiting", "Defects"])
    impacto_bto = st.selectbox("Impacto BTO:",
                               ["Safe and Sustainable", "People & Culture",
                                "Network Optimisation", "Supply Chain and Manufacturing Excellence"])

    context = st.text_area("Contexto adicional:", height=80)

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

        with st.spinner("🧠 Generando documento Kaizen..."):
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
# PANTALLA DE REVISION UNIFICADA
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
    st.info("💡 Revise cada campo, realice correcciones y genere el documento final.")

    # Corrector ortografico inline
    if config.get("spell_check", True):
        st.markdown('<div class="spell-hint">🔧 <strong>Corrector ortográfico:</strong> Puede usar el botón "Corregir ortografía" en cada campo para pulir la redacción automáticamente.</div>', unsafe_allow_html=True)

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
        text = st.text_area(label, value=value, height=100, key=f"{key_prefix}_field")
    with col2:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("✨ Corregir", key=f"{key_prefix}_spell", help="Corregir ortografía con IA"):
            with st.spinner("Corrigiendo..."):
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

        col1, col2, col3 = st.columns(3)
        with col1:
            if st.button("🇪🇸 Generar Español", type="primary", use_container_width=True):
                _finalize_document(data, meta, images, "es", "moc")
        with col2:
            if st.button("🇺🇸 Generar Inglés", type="primary", use_container_width=True):
                _finalize_document(data, meta, images, "en", "moc")
        with col3:
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
        col1, col2 = st.columns(2)
        with col1:
            if st.button("✅ Generar A3", type="primary", use_container_width=True):
                _finalize_document(data, meta, images, "es", "a3")
        with col2:
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
        col1, col2 = st.columns(2)
        with col1:
            if st.button("✅ Generar Kaizen", type="primary", use_container_width=True):
                _finalize_document(data, meta, images, "es", "kaizen")
        with col2:
            if st.button("🔄 Regenerar", use_container_width=True):
                st.session_state.page = "nuevo_kaizen"
                st.rerun()

    st.session_state.generated_data = data
    st.session_state.doc_meta = meta


def _finalize_document(data, meta, images, language, doc_type):
    """Finaliza y genera documento"""
    config = st.session_state.config
    gemini = GeminiService(config.get("gemini_api_key", ""), config.get("gemini_model", "gemini-1.5-pro"))

    with st.spinner(f"📄 Generando documento {'en Inglés' if language == 'en' else 'en Español'}..."):
        final_data = {**meta, **data}

        if language == "en" and doc_type == "moc":
            st.info("🌐 Traduciendo...")
            final_data = gemini.translate_document(final_data)

        generator = DocumentGenerator()

        if doc_type == "moc":
            buffer = generator.generate_moc(final_data, images, st.session_state.get("template_moc_bytes"))
            ext = "pptx"
            mime = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        elif doc_type == "a3":
            buffer = generator.generate_a3(final_data, images, st.session_state.get("template_a3_bytes"))
            ext = "docx"
            mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        else:
            buffer = generator.generate_kaizen(final_data, images, st.session_state.get("template_kaizen_bytes"))
            ext = "pptx"
            mime = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

        if buffer is None:
            return

        filename = f"{meta.get('moc_number', meta.get('doc_number', 'DOC'))}.{ext}"

        # Guardar en historial
        doc_info = {
            "type": doc_type,
            "title": meta.get("moc_title", meta.get("titulo", "Sin título")),
            "number": meta.get("moc_number", meta.get("doc_number", "")),
            "language": language,
            "filename": filename
        }
        Utils.add_to_history(doc_info)

        st.success(f"✅ Documento generado: {filename}")

        st.download_button(
            label=f"📥 Descargar {ext.upper()}",
            data=buffer,
            file_name=filename,
            mime=mime,
            use_container_width=True
        )

        st.info("💡 Para obtener PDF: abra el archivo descargado en Microsoft Office/LibreOffice y exporte a PDF.")


# =============================================================================
# HISTORIAL Y CONFIGURACION
# =============================================================================

def render_history():
    """Renderiza historial de documentos con exportacion/importacion"""
    st.markdown('<div class="section-header"><h3>📁 Historial de Documentos Generados</h3></div>', unsafe_allow_html=True)

    docs = st.session_state.history.get("documents", [])

    # Exportar/Importar historial
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("#### 💾 Exportar Configuración")
        export_data = {
            "config": st.session_state.config,
            "history": st.session_state.history,
            "export_date": datetime.now().isoformat()
        }
        export_json = json.dumps(export_data, indent=2, ensure_ascii=False)
        st.download_button(
            label="📥 Descargar backup (JSON)",
            data=export_json,
            file_name=f"gestion_documental_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
            mime="application/json",
            use_container_width=True
        )
        st.caption("Guarde este archivo para restaurar su configuración e historial en otra sesión.")

    with col2:
        st.markdown("#### 📂 Importar Configuración")
        uploaded_backup = st.file_uploader("Seleccione archivo de backup (.json):", type=["json"], key="import_backup")
        if uploaded_backup:
            try:
                backup_data = json.loads(uploaded_backup.read())
                if "config" in backup_data:
                    st.session_state.config = backup_data["config"]
                if "history" in backup_data:
                    st.session_state.history = backup_data["history"]
                st.success("✅ Configuración e historial restaurados correctamente.")
                st.rerun()
            except Exception as e:
                st.error(f"❌ Error al importar: {e}")

    st.markdown("<hr>", unsafe_allow_html=True)

    if not docs:
        st.info("📭 No hay documentos generados aún.")
        return

    # Filtros
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

    st.markdown(f"**Mostrando {len(filtered)} documento(s)**")

    for doc in filtered:
        type_emoji = {"moc": "📋", "a3": "📊", "kaizen": "⚡"}.get(doc.get("type"), "📄")
        type_label = {"moc": "MoC", "a3": "A3", "kaizen": "Kaizen"}.get(doc.get("type"), "Doc")
        lang_flag = "🇪🇸" if doc.get("language") == "es" else "🇺🇸"

        st.markdown(f"""
        <div class="history-item">
            <h4 style="margin: 0; color: #1e293b;">{type_emoji} {doc.get('title', 'Sin título')}</h4>
            <p style="margin: 0.25rem 0; color: #64748b; font-size: 0.9rem;">
                {type_label} · {doc.get('number', '')} · {lang_flag} · {doc.get('timestamp', '')[:10]}
            </p>
        </div>
        """, unsafe_allow_html=True)

        if st.button("🗑️ Eliminar", key=f"del_{doc.get('id', 'x')}"):
            st.session_state.history["documents"] = [d for d in st.session_state.history["documents"] 
                                                      if d.get("id") != doc.get("id")]
            st.rerun()


def render_settings():
    """Renderiza configuracion completa"""
    st.markdown('<div class="section-header"><h3>⚙️ Configuración del Sistema</h3></div>', unsafe_allow_html=True)

    config = st.session_state.config

    tabs = st.tabs(["🔑 API Gemini", "🏢 Empresa", "📄 Templates", "💾 Backup"])

    with tabs[0]:
        st.markdown("#### API Key Gemini")
        st.info("💡 Obtenga su API Key gratuita en [Google AI Studio](https://aistudio.google.com/)")

        api_key = st.text_input("API Key:", value=config.get("gemini_api_key", ""), type="password")

        st.markdown("#### Selección de Modelo")
        current_model = config.get("gemini_model", "gemini-1.5-pro")

        col1, col2, col3 = st.columns(3)
        models = [
            ("gemini-1.5-flash-lite", "⚡ 3.1 Flash-Lite", "Respuestas rápidas", "Nuevo"),
            ("gemini-1.5-flash", "🔥 3.5 Flash", "Ayuda completa", "Nuevo"),
            ("gemini-1.5-pro", "🧠 3.1 Pro", "Advanced math and code", "Recomendado"),
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
                    st.success(f"✅ Modelo: {name}")
                    st.rerun()

        st.markdown("#### Nivel de Pensamiento")
        thinking = st.select_slider("Profundidad:", ["Básico", "Estándar", "Profundo"],
                                    value=config.get("thinking_level", "Estándar"))

        st.markdown("#### Corrección Ortográfica")
        spell = st.toggle("Corrección automática activa", value=config.get("spell_check", True))

        if st.button("💾 Guardar Configuración API", type="primary", use_container_width=True):
            config["gemini_api_key"] = api_key
            config["thinking_level"] = thinking
            config["spell_check"] = spell
            st.session_state.config = config
            st.success("✅ Guardado")

    with tabs[1]:
        st.markdown("#### Datos de la Empresa")
        company = st.text_input("Empresa:", value=config.get("company_name", ""), placeholder="Ej: Orica Peru")
        dept = st.text_input("Departamento:", value=config.get("department", ""), placeholder="Ej: Mantenimiento")
        author = st.text_input("Autor por defecto:", value=config.get("default_author", ""))
        area = st.text_input("Área por defecto:", value=config.get("default_area", ""), placeholder="Ej: Planta Lurín")

        if st.button("💾 Guardar Datos", type="primary", use_container_width=True):
            config["company_name"] = company
            config["department"] = dept
            config["default_author"] = author
            config["default_area"] = area
            st.session_state.config = config
            st.success("✅ Guardado")

    with tabs[2]:
        st.markdown("#### Carga de Templates Oficiales")
        st.warning("⚠️ **Importante:** Los templates son los formatos oficiales de su empresa. Suba los archivos .pptx (MoC, Kaizen) y .docx (A3).")

        # Estado actual
        moc_ok = st.session_state.get("template_moc_bytes") is not None
        a3_ok = st.session_state.get("template_a3_bytes") is not None
        kzn_ok = st.session_state.get("template_kaizen_bytes") is not None

        col1, col2, col3 = st.columns(3)

        with col1:
            status_moc = "✅ Cargado" if moc_ok else "❌ Pendiente"
            st.markdown(f"**Template MoC**\n{status_moc}")
            moc_file = st.file_uploader("Subir MoC (.pptx)", type=["pptx"], key="upload_moc")

        with col2:
            status_a3 = "✅ Cargado" if a3_ok else "❌ Pendiente"
            st.markdown(f"**Template A3**\n{status_a3}")
            a3_file = st.file_uploader("Subir A3 (.docx)", type=["docx"], key="upload_a3")

        with col3:
            status_kzn = "✅ Cargado" if kzn_ok else "❌ Pendiente"
            st.markdown(f"**Template Kaizen**\n{status_kzn}")
            kzn_file = st.file_uploader("Subir Kaizen (.pptx)", type=["pptx"], key="upload_kzn")

        st.markdown("<br>", unsafe_allow_html=True)

        # Botón único para guardar todos los templates seleccionados
        if st.button("💾 Guardar Templates", type="primary", use_container_width=True):
            saved_any = False

            if moc_file is not None:
                st.session_state.template_moc_bytes = moc_file.getvalue()
                st.success("✅ Template MoC guardado")
                saved_any = True

            if a3_file is not None:
                st.session_state.template_a3_bytes = a3_file.getvalue()
                st.success("✅ Template A3 guardado")
                saved_any = True

            if kzn_file is not None:
                st.session_state.template_kaizen_bytes = kzn_file.getvalue()
                st.success("✅ Template Kaizen guardado")
                saved_any = True

            if not saved_any:
                st.warning("⚠️ No se seleccionó ningún archivo nuevo.")
            else:
                # Solo hacer rerun una vez al final, después de guardar todo
                st.rerun()

        if moc_ok and a3_ok and kzn_ok:
            st.balloons()
            st.success("🎉 ¡Todos los templates están cargados! Puede comenzar a generar documentos.")

    with tabs[3]:
        st.markdown("#### Backup y Restauración")
        st.info("💡 Exporte su configuración e historial para no perderlos al cerrar el navegador.")

        export_data = {
            "config": st.session_state.config,
            "history": st.session_state.history,
            "export_date": datetime.now().isoformat(),
            "version": "3.0.0"
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
                if "history" in restore_data:
                    st.session_state.history = restore_data["history"]
                st.success("✅ Restauración completada. Recargando...")
                st.rerun()
            except Exception as e:
                st.error(f"❌ Error: {e}")


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
        <p>Diseñado por <strong>Roger Huamani</strong> | Sistema de Gestión Documental v3.0.0</p>
        <p style="font-size: 0.75rem; color: #94a3b8;">
            Software empresarial para automatización de documentos MoC, A3 y Kaizen.<br>
            Mantiene los formatos oficiales de la empresa sin modificaciones.
        </p>
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
