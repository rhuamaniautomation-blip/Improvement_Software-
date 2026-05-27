# 🏭 Sistema de Gestión Documental - MoC | A3 | Kaizen

## 🌐 Versión Streamlit Web

Software empresarial profesional para la automatización inteligente de documentos de gestión de cambios, ejecutable online mediante Streamlit.

---

## ✨ Características Principales

### 🤖 Inteligencia Artificial (Gemini)
- **Selección de modelo**: 3.1 Flash-Lite, 3.5 Flash, 3.1 Pro
- **Corrección ortográfica automática**: La IA corrige errores en tiempo real
- **Generación completa**: Describe el problema, la IA genera TODO el documento
- **Traducción al inglés**: Para envío al Panel de Expertos

### 📄 Documentos
- **MoC**: PowerPoint (.pptx) + PDF
- **A3**: Word (.docx) + PDF
- **Kaizen**: PowerPoint (.pptx) + PDF
- **Formatos oficiales preservados**: Sin modificaciones a los templates

### 🖼️ Imágenes
- Carga múltiple con numeración correlativa automática (Figura 1, 2, 3...)
- Descripción profesional por cada imagen

### 📁 Gestión
- **Historial completo** de todos los documentos generados
- **Números correlativos** automáticos por tipo (MOC-20260526-0001, A3-..., KZN-...)
- **Filtros** por tipo, idioma, búsqueda
- **Persistencia** de configuración entre sesiones

### ⚙️ Configuración
- API Key Gemini
- Modelo de IA seleccionable
- Datos de empresa (autor, área, departamento)
- Templates actualizables
- Header/Footer personalizable

---

## 🚀 Instalación y Ejecución

### Requisitos
- Python 3.9+
- pip

### Pasos

```bash
# 1. Clonar o descargar el proyecto
cd gestion_moc_streamlit

# 2. Instalar dependencias
pip install -r requirements.txt

# 3. Ejecutar la aplicación
streamlit run app.py

# 4. Abrir en navegador
# http://localhost:8501
```

### Despliegue Online (Streamlit Cloud)

1. Subir a repositorio GitHub
2. Conectar con [Streamlit Cloud](https://streamlit.io/cloud)
3. ¡Listo! Cualquier usuario con el enlace puede acceder

---

## 📋 Flujo de Uso

### Crear MoC
1. Click **"Nueva MoC"**
2. Complete información general
3. Describa el problema con sus palabras
4. Click **"Generar con IA"**
5. Revise y edite todos los campos
6. Genere en **Español** o **Inglés**
7. Descargue PowerPoint o PDF

### Crear A3
1. Click **"Nueva Mejora A3"**
2. Describa el problema actual
3. Agregue imágenes de soporte
4. Genere y revise
5. Descargue Word o PDF

### Crear Kaizen
1. Click **"Nuevo Kaizen"**
2. Describa la actividad realizada
3. Seleccione tipo de desperdicio e impacto BTO
4. Genere y descargue

---

## 🔧 Configuración

### API Gemini (Recomendado)
1. Visite [Google AI Studio](https://aistudio.google.com/)
2. Genere API Key gratuita
3. Ingresela en **Configuración > API Gemini**
4. Seleccione modelo (recomendado: **3.1 Pro**)

### Sin API Key
El software funciona con generación local incluida, aunque la calidad es básica.

---

## 🎨 Diseño

- **Colores empresariales**: Tonos claros, azul corporativo #1a5f7a
- **Interfaz intuitiva**: Para personal técnico no especialista en documentación
- **Responsive**: Funciona en desktop, tablet y móvil
- **Sidebar oscuro**: Navegación clara y profesional

---

## 👨‍💻 Créditos

**CAVA** - Especialistas en Robótica y Automatización  
**Diseñado por:** Roger Huamani  
**Versión:** 2.0.0 | Mayo 2026

---

## 📄 Licencia
Uso interno empresarial. Confidencial.
