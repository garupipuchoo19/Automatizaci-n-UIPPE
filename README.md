# 📊 Sistema de Automatización de Reportes Administrativos - UIPPE

**OPERAGUA Cuautitlán Izcalli**  
*Unidad de Información, Planeación, Programación y Evaluación (UIPPE)*

---

## 📄 Descripción del Proyecto

Aplicación de escritorio modular desarrollada en Python para la Unidad de Información, Planeación, Programación y Evaluación (UIPPE) del organismo público descentralizado OPERAGUA Cuautitlán Izcalli.

El sistema tiene como objetivo automatizar la lectura, la validación normativa, consolidación e inyección de datos para reportes trimestrales de evaluación de metas e indicadores (PbRM e ITSP). Integra procesamiento de hojas de cálculo de Excel, generación de informes ejecutivos en Word y síntesis cualitativa mediante la API de Google Gemini.

---

## 🛠️ Arquitectura y Stack Tecnológico

La aplicación está construida bajo una arquitectura modular en capas en Python 3.11+, separando la interfaz de usuario, los controladores de flujo, los motores de negocio y los servicios de soporte.

* Lenguaje: Python 3.11.9
* Interfaz Gráfica (GUI): CustomTkinter (soporte multihilo con threading para evitar congelamiento de la interfaz)
* Procesamiento de Datos: pandas / openpyxl
* Generación de Documentos: python-docx / reportlab
* Lectura e Ingesta: pdfplumber / python-docx
* Integración de IA: Google Gemini API (google-generativeai)
* Gestión de Entorno: python-dotenv
* Compilador Binario: Nuitka

---

## 🚀 Instalación y Configuración

### 1. Clonar el repositorio
git clone https://github.com/garupipuchoo19/Automatizaci-n-UIPPE.git
cd Automatizaci-n-UIPPE

### 2. Crear y activar el entorno virtual
# Crear entorno virtual
python -m venv venv

# Activar en PowerShell (Windows)
.\venv\Scripts\Activate.ps1

### 3. Instalar dependencias
pip install -r requirements.txt

### 4. Configurar variables de entorno (.env)
1. Copia el archivo de ejemplo .env.example y renómbralo a .env:
   Copy-Item .env.example .env

2. Abre el archivo .env en tu editor e ingresa tu API Key:
   GEMINI_API_KEY=tu_clave_de_google_ai_studio_aqui

Nota de Seguridad: El archivo .env se encuentra excluido del control de versiones mediante .gitignore para prevenir la exposición de credenciales en repositorios públicos.

---

## 🔥 Características Principales y Funcionalidad

1. Semaforización Normativa (PbRM):
   * Clasificación automática del porcentaje de cumplimiento por meta individual e institucional en 5 niveles:
     * 🟢 Adecuado: 95.0% - 105.0%
     * 🟠 Crítico: 70.0% - 94.9%
     * 🔵 Sobrepasado: > 105.0%
     * 🟡 Regular: 50.0% - 69.9%
     * 🔴 Deficiente: < 50.0%

2. Captura Rápida e Inyección Directa en Excel:
   * Formulario intuitivo para actualizar avances físicos/financieros y justificaciones por Dirección y Área operativa, inyectándolos directamente en la sábana maestra CUAUTITLAN METAS E INDICADORES 2026.xlsx preservando las fórmulas existentes.

3. Generación de Concentrado Ejecutivo en Word:
   * Sustituye el vaciado masivo de registros por un Concentrado Ejecutivo Sintético de 1 página en la plantilla Word oficial (ITSP 1 TRIMESTRE 2026.docx / ITSP 2 TRIMESTRE 2026.docx), incorporando métricas globales, porcentajes e identificadores institucionales de color.

4. Síntesis Cualitativa Inteligente:
   * Integración con la API de Gemini (con lógica de reintentos y exponential backoff) para analizar desvíos cuantitativos y generar explicaciones cualitativas consolidadas.

5. Validación de Insumos e Inspección:
   * Motor de inspección para identificar celdas vacías, inconsistencias en fórmulas, registros incompletos y duplicados antes de emitir el reporte final.

6. Bitácora Automatizada de Eventos:
   * Registro detallado en tiempo real dentro de la carpeta logs/, garantizando rastreabilidad de operaciones y depuración expedita.

---

## 📂 Estructura del Proyecto

Automatizacion_UIPPE/
│
├── .env                        # Archivo local con la API Key (NO se sube a Git)
├── .env.example                # Plantilla pública de variables de entorno
├── .gitignore                  # Reglas de exclusión para Git
├── main.py                     # Punto de entrada principal (GUI CustomTkinter)
├── generar_docs.py             # Script auxiliar de formato y exportación de reportes Word
├── test_prueba.py              # Script de prueba de integración y validación CLI
├── requirements.txt            # Inventario de dependencias del proyecto
├── README.md                   # Documentación técnica y guía de uso
│
├── entradas/                   # Archivos de insumo recibidos (contiene CUAUTITLAN METAS E INDICADORES 2026.xlsx)
├── salidas/                    # Reportes procesados, consolidados e inyectados (.docx / .xlsx)
├── plantillas/                 # Formatos base oficiales e insumos estáticos (.docx de ITSP 1° y 2° Trimestre)
├── logs/                       # Historial de ejecuciones y registro físico de eventos (log_uippe_YYYYMM.log)
├── respaldos/                  # Copias de seguridad de contexto e historial de salidas
│   ├── contexto_word/          # Archivos Word consolidados de contexto previo
│   └── historico_salidas/      # Resguardo de versiones anteriores procesadas
│
└── src/                        # Código fuente modular
    ├── __init__.py
    │
    ├── controllers/            # Controladores de flujo y orquestación de procesos
    │   ├── __init__.py
    │   └── procesador_uippe.py # Orquestador general del pipeline (Engine, Validador y Generador)
    │
    ├── core/                   # Lógica de negocio, motores de cálculo e inyección
    │   ├── __init__.py
    │   ├── excel_engine.py     # Motor de lectura, recálculo de sábanas y evaluación por área
    │   ├── generador_doc.py    # Exportación a Excel estilizado (.xlsx) y Concentrado Ejecutivo en Word (.docx)
    │   ├── generador_excel_pbrm.py # Motor de inyección directa de avances en Excel sin romper fórmulas
    │   ├── integracion_gemini.py   # Conector e integración con la API de Gemini para análisis cualitativo
    │   ├── lector_excel.py     # Lectura y parsing especializado de hojas de trabajo de Excel
    │   ├── lector_pdf.py       # Extracción de tablas y texto estructurado desde archivos PDF
    │   └── validador_reg.py    # Clasificador de desempeño por semáforo oficial (Púrpura, Verde, Amarillo, Naranja, Rojo)
    │
    ├── gui/                    # Módulo de Interfaz Gráfica (CustomTkinter)
    │   ├── __init__.py
    │   ├── app.py              # Ventana principal con navegación por pestañas (Captura + Consolidador)
    │   ├── captura_view.py     # VistaCapturaRapida (Interfaz de formulario por Dirección y Área)
    │   └── components.py       # Componentes reutilizables de UI (Tarjetas KPI y Opciones de Exportación)
    │
    └── utils/                  # Herramientas y servicios auxiliares
        ├── __init__.py
        ├── helpers.py          # Manejo de rutas absolutas (BASE_DIR), auto-creación de carpetas y fechas oficiales
        └── logger.py           # Sistema de registro e historial automatizado de eventos hacia la carpeta logs/
