import os
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def aplicar_sombreado_celda(cell, color_hex):
    """Aplica color de fondo a una celda de tabla."""
    tcPr = cell._element.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    tcPr.append(shd)

def corregir_ajuste_tabla(tabla):
    """Garantiza alineación centrada sin desbordes de texto a los costados."""
    tabla.alignment = WD_TABLE_ALIGNMENT.CENTER

def crear_documentacion():
    doc = Document()

    # --- CONFIGURACIÓN DE MÁRGENES ---
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # --- COLORES INSTITUCIONALES OPERAGUA / UIPPE ---
    AZUL_OPERAGUA = RGBColor(31, 78, 120)  # #1F4E78
    GRIS_TEXTO = RGBColor(89, 89, 89)      # #595959
    HEX_AZUL = "1F4E78"
    HEX_GRIS_CLARO = "F2F2F2"

    # --- ENCABEZADO Y TÍTULO ---
    p_encabezado = doc.add_paragraph()
    r_encabezado = p_encabezado.add_run(
        "ORGANISMO PÚBLICO DESCENTRALIZADO OPERAGUA CUAUTITLÁN IZCALLI\n"
        "UIPPE - Unidad de Información, Planeación, Programación y Evaluación"
    )
    r_encabezado.font.size = Pt(9)
    r_encabezado.font.bold = True
    r_encabezado.font.color.rgb = GRIS_TEXTO

    p_titulo = doc.add_paragraph()
    p_titulo.paragraph_format.space_before = Pt(20)
    p_titulo.paragraph_format.space_after = Pt(10)
    r_titulo = p_titulo.add_run("DOCUMENTACIÓN TÉCNICA Y ARQUITECTURA DE SOFTWARE")
    r_titulo.font.size = Pt(20)
    r_titulo.font.bold = True
    r_titulo.font.color.rgb = AZUL_OPERAGUA

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_after = Pt(20)
    r_sub = p_sub.add_run(
        "Proyecto: Sistema de Automatización, Inyección y Semaforización PbRM (Fase 2: Consolidación Ejecutiva, IA y Memoria Histórica)"
    )
    r_sub.font.size = Pt(11)
    r_sub.font.italic = True
    r_sub.font.color.rgb = GRIS_TEXTO

    # --- SECCIÓN 1: RESUMEN EJECUTIVO Y EVOLUCIÓN ---
    h1 = doc.add_heading("1. Resumen Ejecutivo y Evolución del Sistema", level=1)
    h1.runs[0].font.color.rgb = AZUL_OPERAGUA

    p_body1 = doc.add_paragraph(
        "El sistema Automatizacion_UIPPE ha sido desarrollado para optimizar la ingesta, evaluación, semaforización "
        "y redacción de informes del Presupuesto basado en Resultados Municipal (PbRM) en OPERAGUA Cuautitlán Izcalli.\n\n"
        "Evolución del Proyecto:\n"
        "• Funcionalidades Base (Iniciales): Captura rápida de avances por Dirección/Área, extracción estructurada desde archivos Excel/PDF y semaforización básica de cumplimiento.\n"
        "• Funcionalidades Actuales (Fase 2): Inyección directa de datos en sábanas sin romper fórmulas integradas, reemplazo directo de Sábana Maestra desde la GUI para nuevos ciclos fiscales, "
        "selección estilizada de fechas vía CTkDatePicker, carga automática de credenciales de IA (.env), generación de análisis cualitativos ejecutivos vía Google Gemini API "
        "con lectura de contexto histórico (.docx) de hasta 5 trimestres, corrección de ambigüedades en DataFrames de Pandas y un ciclo automatizado de rotación/depuración de almacenamiento."
    )
    p_body1.paragraph_format.line_spacing = 1.15
    p_body1.paragraph_format.space_after = Pt(12)

    # --- SECCIÓN 2: ENTORNO TÉCNICO Y DEPENDENCIAS ---
    h2 = doc.add_heading("2. Especificaciones del Entorno Técnico", level=1)
    h2.runs[0].font.color.rgb = AZUL_OPERAGUA

    tabla_env = doc.add_table(rows=8, cols=2)
    corregir_ajuste_tabla(tabla_env)
    tabla_env.style = 'Table Grid'

    datos_env = [
        ("Componente", "Especificación / Versión"),
        ("Sistema Operativo", "Windows 10 / 11 (64-bit)"),
        ("Lenguaje de Programación", "Python 3.11+"),
        ("Interfaz Gráfica (GUI)", "CustomTkinter + tkcalendar (DatePicker, Pestañas, Captura y KPIs)"),
        ("Procesamiento de Archivos", "Pandas, openpyxl, python-docx, pdfplumber, shutil"),
        ("Inteligencia Artificial", "Google Gemini API (Análisis cualitativo con contexto histórico y credenciales .env)"),
        ("Gestión de Almacenamiento", "Rotación y depuración automática (Máx. 5 reportes en contexto, 10 en historial)"),
        ("Compilación Binaria", "Nuitka / PyInstaller (Ejecutable Standalone .exe)")
    ]

    for row_idx, (c1, c2) in enumerate(datos_env):
        row = tabla_env.rows[row_idx]
        row.cells[0].text = c1
        row.cells[1].text = c2
        if row_idx == 0:
            aplicar_sombreado_celda(row.cells[0], HEX_AZUL)
            aplicar_sombreado_celda(row.cells[1], HEX_AZUL)
            row.cells[0].paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
            row.cells[0].paragraphs[0].runs[0].font.bold = True
            row.cells[1].paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
            row.cells[1].paragraphs[0].runs[0].font.bold = True
        elif row_idx % 2 == 1:
            aplicar_sombreado_celda(row.cells[0], HEX_GRIS_CLARO)
            aplicar_sombreado_celda(row.cells[1], HEX_GRIS_CLARO)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # --- SECCIÓN 3: SEMAFORIZACIÓN NORMATIVA ---
    h3_sem = doc.add_heading("3. Reglas de Semaforización Normativa (PbRM / OSFEM)", level=1)
    h3_sem.runs[0].font.color.rgb = AZUL_OPERAGUA

    p_sem = doc.add_paragraph(
        "El motor de validación clasifica el porcentaje de cumplimiento programático de cada meta en 5 niveles oficiales de la matriz de evaluación:"
    )

    tabla_sem = doc.add_table(rows=6, cols=3)
    corregir_ajuste_tabla(tabla_sem)
    tabla_sem.style = 'Table Grid'

    datos_sem = [
        ("Estatus / Semáforo", "Rango de Cumplimiento", "Criterio de Evaluación UIPPE"),
        ("Crítico (C)", "< 70.0%", "Incumplimiento crítico o desfase severo en avance"),
        ("Deficiente (D)", "70.0% - 84.9%", "Subejercicio significativo respecto a lo programado"),
        ("Regular (R)", "85.0% - 94.9%", "Cumplimiento parcial cercano al margen de tolerancia"),
        ("Adecuado (A)", "95.0% - 110.0%", "Cumplimiento óptimo dentro de parámetros programados"),
        ("Sobrepasado (S)", "> 110.0%", "Superación del margen de meta estimada")
    ]

    colores_hex_sem = ["1F4E78", "E74C3C", "E67E22", "F1C40F", "27AE60", "2980B9"]

    for row_idx, (c1, c2, c3) in enumerate(datos_sem):
        row = tabla_sem.rows[row_idx]
        row.cells[0].text = c1
        row.cells[1].text = c2
        row.cells[2].text = c3

        if row_idx == 0:
            aplicar_sombreado_celda(row.cells[0], colores_hex_sem[0])
            aplicar_sombreado_celda(row.cells[1], colores_hex_sem[0])
            aplicar_sombreado_celda(row.cells[2], colores_hex_sem[0])
            for cell in row.cells:
                cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
                cell.paragraphs[0].runs[0].font.bold = True
        else:
            aplicar_sombreado_celda(row.cells[0], colores_hex_sem[row_idx])
            row.cells[0].paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
            row.cells[0].paragraphs[0].runs[0].font.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # --- SECCIÓN 4: ARQUITECTURA DETALLADA DEL PROYECTO ---
    h4 = doc.add_heading("4. Arquitectura Modular del Sistema", level=1)
    h4.runs[0].font.color.rgb = AZUL_OPERAGUA

    modulos = [
        ("main.py", "Punto de entrada principal. Inicializa carpetas clave y lanza la interfaz gráfica CustomTkinter."),
        ("generar_docs.py", "Generador dinámico en ejecución aislada para actualizar la Documentación Técnica Oficial en Word."),
        ("src/controllers/procesador_uippe.py", "Orquestador general que coordina el Engine, Validador, IA y Generadores."),
        ("src/gui/app.py", "Ventana principal con gestión de ciclo fiscal, actualización de Sábana Maestra, selector CTkDatePicker y consola de bitácora."),
        ("src/gui/captura_view.py", "VistaCapturaRapida para el ingreso de avances y justificaciones por Dirección y Área."),
        ("src/gui/components.py", "Componentes reutilizables de UI (Tarjetas KPI y Opciones de Exportación)."),
        ("src/core/excel_engine.py", "Motor de lectura, recálculo de sábanas y evaluación consolidada por área."),
        ("src/core/generador_doc.py", "Generador de reportes en Word (.docx) con inyección sintáctica limpia, matrices por Dirección e IA."),
        ("src/core/generador_excel_pbrm.py", "Motor de inyección directa de avances y justificaciones en Excel sin romper fórmulas."),
        ("src/core/integracion_gemini.py", "Módulo de conexión con Gemini API mediante entorno seguro (.env) para redacción de análisis cualitativo con contexto acumulado."),
        ("src/core/lector_excel.py", "Lectura y parsing especializado de hojas de trabajo de Excel sin ambigüedades en DataFrames."),
        ("src/core/lector_pdf.py", "Extracción de tablas y texto estructurado desde archivos PDF."),
        ("src/core/validador_reg.py", "Clasificador de desempeño por semáforo oficial PbRM/OSFEM (C, D, R, A, S)."),
        ("src/utils/helpers.py", "Manejo de rutas absolutas, depuración automática de contexto Word (máx. 5) y respaldos de salidas."),
        ("src/utils/logger.py", "Sistema de registro e historial automatizado de eventos hacia la carpeta logs/.")
    ]

    for mod, desc in modulos:
        p_mod = doc.add_paragraph(style='List Bullet')
        r_mod_name = p_mod.add_run(f"{mod}: ")
        r_mod_name.bold = True
        p_mod.add_run(desc)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # --- SECCIÓN 5: POLÍTICA DE ALMACENAMIENTO ---
    h5_maint = doc.add_heading("5. Estrategia de Almacenamiento y Mantenimiento", level=1)
    h5_maint.runs[0].font.color.rgb = AZUL_OPERAGUA

    p_maint = doc.add_paragraph(
        "Para evitar la saturación gradual del almacenamiento, el sistema aplica un ciclo de rotación automatizado en helpers.py:\n"
        "1. Memoria Contextual (respaldos/contexto_word/): Conserva un tope estrictamente configurado de 5 reportes (.docx) "
        "(correspondientes a los 4 trimestres del año actual + 1 del año previo), eliminando automáticamente las versiones excedentes más antiguas.\n"
        "2. Histórico de Salidas (respaldos/historico_salidas/): Mantiene un límite de 10 versiones para respaldo de seguridad sin duplicar espacio en disco.\n"
        "3. Actualización de Plantilla Maestra: Permite el reemplazo interactivo mediante GUI manteniendo una nomenclatura dinámica de acuerdo al ciclo seleccionado."
    )
    p_maint.paragraph_format.line_spacing = 1.15
    p_maint.paragraph_format.space_after = Pt(12)

    # --- SECCIÓN 6: MATRIZ DE PRUEBAS Y CONTROL DE ERRORES ---
    h6_err = doc.add_heading("6. Matriz de Pruebas y Control de Errores", level=1)
    h6_err.runs[0].font.color.rgb = AZUL_OPERAGUA

    tabla_err = doc.add_table(rows=7, cols=3)
    corregir_ajuste_tabla(tabla_err)
    tabla_err.style = 'Table Grid'

    datos_err = [
        ("Escenario de Error / Prueba", "Causa Raíz Detectada", "Estrategia de Solución Implementada"),
        ("Estructura Excel Inválida", "Columnas o nombres de pestañas modificados manualmente por el usuario", "Validación previa en lector_excel.py y registro limpio en logs/ sin tumbar la app."),
        ("Evaluación Ambigua en Pandas", "Uso directo de condiciones booleanas (if df:) sobre DataFrames", "Uso explícito de .empty e isinstance() para validación lógica de matrices."),
        ("Falla de Conexión en Gemini API", "Corte de red, API Key no configurada o límite de cuota (HTTP 429/503)", "Mecanismo de fallback con texto estandarizado que permite generar el Word sin bloquear la salida."),
        ("Fórmulas Rotas en Excel", "Sobrescritura directa de celdas calculadas durante la inyección de avances", "Inyección celda a celda mediante openpyxl respetando celdas con fórmulas primarias."),
        ("Actualización de Sábana Activa", "Cambio de ciclo fiscal sin modificar archivo base en entradas/", "Función actualizar_excel_maestro() en GUI que reemplaza y recarga instancias en tiempo real."),
        ("Saturación de Almacenamiento", "Acumulación indeterminada de reportes temporales o historiales", "Ejecución de depurar_contexto_historico() y depurar_historico_salidas() automática en cada guardado.")
    ]

    for row_idx, (c1, c2, c3) in enumerate(datos_err):
        row = tabla_err.rows[row_idx]
        row.cells[0].text = c1
        row.cells[1].text = c2
        row.cells[2].text = c3

        if row_idx == 0:
            aplicar_sombreado_celda(row.cells[0], HEX_AZUL)
            aplicar_sombreado_celda(row.cells[1], HEX_AZUL)
            aplicar_sombreado_celda(row.cells[2], HEX_AZUL)
            for cell in row.cells:
                cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
                cell.paragraphs[0].runs[0].font.bold = True
        elif row_idx % 2 == 1:
            aplicar_sombreado_celda(row.cells[0], HEX_GRIS_CLARO)
            aplicar_sombreado_celda(row.cells[1], HEX_GRIS_CLARO)
            aplicar_sombreado_celda(row.cells[2], HEX_GRIS_CLARO)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # --- SECCIÓN 7: ESTADO Y VALIDACIÓN ---
    h7 = doc.add_heading("7. Estado de Liberación y Flujo de Insumos", level=1)
    h7.runs[0].font.color.rgb = AZUL_OPERAGUA

    p_estado = doc.add_paragraph(
        "El sistema se encuentra consolidado y listo para pruebas de campo. "
        "Los insumos procesados desde entradas/ se inyectan en la plantilla institucional de plantillas/ "
        "y se depositan formateados en salidas/. Todo evento o aviso de validación queda registrado en logs/."
    )
    p_estado.paragraph_format.space_after = Pt(20)

    # Pie de página
    section = doc.sections[0]
    footer = section.footer
    p_ftr = footer.paragraphs[0]
    p_ftr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_ftr = p_ftr.add_run("Documentación Técnica UIPPE | Conservar como Historial de Software")
    r_ftr.font.size = Pt(8)
    r_ftr.font.color.rgb = GRIS_TEXTO

    # Guardar documento en la carpeta salidas/
    ruta_salida = os.path.join(os.path.dirname(__file__), "salidas", "Documentacion_Tecnica_UIPPE.docx")
    os.makedirs(os.path.dirname(ruta_salida), exist_ok=True)
    doc.save(ruta_salida)
    print(f"✔ Documentación creada exitosamente en: {ruta_salida}")

if __name__ == "__main__":
    crear_documentacion()