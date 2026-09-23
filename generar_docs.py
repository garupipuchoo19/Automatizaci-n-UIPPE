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
    p_titulo.paragraph_format.space_before = Pt(18)
    p_titulo.paragraph_format.space_after = Pt(8)
    r_titulo = p_titulo.add_run("DOCUMENTACIÓN TÉCNICA Y ARQUITECTURA DE SOFTWARE")
    r_titulo.font.size = Pt(20)
    r_titulo.font.bold = True
    r_titulo.font.color.rgb = AZUL_OPERAGUA

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_after = Pt(18)
    r_sub = p_sub.add_run(
        "Proyecto: Sistema de Automatización, Inyección y Semaforización PbRM\n"
        "Consolidador Ejecutivo, Análisis Cualitativo con IA y Memoria Histórica Multitrimestral"
    )
    r_sub.font.size = Pt(11)
    r_sub.font.italic = True
    r_sub.font.color.rgb = GRIS_TEXTO

    # --- SECCIÓN 1: RESUMEN EJECUTIVO Y EVOLUCIÓN ---
    h1 = doc.add_heading("1. Resumen Ejecutivo y Evolución del Sistema", level=1)
    h1.runs[0].font.color.rgb = AZUL_OPERAGUA

    p_body1 = doc.add_paragraph(
        "El sistema Automatizacion_UIPPE ha sido diseñado para transformar la gestión del Presupuesto basado en Resultados "
        "Municipal (PbRM) en OPERAGUA Cuautitlán Izcalli. Automatiza de extremo a extremo la captura de metas, la evaluación "
        "programática, el recálculo de Sábanas Maestras de Excel sin alteración de fórmulas, la semaforización bajo normativa del "
        "Órgano Superior de Fiscalización del Estado de México (OSFEM) y la redacción asistida por Inteligencia Artificial para la "
        "emisión de reportes ejecutivos institucionales."
    )
    p_body1.paragraph_format.line_spacing = 1.15
    p_body1.paragraph_format.space_after = Pt(10)

    # --- SECCIÓN 2: MARCO INSTITUCIONAL Y GLOSARIO ---
    h2_glos = doc.add_heading("2. Marco Institucional y Glosario PbRM / OSFEM", level=1)
    h2_glos.runs[0].font.color.rgb = AZUL_OPERAGUA

    glosario_items = [
        ("PbRM (Presupuesto basado en Resultados Municipal)", "Metodología de gestión pública en el Estado de México que vincula la asignación de recursos al cumplimiento de metas e indicadores medibles."),
        ("OSFEM", "Órgano Superior de Fiscalización del Estado de México. Ente fiscalizador que determina los criterios normativos y rangos de semaforización programática."),
        ("UIPPE", "Unidad de Información, Planeación, Programación y Evaluación del Organismo OPERAGUA. Instancia encargada de auditar, consolidar y reportar el avance trimestral."),
        ("Sábana Maestra", "Matriz institucional en Microsoft Excel (.xlsx) que consolida las metas, programados, alcanzados, fórmulas integradas y justificaciones de todas las Direcciones y Áreas del organismo."),
        ("Memoria Contextual", "Mecanismo que analiza los informes trimestrales precedentes (.docx) para identificar tendencias, recurrencias y causas raíz en el comportamiento de las metas.")
    ]

    for termino, defn in glosario_items:
        p_g = doc.add_paragraph(style='List Bullet')
        r_g = p_g.add_run(f"{termino}: ")
        r_g.bold = True
        p_g.add_run(defn)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # --- SECCIÓN 3: ENTORNO TÉCNICO Y ARQUITECTURA DE DIRECTORIOS ---
    h3_env = doc.add_heading("3. Especificaciones Técnicas y Estructura de Directorios", level=1)
    h3_env.runs[0].font.color.rgb = AZUL_OPERAGUA

    tabla_env = doc.add_table(rows=8, cols=2)
    corregir_ajuste_tabla(tabla_env)
    tabla_env.style = 'Table Grid'

    datos_env = [
        ("Componente", "Especificación / Versión"),
        ("Sistema Operativo", "Windows 10 / 11 (64-bit)"),
        ("Entorno de Ejecución", "Python 3.11+ / Entorno Virtual (venv)"),
        ("Interfaz Gráfica (GUI)", "CustomTkinter + tkcalendar (DatePicker, Tabview, KPIs, Log Viewer)"),
        ("Procesamiento de Datos", "Pandas, openpyxl, python-docx, pdfplumber, shutil"),
        ("Motor de Inteligencia Artificial", "Google Gemini API (gemini-1.5-flash / gemini-2.0-flash via python-dotenv)"),
        ("Almacenamiento y Respaldos", "Rotación automática de contexto (máx. 5 .docx) e historial de salidas (máx. 10 .docx)"),
        ("Empaquetado y Distribución", "Nuitka / PyInstaller (Compilación binaria nativa .exe)")
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

    p_dirs_title = doc.add_paragraph()
    p_dirs_title.paragraph_format.space_before = Pt(12)
    r_dt = p_dirs_title.add_run("Estructura de Directorios del Proyecto:")
    r_dt.bold = True

    directorios = [
        ("entradas/", "Almacena el archivo maestro de Excel del ciclo fiscal activo (ej. CUAUTITLAN METAS E INDICADORES 2026.xlsx)."),
        ("plantillas/", "Contiene la plantilla institucional en Word (.docx) con encabezados, estilos y marcadores de posición."),
        ("salidas/", "Ubicación final de los reportes generados (.docx) y archivos de documentación compilados."),
        ("logs/", "Registros automatizados de auditoría y bitácora de eventos del sistema (app.log)."),
        ("respaldos/contexto_word/", "Repositorio de memoria histórica que conserva hasta 5 reportes trimestrales previos para lectura de la IA."),
        ("respaldos/historico_salidas/", "Resguardo rotativo de las últimas 10 generaciones de informes ejecutivos.")
    ]

    for d, d_desc in directorios:
        p_d = doc.add_paragraph(style='List Bullet')
        r_dn = p_d.add_run(f"{d}: ")
        r_dn.bold = True
        p_d.add_run(d_desc)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # --- SECCIÓN 4: SEMAFORIZACIÓN NORMATIVA OSFEM ---
    h4_sem = doc.add_heading("4. Criterios de Evaluación y Semaforización OSFEM", level=1)
    h4_sem.runs[0].font.color.rgb = AZUL_OPERAGUA

    p_sem = doc.add_paragraph(
        "El motor de validación (validador_reg.py) evalúa el porcentaje de avance programático mediante las siguientes reglas oficiales:"
    )

    tabla_sem = doc.add_table(rows=6, cols=3)
    corregir_ajuste_tabla(tabla_sem)
    tabla_sem.style = 'Table Grid'

    datos_sem = [
        ("Código / Semáforo", "Rango de Cumplimiento", "Criterio y Acción Normativa UIPPE"),
        ("Crítico (C)", "< 70.0%", "Incumplimiento crítico o desfase severo. Requiere justificación obligatoria."),
        ("Deficiente (D)", "70.0% - 84.9%", "Subejercicio significativo. Requiere plan de regularización trimestral."),
        ("Regular (R)", "85.0% - 94.9%", "Cumplimiento parcial dentro del margen de tolerancia aceptable."),
        ("Adecuado (A)", "95.0% - 110.0%", "Cumplimiento óptimo dentro de los parámetros programados."),
        ("Sobrepasado (S)", "> 110.0%", "Superación del margen estimado. Requiere justificación por sobremeta.")
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

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # --- SECCIÓN 5: ARQUITECTURA MODULAR Y PIPELINE DE DATOS ---
    h5_arch = doc.add_heading("5. Arquitectura Modular del Software", level=1)
    h5_arch.runs[0].font.color.rgb = AZUL_OPERAGUA

    modulos = [
        ("main.py", "Punto de entrada. Carga variables de entorno (.env), verifica carpetas del sistema e inicia la GUI."),
        ("generar_docs.py", "Generador dinámico en ejecución aislada para compilar la Documentación Técnica Oficial."),
        ("generar_manual_usuario.py", "Generador dinámico para compilar el Manual de Usuario en formato Word."),
        ("src/controllers/procesador_uippe.py", "Orquestador central que coordina la carga, semaforización, inyección, análisis IA y reportes."),
        ("src/gui/app.py", "Interfaz principal CustomTkinter con selectores de ciclo fiscal, reemplazo de Sábana, CTkDatePicker y consola."),
        ("src/gui/captura_view.py", "Módulo VistaCapturaRapida para la ingesta rápida de avances y justificaciones por Dirección y Área."),
        ("src/gui/components.py", "Componentes visuales reutilizables (Tarjetas de KPIs globales y controles de exportación)."),
        ("src/core/excel_engine.py", "Motor de lectura, consolidación matricial y recálculo de sábanas en memoria."),
        ("src/core/generador_doc.py", "Generador de reportes .docx con inyección sintáctica, tablas formateadas por Dirección e IA."),
        ("src/core/generador_excel_pbrm.py", "Motor de escritura directa mediante openpyxl que preserva la integridad de fórmulas primarias."),
        ("src/core/integracion_gemini.py", "Módulo de conexión con Gemini API mediante entorno seguro para análisis cualitativo contextualized."),
        ("src/core/lector_excel.py", "Parsing especializado de hojas Excel, evitando ambigüedades lógicas en DataFrames de Pandas."),
        ("src/core/lector_pdf.py", "Extractor de tablas y texto no estructurado desde reportes en formato PDF."),
        ("src/core/validador_reg.py", "Evaluador de desempeño programático bajo norma OSFEM (C, D, R, A, S)."),
        ("src/utils/helpers.py", "Manejo de rutas relativas/absolutas y depuración rotativa de respaldos y memorias Word."),
        ("src/utils/logger.py", "Sistema de logging y auditoría continua hacia logs/app.log.")
    ]

    for mod, desc in modulos:
        p_mod = doc.add_paragraph(style='List Bullet')
        r_mod_name = p_mod.add_run(f"{mod}: ")
        r_mod_name.bold = True
        p_mod.add_run(desc)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # --- SECCIÓN 6: INTEGRACIÓN DE IA Y RESILIENCIA ---
    h6_ia = doc.add_heading("6. Arquitectura de Inteligencia Artificial y Resiliencia", level=1)
    h6_ia.runs[0].font.color.rgb = AZUL_OPERAGUA

    p_ia = doc.add_paragraph(
        "El módulo de IA (integracion_gemini.py) implementa una estrategia de análisis cualitativo avanzado:\n"
        "1. Carga Segura de Credenciales: Lee la clave GEMINI_API_KEY desde el archivo .env sin exponerla en la interfaz gráfica.\n"
        "2. Ingesta Multitrimestral de Contexto: Examina los últimos 5 archivos .docx almacenados en respaldos/contexto_word/ "
        "para identificar si las justificaciones actuales corresponden a problemas sistémicos o incidentes aislados.\n"
        "3. Fallback y Resiliencia Institucional: Ante fallas de conexión o límites de cuota (HTTP 429/503), el sistema activa "
        "un motor de redacción sintética basada en reglas que genera el informe Word sin detener la operación."
    )
    p_ia.paragraph_format.line_spacing = 1.15
    p_ia.paragraph_format.space_after = Pt(10)

    # --- SECCIÓN 7: MATRIZ DE CONTROL DE ERRORES ---
    h7_err = doc.add_heading("7. Matriz de Control de Errores y Calidad", level=1)
    h7_err.runs[0].font.color.rgb = AZUL_OPERAGUA

    tabla_err = doc.add_table(rows=7, cols=3)
    corregir_ajuste_tabla(tabla_err)
    tabla_err.style = 'Table Grid'

    datos_err = [
        ("Escenario / Prueba", "Causa Raíz Detectada", "Estrategia de Solución Implementada"),
        ("Estructura Excel Inválida", "Pestañas o columnas renombradas manualmente por el usuario", "Validación previa en lector_excel.py con registro limpio en logs/ sin caídas de la app."),
        ("Evaluación Ambigua Pandas", "Uso directo de condiciones booleanas (if df:) sobre DataFrames", "Uso explícito de .empty e isinstance() para validaciones numéricas y matriciales."),
        ("Interrupción de Gemini API", "Falta de conexión a red o API Key no configurada en .env", "Fallback automático a plantilla de redacción cualitativa estándar UIPPE."),
        ("Ruptura de Fórmulas Excel", "Escritura directa sobre celdas calculadas durante inyección", "Inyección controlada celda a celda mediante openpyxl respetando celdas de fórmula."),
        ("Reemplazo de Sábana Activa", "Cambio de ciclo fiscal sin reiniciar el aplicativo", "Función actualizar_excel_maestro() que copia, reemplaza y recarga instancias en caliente."),
        ("Saturación de Disco", "Acumulación indeterminada de reportes temporales e historiales", "Ejecución de depurar_contexto_historico() y depurar_historico_salidas() automática.")
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

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # --- SECCIÓN 8: ESTADO Y VALIDACIÓN ---
    h8 = doc.add_heading("8. Estado de Liberación y Control de Cambios", level=1)
    h8.runs[0].font.color.rgb = AZUL_OPERAGUA

    p_estado = doc.add_paragraph(
        "El proyecto se encuentra auditado y consolidado en su Fase 2. Las pruebas unitarias e integrales "
        "confirman el correcto flujo entre la captura, la inyección a la Sábana Maestra, el recálculo semafórico "
        "y la generación de salidas institucionales."
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
    print(f"✔ Documentación Técnica creada exitosamente en: {ruta_salida}")

if __name__ == "__main__":
    crear_documentacion()