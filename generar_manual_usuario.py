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

def crear_manual_usuario():
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
    r_titulo = p_titulo.add_run("MANUAL DE USUARIO Y GUÍA OPERATIVA DE CAMPO")
    r_titulo.font.size = Pt(20)
    r_titulo.font.bold = True
    r_titulo.font.color.rgb = AZUL_OPERAGUA

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_after = Pt(18)
    r_sub = p_sub.add_run(
        "Sistema de Automatización, Inyección y Semaforización PbRM\n"
        "Guía Paso a Paso para la Captura, Evaluación y Generación de Informes"
    )
    r_sub.font.size = Pt(11)
    r_sub.font.italic = True
    r_sub.font.color.rgb = GRIS_TEXTO

    # --- SECCIÓN 1: INTRODUCCIÓN Y OBJETIVO ---
    h1 = doc.add_heading("1. Introducción y Objetivo de la Guía", level=1)
    h1.runs[0].font.color.rgb = AZUL_OPERAGUA

    p_body1 = doc.add_paragraph(
        "Este manual de usuario ofrece una guía clara y paso a paso para el personal operativo, analistas y "
        "directivos de la UIPPE en OPERAGUA Cuautitlán Izcalli. Su propósito es orientar el uso del sistema "
        "para realizar la captura rápida de avances trimestrales, actualizar la Sábana Maestra de Excel sin alterar sus fórmulas, "
        "evaluar la semaforización oficial según la normativa del OSFEM y emitir reportes ejecutivos en Microsoft Word respaldados por Inteligencia Artificial."
    )
    p_body1.paragraph_format.line_spacing = 1.15
    p_body1.paragraph_format.space_after = Pt(10)

    # --- SECCIÓN 2: FLUJO DE TRABAJO RECOMENDADO ---
    h2_flujo = doc.add_heading("2. Flujo de Trabajo Recomendado (Paso a Paso)", level=1)
    h2_flujo.runs[0].font.color.rgb = AZUL_OPERAGUA

    p_flujo_intro = doc.add_paragraph(
        "Para garantizar la consistencia en el seguimiento del PbRM, se recomienda seguir el siguiente orden operativo:"
    )

    pasos_flujo = [
        ("Paso 1: Configurar Ciclo y Fecha", "Verificar que el ciclo fiscal (año) y la fecha de corte seleccionados coincidan con el periodo de evaluación."),
        ("Paso 2: Validar o Reemplazar Sábana Maestra", "Si inicia un nuevo ciclo anual o el área de planeación actualizó la matriz base, cargar la nueva Sábana Maestra desde la interfaz."),
        ("Paso 3: Capturar Avances y Justificaciones", "Seleccionar Dirección/Área, registrar el avance cuantitativo y capturar la justificación en caso de desvíos en la meta."),
        ("Paso 4: Inyectar Datos en Excel", "Guardar e inyectar la información. El sistema escribirá directamente en el Excel respetando las fórmulas integradas."),
        ("Paso 5: Revisar KPIs y Consolidado", "Verificar la distribución del semáforo global (Crítico, Deficiente, Regular, Adecuado, Sobrepasado) en la pestaña Consolidador."),
        ("Paso 6: Generar Reporte Ejecutivo Word", "Activar el análisis con IA (Gemini) para evaluar el contexto histórico y exportar el reporte institucional listo para firma.")
    ]

    for titulo, desc in pasos_flujo:
        p = doc.add_paragraph(style='List Bullet')
        r_t = p.add_run(f"{titulo}: ")
        r_t.bold = True
        p.add_run(desc)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # --- SECCIÓN 3: PANTALLA PRINCIPAL Y CONFIGURACIÓN ---
    h3_gui = doc.add_heading("3. Interfaz Gráfica y Opciones Globales", level=1)
    h3_gui.runs[0].font.color.rgb = AZUL_OPERAGUA

    p_gui = doc.add_paragraph(
        "La interfaz principal integra un panel superior de control global y un visor de bitácora en tiempo real:"
    )

    elementos_gui = [
        ("Selector de Ciclo Fiscal (Año)", "Menú desplegable que permite cambiar entre años de gestión. Ajusta automáticamente la Sábana Maestra vinculada (ej. CUAUTITLAN METAS E INDICADORES 2026.xlsx)."),
        ("Selector de Fecha (CTkDatePicker)", "Campo con calendario desplegable integrado para definir la fecha oficial de emisión del reporte que figurará en el encabezado del documento Word."),
        ("Botón Reemplazar Sábana Maestra", "Permite seleccionar un archivo Excel externo (.xlsx) desde el explorador de archivos. El sistema copia la nueva matriz a entradas/ y actualiza las instancias activas."),
        ("Visor de Bitácora y Logs", "Consola inferior que muestra la confirmación de cada acción, advertencias de validación o alertas en caso de algún archivo faltante.")
    ]

    for elem, e_desc in elementos_gui:
        p_e = doc.add_paragraph(style='List Bullet')
        r_en = p_e.add_run(f"{elem}: ")
        r_en.bold = True
        p_e.add_run(e_desc)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # --- SECCIÓN 4: CAPTURA RÁPIDA E INYECCIÓN EN EXCEL ---
    h4_cap = doc.add_heading("4. Módulo de Captura Rápida de Avances", level=1)
    h4_cap.runs[0].font.color.rgb = AZUL_OPERAGUA

    p_cap = doc.add_paragraph(
        "En la pestaña 'Captura de Avances', el usuario interactúa de forma directa con los indicadores programados:"
    )

    pasos_captura_det = [
        ("Filtro por Dirección y Área", "Al seleccionar la Dirección, el segundo menú cargará únicamente las Áreas que le corresponden."),
        ("Muestra de Metas Programadas", "El sistema muestra de manera automática la meta programada del trimestre para brindar una referencia visual inmediata."),
        ("Ingreso de Avance y Semáforo en Vivo", "Al escribir el valor alcanzado, la aplicación calcula al instante el porcentaje de cumplimiento y asigna el color de semáforo OSFEM correspondiente."),
        ("Captura de Justificación Cualitativa", "Si la meta presenta un avance en semáforo Crítico, Deficiente o Sobrepasado, se habilita el cuadro de texto para ingresar las causas y medidas correctivas."),
        ("Inyección Segura", "Al hacer clic en 'Guardar / Inyectar en Excel', los cambios se guardan directamente en el archivo .xlsx de entradas/ sin sobreescribir las celdas con fórmulas primarias.")
    ]

    for titulo, desc in pasos_captura_det:
        p = doc.add_paragraph(style='List Bullet')
        r_t = p.add_run(f"{titulo}: ")
        r_t.bold = True
        p.add_run(desc)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # --- SECCIÓN 5: SEMAFORIZACIÓN NORMATIVA Y INTERPRETACÓN ---
    h5_sem = doc.add_heading("5. Criterios de Semaforización (Normativa OSFEM / PbRM)", level=1)
    h5_sem.runs[0].font.color.rgb = AZUL_OPERAGUA

    tabla_sem = doc.add_table(rows=6, cols=3)
    corregir_ajuste_tabla(tabla_sem)
    tabla_sem.style = 'Table Grid'

    datos_sem = [
        ("Semáforo", "Rango de Avance", "Significado y Acción Requerida"),
        ("Crítico (Rojo)", "< 70.0%", "Incumplimiento severo. Requiere justificación obligatoria y plan de contingencia."),
        ("Deficiente (Naranja)", "70.0% - 84.9%", "Subejercicio significativo. Se debe justificar la desviación en la captura."),
        ("Regular (Amarillo)", "85.0% - 94.9%", "Cumplimiento cercano a la meta. Se encuentra dentro del rango de tolerancia."),
        ("Adecuado (Verde)", "95.0% - 110.0%", "Cumplimiento óptimo de lo estimado. No requiere justificación adiciona."),
        ("Sobrepasado (Azul)", "> 110.0%", "Meta rebasada. Requiere aclaración técnica por rebasamiento de lo programado.")
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

    # --- SECCIÓN 6: CONSOLIDADOR Y REPORTES CON IA ---
    h6_cons = doc.add_heading("6. Consolidador Ejecutivo y Generación de Reportes Word", level=1)
    h6_cons.runs[0].font.color.rgb = AZUL_OPERAGUA

    p_cons = doc.add_paragraph(
        "En la pestaña 'Consolidador y KPIs', el usuario puede revisar las métricas globales del organismo y emitir el informe final:"
    )

    pasos_cons_det = [
        ("Resumen de Tarjetas KPI", "Visualización instantánea del porcentaje general de cumplimiento del organismo y conteo de metas por color de semáforo."),
        ("Análisis Cualitativo con IA (Gemini API)", "Al presionar 'Generar Informe Word', el sistema evalúa la información actual junto con la memoria contextual acumulada en respaldos/contexto_word/ (hasta 5 trimestres anteriores)."),
        ("Redacción Autónoma de Conclusiones", "La IA redacta de forma ejecutiva la síntesis del trimestre, destacando áreas con avance sobresaliente, causas comunes de retraso y recomendaciones operativas."),
        ("Ubicación del Documento Generado", "El reporte final se deposita en la carpeta salidas/ en formato Word (.docx) aplicando los colores y estilos institucionales de OPERAGUA.")
    ]

    for titulo, desc in pasos_cons_det:
        p = doc.add_paragraph(style='List Bullet')
        r_t = p.add_run(f"{titulo}: ")
        r_t.bold = True
        p.add_run(desc)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # --- SECCIÓN 7: SOLUCIÓN DE PROBLEMAS Y PREGUNTAS FRECUENTES ---
    h7_faq = doc.add_heading("7. Solución de Problemas Frecuentes (FAQ)", level=1)
    h7_faq.runs[0].font.color.rgb = AZUL_OPERAGUA

    faqs = [
        ("¿Qué sucede si no hay conexión a Internet o falla la API Key de Gemini?", "El sistema cuenta con un motor de respaldo (fallback) que detecta la interrupción y genera el informe Word utilizando plantillas de redacción estandarizada sin detener el proceso ni perder datos."),
        ("¿Cómo se configuran las credenciales de IA (API Key)?", "La clave GEMINI_API_KEY se almacena de forma segura en un archivo de configuración .env en la raíz del programa, evitando tener que escribirla en la interfaz cada vez que se usa."),
        ("¿El programa borra o sobrescribe los reportes de trimestres anteriores?", "No los borra de forma permanente. El sistema aplica un ciclo de rotación que conserva automáticamente los 5 reportes trimestrales más recientes en respaldos/contexto_word/ como memoria contextual."),
        ("¿Qué hago si se actualizó la Sábana Maestra durante el proceso?", "Haga clic en el botón 'Reemplazar Sábana Maestra' en la barra superior de la app, seleccione el nuevo archivo Excel y confirme. El sistema recargará los menús y las referencias en tiempo real."),
        ("¿Es posible editar el reporte Word generado?", "Sí. El documento resultante en salidas/ es un archivo estándar de Microsoft Word (.docx) totalmente editable para ajustes de formato o firmas institucionales.")
    ]

    for preg, resp in faqs:
        p = doc.add_paragraph()
        r_p = p.add_run(f"• {preg}\n")
        r_p.bold = True
        r_p.font.color.rgb = AZUL_OPERAGUA
        p.add_run(resp)
        p.paragraph_format.space_after = Pt(6)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # Pie de página
    section = doc.sections[0]
    footer = section.footer
    p_ftr = footer.paragraphs[0]
    p_ftr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_ftr = p_ftr.add_run("Manual de Usuario | OPERAGUA UIPPE")
    r_ftr.font.size = Pt(8)
    r_ftr.font.color.rgb = GRIS_TEXTO

    # Guardar documento en la carpeta salidas/
    ruta_salida = os.path.join(os.path.dirname(__file__), "salidas", "Manual_de_Usuario_UIPPE.docx")
    os.makedirs(os.path.dirname(ruta_salida), exist_ok=True)
    doc.save(ruta_salida)
    print(f"✔ Manual de Usuario creado exitosamente en: {ruta_salida}")

if __name__ == "__main__":
    crear_manual_usuario()