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
    p_titulo.paragraph_format.space_before = Pt(20)
    p_titulo.paragraph_format.space_after = Pt(10)
    r_titulo = p_titulo.add_run("MANUAL DE USUARIO Y GUÍA DE OPERACIÓN")
    r_titulo.font.size = Pt(20)
    r_titulo.font.bold = True
    r_titulo.font.color.rgb = AZUL_OPERAGUA

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_after = Pt(20)
    r_sub = p_sub.add_run(
        "Sistema de Automatización, Inyección y Semaforización PbRM"
    )
    r_sub.font.size = Pt(12)
    r_sub.font.italic = True
    r_sub.font.color.rgb = GRIS_TEXTO

    # --- SECCIÓN 1: INTRODUCCIÓN ---
    h1 = doc.add_heading("1. Introducción y Objetivo", level=1)
    h1.runs[0].font.color.rgb = AZUL_OPERAGUA

    p_body1 = doc.add_paragraph(
        "El presente manual tiene como objetivo guiar al personal analista y administrativo de la UIPPE en el uso "
        "del Sistema de Automatización PbRM. Esta herramienta facilita la captura de avances, la actualización de "
        "Sábanas Maestras de Excel, el cálculo automatizado del semáforo de cumplimiento programático y la "
        "generación de reportes ejecutivos en Word respaldados por Inteligencia Artificial."
    )
    p_body1.paragraph_format.line_spacing = 1.15
    p_body1.paragraph_format.space_after = Pt(12)

    # --- SECCIÓN 2: INICIO Y CONFIGURACIÓN INICIAL ---
    h2 = doc.add_heading("2. Inicio del Sistema y Configuración Global", level=1)
    h2.runs[0].font.color.rgb = AZUL_OPERAGUA

    pasos_inicio = [
        ("Apertura del Aplicativo", "Ejecute el archivo principal (.exe o script). Al iniciar, la aplicación cargará automáticamente las credenciales seguras y la estructura básica del proyecto."),
        ("Selección del Ciclo Fiscal (Año)", "En la barra superior de la pantalla principal, seleccione el año de gestión (ej. 2026). Esto actualizará automáticamente la Sábana Maestra vinculada."),
        ("Fecha de Emisión", "Haga clic en el campo de fecha o presione el botón del calendario (CTkDatePicker) para seleccionar la fecha de corte oficial para los reportes."),
        ("Reemplazo de Sábana Maestra (Si aplica)", "Si inicia un nuevo año fiscal o requiere actualizar el archivo base, presione el botón 'Reemplazar Sábana Maestra', seleccione el archivo .xlsx correspondiente y confirme el reemplazo.")
    ]

    for titulo, desc in pasos_inicio:
        p = doc.add_paragraph(style='List Bullet')
        r_t = p.add_run(f"{titulo}: ")
        r_t.bold = True
        p.add_run(desc)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # --- SECCIÓN 3: PESTAÑA DE CAPTURA RÁPIDA ---
    h3 = doc.add_heading("3. Módulo de Captura Rápida de Avances", level=1)
    h3.runs[0].font.color.rgb = AZUL_OPERAGUA

    p_captura = doc.add_paragraph(
        "En la pestaña 'Captura de Avances' podrá registrar las metas reportadas por cada unidad administrativa:"
    )

    pasos_captura = [
        ("Selección de Área", "Elija la Dirección y el Área correspondiente utilizando los menús desplegables."),
        ("Ingreso de Avance", "Introduzca el valor del avance alcanzado en el trimestre. El sistema actualizará inmediatamente la semaforización de la meta."),
        ("Justificación Cualitativa", "En caso de desfases o desviaciones significativas en la meta, capture la justificación administrativa correspondiente en la caja de texto."),
        ("Inyección a Excel", "Presione el botón 'Guardar / Inyectar en Excel'. Los datos se escribirán directamente en la Sábana Maestra preservando todas las fórmulas y cálculos existentes.")
    ]

    for titulo, desc in pasos_captura:
        p = doc.add_paragraph(style='List Bullet')
        r_t = p.add_run(f"{titulo}: ")
        r_t.bold = True
        p.add_run(desc)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # --- SECCIÓN 4: CONSOLIDADOR Y GENERACIÓN DE REPORTES ---
    h4 = doc.add_heading("4. Módulo Consolidador y Generación de Reportes", level=1)
    h4.runs[0].font.color.rgb = AZUL_OPERAGUA

    p_cons = doc.add_paragraph(
        "En la pestaña 'Consolidador y KPIs' podrá evaluar el avance general del organismo y emitir los informes institucionales:"
    )

    pasos_cons = [
        ("Vista de KPIs", "Visualice las tarjetas resumen que indican el porcentaje global de cumplimiento y la distribución de metas por color de semáforo."),
        ("Análisis con IA (Gemini)", "Active la casilla de evaluación con Inteligencia Artificial. El sistema analizará el desempeño actual comparándolo con la memoria histórica acumulada (hasta 5 trimestres previos)."),
        ("Generación de Reporte Word", "Presione 'Generar Informe Word'. Se creará un documento .docx formateado con tablas institucionales, gráficos cualitativos y observaciones redactadas."),
        ("Ubicación de Salida", "Todos los archivos generados se guardarán automáticamente en la carpeta 'salidas/' con respaldo de seguridad.")
    ]

    for titulo, desc in pasos_cons:
        p = doc.add_paragraph(style='List Bullet')
        r_t = p.add_run(f"{titulo}: ")
        r_t.bold = True
        p.add_run(desc)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # --- SECCIÓN 5: INTERPRETACIÓN DE SEMÁFOROS ---
    h5 = doc.add_heading("5. Interpretación de la Semaforización (OSFEM / PbRM)", level=1)
    h5.runs[0].font.color.rgb = AZUL_OPERAGUA

    tabla_sem = doc.add_table(rows=6, cols=3)
    corregir_ajuste_tabla(tabla_sem)
    tabla_sem.style = 'Table Grid'

    datos_sem = [
        ("Semáforo", "Rango de Avance", "Significado / Acción Requerida"),
        ("Crítico (Rojo)", "< 70.0%", "Incumplimiento severo. Requiere justificación obligatoria e intervención."),
        ("Deficiente (Naranja)", "70.0% - 84.9%", "Subejercicio. Requiere plan de regularización."),
        ("Regular (Amarillo)", "85.0% - 94.9%", "Cumplimiento aceptable con margen de mejora."),
        ("Adecuado (Verde)", "95.0% - 110.0%", "Cumplimiento óptimo de lo programado."),
        ("Sobrepasado (Azul)", "> 110.0%", "Meta rebasada. Requiere revisión de programación inicial.")
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

    # --- SECCIÓN 6: PREGUNTAS FRECUENTES Y SOLUCIÓN DE PROBLEMAS ---
    h6 = doc.add_heading("6. Solución de Problemas Frecuentes", level=1)
    h6.runs[0].font.color.rgb = AZUL_OPERAGUA

    faqs = [
        ("¿Qué pasa si falla la conexión a Internet durante el análisis de IA?", "El sistema cuenta con un respaldo automático. Generará el reporte en Word utilizando plantillas de redacción institucional sin interrumpir el proceso."),
        ("¿El sistema borra los reportes de trimestres anteriores?", "No los elimina por completo; mantiene de forma automática los 5 reportes trimestrales más recientes como memoria contextual en 'respaldos/contexto_word/' para no saturar el equipo."),
        ("¿Puedo modificar la Sábana Maestra directamente en Excel?", "Sí, pero se recomienda realizar las capturas desde el sistema para mantener la coherencia de datos y el registro en la bitácora de eventos.")
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