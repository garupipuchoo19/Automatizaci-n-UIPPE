import os
import re
from datetime import datetime
import pandas as pd
import docx
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

try:
    import matplotlib.pyplot as plt
    MATPLOTLIB_DISPONIBLE = True
except ImportError:
    MATPLOTLIB_DISPONIBLE = False

from src.utils.helpers import obtener_nombre_salida, obtener_fecha_formal


def aplicar_sombreado(celda, color_hex):
    """Aplica un color de fondo (HEX) a una celda de tabla en Word."""
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    celda._tc.get_or_add_tcPr().append(shading_elm)


def corregir_ajuste_tabla(tabla):
    """Fuerza a la tabla a estar 'En línea con el texto' y centrada respetando la plantilla."""
    tblPr = tabla._tbl.tblPr
    tblOverlap = tblPr.find(qn('w:tblOverlap'))
    if tblOverlap is not None:
        tblPr.remove(tblOverlap)
    tabla.alignment = WD_TABLE_ALIGNMENT.CENTER


def agregar_parrafo_con_markdown(doc_o_celda, texto_markdown):
    """Convierte sintaxis básica de Markdown (**negrita**, *cursiva*) respetando el formato."""
    p = doc_o_celda.add_paragraph()
    p.paragraph_format.line_spacing = 1.15
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.space_before = Pt(0)

    patron = re.compile(r'(\*\*.*?\*\*|\*.*?\*)')
    partes = patron.split(texto_markdown)

    for parte in partes:
        if not parte:
            continue
        if parte.startswith('**') and parte.endswith('**'):
            run = p.add_run(parte[2:-2])
            run.font.bold = True
        elif parte.startswith('*') and parte.endswith('*'):
            run = p.add_run(parte[1:-1])
            run.font.italic = True
        else:
            p.add_run(parte)

    return p


def validar_plantilla_word(ruta_plantilla, marcadores_requeridos=None):
    """Valida si la plantilla Word existe y contiene marcadores de inyección."""
    if marcadores_requeridos is None:
        marcadores_requeridos = ["{TITULO_REPORTE}", "{TRIMESTRE}", "{ANALISIS_CUALITATIVO}"]

    if not ruta_plantilla or not os.path.exists(ruta_plantilla):
        return True, "No se especificó plantilla existente. Se generará un documento desde cero."

    try:
        doc = docx.Document(ruta_plantilla)
        texto_completo = []
        for p in doc.paragraphs:
            texto_completo.append(p.text)
        for tabla in doc.tables:
            for fila in tabla.rows:
                for celda in fila.cells:
                    texto_completo.append(celda.text)
                    
        contenido = " ".join(texto_completo)
        faltantes = [m for m in marcadores_requeridos if m not in contenido]
        
        if faltantes:
            return False, f"Atención: Faltan marcadores en la plantilla Word: {', '.join(faltantes)}"
            
        return True, "Plantilla válida."
        
    except Exception as e:
        return False, f"Error al abrir la plantilla Word: {str(e)}"


class GeneradorReporte:
    def __init__(self, dataframe_procesado=None, ruta_plantilla_word=None, trimestre_num=2, anio=None, fecha_emision=None):
        if dataframe_procesado is not None and isinstance(dataframe_procesado, pd.DataFrame):
            self.df = dataframe_procesado
        else:
            self.df = pd.DataFrame()
            
        self.ruta_plantilla = ruta_plantilla_word
        self.trimestre_num = trimestre_num
        self.anio = anio if anio else datetime.now().year
        self.fecha_emision = fecha_emision if fecha_emision else datetime.now().strftime("%d/%m/%Y")

        self.AZUL_OPERAGUA = RGBColor(31, 78, 120)
        self.HEX_AZUL = "1F4E78"
        self.HEX_GRIS_CLARO = "F2F2F2"

        self.SEMAFORO_PBRM = {
            'C': {'nombre': 'Crítico', 'hex': '#E74C3C'},
            'D': {'nombre': 'Deficiente', 'hex': '#E67E22'},
            'R': {'nombre': 'Regular', 'hex': '#F1C40F'},
            'A': {'nombre': 'Adecuado', 'hex': '#27AE60'},
            'S': {'nombre': 'Sobrepasado', 'hex': '#2980B9'}
        }

    def exportar_excel_consolidado(self, prefijo="Reporte_UIPPE_Consolidado"):
        prefijo_trim = f"{prefijo}_T{self.trimestre_num}_{self.anio}"
        ruta_salida = obtener_nombre_salida(prefix=prefijo_trim, extension="xlsx")
        with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
            self.df.to_excel(writer, index=False, sheet_name="Reporte UIPPE")
        return ruta_salida

    def generar_grafica_semaforo(self, conteos_dict, titulo_grafica="Distribución por Semáforo"):
        if not MATPLOTLIB_DISPONIBLE:
            return None

        categorias = ['Crítico (C)', 'Deficiente (D)', 'Regular (R)', 'Adecuado (A)', 'Sobrepasado (S)']
        claves = ['C', 'D', 'R', 'A', 'S']
        valores = [conteos_dict.get(k, 0) for k in claves]
        colores = [self.SEMAFORO_PBRM[k]['hex'] for k in claves]

        plt.figure(figsize=(6, 2.8))
        bars = plt.bar(categorias, valores, color=colores, edgecolor='black', linewidth=0.6)

        plt.title(titulo_grafica, fontsize=9, fontweight='bold', pad=8, color='#1F4E78')
        plt.ylabel("Cantidad", fontsize=8)
        plt.xticks(fontsize=7.5, rotation=10)
        plt.yticks(fontsize=7.5)
        plt.grid(axis='y', linestyle='--', alpha=0.4)

        for bar in bars:
            yval = bar.get_height()
            if yval > 0:
                plt.text(bar.get_x() + bar.get_width()/2.0, yval + 0.1, str(int(yval)), ha='center', va='bottom', fontsize=8, fontweight='bold')

        plt.tight_layout()
        
        ruta_img = f"temp_grafica_{datetime.now().strftime('%H%M%S%f')}.png"
        plt.savefig(ruta_img, dpi=180)
        plt.close()
        return ruta_img

    def construir_tabla_matriz_itsp(self, doc, titulo_subseccion, datos_unidades, estatus_global, pct_global, estilo_tabla=None, generar_grafica=False):
        p_sub = doc.add_paragraph()
        p_sub.paragraph_format.space_before = Pt(8)
        p_sub.paragraph_format.space_after = Pt(4)
        r_sub = p_sub.add_run(titulo_subseccion)
        r_sub.bold = True
        r_sub.font.size = Pt(10)

        tabla = doc.add_table(rows=1, cols=7)
        corregir_ajuste_tabla(tabla)
        
        # Asignar estilo nativo de la plantilla si existe
        if estilo_tabla:
            tabla.style = estilo_tabla
        else:
            tabla.style = 'Table Grid'

        anchos = [3.0, 0.5, 0.5, 0.5, 0.5, 0.5, 1.0]
        hdr_cells = tabla.rows[0].cells
        headers = ["Unidad Administrativa", "C", "D", "R", "A", "S", "VALORACIÓN %"]
        
        conteos_totales = {'C': 0, 'D': 0, 'R': 0, 'A': 0, 'S': 0}

        for idx, text in enumerate(headers):
            hdr_cells[idx].text = text
            hdr_cells[idx].width = Inches(anchos[idx])
            aplicar_sombreado(hdr_cells[idx], self.HEX_AZUL)
            p = hdr_cells[idx].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if p.runs:
                p.runs[0].font.bold = True
                p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
                p.runs[0].font.size = Pt(9)

        for u in datos_unidades:
            row_cells = tabla.add_row().cells
            row_cells[0].text = str(u.get("unidad", ""))
            
            for idx, k in enumerate(['C', 'D', 'R', 'A', 'S'], start=1):
                val = u.get(k, "") or ""
                row_cells[idx].text = str(val)
                if val:
                    try:
                        conteos_totales[k] += int(val)
                    except ValueError:
                        pass

            val_pct = u.get('valoracion', 0.0)
            row_cells[6].text = f"{val_pct:.2f}%" if isinstance(val_pct, (int, float)) else str(val_pct)

            for i, w in enumerate(anchos):
                row_cells[i].width = Inches(w)
                p = row_cells[i].paragraphs[0]
                if p.runs:
                    p.runs[0].font.size = Pt(8.5)
                if i > 0:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        row_tot = tabla.add_row().cells
        row_tot[0].text = "DESEMPEÑO A NIVEL DE DIRECCIÓN"
        row_tot[1].text = str(estatus_global).upper()
        row_tot[6].text = f"{pct_global:.2f}%" if isinstance(pct_global, (int, float)) else str(pct_global)

        for i, w in enumerate(anchos):
            row_tot[i].width = Inches(w)
            aplicar_sombreado(row_tot[i], self.HEX_GRIS_CLARO)
            p = row_tot[i].paragraphs[0]
            if p.runs:
                p.runs[0].font.bold = True
                p.runs[0].font.size = Pt(9)
            if i in [1, 6]:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        if generar_grafica and MATPLOTLIB_DISPONIBLE:
            try:
                ruta_grafica = self.generar_grafica_semaforo(conteos_totales, f"Semaforización - {titulo_subseccion.replace(':', '')}")
                if ruta_grafica and os.path.exists(ruta_grafica):
                    p_img = doc.add_paragraph()
                    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_img.paragraph_format.space_before = Pt(4)
                    p_img.paragraph_format.space_after = Pt(8)
                    p_img.add_run().add_picture(ruta_grafica, width=Inches(5.0))
                    os.remove(ruta_grafica)
            except Exception:
                pass

        doc.add_paragraph().paragraph_format.space_after = Pt(6)

    def exportar_word_ejecutivo(self, titulo=None, datos_direcciones=None, analisis_cualitativo=None, incluir_graficas=False):
        if not titulo:
            titulo = f"INFORME TRIMESTRAL DE SEGUIMIENTO PROGRAMÁTICO (ITSP)\n{self.trimestre_num}° TRIMESTRE {self.anio}"

        prefijo_salida = f"Reporte_Ejecutivo_ITSP_T{self.trimestre_num}_{self.anio}"
        ruta_salida = obtener_nombre_salida(prefix=prefijo_salida, extension="docx")
        
        # Cargar documento base desde carpeta plantillas/
        estilo_tabla_plantilla = None
        if self.ruta_plantilla and os.path.exists(self.ruta_plantilla):
            doc = Document(self.ruta_plantilla)
            if len(doc.tables) > 0:
                estilo_tabla_plantilla = doc.tables[0].style
        else:
            doc = Document()

        for t in doc.tables:
            corregir_ajuste_tabla(t)

        mapa_ordinales = {1: "PRIMER", 2: "SEGUNDO", 3: "TERCER", 4: "CUARTO"}
        texto_ordinal = mapa_ordinales.get(self.trimestre_num, "SEGUNDO")

        def reemplazar_marcadores_texto(texto):
            texto = texto.replace("{TITULO_REPORTE}", titulo)
            texto = texto.replace("{TRIMESTRE}", f"{self.trimestre_num}° TRIMESTRE")
            texto = texto.replace("{ANIO}", str(self.anio))
            texto = texto.replace("{FECHA_EMISION}", self.fecha_emision)
            
            for anio_fijo in ["2026", "2025", "2024"]:
                if str(self.anio) != anio_fijo:
                    texto = texto.replace(anio_fijo, str(self.anio))
                    
            if "SEGUNDO TRIMESTRE" in texto and texto_ordinal != "SEGUNDO":
                texto = texto.replace("SEGUNDO TRIMESTRE", f"{texto_ordinal} TRIMESTRE")
            if "2° TRIMESTRE" in texto and self.trimestre_num != 2:
                texto = texto.replace("2° TRIMESTRE", f"{self.trimestre_num}° TRIMESTRE")
            return texto

        # Sustitución limpia respetando formato de párrafos de la plantilla
        analisis_reemplazado = False
        for p in doc.paragraphs:
            if "{ANALISIS_CUALITATIVO}" in p.text:
                p.text = ""
                if analisis_cualitativo:
                    lineas = [l.strip() for l in analisis_cualitativo.split("\n") if l.strip()]
                    for l in lineas:
                        agregar_parrafo_con_markdown(doc, l)
                analisis_reemplazado = True
            elif any(m in p.text for m in ["{TITULO_REPORTE}", "{TRIMESTRE}", "{ANIO}", "{FECHA_EMISION}"]):
                # Preserva los runs originales y reemplaza dentro de los mismos para NO perder la fuente/color de la plantilla
                for r in p.runs:
                    r.text = reemplazar_marcadores_texto(r.text)

        # Reemplazo en las celdas de las tablas de la plantilla
        for t in doc.tables:
            for fila in t.rows:
                for celda in fila.cells:
                    for p in celda.paragraphs:
                        for r in p.runs:
                            r.text = reemplazar_marcadores_texto(r.text)

        if analisis_cualitativo and not analisis_reemplazado:
            h_analisis = doc.add_heading(f"Análisis Cualitativo Ejecutivo - {self.trimestre_num}° Trimestre {self.anio}", level=2)
            if h_analisis.runs:
                h_analisis.runs[0].font.color.rgb = self.AZUL_OPERAGUA
                h_analisis.runs[0].font.size = Pt(13)
            
            lineas = [l.strip() for l in analisis_cualitativo.split("\n") if l.strip()]
            for l in lineas:
                agregar_parrafo_con_markdown(doc, l)

        tiene_direcciones = False
        lista_direcciones = []

        if isinstance(datos_direcciones, pd.DataFrame):
            if not datos_direcciones.empty:
                tiene_direcciones = True
                col_dir = 'clv' if 'clv' in datos_direcciones.columns else datos_direcciones.columns[0]
                for direccion, grupo in datos_direcciones.groupby(col_dir):
                    metas_u = []
                    for _, row in grupo.iterrows():
                        metas_u.append({
                            "unidad": row.get('Etiquetas de fila', ''),
                            "C": row.get('CRÍTICO', ''),
                            "D": row.get('DEFICIENTE', ''),
                            "R": row.get('REGULAR', ''),
                            "A": row.get('ADECUADO', ''),
                            "S": row.get('SOBREPASADO', ''),
                            "valoracion": row.get('PORCENTAJE', 0.0)
                        })
                    
                    pct_global = round(grupo['PORCENTAJE'].mean(), 2) if 'PORCENTAJE' in grupo.columns else 0.0
                    estatus_global = grupo['NIVEL'].iloc[0] if 'NIVEL' in grupo.columns else "ADECUADO"
                    
                    lista_direcciones.append({
                        "nombre": str(direccion),
                        "metas_unidades": metas_u,
                        "estatus_metas": estatus_global,
                        "valoracion_metas_global": pct_global
                    })
        elif isinstance(datos_direcciones, list) and len(datos_direcciones) > 0:
            tiene_direcciones = True
            lista_direcciones = datos_direcciones

        if tiene_direcciones:
            h_res = doc.add_heading("RESULTADOS DEL SEGUIMIENTO PROGRAMÁTICO", level=1)
            if h_res.runs:
                h_res.runs[0].font.color.rgb = self.AZUL_OPERAGUA

            for dir_data in lista_direcciones:
                p_dir = doc.add_paragraph()
                p_dir.paragraph_format.space_before = Pt(12)
                r_dir = p_dir.add_run(dir_data.get("nombre", "").upper())
                r_dir.font.bold = True
                r_dir.font.size = Pt(12)
                r_dir.font.color.rgb = self.AZUL_OPERAGUA

                self.construir_tabla_matriz_itsp(
                    doc=doc,
                    titulo_subseccion="Resultados Trimestrales de Metas:",
                    datos_unidades=dir_data.get("metas_unidades", []),
                    estatus_global=dir_data.get("estatus_metas", "ADECUADO"),
                    pct_global=dir_data.get("valoracion_metas_global", 0.0),
                    estilo_tabla=estilo_tabla_plantilla,
                    generar_grafica=incluir_graficas
                )

                if "indicadores_unidades" in dir_data:
                    self.construir_tabla_matriz_itsp(
                        doc=doc,
                        titulo_subseccion="Resultados Trimestrales de Indicadores:",
                        datos_unidades=dir_data.get("indicadores_unidades", []),
                        estatus_global=dir_data.get("estatus_indicadores", "ADECUADO"),
                        pct_global=dir_data.get("valoracion_indicadores_global", 0.0),
                        estilo_tabla=estilo_tabla_plantilla,
                        generar_grafica=incluir_graficas
                    )

        os.makedirs(os.path.dirname(ruta_salida), exist_ok=True)
        doc.save(ruta_salida)
        return ruta_salida