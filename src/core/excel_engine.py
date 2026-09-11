import os
import pandas as pd
import openpyxl
from google import genai
from src.utils.helpers import ENTRADAS_DIR, obtener_ruta_entrada

class EngineExcelUIPPE:
    """Motor para leer, validar y calcular la sábana consolidada en Excel."""

    def __init__(self, ruta_excel_maestro):
        # Si recibe solo el nombre del archivo, resuelve la ruta completa dentro de entradas/
        if not os.path.isabs(ruta_excel_maestro):
            self.ruta_excel = obtener_ruta_entrada(ruta_excel_maestro)
        else:
            self.ruta_excel = ruta_excel_maestro

    def validar_captura_trimestre(self, trimestre: int):
        """
        [Opción 1] Valida que el trimestre objetivo tenga avances capturados en Excel.
        Previene falsos negativos/semáforos críticos si el periodo aún no concluye o no se ha reportado.
        """
        if not os.path.exists(self.ruta_excel):
            raise FileNotFoundError(f"No se encontró el archivo de insumo en: {self.ruta_excel}")

        df_raw = pd.read_excel(self.ruta_excel, sheet_name='METAS', header=None)
        df = df_raw.iloc[3:].copy()
        df.columns = df_raw.iloc[2].values

        # Mapeo de nombres de columna por trimestre
        col_progr = [c for c in df.columns if f"PROGR. {trimestre}" in str(c)][0]
        col_avance = [c for c in df.columns if f"AVANCE {trimestre}" in str(c)][0]

        df[col_progr] = pd.to_numeric(df[col_progr], errors='coerce').fillna(0)
        df[col_avance] = pd.to_numeric(df[col_avance], errors='coerce').fillna(0)

        total_programado = df[col_progr].sum()
        total_avance = df[col_avance].sum()

        if total_avance == 0 and total_programado > 0:
            raise ValueError(
                f"⚠️ VALIDACIÓN SUSPENDIDA: El Trimestre {trimestre} en '{os.path.basename(self.ruta_excel)}' "
                f"no registra avances capturados (suma de avances es 0.0)."
            )

        return True

    def procesar_hoja_metas(self, col_semaforo="2°_SEMAFORO"):
        """Lee la hoja METAS del Excel y calcula el concentrado por Área/Dirección."""
        if not os.path.exists(self.ruta_excel):
            raise FileNotFoundError(f"No se encontró el archivo en: {self.ruta_excel}")

        # Leer saltando los encabezados de adorno (header en fila 2)
        df_metas = pd.read_excel(self.ruta_excel, sheet_name='METAS', header=2)
        
        # Filtrar filas vacías o de totales
        df_metas = df_metas.dropna(subset=['CLAVE', 'ÁREA ENCARGADA'])

        resumen_areas = []

        # Agrupar por Dirección General y Área Encargada
        grupos = df_metas.groupby(['DEPENDENCIA GENERAL', 'ÁREA ENCARGADA'])

        for (direccion, area), grupo in grupos:
            conteos = grupo[col_semaforo].value_counts().to_dict()

            critico = conteos.get('CRÍTICO', 0)
            deficiente = conteos.get('DEFICIENTE', 0)
            regular = conteos.get('REGULAR', 0)
            adecuado = conteos.get('ADECUADO', 0)
            sobrepasado = conteos.get('SOBREPASADO', 0)

            total_metas = len(grupo)

            # Ponderación oficial UIPPE
            puntuacion = (sobrepasado * 10) + (adecuado * 10) + (regular * 8) + (deficiente * 6) + (critico * 0)
            
            porcentaje_area = (puntuacion / (total_metas * 10.0)) * 100 if total_metas > 0 else 0

            # Determinar Nivel de Desempeño
            if porcentaje_area >= 111.0: nivel = "SOBREPASADO"
            elif porcentaje_area >= 90.0: nivel = "ADECUADO"
            elif porcentaje_area >= 70.0: nivel = "REGULAR"
            elif porcentaje_area >= 50.0: nivel = "DEFICIENTE"
            else: nivel = "CRÍTICO"

            resumen_areas.append({
                'clv': direccion,
                'Etiquetas de fila': area,
                'CRÍTICO': critico if critico > 0 else None,
                'DEFICIENTE': deficiente if deficiente > 0 else None,
                'REGULAR': regular if regular > 0 else None,
                'ADECUADO': adecuado if adecuado > 0 else None,
                'SOBREPASADO': sobrepasado if sobrepasado > 0 else None,
                'PORCENTAJE': round(porcentaje_area, 2),
                'NO. METAS': total_metas,
                'PUNTUACIÓN': puntuacion,
                'NIVEL': nivel
            })

        return pd.DataFrame(resumen_areas)

    def actualizar_hoja_resultados(self, df_resultados, nombre_hoja="resultados general"):
        """Escribe el resumen recalculado de vuelta en el libro Excel Maestro."""
        with pd.ExcelWriter(self.ruta_excel, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df_resultados.to_excel(writer, sheet_name=nombre_hoja, index=False)


class GeminiAnalyst:
    """[Opción 3] Servicio de integración con la API de Gemini para análisis cualitativo."""

    def __init__(self, api_key: str):
        self.client = genai.Client(api_key=api_key)

    def generar_analisis_ejecutivo(self, trimestre: int, df_resumen_areas: pd.DataFrame, contexto_previo: str = None) -> str:
        """Sintetiza la tabla de resultados por área y solicita una redacción cualitativa ejecutiva."""
        # Validación segura del DataFrame para evitar ambigüedad en Pandas
        if df_resumen_areas is None or (isinstance(df_resumen_areas, pd.DataFrame) and df_resumen_areas.empty):
            total_metas = 0
            promedio_cumplimiento = 0.0
            niveles = {}
            num_areas = 0
        else:
            total_metas = df_resumen_areas['NO. METAS'].sum() if 'NO. METAS' in df_resumen_areas.columns else 0
            promedio_cumplimiento = round(df_resumen_areas['PORCENTAJE'].mean(), 2) if 'PORCENTAJE' in df_resumen_areas.columns else 0.0
            niveles = df_resumen_areas['NIVEL'].value_counts().to_dict() if 'NIVEL' in df_resumen_areas.columns else {}
            num_areas = len(df_resumen_areas)

        bloque_contexto = f"--- INICIO REFERENCIA Y ESTILO DE TRIMESTRES ANTERIORES ---\n{contexto_previo}\n--- FIN REFERENCIA ---" if contexto_previo else "No hay documentos previos de referencia."

        prompt = f"""
Eres un analista experto en administración pública municipal y evaluación de programas (PbRM/UIPPE).

A continuación te presento informes de trimestres pasados para usar como ESTILO, TONO INSTITUCIONAL Y ESTRUCTURA DE REDACCIÓN:
{bloque_contexto}

TAREA:
Redacta el **Análisis Cualitativo Institucional** para el Informe del {trimestre}° Trimestre 2026.

Datos consolidados del municipio:
- Trimestre Evaluado: {trimestre}° Trimestre
- Áreas / Unidades Evaluadas: {num_areas}
- Total de Metas Físicas Evaluadas: {total_metas}
- Porcentaje Promedio de Cumplimiento Institucional: {promedio_cumplimiento}%

Distribución por Nivel de Desempeño:
- Sobrepasado: {niveles.get('SOBREPASADO', 0)} áreas
- Adecuado: {niveles.get('ADECUADO', 0)} áreas
- Regular: {niveles.get('REGULAR', 0)} áreas
- Deficiente: {niveles.get('DEFICIENTE', 0)} áreas
- Crítico: {niveles.get('CRÍTICO', 0)} áreas

Lineamientos de redacción:
1. Utiliza un estilo oficial, formal y ejecutivo acorde a los antecedentes provistos.
2. Sintetiza el rendimiento global de la administración en el periodo.
3. Comenta brevemente sobre la distribución por semáforos de cumplimiento.
4. Agrega una recomendación técnica preventiva para las Unidades Administrativas.
5. Límite de extensión: máximo 3 párrafos concisos.
"""

        # Modelo actualizado a la versión vigente en el SDK
        response = self.client.models.generate_content(
            model="gemini-2.5-flash",
            contents=prompt
        )

        return response.text