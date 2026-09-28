import os
from datetime import datetime
from src.core.lector_excel import LectorExcel
from src.core.validador_reg import ValidadorReglamento
from src.core.generador_doc import GeneradorReporte
from src.core.excel_engine import EngineExcelUIPPE, GeminiAnalyst
from src.utils.helpers import (
    PLANTILLAS_DIR, 
    ENTRADAS_DIR, 
    obtener_ruta_entrada,
    obtener_contexto_historico_word,
    procesar_respaldos_salida
)

class ProcesadorUIPPE:
    """Orquesta la lectura de Excel, validación de metas e indicadores, generación con Gemini y exportación Word."""

    @staticmethod
    def procesar_archivo_excel(ruta_excel, trimestre_num=2, anio=None, fecha_emision=None, ruta_plantilla_word=None, solo_programadas=True, api_key_gemini=None, tipo_registro="Metas"):
        if not anio:
            anio = datetime.now().year

        mapa_cols_avance = {
            1: '% AVANCE 1er. TRIM',
            2: '% AVANCE 2o. TRIM',
            3: '% AVANCE 3o. TRIM',
            4: '% AVANCE 4o. TRIM'
        }
        mapa_cols_prog = {
            1: 'PROGR. 1° TRIM',
            2: 'PROGR. 2° TRIM',
            3: 'PROGR. 3° TRIM',
            4: 'PROGR. 4° TRIM'
        }

        col_avance_sel = mapa_cols_avance.get(trimestre_num, '% AVANCE 2o. TRIM')
        col_prog_sel = mapa_cols_prog.get(trimestre_num, 'PROGR. 2° TRIM')

        # Plantilla Word oficial según año
        if not ruta_plantilla_word:
            nombre_plantilla = f"ITSP {trimestre_num} TRIMESTRE {anio}.docx"
            plantilla_especifica = os.path.join(PLANTILLAS_DIR, nombre_plantilla)
            if os.path.exists(plantilla_especifica):
                ruta_plantilla_word = plantilla_especifica
            else:
                ruta_plantilla_word = os.path.join(PLANTILLAS_DIR, "Plantilla_Oficial_UIPPE.docx")

        # 0. Instanciar Motor de Excel
        engine = EngineExcelUIPPE(ruta_excel)

        # 1. Validar captura del trimestre
        engine.validar_captura_trimestre(trimestre_num)

        # 2. Leer Excel maestro mediante LectorExcel
        lector = LectorExcel(ruta_excel)
        df_registros = lector.cargar_metas() if tipo_registro == "Metas" else lector.cargar_indicadores()

        # 3. Filtrar registros programados
        if solo_programadas:
            if col_prog_sel not in df_registros.columns:
                coincidencias_prog = [c for c in df_registros.columns if f"{trimestre_num}" in c and ("PROG" in c.upper() or "CAN" in c.upper())]
                if coincidencias_prog:
                    col_prog_sel = coincidencias_prog[0]

            if col_prog_sel in df_registros.columns:
                df_registros = df_registros[df_registros[col_prog_sel].notna() & (df_registros[col_prog_sel] > 0)].copy()

        # 4. Validar cumplimiento por semáforo
        validador = ValidadorReglamento(df_registros)
        
        if col_avance_sel not in df_registros.columns:
            coincidencias = [c for c in df_registros.columns if f"{trimestre_num}" in c and "AVANCE" in c.upper()]
            if coincidencias:
                col_avance_sel = coincidencias[0]

        df_procesado, hallazgos = validador.aplicar_reglas_cumplimiento(columna_avance=col_avance_sel)

        # 5. Cargar y recalcular resumen por Direcciones y Áreas
        col_semaforo = f"{trimestre_num}°_SEMAFORO"
        df_resumen = engine.procesar_hoja_metas(col_semaforo=col_semaforo)

        # 6. Generar Análisis Cualitativo con Gemini API
        analisis_narrativo = None
        if api_key_gemini:
            try:
                if df_resumen is not None and not df_resumen.empty:
                    contexto_historico = obtener_contexto_historico_word()
                    
                    analyst = GeminiAnalyst(api_key=api_key_gemini)
                    analisis_narrativo = analyst.generar_analisis_ejecutivo(
                        trimestre=trimestre_num, 
                        df_resumen_areas=df_resumen,
                        contexto_previo=contexto_historico
                    )
                    hallazgos.append("✔ Redacción cualitativa generada con éxito con Gemini AI.")
                else:
                    hallazgos.append("⚠️ El resumen recalculado está vacío; no se invocó a Gemini.")
            except Exception as e:
                hallazgos.append(f"⚠️ No se pudo generar la narrativa con Gemini: {str(e)}")

        # 7. Generar el reporte consolidado en Word
        generador = GeneradorReporte(
            dataframe_procesado=df_procesado,
            ruta_plantilla_word=ruta_plantilla_word,
            trimestre_num=trimestre_num,
            anio=anio,
            fecha_emision=fecha_emision
        )
        
        titulo_reporte = f"EVALUACIÓN DEL CUMPLIMIENTO PROGRAMÁTICO DE {tipo_registro.upper()} - {trimestre_num}° TRIMESTRE {anio}"
        ruta_doc = generador.exportar_word_ejecutivo(
            titulo=titulo_reporte,
            datos_direcciones=df_resumen,
            analisis_cualitativo=analisis_narrativo
        )

        # 8. Sincronización de respaldos
        ruta_backup, ruta_contexto = procesar_respaldos_salida(ruta_doc)
        if ruta_backup:
            hallazgos.append("✔ Respaldos sincronizados en histórico y base de contexto Word.")

        return ruta_doc, hallazgos