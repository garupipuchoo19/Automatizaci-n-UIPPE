from src.core.excel_engine import EngineExcelUIPPE, GeminiAnalyst

API_KEY_GEMINI = ""  # Coloca tu clave de Gemini aquí o usa la variable de entorno GEMINI_API_KEY
ARCHIVO_EXCEL = "CUAUTITLAN METAS E INDICADORES 2026.xlsx"
TRIMESTRE = 2  # Cambia a 1, 2, 3 o 4 según el periodo a evaluar

if __name__ == "__main__":
    # 1. Inicializar motor apuntando al archivo en entradas/
    engine = EngineExcelUIPPE(ARCHIVO_EXCEL)

    try:
        # 2. Validación Previa (Opción 1)
        engine.validar_captura_trimestre(TRIMESTRE)
        print(f"✓ Validación de capturas para Trimestre {TRIMESTRE} aprobada.")

        # 3. Procesar tabla por Dirección/Área
        col_sem = f"{TRIMESTRE}°_SEMAFORO"
        df_resumen = engine.procesar_hoja_metas(col_semaforo=col_sem)
        print("✓ Concentrado por áreas recalculado exitosamente.")

        # 4. Solicitud a Gemini (Opción 3)
        analyst = GeminiAnalyst(api_key=API_KEY_GEMINI)
        print("🤖 Solicitando redacción cualitativa a Gemini...")
        analisis_texto = analyst.generar_analisis_ejecutivo(TRIMESTRE, df_resumen)

        print("\n--- REDACCIÓN CUALITATIVA GENERADA ---")
        print(analisis_texto)

    except ValueError as e:
        print(f"\n[VALIDACIÓN]: {e}")
    except Exception as e:
        print(f"\n[ERROR]: {e}")