import time
from google import genai
from src.utils.helpers import obtener_contexto_historico_word
from src.utils.logger import registrar_evento

def generar_analisis_cualitativo_gemini(df_resumen_direcciones, trimestre_num=2, api_key=None, max_reintentos=3, espera_inicial=2):
    """
    Genera el análisis cualitativo oficial para el informe UIPPE usando la API de Gemini,
    incorporando contexto de reportes trimestrales previos y reintentos automáticos con Exponential Backoff.
    """
    if not api_key:
        registrar_evento("Gemini API Key no proporcionada. Se omitirá la generación por Inteligencia Artificial.")
        return None

    # 1. Recuperar contexto histórico acumulado (.docx)
    contexto_previo = obtener_contexto_historico_word()

    # 2. Construir datos resumidos en texto plano
    resumen_texto = ""
    if hasattr(df_resumen_direcciones, 'to_string'):
        resumen_texto = df_resumen_direcciones.to_string()
    else:
        resumen_texto = str(df_resumen_direcciones)

    # 3. Formular prompt del sistema
    prompt = f"""
    Eres un consultor experto en evaluación de la gestión pública municipal del Presupuesto basado en Resultados (PbRM) para OPERAGUA Cuautitlán Izcalli.

    Genera un Análisis Cualitativo Ejecutivo para el **{trimestre_num}° Trimestre**, redactado en un tono formal, institucional y analítico.

    DOCUMENTOS Y ANTECEDENTES HISTÓRICOS (Para coherencia y comparación):
    {contexto_previo if contexto_previo else "No hay reportes de trimestres anteriores registrados aún."}

    DATOS ACTUALES DEL {trimestre_num}° TRIMESTRE:
    {resumen_texto}

    Estructura la respuesta en los siguientes puntos utilizando negritas para los títulos principales:
    1. **Resumen Ejecutivo de Desempeño General**: Resumen de los avances y cumplimiento global del trimestre.
    2. **Análisis por Semáforo de Cumplimiento**: Comentario analítico sobre las metas en estado Crítico, Deficiente, Regular, Adecuado y Sobrepasado.
    3. **Evolución y Contexto Comparativo**: Breve comparación respecto a la tendencia observada en trimestres anteriores (si aplica).
    4. **Recomendaciones para las Unidades Administrativas**: Sugerencias claras para corregir subejercicios o mantener la eficiencia.

    No inventes cifras fuera de los datos proporcionados.
    """

    # 4. Intentar llamada con Exponential Backoff
    espera = espera_inicial
    for intento in range(1, max_reintentos + 1):
        try:
            registrar_evento(f"Consultando Gemini API (Intento {intento}/{max_reintentos})...")
            client = genai.Client(api_key=api_key)
            
            # Usar gemini-2.5-flash
            response = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=prompt
            )

            if response and response.text:
                registrar_evento("✔ Análisis cualitativo generado exitosamente con Gemini API.")
                return response.text.strip()

        except Exception as e:
            msg_error = str(e)
            registrar_evento(f"Aviso: Intento {intento} fallido con Gemini API: {msg_error}")

            if intento == max_reintentos:
                registrar_evento("❌ Se agotaron los reintentos con Gemini API. Se aplicará respuesta de respaldo.")
                break

            registrar_evento(f"Reintentando conexión con Gemini en {espera} segundos...")
            time.sleep(espera)
            espera *= 2  # Exponential backoff (2s -> 4s -> 8s)

    # Texto fallback en caso de error o falta de red
    return (
        "**Análisis Cualitativo Ejecutivo - Resguardo de Fallback**\n\n"
        "El análisis automático por IA no pudo completarse debido a una interrupción en la conexión "
        "con los servicios de Google Gemini API o un límite temporal de cuota. "
        "Se sugiere validar la conectividad de red o la validez de la API Key e integrar la redacción manualmente."
    )