import os
import shutil
import docx
from datetime import datetime

# Directorio raíz del proyecto (Automatizacion_UIPPE)
BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

ENTRADAS_DIR = os.path.join(BASE_DIR, "entradas")
SALIDAS_DIR = os.path.join(BASE_DIR, "salidas")
PLANTILLAS_DIR = os.path.join(BASE_DIR, "plantillas")
LOGS_DIR = os.path.join(BASE_DIR, "logs")

# Rutas adicionales para respaldos y memoria de Gemini
RESPALDOS_DIR = os.path.join(BASE_DIR, "respaldos")
CONTEXTO_WORD_DIR = os.path.join(RESPALDOS_DIR, "contexto_word")
HISTORICO_SALIDAS_DIR = os.path.join(RESPALDOS_DIR, "historico_salidas")


def inicializar_directorios():
    """Garantiza la existencia de las carpetas del sistema al arrancar."""
    carpetas = [
        ENTRADAS_DIR, SALIDAS_DIR, PLANTILLAS_DIR, LOGS_DIR,
        RESPALDOS_DIR, CONTEXTO_WORD_DIR, HISTORICO_SALIDAS_DIR
    ]
    for carpeta in carpetas:
        os.makedirs(carpeta, exist_ok=True)


def obtener_ruta_entrada(nombre_archivo):
    """Devuelve la ruta absoluta de un archivo dentro de la carpeta entradas/."""
    return os.path.join(ENTRADAS_DIR, nombre_archivo)


def obtener_nombre_salida(prefix="Reporte_UIPPE", extension="xlsx"):
    """Genera un nombre único con sello de fecha y hora."""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{prefix}_{timestamp}.{extension}"
    return os.path.join(SALIDAS_DIR, filename)


def obtener_fecha_formal():
    """Devuelve la fecha en formato formal de redacción oficial."""
    meses = [
        "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
        "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
    ]
    ahora = datetime.now()
    return f"Cuautitlán Izcalli, Estado de México a {ahora.day} de {meses[ahora.month - 1]} de {ahora.year}"


def extraer_texto_word(ruta_docx):
    """Extrae el texto de un documento Word para ser utilizado como contexto histórico por Gemini."""
    if not os.path.exists(ruta_docx):
        return ""
    try:
        doc = docx.Document(ruta_docx)
        parrafos = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        return "\n".join(parrafos)
    except Exception:
        return ""


def obtener_contexto_historico_word():
    """Escanea la carpeta respaldos/contexto_word/ y concatena el texto de los Word previos."""
    textos = []
    if os.path.exists(CONTEXTO_WORD_DIR):
        archivos = sorted(os.listdir(CONTEXTO_WORD_DIR))
        for archivo in archivos:
            if archivo.endswith(".docx") and not archivo.startswith("~$"):
                ruta = os.path.join(CONTEXTO_WORD_DIR, archivo)
                contenido = extraer_texto_word(ruta)
                if contenido:
                    textos.append(f"--- REFERENCIA HISTÓRICA: {archivo} ---\n{contenido}\n")
    return "\n".join(textos)


def depurar_contexto_historico(max_archivos=5):
    """
    Mantiene únicamente los reportes Word más recientes dentro de contexto_word/
    (por defecto 5: los 4 trimestres del año actual + 1 del año anterior)
    para optimizar almacenamiento y contexto de Gemini.
    """
    if not os.path.exists(CONTEXTO_WORD_DIR):
        return

    archivos_word = [
        os.path.join(CONTEXTO_WORD_DIR, f)
        for f in os.listdir(CONTEXTO_WORD_DIR)
        if f.endswith(".docx") and not f.startswith("~$")
    ]

    # Ordenar por fecha de última modificación (del más reciente al más antiguo)
    archivos_word.sort(key=lambda x: os.path.getmtime(x), reverse=True)

    # Eliminar excedentes si supera el límite permitido
    if len(archivos_word) > max_archivos:
        for archivo in archivos_word[max_archivos:]:
            try:
                os.remove(archivo)
            except Exception:
                pass


def depurar_historico_salidas(max_archivos=10):
    """
    Mantiene un tope de respaldos de salida antiguos en historico_salidas/
    para evitar acumulación indeterminada de archivos temporales.
    """
    if not os.path.exists(HISTORICO_SALIDAS_DIR):
        return

    archivos = [
        os.path.join(HISTORICO_SALIDAS_DIR, f)
        for f in os.listdir(HISTORICO_SALIDAS_DIR)
        if os.path.isfile(os.path.join(HISTORICO_SALIDAS_DIR, f)) and not f.startswith("~$")
    ]

    archivos.sort(key=lambda x: os.path.getmtime(x), reverse=True)

    if len(archivos) > max_archivos:
        for archivo in archivos[max_archivos:]:
            try:
                os.remove(archivo)
            except Exception:
                pass


def procesar_respaldos_salida(ruta_archivo_salida):
    """
    Guarda automáticamente:
    1. Una copia fechada en respaldos/historico_salidas/
    2. Si es .docx, una copia limpia en respaldos/contexto_word/ para alimentar consultas futuras de Gemini.
    3. Ejecuta la rotación/depuración automática de almacenamiento.
    """
    inicializar_directorios()
    if not os.path.exists(ruta_archivo_salida):
        return None, None

    nombre_base = os.path.basename(ruta_archivo_salida)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    # 1. Copia de respaldo histórico (inmutable)
    nombre_backup = f"{timestamp}_{nombre_base}"
    ruta_historico = os.path.join(HISTORICO_SALIDAS_DIR, nombre_backup)
    shutil.copy2(ruta_archivo_salida, ruta_historico)

    # 2. Actualización de la memoria de contexto (.docx)
    ruta_contexto = None
    if ruta_archivo_salida.endswith(".docx"):
        ruta_contexto = os.path.join(CONTEXTO_WORD_DIR, nombre_base)
        shutil.copy2(ruta_archivo_salida, ruta_contexto)

    # 3. Mantenimiento y depuración automática de carpetas
    depurar_contexto_historico(max_archivos=5)
    depurar_historico_salidas(max_archivos=10)

    return ruta_historico, ruta_contexto


# Auto-ejecución al importar el módulo
inicializar_directorios()