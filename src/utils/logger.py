import os
import logging
from datetime import datetime
from src.utils.helpers import LOGS_DIR

def registrar_evento(mensaje, nivel="INFO"):
    """Registra eventos y errores tanto en consola como en archivo log dentro de logs/."""
    archivo_log = os.path.join(LOGS_DIR, f"log_uippe_{datetime.now().strftime('%Y%m')}.log")
    
    logging.basicConfig(
        filename=archivo_log,
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        encoding="utf-8"
    )
    
    if nivel.upper() == "ERROR":
        logging.error(mensaje)
    elif nivel.upper() == "WARNING":
        logging.warning(mensaje)
    else:
        logging.info(mensaje)