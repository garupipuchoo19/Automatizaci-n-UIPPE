import openpyxl
import pandas as pd
import os
from datetime import datetime

class InyectorExcelPbRM:
    def __init__(self, ruta_plantilla_maestra: str):
        """
        Inicializa el motor con la ruta del archivo Excel plantilla.
        :param ruta_plantilla_maestra: Ruta al archivo Excel institucional base.
        """
        if not os.path.exists(ruta_plantilla_maestra):
            raise FileNotFoundError(f"No se encontró la plantilla maestra en: {ruta_plantilla_maestra}")
        
        self.ruta_plantilla = ruta_plantilla_maestra

        # Mapeo oficial de columnas según la estructura examinada
        self.MAPEO_COLUMNAS = {
            1: {"avance": 17, "justificacion": 43},
            2: {"avance": 21, "justificacion": 44},
            3: {"avance": 25, "justificacion": 45},
            4: {"avance": 29, "justificacion": 46}
        }

    def inyectar_captura(self, df_captura: pd.DataFrame, trimestre: int, ruta_salida: str = None) -> str:
        """
        Escribe los avances y justificaciones en el libro Excel respetando celdas calculadas.
        
        :param df_captura: DataFrame con columnas ['CLAVE', 'AVANCE', 'JUSTIFICACION']
        :param trimestre: Número del trimestre (1, 2, 3 o 4)
        :param ruta_salida: Ruta donde se guardará el nuevo archivo. Si es None, sobreescribe la plantilla maestra.
        :return: Ruta del archivo generado/actualizado
        """
        if trimestre not in self.MAPEO_COLUMNAS:
            raise ValueError(f"Trimestre {trimestre} inválido. Debe ser 1, 2, 3 o 4.")

        col_av = self.MAPEO_COLUMNAS[trimestre]["avance"]
        col_just = self.MAPEO_COLUMNAS[trimestre]["justificacion"]

        # Cargar libro conservando fórmulas originales (data_only=False)
        wb = openpyxl.load_workbook(self.ruta_plantilla, data_only=False)
        
        if "METAS" not in wb.sheetnames:
            raise KeyError("La plantilla no contiene la pestaña obligatoria 'METAS'")
            
        ws_metas = wb["METAS"]

        # Indexar la captura por CLAVE para acceso veloz O(1)
        # Convertimos las claves a string y eliminamos espacios
        df_captura['CLAVE'] = df_captura['CLAVE'].astype(str).str.strip()
        mapa_captura = df_captura.set_index('CLAVE').to_dict(orient='index')

        registros_actualizados = 0

        # Recorrer filas a partir de la fila 4 (donde inician los registros de metas)
        for row in range(4, ws_metas.max_row + 1):
            val_est = ws_metas.cell(row=row, column=7).value
            val_num = ws_metas.cell(row=row, column=12).value

            if val_est and val_num is not None:
                clave_calculada = f"{str(val_est).strip()}{str(val_num).strip()}"

                if clave_calculada in mapa_captura:
                    datos = mapa_captura[clave_calculada]

                    # 1. Inyectar Avance Real
                    if "AVANCE" in datos and datos["AVANCE"] is not None:
                        try:
                            ws_metas.cell(row=row, column=col_av).value = float(datos["AVANCE"])
                        except (ValueError, TypeError):
                            ws_metas.cell(row=row, column=col_av).value = 0.0

                    # 2. Inyectar Justificación
                    if "JUSTIFICACION" in datos and datos["JUSTIFICACION"]:
                        ws_metas.cell(row=row, column=col_just).value = str(datos["JUSTIFICACION"]).strip()

                    registros_actualizados += 1

        # Si no se define una ruta explícita, se guarda sobreescritura directa en la plantilla maestra
        destino = ruta_salida if ruta_salida else self.ruta_plantilla

        wb.save(destino)
        print(f"✔ Éxito: Se inyectaron {registros_actualizados} metas en Trimestre {trimestre}.")
        print(f"📁 Archivo actualizado: {destino}")

        return destino