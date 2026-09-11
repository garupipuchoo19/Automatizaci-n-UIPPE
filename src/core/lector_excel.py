import pandas as pd
import os

class LectorExcel:
    def __init__(self, ruta_archivo):
        self.ruta_archivo = ruta_archivo
        self.datos = None

    def cargar_datos(self, nombre_hoja=0, header=0):
        if not os.path.exists(self.ruta_archivo):
            raise FileNotFoundError(f"El archivo {self.ruta_archivo} no existe.")
        try:
            self.datos = pd.read_excel(self.ruta_archivo, sheet_name=nombre_hoja, header=header)
            return self.datos
        except Exception as e:
            raise RuntimeError(f"Error al leer la hoja de Excel: {str(e)}")

    def obtener_resumen(self):
        if self.datos is not None and not self.datos.empty:
            return {"filas": self.datos.shape[0], "columnas": list(self.datos.columns)}
        return None

    # --- MÉTODOS ESPECÍFICOS PARA UIPPE ---

    def cargar_metas(self):
        """Carga la hoja METAS omitiendo el título superior (header en la fila 3)."""
        df = self.cargar_datos(nombre_hoja='METAS', header=2)
        # Validación explícita de Pandas para evitar la advertencia de ambigüedad
        if df is not None and not df.empty:
            df = df.dropna(subset=['CLAVE', 'ÁREA ENCARGADA'])
        return df

    def cargar_indicadores(self):
        """Carga la hoja INDICADORES con el encabezado estructurado."""
        df = self.cargar_datos(nombre_hoja='INDICADORES', header=2)
        if df is not None and not df.empty:
            df = df.dropna(subset=['CLAVE', 'ÁREA ENCARGADA'])
        return df

    def cargar_resultados_general(self):
        """Carga la hoja de resultados calculados del Excel maestro."""
        return self.cargar_datos(nombre_hoja='resultados general', header=0)