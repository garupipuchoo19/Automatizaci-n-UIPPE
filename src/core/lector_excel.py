import pandas as pd
import os

class LectorExcel:
    def __init__(self, ruta_archivo):
        self.ruta_archivo = ruta_archivo
        self.datos = None

    def cargar_datos(self, nombre_hoja=0, header=1):
        if not os.path.exists(self.ruta_archivo):
            raise FileNotFoundError(f"El archivo {self.ruta_archivo} no existe.")
        try:
            # Los encabezados oficiales están en la fila índice 1 del archivo de Cuautitlán
            df_raw = pd.read_excel(self.ruta_archivo, sheet_name=nombre_hoja, header=header)
            
            # Limpiar espacios en los nombres de las columnas
            df_raw.columns = df_raw.columns.astype(str).str.strip()
            self.datos = df_raw
            return self.datos
        except Exception as e:
            raise RuntimeError(f"Error al leer la hoja de Excel '{nombre_hoja}': {str(e)}")

    def obtener_resumen(self):
        if self.datos is not None and not self.datos.empty:
            return {"filas": self.datos.shape[0], "columnas": list(self.datos.columns)}
        return None

    # --- MÉTODOS ESPECÍFICOS PARA UIPPE ---

    def cargar_metas(self):
        """Carga la hoja METAS de forma tolerante a columnas faltantes o variaciones."""
        try:
            df = self.cargar_datos(nombre_hoja='METAS', header=1)
        except Exception:
            # Si la hoja 'METAS' no existe por nombre, intenta leer la primera hoja disponible
            df = self.cargar_datos(nombre_hoja=0, header=1)

        if df is not None and not df.empty:
            # Identificar qué columnas requeridas existen realmente
            cols_filtro = [col for col in ['CLAVE', 'ESTRUCTURA'] if col in df.columns]
            
            # Solo filtra filas vacías sobre las columnas que realmente existan en la tabla
            if cols_filtro:
                df = df.dropna(subset=cols_filtro)
                
        return df

    def cargar_indicadores(self):
        """Carga la hoja INDICADORES de forma tolerante a columnas faltantes o variaciones."""
        try:
            df = self.cargar_datos(nombre_hoja='INDICADORES', header=1)
        except Exception:
            df = self.cargar_datos(nombre_hoja=0, header=1)

        if df is not None and not df.empty:
            cols_filtro = [col for col in ['CLAVE', 'ESTRUCTURA'] if col in df.columns]
            if cols_filtro:
                df = df.dropna(subset=cols_filtro)
                
        return df

    def cargar_resultados_general(self):
        """Carga la hoja de resultados calculados del Excel maestro."""
        try:
            return self.cargar_datos(nombre_hoja='resultados general', header=0)
        except Exception:
            return None

    def obtener_catalogo_estructuras(self, tipo_registro="METAS"):
        """
        Retorna la lista ordenada de Estructuras (códigos de área tipo A002120103010103)
        y un diccionario mapeador de Estructura -> Área Encargada.
        """
        df = self.cargar_metas() if tipo_registro == "METAS" else self.cargar_indicadores()
        
        if df is None or df.empty:
            return [], {}

        # Verificar si existen ambas columnas antes de extraer catálogo
        if 'ESTRUCTURA' in df.columns and 'ÁREA ENCARGADA' in df.columns:
            df_sub = df[['ESTRUCTURA', 'ÁREA ENCARGADA']].dropna(subset=['ESTRUCTURA']).drop_duplicates()
            estructuras = sorted(df_sub['ESTRUCTURA'].astype(str).unique().tolist())
            mapa_areas = dict(zip(df_sub['ESTRUCTURA'].astype(str), df_sub['ÁREA ENCARGADA'].astype(str)))
            return estructuras, mapa_areas

        return [], {}