import pandas as pd

class ValidadorReglamento:
    """Motor de validación y cálculo de cumplimiento programático PbRM de la UIPPE."""
    
    def __init__(self, dataframe_datos):
        self.df = dataframe_datos
        self.hallazgos = []

    @staticmethod
    def clasificar_desempeno(porcentaje):
        """Clasifica el porcentaje según los parámetros oficiales del ITSP de la UIPPE."""
        try:
            val = float(porcentaje)
            if val >= 111.0:
                return "Sobrepasado", "S", "#800080"  # Púrpura
            elif 90.0 <= val <= 110.99:
                return "Adecuado", "A", "#008000"    # Verde
            elif 70.0 <= val <= 89.99:
                return "Regular", "R", "#FFFF00"     # Amarillo
            elif 50.0 <= val <= 69.99:
                return "Deficiente", "D", "#FFA500"  # Naranja
            else:
                return "Crítico", "C", "#FF0000"     # Rojo
        except (ValueError, TypeError):
            return "No Evaluado", "N/E", "#808080"

    def validar_campos_obligatorios(self, columnas_requeridas):
        if self.df is None or self.df.empty:
            self.hallazgos.append("El archivo no contiene registros o no se pudieron extraer datos.")
            return False

        columnas_actuales = set(self.df.columns)
        faltantes = [col for col in columnas_requeridas if col not in columnas_actuales]
        if faltantes:
            self.hallazgos.append(f"Faltan columnas obligatorias: {', '.join(faltantes)}")
            return False
        return True

    def verificar_limpieza_general(self):
        """Revisión universal de registros vacíos y duplicados."""
        if self.df is None or self.df.empty:
            return

        filas_vacias = self.df.isnull().all(axis=1).sum()
        if filas_vacias > 0:
            self.hallazgos.append(f"Se omitieron {filas_vacias} filas totalmente vacías.")

        duplicados = self.df.duplicated().sum()
        if duplicados > 0:
            self.hallazgos.append(f"Atención: Se identificaron {duplicados} filas exactamente iguales.")

    def aplicar_reglas_cumplimiento(self, columna_avance=None):
        """Aplica la lógica de evaluación programática al DataFrame."""
        if self.df is None or self.df.empty:
            return pd.DataFrame(), self.hallazgos

        self.verificar_limpieza_general()
        df_procesado = self.df.copy()

        # Si se especifica una columna de porcentaje de avance, clasificamos cada fila
        if columna_avance and columna_avance in df_procesado.columns:
            df_procesado['Estatus_UIPPE'], df_procesado['Codigo_UIPPE'], df_procesado['Color_Hex'] = zip(
                *df_procesado[columna_avance].apply(self.clasificar_desempeno)
            )
            
            # Resumen de conteos por categoría
            resumen = df_procesado['Estatus_UIPPE'].value_counts().to_dict()
            for estatus, conteo in resumen.items():
                self.hallazgos.append(f"Categoría '{estatus}': {conteo} registros.")

        return df_procesado, self.hallazgos