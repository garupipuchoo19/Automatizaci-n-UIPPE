import pdfplumber
import os
import pandas as pd

class LectorPDF:
    def __init__(self, ruta_archivo):
        self.ruta_archivo = ruta_archivo

    def extraer_texto_completo(self):
        """Extrae todo el texto plano del PDF página por página."""
        if not os.path.exists(self.ruta_archivo):
            raise FileNotFoundError(f"El archivo {self.ruta_archivo} no existe.")

        texto_paginas = []
        with pdfplumber.open(self.ruta_archivo) as pdf:
            for i, pagina in enumerate(pdf.pages):
                contenido = pagina.extract_text()
                if contenido:
                    texto_paginas.append(f"--- Página {i+1} ---\n{contenido}")
        
        return "\n\n".join(texto_paginas)

    def extraer_tablas_a_dataframe(self):
        """Busca tablas dentro del PDF y las convierte en un DataFrame de Pandas."""
        if not os.path.exists(self.ruta_archivo):
            raise FileNotFoundError(f"El archivo {self.ruta_archivo} no existe.")

        todas_las_filas = []
        with pdfplumber.open(self.ruta_archivo) as pdf:
            for pagina in pdf.pages:
                tablas = pagina.extract_tables()
                for tabla in tablas:
                    for fila in tabla:
                        todas_las_filas.append(fila)

        if todas_las_filas:
            # La primera fila suele ser el encabezado
            encabezados = todas_las_filas[0]
            datos = todas_las_filas[1:]
            return pd.DataFrame(datos, columns=encabezados)
        
        return pd.DataFrame()