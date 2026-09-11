import customtkinter as ctk

class TarjetaKpi(ctk.CTkFrame):
    """Componente para mostrar métricas breves en la interfaz."""
    def __init__(self, parent, titulo="Métrica", valor="0", color_fondo="#2b2b2b", **kwargs):
        super().__init__(parent, fg_color=color_fondo, corner_radius=8, **kwargs)

        self.lbl_titulo = ctk.CTkLabel(
            self, text=titulo.upper(), font=ctk.CTkFont(size=10, weight="bold"), text_color="#a1a1a1"
        )
        self.lbl_titulo.pack(padx=15, pady=(10, 2), anchor="w")

        self.lbl_valor = ctk.CTkLabel(
            self, text=valor, font=ctk.CTkFont(size=22, weight="bold")
        )
        self.lbl_valor.pack(padx=15, pady=(0, 10), anchor="w")

    def actualizar(self, nuevo_valor):
        self.lbl_valor.configure(text=str(nuevo_valor))


class PanelOpcionesExportacion(ctk.CTkFrame):
    """Panel de radio-buttons y switches para elegir el formato de salida."""
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)

        self.lbl_titulo = ctk.CTkLabel(self, text="Opciones de Procesamiento", font=ctk.CTkFont(weight="bold"))
        self.lbl_titulo.pack(padx=15, pady=(10, 5), anchor="w")

        self.var_formato = ctk.StringVar(value="excel")

        self.rb_excel = ctk.CTkRadioButton(
            self, text="Generar Consolidado Excel (.xlsx)", variable=self.var_formato, value="excel"
        )
        self.rb_excel.pack(padx=20, pady=5, anchor="w")

        self.rb_word = ctk.CTkRadioButton(
            self, text="Generar Reporte Ejecutivo Word (.docx)", variable=self.var_formato, value="word"
        )
        self.rb_word.pack(padx=20, pady=5, anchor="w")

        self.switch_validar = ctk.CTkSwitch(self, text="Validar estricto según reglamento UIPPE")
        self.switch_validar.select()
        self.switch_validar.pack(padx=20, pady=(10, 15), anchor="w")

    def obtener_formato(self):
        return self.var_formato.get()

    def es_estricto(self):
        return bool(self.switch_validar.get())