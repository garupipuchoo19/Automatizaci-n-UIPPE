import os
import threading
from datetime import datetime
import customtkinter as ctk
from tkinter import filedialog, messagebox

# Intentar importar tkcalendar para el widget desplegable
try:
    from tkcalendar import Calendar
    TIENE_TKCALENDAR = True
except ImportError:
    TIENE_TKCALENDAR = False

# Controladores y Backend
from src.controllers.procesador_uippe import ProcesadorUIPPE
from src.core.generador_excel_pbrm import InyectorExcelPbRM
from src.core.lector_excel import LectorExcel
from src.core.generador_doc import validar_plantilla_word

# Interfaz y Componentes GUI
from src.gui.captura_view import VistaCapturaRapida
from src.gui.components import TarjetaKpi, PanelOpcionesExportacion
from src.utils.helpers import ENTRADAS_DIR, PLANTILLAS_DIR
from src.utils.logger import registrar_evento

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


class CTkDatePicker(ctk.CTkFrame):
    """Componente de fecha con diseño 100% nativo e integrado con CustomTkinter."""
    def __init__(self, master, width=140, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        
        self.entry_fecha = ctk.CTkEntry(self, width=width - 35, placeholder_text="DD/MM/YYYY")
        self.entry_fecha.pack(side="left", padx=(0, 2))
        self.entry_fecha.insert(0, datetime.now().strftime("%d/%m/%Y"))

        self.btn_calendar = ctk.CTkButton(
            self, 
            text="📅", 
            width=30, 
            fg_color="transparent", 
            border_width=1,
            text_color=("black", "white"),
            command=self._abrir_calendario
        )
        self.btn_calendar.pack(side="left")

    def _abrir_calendario(self):
        if not TIENE_TKCALENDAR:
            return

        top = ctk.CTkToplevel(self)
        top.title("Seleccionar Fecha")
        top.geometry("260x250")
        top.grab_set()
        top.resizable(False, False)

        cal = Calendar(
            top, 
            selectmode='day', 
            date_pattern='dd/mm/yyyy',
            headersbackground='#1f4e78',
            selectbackground='#1f6aa5'
        )
        cal.pack(padx=10, pady=10, fill="both", expand=True)

        def seleccionar():
            self.entry_fecha.delete(0, "end")
            self.entry_fecha.insert(0, cal.get_date())
            top.destroy()

        btn_aceptar = ctk.CTkButton(top, text="Seleccionar", command=seleccionar, height=28)
        btn_aceptar.pack(pady=(0, 10))

    def get(self):
        return self.entry_fecha.get().strip()


class AppUIPPE(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Sistema de Automatización y Evaluación PbRM - UIPPE OPERAGUA")
        self.geometry("1150x820")
        self.minsize(950, 680)

        self.archivo_seleccionado = None
        self.ruta_excel_maestro = os.path.join(ENTRADAS_DIR, f"CUAUTITLAN METAS E INDICADORES {datetime.now().year}.xlsx")
        self.ruta_plantilla_word = os.path.join(PLANTILLAS_DIR, "Plantilla_Oficial_UIPPE.docx")

        self.inyector = InyectorExcelPbRM(self.ruta_excel_maestro)

        self.lbl_titulo = ctk.CTkLabel(
            self, 
            text="OPERAGUA Cuautitlán Izcalli - UIPPE", 
            font=ctk.CTkFont(size=18, weight="bold")
        )
        self.lbl_titulo.pack(padx=20, pady=(15, 2))

        self.lbl_subtitulo = ctk.CTkLabel(
            self, 
            text="Unidad de Información, Planeación, Programación y Evaluación", 
            font=ctk.CTkFont(size=12),
            text_color="#8d8d8d"
        )
        self.lbl_subtitulo.pack(padx=20, pady=(0, 10))

        self.tabview = ctk.CTkTabview(self)
        self.tabview.pack(padx=20, pady=(0, 15), fill="both", expand=True)

        self.tab_captura = self.tabview.add("📝 Captura Rápida por Área")
        self.tab_reportes = self.tabview.add("📊 Consolidación y Reportes UIPPE")

        self.vista_captura = VistaCapturaRapida(
            parent=self.tab_captura,
            ruta_excel_maestro=self.ruta_excel_maestro,
            al_guardar_callback=self.procesar_guardado_captura
        )
        self.vista_captura.pack(fill="both", expand=True)

        self.construir_pestaña_reportes()

    def construir_pestaña_reportes(self):
        self.frame_file = ctk.CTkFrame(self.tab_reportes)
        self.frame_file.pack(padx=10, pady=5, fill="x")

        self.btn_seleccionar = ctk.CTkButton(
            self.frame_file, 
            text="Cargar Insumo (Excel/Word/PDF)", 
            command=self.seleccionar_archivo
        )
        self.btn_seleccionar.pack(side="left", padx=15, pady=15)

        self.lbl_archivo = ctk.CTkLabel(
            self.frame_file, 
            text="Sin archivo seleccionado...", 
            anchor="w",
            text_color="#a1a1a1"
        )
        self.lbl_archivo.pack(side="left", padx=10, fill="x", expand=True)

        # --- CONFIGURACIÓN DE PARÁMETROS Y FECHAS DINÁMICAS ---
        self.frame_config = ctk.CTkFrame(self.tab_reportes)
        self.frame_config.pack(padx=10, pady=5, fill="x")

        # Trimestre
        lbl_trim = ctk.CTkLabel(self.frame_config, text="Trimestre:", font=ctk.CTkFont(weight="bold"))
        lbl_trim.pack(side="left", padx=(15, 5), pady=10)

        self.combo_trimestre_reporte = ctk.CTkComboBox(
            self.frame_config,
            values=["1° TRIMESTRE", "2° TRIMESTRE", "3° TRIMESTRE", "4° TRIMESTRE"],
            state="readonly",
            width=140
        )
        self.combo_trimestre_reporte.set("1° TRIMESTRE")
        self.combo_trimestre_reporte.pack(side="left", padx=5, pady=10)

        # Selector de Año Dinámico
        lbl_anio = ctk.CTkLabel(self.frame_config, text="Año:", font=ctk.CTkFont(weight="bold"))
        lbl_anio.pack(side="left", padx=(20, 5), pady=10)

        anio_actual = str(datetime.now().year)
        anios_opciones = [str(a) for a in range(int(anio_actual) - 2, int(anio_actual) + 3)]
        
        self.combo_anio = ctk.CTkComboBox(
            self.frame_config,
            values=anios_opciones,
            state="readonly",
            width=100
        )
        self.combo_anio.set(anio_actual)
        self.combo_anio.pack(side="left", padx=5, pady=10)

        # Selector de Fecha de Emisión (Date Picker Estilizado)
        lbl_fecha = ctk.CTkLabel(self.frame_config, text="Fecha Emisión:", font=ctk.CTkFont(weight="bold"))
        lbl_fecha.pack(side="left", padx=(20, 5), pady=10)

        self.picker_fecha = CTkDatePicker(self.frame_config, width=140)
        self.picker_fecha.pack(side="left", padx=5, pady=10)

        self.frame_kpis = ctk.CTkFrame(self.tab_reportes, fg_color="transparent")
        self.frame_kpis.pack(padx=10, pady=10, fill="x")

        self.kpi_filas = TarjetaKpi(self.frame_kpis, titulo="Registros Cargados", valor="0")
        self.kpi_filas.pack(side="left", padx=(0, 5), expand=True, fill="x")

        self.kpi_estado = TarjetaKpi(self.frame_kpis, titulo="Estatus Insumo", valor="Pendiente", color_fondo="#1f2937")
        self.kpi_estado.pack(side="left", padx=(5, 0), expand=True, fill="x")

        self.opciones = PanelOpcionesExportacion(self.tab_reportes)
        self.opciones.pack(padx=10, pady=10, fill="x")

        self.btn_procesar = ctk.CTkButton(
            self.tab_reportes, 
            text="▶ Ejecutar Validación y Generar Reporte", 
            state="disabled",
            height=40,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#1f6aa5",
            command=self.ejecutar_procesamiento
        )
        self.btn_procesar.pack(padx=10, pady=10, fill="x")

        self.lbl_log = ctk.CTkLabel(self.tab_reportes, text="Bitácora de Procesamiento:", font=ctk.CTkFont(size=11, weight="bold"))
        self.lbl_log.pack(padx=10, pady=(5, 0), anchor="w")

        self.txt_log = ctk.CTkTextbox(self.tab_reportes, height=120)
        self.txt_log.pack(padx=10, pady=(2, 10), fill="both", expand=True)
        
        self.log("Sistema inicializado correctamente. En espera de archivo de insumo...")

    def obtener_fecha_seleccionada(self):
        """Obtiene la fecha formateada en cadena DD/MM/YYYY."""
        return self.picker_fecha.get()

    def procesar_guardado_captura(self, df_captura, trimestre_num):
        try:
            ruta_generada = self.inyector.inyectar_captura(
                df_captura, 
                trimestre=trimestre_num, 
                ruta_salida=self.ruta_excel_maestro
            )
            self.vista_captura.cargar_datos_base()
            self.vista_captura.actualizar_tabla_metas()

            messagebox.showinfo(
                "¡Inyección Exitosa!", 
                f"Se actualizaron {len(df_captura)} metas del Trimestre {trimestre_num} correctamente."
            )
            self.log(f"✔ Captura guardada permanentemente en plantilla maestra ({len(df_captura)} metas en T{trimestre_num}).")
        except Exception as e:
            messagebox.showerror("Error de Inyección", f"Ocurrió un problema al guardar en Excel:\n{e}")
            self.log(f"ERROR EN INYECCIÓN: {str(e)}")

    def seleccionar_archivo(self):
        tipos = [
            ("Archivos de Insumo", "*.xlsx *.xls *.docx *.pdf"),
            ("Hojas de Cálculo Excel", "*.xlsx *.xls"),
            ("Documentos Word", "*.docx"),
            ("Archivos PDF", "*.pdf"),
            ("Todos los archivos", "*.*")
        ]
        ruta = filedialog.askopenfilename(initialdir=ENTRADAS_DIR, filetypes=tipos)
        
        if ruta:
            self.archivo_seleccionado = ruta
            nombre_base = os.path.basename(ruta)
            self.lbl_archivo.configure(text=nombre_base, text_color="#ffffff")
            self.btn_procesar.configure(state="normal")
            
            ext = os.path.splitext(ruta)[1].lower()
            if ext in [".xlsx", ".xls"]:
                try:
                    lector = LectorExcel(ruta)
                    df_metas = lector.cargar_metas()
                    registros = len(df_metas) if df_metas is not None else 0
                    self.kpi_filas.actualizar(str(registros))
                except Exception:
                    self.kpi_filas.actualizar("Cargado")
            
            self.kpi_estado.actualizar("Cargado")
            self.log(f"Archivo seleccionado: {nombre_base}")

    def ejecutar_procesamiento(self):
        if not self.archivo_seleccionado:
            return

        valida, msj = validar_plantilla_word(self.ruta_plantilla_word)
        if not valida:
            messagebox.showwarning("Validación de Plantilla", msj)
            self.log(f"⚠️ {msj}")

        self.btn_procesar.configure(state="disabled", text="⌛ Procesando... Por favor espere")
        self.kpi_estado.actualizar("Procesando...")

        hilo = threading.Thread(target=self._tarea_procesamiento_background, daemon=True)
        hilo.start()

    def _tarea_procesamiento_background(self):
        ext = os.path.splitext(self.archivo_seleccionado)[1].lower()
        trimestre_num = int(self.combo_trimestre_reporte.get()[0])
        anio_sel = int(self.combo_anio.get())
        fecha_emision = self.obtener_fecha_seleccionada()

        # Cargar API key directamente desde las variables de entorno (.env)
        api_key = os.getenv("GEMINI_API_KEY", "").strip()

        self.log("----------------------------------------")
        self.log(f"Iniciando procesamiento de metas ({trimestre_num}° Trimestre {anio_sel})...")

        try:
            if ext in [".xlsx", ".xls"]:
                ruta_salida, hallazgos = ProcesadorUIPPE.procesar_archivo_excel(
                    ruta_excel=self.archivo_seleccionado,
                    trimestre_num=trimestre_num,
                    anio=anio_sel,
                    fecha_emision=fecha_emision,
                    solo_programadas=True,
                    api_key_gemini=api_key if len(api_key) > 0 else None
                )

                for h in hallazgos:
                    self.log(f"  • {h}")

                self.after(0, self._finalizar_procesamiento, True, f"Reporte UIPPE del {trimestre_num}° Trimestre {anio_sel} generado correctamente:\n\n{os.path.basename(ruta_salida)}", ruta_salida, "Completado")

        except ValueError as val_err:
            self.after(0, self._finalizar_procesamiento, False, str(val_err), None, "Incongruente")

        except Exception as e:
            self.after(0, self._finalizar_procesamiento, False, str(e), None, "Error")

    def _finalizar_procesamiento(self, exito, mensaje, ruta_salida=None, estado_kpi="Completado"):
        self.btn_procesar.configure(state="normal", text="▶ Ejecutar Validación y Generar Reporte")
        self.kpi_estado.actualizar(estado_kpi)

        if exito:
            self.log("✔ Reporte generado con éxito.")
            if ruta_salida:
                self.log(f"Ubicación: {ruta_salida}")
            messagebox.showinfo("Proceso Exitoso", mensaje)
        else:
            if estado_kpi == "Incongruente":
                self.log(f"AVISO DE VALIDACIÓN: {mensaje}")
                messagebox.showwarning("Incongruencia de Datos", mensaje)
            else:
                self.log(f"ERROR CRÍTICO: {mensaje}")
                messagebox.showerror("Error de Procesamiento", mensaje)

    def log(self, mensaje):
        self.txt_log.insert("end", f"> {mensaje}\n")
        self.txt_log.see("end")
        registrar_evento(mensaje)


if __name__ == "__main__":
    app = AppUIPPE()
    app.mainloop()