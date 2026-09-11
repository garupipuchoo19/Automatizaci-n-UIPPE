import os
import threading
import customtkinter as ctk
from tkinter import filedialog, messagebox

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

class AppUIPPE(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Sistema de Automatización y Evaluación PbRM - UIPPE OPERAGUA")
        self.geometry("1150x780")
        self.minsize(900, 650)

        self.archivo_seleccionado = None
        self.ruta_excel_maestro = os.path.join(ENTRADAS_DIR, "CUAUTITLAN METAS E INDICADORES 2026.xlsx")
        self.ruta_plantilla_word = os.path.join(PLANTILLAS_DIR, "Plantilla_Oficial_UIPPE.docx")

        # Instanciar el motor backend de inyección
        self.inyector = InyectorExcelPbRM(self.ruta_excel_maestro)

        # Encabezado Institucional Principal
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

        # --- NAVEGACIÓN POR PESTAÑAS (TABVIEW) ---
        self.tabview = ctk.CTkTabview(self)
        self.tabview.pack(padx=20, pady=(0, 15), fill="both", expand=True)

        self.tab_captura = self.tabview.add("📝 Captura Rápida por Área")
        self.tab_reportes = self.tabview.add("📊 Consolidación y Reportes UIPPE")

        # --- PESTAÑA 1: CAPTURA RÁPIDA ---
        self.vista_captura = VistaCapturaRapida(
            parent=self.tab_captura,
            ruta_excel_maestro=self.ruta_excel_maestro,
            al_guardar_callback=self.procesar_guardado_captura
        )
        self.vista_captura.pack(fill="both", expand=True)

        # --- PESTAÑA 2: REPORTEADOR Y CONSOLIDACIÓN ---
        self.construir_pestaña_reportes()

    def construir_pestaña_reportes(self):
        """Construye la vista de carga de insumos, métricas y log de reportes."""
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

        # --- CONFIGURACIÓN DE PARÁMETROS Y API KEY ---
        self.frame_config = ctk.CTkFrame(self.tab_reportes)
        self.frame_config.pack(padx=10, pady=5, fill="x")

        lbl_trim = ctk.CTkLabel(
            self.frame_config, 
            text="Trimestre:", 
            font=ctk.CTkFont(weight="bold")
        )
        lbl_trim.pack(side="left", padx=(15, 5), pady=10)

        self.combo_trimestre_reporte = ctk.CTkComboBox(
            self.frame_config,
            values=["1° TRIMESTRE", "2° TRIMESTRE", "3° TRIMESTRE", "4° TRIMESTRE"],
            state="readonly",
            width=140
        )
        self.combo_trimestre_reporte.set("1° TRIMESTRE")
        self.combo_trimestre_reporte.pack(side="left", padx=5, pady=10)

        lbl_api = ctk.CTkLabel(
            self.frame_config, 
            text="Gemini API Key:", 
            font=ctk.CTkFont(weight="bold")
        )
        lbl_api.pack(side="left", padx=(20, 5), pady=10)

        self.txt_api_key = ctk.CTkEntry(
            self.frame_config,
            placeholder_text="Pega tu clave de Google AI Studio aquí...",
            show="*",
            width=280
        )
        self.txt_api_key.pack(side="left", padx=5, pady=10, fill="x", expand=True)

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

    def procesar_guardado_captura(self, df_captura, trimestre_num):
        """Callback invocado por VistaCapturaRapida para inyectar datos en el Excel maestro."""
        try:
            ruta_generada = self.inyector.inyectar_captura(df_captura, trimestre=trimestre_num)
            
            messagebox.showinfo(
                "¡Inyección Exitosa!", 
                f"Se actualizaron {len(df_captura)} metas del Trimestre {trimestre_num} correctamente.\n\n"
                f"Archivo generado en:\n{ruta_generada}"
            )
            self.log(f"✔ Captura inyectada exitosamente para Trimestre {trimestre_num} ({len(df_captura)} metas).")
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
        """Lanza la ejecución en un hilo secundario para evitar congelar la interfaz."""
        if not self.archivo_seleccionado:
            return

        # 1. Validar plantilla Word antes de empezar
        valida, msj = validar_plantilla_word(self.ruta_plantilla_word)
        if not valida:
            messagebox.showwarning("Validación de Plantilla", msj)
            self.log(f"⚠️ {msj}")

        # 2. Bloquear botón y actualizar estado visual
        self.btn_procesar.configure(state="disabled", text="⌛ Procesando... Por favor espere")
        self.kpi_estado.actualizar("Procesando...")

        # 3. Lanzar procesamiento en segundo plano
        hilo = threading.Thread(target=self._tarea_procesamiento_background, daemon=True)
        hilo.start()

    def _tarea_procesamiento_background(self):
        """Lógica pesada ejecutada fuera del hilo de la interfaz gráfica."""
        ext = os.path.splitext(self.archivo_seleccionado)[1].lower()
        trimestre_num = int(self.combo_trimestre_reporte.get()[0])
        api_key = self.txt_api_key.get().strip()

        self.log("----------------------------------------")
        self.log(f"Iniciando procesamiento de metas programadas para el {trimestre_num}° Trimestre...")

        try:
            if ext in [".xlsx", ".xls"]:
                ruta_salida, hallazgos = ProcesadorUIPPE.procesar_archivo_excel(
                    ruta_excel=self.archivo_seleccionado,
                    trimestre_num=trimestre_num,
                    solo_programadas=True,
                    api_key_gemini=api_key if len(api_key) > 0 else None
                )

                for h in hallazgos:
                    self.log(f"  • {h}")

                # Notificar éxito a la interfaz
                self.after(0, self._finalizar_procesamiento, True, f"Reporte UIPPE del {trimestre_num}° Trimestre generado correctamente:\n\n{os.path.basename(ruta_salida)}", ruta_salida, "Completado")

        except ValueError as val_err:
            self.after(0, self._finalizar_procesamiento, False, str(val_err), None, "Incongruente")

        except Exception as e:
            self.after(0, self._finalizar_procesamiento, False, str(e), None, "Error")

    def _finalizar_procesamiento(self, exito, mensaje, ruta_salida=None, estado_kpi="Completado"):
        """Restaura los controles de la interfaz y muestra el resultado en el hilo principal."""
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