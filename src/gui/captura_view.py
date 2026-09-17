import customtkinter as ctk
import pandas as pd
from tkinter import messagebox, ttk

class VistaCapturaRapida(ctk.CTkFrame):
    def __init__(self, parent, ruta_excel_maestro, al_guardar_callback=None):
        super().__init__(parent)
        self.ruta_excel = ruta_excel_maestro
        self.al_guardar_callback = al_guardar_callback
        
        self.df_metas = None
        self.df_areas = None
        self.entradas_captura = {}  # Guarda los widgets de entrada por CLAVE de meta
        
        self.cargar_datos_base()
        self.crear_interfaz()

    def cargar_datos_base(self):
        """Carga la hoja METAS del libro maestro para extraer áreas y catálogos."""
        try:
            self.df_metas = pd.read_excel(self.ruta_excel, sheet_name="METAS", header=2)
            # Limpiar nombres de columnas eliminando espacios innecesarios
            self.df_metas.columns = self.df_metas.columns.str.strip()
            
            # Extraer mapa de Dependencias y Áreas auxiliares
            self.df_areas = self.df_metas[['DEPENDENCIA GENERAL', 'AUX']].drop_duplicates().dropna()
        except Exception as e:
            messagebox.showerror("Error de Carga", f"No se pudo leer el archivo maestro:\n{e}")

    def crear_interfaz(self):
        # --- CABECERA ---
        lbl_titulo = ctk.CTkLabel(
            self, 
            text="Módulo de Captura Rápida de Avances Trimestrales", 
            font=ctk.CTkFont(size=18, weight="bold")
        )
        lbl_titulo.pack(pady=(15, 5), padx=20, anchor="w")

        lbl_sub = ctk.CTkLabel(
            self, 
            text="Seleccione el trimestre y área correspondiente para registrar los avances físicos.",
            font=ctk.CTkFont(size=12),
            text_color="gray"
        )
        lbl_sub.pack(pady=(0, 15), padx=20, anchor="w")

        # --- PANEL DE FILTROS Y SELECCIÓN ---
        frame_filtros = ctk.CTkFrame(self)
        frame_filtros.pack(fill="x", padx=20, pady=5)

        # 1. Trimestre
        ctk.CTkLabel(frame_filtros, text="Trimestre:", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.combo_trimestre = ctk.CTkComboBox(
            frame_filtros, 
            values=["1° TRIMESTRE", "2° TRIMESTRE", "3° TRIMESTRE", "4° TRIMESTRE"],
            state="readonly",
            command=lambda _: self.actualizar_tabla_metas()
        )
        self.combo_trimestre.set("1° TRIMESTRE")
        self.combo_trimestre.grid(row=0, column=1, padx=10, pady=10)

        # 2. Dependencia General
        ctk.CTkLabel(frame_filtros, text="Dependencia:", font=ctk.CTkFont(weight="bold")).grid(row=0, column=2, padx=10, pady=10, sticky="w")
        deps_unicas = sorted(self.df_areas['DEPENDENCIA GENERAL'].unique().tolist()) if self.df_areas is not None else []
        self.combo_dep = ctk.CTkComboBox(
            frame_filtros, 
            values=deps_unicas,
            width=220,
            state="readonly",
            command=self.al_cambiar_dependencia
        )
        self.combo_dep.grid(row=0, column=3, padx=10, pady=10)

        # 3. Área / Departamento Encargado
        ctk.CTkLabel(frame_filtros, text="Área Encargada:", font=ctk.CTkFont(weight="bold")).grid(row=0, column=4, padx=10, pady=10, sticky="w")
        self.combo_area = ctk.CTkComboBox(
            frame_filtros, 
            values=[],
            width=250,
            state="readonly",
            command=lambda _: self.actualizar_tabla_metas()
        )
        self.combo_area.grid(row=0, column=5, padx=10, pady=10)

        # --- CONTENEDOR SCROLLABLE PARA LA TABLA DE CAPTURA ---
        self.scroll_frame = ctk.CTkScrollableFrame(self, label_text="Metas Asignadas al Área")
        self.scroll_frame.pack(fill="both", expand=True, padx=20, pady=10)

        # --- BOTONERA DE ACCIÓN ---
        frame_acciones = ctk.CTkFrame(self, fg_color="transparent")
        frame_acciones.pack(fill="x", padx=20, pady=10)

        btn_guardar = ctk.CTkButton(
            frame_acciones, 
            text="💾 Guardar Captura en Excel PbRM", 
            fg_color="#1F4E78", 
            hover_color="#143450",
            font=ctk.CTkFont(weight="bold"),
            command=self.guardar_captura
        )
        btn_guardar.pack(side="right", padx=10)

        # Inicializar combobox de área y cargar datos en la tabla
        if deps_unicas:
            self.combo_dep.set(deps_unicas[0])
            self.al_cambiar_dependencia(deps_unicas[0])

    def al_cambiar_dependencia(self, dep_seleccionada):
        """Filtra las áreas encargadas según la Dependencia General seleccionada."""
        if self.df_areas is None:
            return

        areas_filtradas = sorted(self.df_areas[self.df_areas['DEPENDENCIA GENERAL'] == dep_seleccionada]['AUX'].unique().tolist())
        self.combo_area.configure(values=areas_filtradas)
        if areas_filtradas:
            self.combo_area.set(areas_filtradas[0])
        else:
            self.combo_area.set("")
        self.actualizar_tabla_metas()

    def obtener_columna_programada(self, trimestre_str):
        mapa_cols = {
            "1° TRIMESTRE": "PROGR. 1° TRIM",
            "2° TRIMESTRE": "PROGR. 2° TRIM",
            "3° TRIMESTRE": "PROGR. 3° TRIM",
            "4° TRIMESTRE": "PROGR. 4° TRIM"
        }
        return mapa_cols.get(trimestre_str, "PROGR. 1° TRIM")

    def obtener_columna_avance(self, trimestre_str):
        mapa_cols = {
            "1° TRIMESTRE": "AVANCE 1°. TRIM",
            "2° TRIMESTRE": "AVANCE 2°. TRIM",
            "3° TRIMESTRE": "AVANCE 3°. TRIM",
            "4° TRIMESTRE": "AVANCE 4°. TRIM"
        }
        return mapa_cols.get(trimestre_str, "AVANCE 1°. TRIM")

    def actualizar_tabla_metas(self):
        """Limpia y vuelve a renderizar los campos de entrada de las metas filtradas."""
        if not hasattr(self, 'scroll_frame') or self.scroll_frame is None:
            return

        # Limpiar frame scrollable
        for widget in self.scroll_frame.winfo_children():
            widget.destroy()
        
        self.entradas_captura.clear()

        # Restablecer el scroll a la parte superior (0.0)
        try:
            self.scroll_frame._parent_canvas.yview_moveto(0.0)
        except Exception:
            pass

        if self.df_metas is None:
            lbl_vacio = ctk.CTkLabel(self.scroll_frame, text="No se pudieron cargar los datos de metas.")
            lbl_vacio.pack(pady=20)
            return

        dep_sel = self.combo_dep.get()
        area_sel = self.combo_area.get()
        trim_sel = self.combo_trimestre.get()

        col_progr = self.obtener_columna_programada(trim_sel)
        col_avance = self.obtener_columna_avance(trim_sel)

        # Filtrar metas de la base de datos
        metas_filtradas = self.df_metas[
            (self.df_metas['DEPENDENCIA GENERAL'] == dep_sel) & 
            (self.df_metas['AUX'] == area_sel)
        ]

        if metas_filtradas.empty:
            lbl_vacio = ctk.CTkLabel(self.scroll_frame, text="No hay metas registradas para el área seleccionada.")
            lbl_vacio.pack(pady=20)
            return

        # Encabezados de la Tabla de Captura
        headers = ["#", "Descripción Meta Física", "U. Medida", "Progr.", "Avance Real", "Justificación (Si aplica)"]
        widths = [30, 320, 90, 60, 90, 220]

        for col_idx, (h_text, w) in enumerate(zip(headers, widths)):
            lbl_h = ctk.CTkLabel(
                self.scroll_frame, 
                text=h_text, 
                font=ctk.CTkFont(size=11, weight="bold"),
                width=w,
                anchor="w" if col_idx in [1, 5] else "center"
            )
            lbl_h.grid(row=0, column=col_idx, padx=5, pady=5, sticky="w")

        # Filas de Metas
        for row_idx, (_, row_data) in enumerate(metas_filtradas.iterrows(), start=1):
            clave_meta = str(row_data['CLAVE'])
            num_meta = str(row_data['#'])
            desc_meta = str(row_data['DESCRIPCIÓN DE LA META FÍSICA'])
            u_medida = str(row_data['UNIDAD DE MEDIDA'])
            val_progr = row_data.get(col_progr, 0)
            val_avance_previo = row_data.get(col_avance, "")

            # Si es NaN, convertir a vacío
            if pd.isna(val_avance_previo):
                val_avance_previo = ""

            # Label Num Meta
            ctk.CTkLabel(self.scroll_frame, text=num_meta, width=30).grid(row=row_idx, column=0, padx=5, pady=5)
            
            # Label Descripción (Truncada a 55 chars para visualización limpia)
            desc_fmt = desc_meta[:55] + "..." if len(desc_meta) > 55 else desc_meta
            lbl_desc = ctk.CTkLabel(self.scroll_frame, text=desc_fmt, width=320, anchor="w", justify="left")
            lbl_desc.grid(row=row_idx, column=1, padx=5, pady=5, sticky="w")

            # Label U. Medida
            ctk.CTkLabel(self.scroll_frame, text=u_medida, width=90, anchor="center").grid(row=row_idx, column=2, padx=5, pady=5)

            # Label Valor Programado
            ctk.CTkLabel(self.scroll_frame, text=str(val_progr), font=ctk.CTkFont(weight="bold"), width=60, anchor="center").grid(row=row_idx, column=3, padx=5, pady=5)

            # Input: Avance Real
            txt_avance = ctk.CTkEntry(self.scroll_frame, width=80, placeholder_text="0")
            txt_avance.insert(0, str(val_avance_previo))
            txt_avance.grid(row=row_idx, column=4, padx=5, pady=5)

            # Input: Justificación
            txt_just = ctk.CTkEntry(self.scroll_frame, width=220, placeholder_text="Motivo de desfase (opcional)")
            txt_just.grid(row=row_idx, column=5, padx=5, pady=5)

            # Guardar referencia de inputs mapeados a la CLAVE de meta
            self.entradas_captura[clave_meta] = {
                "input_avance": txt_avance,
                "input_justificacion": txt_just,
                "programado": val_progr
            }

    def guardar_captura(self):
        """Recopila los valores ingresados en la tabla para enviarlos al motor de inyección."""
        datos_capturados = []
        
        for clave, widgets in self.entradas_captura.items():
            val_av = widgets["input_avance"].get().strip()
            val_just = widgets["input_justificacion"].get().strip()

            if val_av != "":
                try:
                    val_av_float = float(val_av)
                except ValueError:
                    messagebox.showwarning("Dato Inválido", f"El avance ingresado para la meta con clave '{clave}' debe ser un número válido.")
                    return

                datos_capturados.append({
                    "CLAVE": clave,
                    "AVANCE": val_av_float,
                    "JUSTIFICACION": val_just
                })

        if not datos_capturados:
            messagebox.showinfo("Sin Datos", "No hay valores de avance ingresados para guardar.")
            return

        df_captura = pd.DataFrame(datos_capturados)
        trimestre_num = int(self.combo_trimestre.get()[0]) # Extrae 1, 2, 3 o 4

        if self.al_guardar_callback:
            self.al_guardar_callback(df_captura, trimestre_num)
        else:
            messagebox.showinfo("Captura Registrada", f"Se capturaron {len(df_captura)} metas exitosamente para el Trimestre {trimestre_num}.")
            self.cargar_datos_base()
            self.actualizar_tabla_metas()