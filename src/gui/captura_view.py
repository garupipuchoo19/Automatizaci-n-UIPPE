import customtkinter as ctk
import pandas as pd
from tkinter import messagebox
import os

class VistaCapturaRapida(ctk.CTkFrame):
    def __init__(self, parent, ruta_excel_maestro, al_guardar_callback=None):
        super().__init__(parent)
        # Asegurar conversión a string para evitar problemas con pathlib.Path
        self.ruta_excel = str(ruta_excel_maestro) if ruta_excel_maestro else ""
        self.al_guardar_callback = al_guardar_callback
        
        self.df_metas = None
        self.df_indicadores = None
        self.mapa_estructuras_metas = {}
        self.mapa_estructuras_ind = {}
        
        self.entradas_captura = {}  # Guarda los widgets de entrada mapeados por CLAVE
        
        self.cargar_datos_base()
        self.crear_interfaz()

    def cargar_datos_base(self):
        """Carga las hojas METAS e INDICADORES ajustando header=2 según la plantilla oficial."""
        if not self.ruta_excel or not os.path.exists(self.ruta_excel):
            messagebox.showerror(
                "Error de Archivo", 
                f"No se encontró el archivo maestro en la ruta:\n{self.ruta_excel}"
            )
            return

        try:
            # 1. Cargar Hoja METAS (Encabezados oficiales en la fila 2)
            df_m = pd.read_excel(self.ruta_excel, sheet_name="METAS", header=2)
            df_m.columns = df_m.columns.astype(str).str.strip()
            
            if 'CLAVE' in df_m.columns and 'ESTRUCTURA' in df_m.columns:
                self.df_metas = df_m.dropna(subset=['CLAVE', 'ESTRUCTURA']).copy()
                self.df_metas['ESTRUCTURA'] = self.df_metas['ESTRUCTURA'].astype(str).str.strip()
                
                col_area = 'ÁREA ENCARGADA' if 'ÁREA ENCARGADA' in self.df_metas.columns else 'DEPENDENCIA GENERAL'
                sub_m = self.df_metas[['ESTRUCTURA', col_area]].drop_duplicates().dropna()
                self.mapa_estructuras_metas = dict(zip(sub_m['ESTRUCTURA'], sub_m[col_area].astype(str)))
            else:
                self.df_metas = pd.DataFrame()

            # 2. Cargar Hoja INDICADORES (Encabezados también en header=2)
            df_i = pd.read_excel(self.ruta_excel, sheet_name="INDICADORES", header=2)
            df_i.columns = df_i.columns.astype(str).str.strip()
            
            if 'CLAVE' in df_i.columns and 'ESTRUCTURA' in df_i.columns:
                self.df_indicadores = df_i.dropna(subset=['CLAVE', 'ESTRUCTURA']).copy()
                self.df_indicadores['ESTRUCTURA'] = self.df_indicadores['ESTRUCTURA'].astype(str).str.strip()
                
                col_area_i = 'ÁREA ENCARGADA' if 'ÁREA ENCARGADA' in self.df_indicadores.columns else 'DEPENDENCIA GENERAL'
                sub_i = self.df_indicadores[['ESTRUCTURA', col_area_i]].drop_duplicates().dropna()
                self.mapa_estructuras_ind = dict(zip(sub_i['ESTRUCTURA'], sub_i[col_area_i].astype(str)))
            else:
                self.df_indicadores = pd.DataFrame()

        except Exception as e:
            messagebox.showerror("Error de Carga", f"Error al procesar el archivo Excel maestro:\n{str(e)}")

    def crear_interfaz(self):
        # --- CABECERA ---
        lbl_titulo = ctk.CTkLabel(
            self, 
            text="Módulo de Captura y Edición por Estructura Programática", 
            font=ctk.CTkFont(size=18, weight="bold")
        )
        lbl_titulo.pack(pady=(15, 5), padx=20, anchor="w")

        lbl_sub = ctk.CTkLabel(
            self, 
            text="Seleccione el tipo de registro (Metas o Indicadores) y la Estructura Programática correspondiente.",
            font=ctk.CTkFont(size=12),
            text_color="gray"
        )
        lbl_sub.pack(pady=(0, 15), padx=20, anchor="w")

        # --- PANEL DE FILTROS Y SELECCIÓN ---
        frame_filtros = ctk.CTkFrame(self)
        frame_filtros.pack(fill="x", padx=20, pady=5)

        # 1. Selector de Modo (Metas o Indicadores)
        ctk.CTkLabel(frame_filtros, text="Modo:", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.combo_tipo = ctk.CTkComboBox(
            frame_filtros, 
            values=["Metas", "Indicadores"],
            width=130,
            state="readonly",
            command=self.al_cambiar_tipo_registro
        )
        self.combo_tipo.set("Metas")
        self.combo_tipo.grid(row=0, column=1, padx=10, pady=10)

        # 2. Trimestre
        ctk.CTkLabel(frame_filtros, text="Trimestre:", font=ctk.CTkFont(weight="bold")).grid(row=0, column=2, padx=10, pady=10, sticky="w")
        self.combo_trimestre = ctk.CTkComboBox(
            frame_filtros, 
            values=["1° TRIMESTRE", "2° TRIMESTRE", "3° TRIMESTRE", "4° TRIMESTRE"],
            width=140,
            state="readonly",
            command=lambda _: self.actualizar_tabla_captura()
        )
        self.combo_trimestre.set("1° TRIMESTRE")
        self.combo_trimestre.grid(row=0, column=3, padx=10, pady=10)

        # 3. Estructura Programática
        ctk.CTkLabel(frame_filtros, text="Estructura:", font=ctk.CTkFont(weight="bold")).grid(row=0, column=4, padx=10, pady=10, sticky="w")
        self.combo_estructura = ctk.CTkComboBox(
            frame_filtros, 
            values=[],
            width=230,
            state="readonly",
            command=lambda _: self.actualizar_tabla_captura()
        )
        self.combo_estructura.grid(row=0, column=5, padx=10, pady=10)

        # --- CONTENEDOR SCROLLABLE PARA LA TABLA ---
        self.scroll_frame = ctk.CTkScrollableFrame(self, label_text="Registros por Estructura y Área Encargada")
        self.scroll_frame.pack(fill="both", expand=True, padx=20, pady=10)

        # --- BOTONERA DE ACCIÓN ---
        frame_acciones = ctk.CTkFrame(self, fg_color="transparent")
        frame_acciones.pack(fill="x", padx=20, pady=10)

        btn_guardar = ctk.CTkButton(
            frame_acciones, 
            text="💾 Guardar Cambios en Excel PbRM", 
            fg_color="#1F4E78", 
            hover_color="#143450",
            font=ctk.CTkFont(weight="bold"),
            command=self.guardar_captura
        )
        btn_guardar.pack(side="right", padx=10)

        # Cargar catálogo de estructuras inicial
        self.al_cambiar_tipo_registro("Metas")

    def al_cambiar_tipo_registro(self, tipo_sel):
        """Actualiza el catálogo de estructuras según el modo seleccionado (Metas o Indicadores)."""
        if tipo_sel == "Metas":
            estructuras = sorted(list(self.mapa_estructuras_metas.keys()))
        else:
            estructuras = sorted(list(self.mapa_estructuras_ind.keys()))

        self.combo_estructura.configure(values=estructuras)
        if estructuras:
            self.combo_estructura.set(estructuras[0])
        else:
            self.combo_estructura.set("")
            
        self.actualizar_tabla_captura()

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

    def actualizar_tabla_captura(self):
        """Renderiza los datos e incluye la columna de Área Encargada en la tabla."""
        if not hasattr(self, 'scroll_frame') or self.scroll_frame is None:
            return

        for widget in self.scroll_frame.winfo_children():
            widget.destroy()
        
        self.entradas_captura.clear()

        try:
            self.scroll_frame._parent_canvas.yview_moveto(0.0)
        except Exception:
            pass

        tipo_sel = self.combo_tipo.get()
        est_sel = self.combo_estructura.get()
        trim_sel = self.combo_trimestre.get()

        df_base = self.df_metas if tipo_sel == "Metas" else self.df_indicadores

        if df_base is None or df_base.empty or est_sel == "":
            lbl_vacio = ctk.CTkLabel(self.scroll_frame, text="No hay registros disponibles.")
            lbl_vacio.pack(pady=20)
            return

        col_progr = self.obtener_columna_programada(trim_sel)
        col_avance = self.obtener_columna_avance(trim_sel)
        col_desc = 'DESCRIPCIÓN DE LA META FÍSICA' if tipo_sel == "Metas" else 'DESCRIPCIÓN DEL INDICADOR'

        # Filtrar por la Estructura seleccionada
        registros_filtrados = df_base[df_base['ESTRUCTURA'].astype(str) == str(est_sel)]

        if registros_filtrados.empty:
            lbl_vacio = ctk.CTkLabel(self.scroll_frame, text="No hay registros asignados a esta estructura.")
            lbl_vacio.pack(pady=20)
            return

        # Encabezados de la Tabla
        headers = ["Estructura", "Área Encargada", "Descripción", "U. Medida", "Progr.", "Avance Real", "Justificación"]
        widths = [130, 220, 280, 80, 50, 80, 180]

        for col_idx, (h_text, w) in enumerate(zip(headers, widths)):
            lbl_h = ctk.CTkLabel(
                self.scroll_frame, 
                text=h_text, 
                font=ctk.CTkFont(size=11, weight="bold"),
                width=w,
                anchor="w" if col_idx in [1, 2, 6] else "center"
            )
            lbl_h.grid(row=0, column=col_idx, padx=4, pady=5, sticky="w")

        # Inyección de Filas
        for row_idx, (_, row_data) in enumerate(registros_filtrados.iterrows(), start=1):
            clave_reg = str(row_data['CLAVE'])
            cod_est = str(row_data['ESTRUCTURA'])
            area_encargada = str(row_data.get('ÁREA ENCARGADA', ''))
            desc_reg = str(row_data.get(col_desc, ''))
            u_medida = str(row_data.get('UNIDAD DE MEDIDA', ''))
            val_progr = row_data.get(col_progr, 0)
            val_avance_previo = row_data.get(col_avance, "")

            if pd.isna(val_avance_previo):
                val_avance_previo = ""

            # Label Estructura
            ctk.CTkLabel(self.scroll_frame, text=cod_est, width=130, anchor="center").grid(row=row_idx, column=0, padx=4, pady=4)
            
            # Label Área Encargada
            area_fmt = area_encargada[:32] + "..." if len(area_encargada) > 32 else area_encargada
            ctk.CTkLabel(self.scroll_frame, text=area_fmt, width=220, anchor="w", justify="left").grid(row=row_idx, column=1, padx=4, pady=4, sticky="w")

            # Label Descripción
            desc_fmt = desc_reg[:45] + "..." if len(desc_reg) > 45 else desc_reg
            ctk.CTkLabel(self.scroll_frame, text=desc_fmt, width=280, anchor="w", justify="left").grid(row=row_idx, column=2, padx=4, pady=4, sticky="w")

            # Label U. Medida
            ctk.CTkLabel(self.scroll_frame, text=u_medida, width=80, anchor="center").grid(row=row_idx, column=3, padx=4, pady=4)

            # Label Programado
            ctk.CTkLabel(self.scroll_frame, text=str(val_progr), font=ctk.CTkFont(weight="bold"), width=50, anchor="center").grid(row=row_idx, column=4, padx=4, pady=4)

            # Input: Avance Real
            txt_avance = ctk.CTkEntry(self.scroll_frame, width=75, placeholder_text="0")
            txt_avance.insert(0, str(val_avance_previo))
            txt_avance.grid(row=row_idx, column=5, padx=4, pady=4)

            # Input: Justificación
            txt_just = ctk.CTkEntry(self.scroll_frame, width=180, placeholder_text="Opcional")
            txt_just.grid(row=row_idx, column=6, padx=4, pady=4)

            self.entradas_captura[clave_reg] = {
                "input_avance": txt_avance,
                "input_justificacion": txt_just,
                "tipo": tipo_sel
            }

    def guardar_captura(self):
        """Recopila y valida la información capturada."""
        datos_capturados = []
        tipo_sel = self.combo_tipo.get()
        
        for clave, widgets in self.entradas_captura.items():
            val_av = widgets["input_avance"].get().strip()
            val_just = widgets["input_justificacion"].get().strip()

            if val_av != "":
                try:
                    val_av_float = float(val_av)
                except ValueError:
                    messagebox.showwarning("Dato Inválido", f"El avance ingresado para la clave '{clave}' debe ser numérico.")
                    return

                datos_capturados.append({
                    "CLAVE": clave,
                    "AVANCE": val_av_float,
                    "JUSTIFICACION": val_just,
                    "TIPO": tipo_sel
                })

        if not datos_capturados:
            messagebox.showinfo("Sin Datos", "No hay avances para guardar.")
            return

        df_captura = pd.DataFrame(datos_capturados)
        trimestre_num = int(self.combo_trimestre.get()[0])

        if self.al_guardar_callback:
            self.al_guardar_callback(df_captura, trimestre_num, tipo_sel)
        else:
            messagebox.showinfo("Captura Registrada", f"Se guardaron {len(df_captura)} registros de {tipo_sel} para el Trimestre {trimestre_num}.")
            self.cargar_datos_base()
            self.actualizar_tabla_captura()