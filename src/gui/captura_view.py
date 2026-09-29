import os
import customtkinter as ctk
import pandas as pd
from tkinter import messagebox
from PIL import Image  # Requerido para el procesamiento de imágenes

class VistaCapturaRapida(ctk.CTkFrame):
    def __init__(self, parent, ruta_excel_maestro, al_guardar_callback=None):
        super().__init__(parent)
        self.ruta_excel = str(ruta_excel_maestro) if ruta_excel_maestro else ""
        self.al_guardar_callback = al_guardar_callback
        
        self.df_metas = None
        self.df_indicadores = None
        self.entradas_captura = {}

        self._bloquear_eventos = False
        
        self.cargar_datos_base()
        self.crear_interfaz()

    def cargar_datos_base(self):
        """Carga METAS e INDICADORES con sus columnas reales."""
        if not self.ruta_excel or not os.path.exists(self.ruta_excel):
            messagebox.showerror(
                "Error de Archivo", 
                f"No se encontró el archivo maestro en la ruta:\n{self.ruta_excel}"
            )
            return

        try:
            # 1. Cargar METAS
            df_m = pd.read_excel(self.ruta_excel, sheet_name="METAS", header=2)
            df_m.columns = df_m.columns.astype(str).str.strip()
            if 'CLAVE' in df_m.columns:
                self.df_metas = df_m.dropna(subset=['CLAVE']).copy()
                for col in ['PROGRAMA', 'PROYECTO', 'AUX', 'ESTRUCTURA']:
                    if col in self.df_metas.columns:
                        self.df_metas[col] = self.df_metas[col].fillna("").astype(str).str.strip()
            else:
                self.df_metas = pd.DataFrame()

            # 2. Cargar INDICADORES
            df_i = pd.read_excel(self.ruta_excel, sheet_name="INDICADORES", header=2)
            df_i.columns = df_i.columns.astype(str).str.strip()
            if 'CLAVE' in df_i.columns:
                self.df_indicadores = df_i.dropna(subset=['CLAVE']).copy()
                for col in ['PROGRAMA', 'PROYECTO', 'AUX', 'ESTRUCTURA', 'NIVEL', 'VARIABLES']:
                    if col in self.df_indicadores.columns:
                        self.df_indicadores[col] = self.df_indicadores[col].fillna("").astype(str).str.strip()
            else:
                self.df_indicadores = pd.DataFrame()

        except Exception as e:
            messagebox.showerror("Error de Carga", f"Error al procesar el archivo Excel maestro:\n{str(e)}")

    def crear_interfaz(self):
        # --- CABECERA PRINCIPAL (SOLO TEXTOS) ---
        lbl_titulo = ctk.CTkLabel(
            self, 
            text="Módulo de Captura y Edición de Avances (PbRM)", 
            font=ctk.CTkFont(size=18, weight="bold")
        )
        lbl_titulo.pack(pady=(15, 2), padx=20, anchor="w")

        lbl_sub = ctk.CTkLabel(
            self, 
            text="Filtre por Programa, Proyecto, Auxiliar o Estructura para capturar los avances físicos del trimestre.",
            font=ctk.CTkFont(size=12),
            text_color="gray"
        )
        lbl_sub.pack(pady=(0, 10), padx=20, anchor="w")

        # --- PANEL DE FILTROS SUPERIOR Y LOGO ---
        frame_filtros = ctk.CTkFrame(self)
        frame_filtros.pack(fill="x", padx=20, pady=5)

        # Fila 0: Modo, Trimestre y Estructura
        ctk.CTkLabel(frame_filtros, text="Modo:", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=(10, 2), pady=8, sticky="w")
        self.combo_tipo = ctk.CTkComboBox(
            frame_filtros, 
            values=["Metas", "Indicadores"],
            width=120,
            state="readonly",
            command=self.al_cambiar_modo
        )
        self.combo_tipo.set("Metas")
        self.combo_tipo.grid(row=0, column=1, padx=8, pady=8, sticky="w")

        ctk.CTkLabel(frame_filtros, text="Trimestre:", font=ctk.CTkFont(weight="bold")).grid(row=0, column=2, padx=(10, 2), pady=8, sticky="w")
        self.combo_trimestre = ctk.CTkComboBox(
            frame_filtros, 
            values=["1° TRIMESTRE", "2° TRIMESTRE", "3° TRIMESTRE", "4° TRIMESTRE"],
            width=130,
            state="readonly",
            command=lambda _: self.actualizar_tabla_captura()
        )
        self.combo_trimestre.set("1° TRIMESTRE")
        self.combo_trimestre.grid(row=0, column=3, padx=8, pady=8, sticky="w")

        ctk.CTkLabel(frame_filtros, text="Estructura:", font=ctk.CTkFont(weight="bold")).grid(row=0, column=4, padx=(10, 2), pady=8, sticky="w")
        self.combo_estructura = ctk.CTkComboBox(
            frame_filtros, 
            values=["TODOS"],
            width=160,
            state="readonly",
            command=lambda val: self.al_filtrar('ESTRUCTURA', val)
        )
        self.combo_estructura.set("TODOS")
        self.combo_estructura.grid(row=0, column=5, padx=8, pady=8, sticky="w")

        # Fila 1: Programa y Proyecto
        ctk.CTkLabel(frame_filtros, text="Programa:", font=ctk.CTkFont(weight="bold")).grid(row=1, column=0, padx=(10, 2), pady=8, sticky="w")
        self.combo_programa = ctk.CTkComboBox(
            frame_filtros, 
            values=["TODOS"],
            width=220,
            state="readonly",
            command=lambda val: self.al_filtrar('PROGRAMA', val)
        )
        self.combo_programa.set("TODOS")
        self.combo_programa.grid(row=1, column=1, columnspan=2, padx=8, pady=8, sticky="w")

        ctk.CTkLabel(frame_filtros, text="Proyecto:", font=ctk.CTkFont(weight="bold")).grid(row=1, column=3, padx=(10, 2), pady=8, sticky="w")
        self.combo_proyecto = ctk.CTkComboBox(
            frame_filtros, 
            values=["TODOS"],
            width=220,
            state="readonly",
            command=lambda val: self.al_filtrar('PROYECTO', val)
        )
        self.combo_proyecto.set("TODOS")
        self.combo_proyecto.grid(row=1, column=4, columnspan=2, padx=8, pady=8, sticky="w")

        # Fila 2: Auxiliar
        ctk.CTkLabel(frame_filtros, text="Auxiliar:", font=ctk.CTkFont(weight="bold")).grid(row=2, column=0, padx=(10, 2), pady=8, sticky="w")
        self.combo_auxiliar = ctk.CTkComboBox(
            frame_filtros, 
            values=["TODOS"],
            width=320,
            state="readonly",
            command=lambda val: self.al_filtrar('AUX', val)
        )
        self.combo_auxiliar.set("TODOS")
        self.combo_auxiliar.grid(row=2, column=1, columnspan=3, padx=8, pady=8, sticky="w")

        # --- CONTENEDOR Y CARGA DEL LOGO (COLUMNA DERECHA DEL FRAME DE FILTROS) ---
        frame_filtros.grid_columnconfigure(6, weight=1)

        ruta_logo = os.path.join(os.path.dirname(__file__), "..", "..", "assets", "logo_operagua.png")
        if os.path.exists(ruta_logo):
            try:
                img_pil = Image.open(ruta_logo)
                
                # Alto máximo fijo en píxeles
                alto_deseado = 65
                
                # Cálculo automático de ancho en función de la relación de aspecto original
                ancho_orig, alto_orig = img_pil.size
                proporcion = alto_deseado / float(alto_orig)
                ancho_calculado = int(ancho_orig * proporcion)
                
                logo_ctk = ctk.CTkImage(
                    light_image=img_pil, 
                    dark_image=img_pil, 
                    size=(ancho_calculado, alto_deseado)
                )
                
                lbl_logo = ctk.CTkLabel(frame_filtros, image=logo_ctk, text="")
                lbl_logo.grid(row=0, column=6, rowspan=3, padx=(0, 25), pady=8, sticky="e")
            except Exception as e:
                print(f"No se pudo cargar el logo de OPERAGUA: {e}")

        # --- CONTENEDOR SCROLLABLE PARA LA TABLA ---
        self.scroll_frame = ctk.CTkScrollableFrame(self, label_text="Registros Filtrados")
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

        self.poblar_catalogos_filtros()

    def al_cambiar_modo(self, _=None):
        """Reinicia todos los catálogos y resetea las selecciones al alternar modo."""
        self.poblar_catalogos_filtros()

    def poblar_catalogos_filtros(self):
        """Carga inicial de valores únicos en los 4 filtros sin restricciones."""
        self._bloquear_eventos = True
        df = self.df_metas if self.combo_tipo.get() == "Metas" else self.df_indicadores

        if df is not None and not df.empty:
            progs = ["TODOS"] + sorted([x for x in df['PROGRAMA'].unique() if str(x).strip()])
            proys = ["TODOS"] + sorted([x for x in df['PROYECTO'].unique() if str(x).strip()])
            auxs = ["TODOS"] + sorted([x for x in df['AUX'].unique() if str(x).strip()])
            ests = ["TODOS"] + sorted([x for x in df['ESTRUCTURA'].unique() if str(x).strip()])
        else:
            progs, proys, auxs, ests = ["TODOS"], ["TODOS"], ["TODOS"], ["TODOS"]

        self.combo_programa.configure(values=progs)
        self.combo_programa.set("TODOS")

        self.combo_proyecto.configure(values=proys)
        self.combo_proyecto.set("TODOS")

        self.combo_auxiliar.configure(values=auxs)
        self.combo_auxiliar.set("TODOS")

        self.combo_estructura.configure(values=ests)
        self.combo_estructura.set("TODOS")

        self._bloquear_eventos = False
        self.actualizar_tabla_captura()

    def al_filtrar(self, campo_modificado, nuevo_valor):
        """
        Recalcula las opciones disponibles en los demás desplegables (filtrado en cascada).
        Mantiene la coherencia cruzada entre Programa, Proyecto, Auxiliar y Estructura.
        """
        if self._bloquear_eventos:
            return

        self._bloquear_eventos = True

        df = self.df_metas if self.combo_tipo.get() == "Metas" else self.df_indicadores
        if df is None or df.empty:
            self._bloquear_eventos = False
            return

        # 1. Obtener valores seleccionados actualmente
        sel_prog = self.combo_programa.get()
        sel_proy = self.combo_proyecto.get()
        sel_aux = self.combo_auxiliar.get()
        sel_est = self.combo_estructura.get()

        # Diccionario para controlar qué filtros están activos
        filtros = {
            'PROGRAMA': sel_prog,
            'PROYECTO': sel_proy,
            'AUX': sel_aux,
            'ESTRUCTURA': sel_est
        }

        # 2. Función auxiliar para recalcular opciones relativas a las demás elecciones
        def obtener_opciones_relacionadas(columna_destino):
            cond = pd.Series(True, index=df.index)
            for col, val in filtros.items():
                if col != columna_destino and val != "TODOS":
                    cond &= (df[col] == val)
            opciones = sorted([x for x in df[cond][columna_destino].unique() if str(x).strip()])
            return ["TODOS"] + opciones

        # 3. Recalcular las opciones de cada ComboBox
        opciones_prog = obtener_opciones_relacionadas('PROGRAMA')
        opciones_proy = obtener_opciones_relacionadas('PROYECTO')
        opciones_aux = obtener_opciones_relacionadas('AUX')
        opciones_est = obtener_opciones_relacionadas('ESTRUCTURA')

        self.combo_programa.configure(values=opciones_prog)
        if sel_prog not in opciones_prog:
            self.combo_programa.set("TODOS")

        self.combo_proyecto.configure(values=opciones_proy)
        if sel_proy not in opciones_proy:
            self.combo_proyecto.set("TODOS")

        self.combo_auxiliar.configure(values=opciones_aux)
        if sel_aux not in opciones_aux:
            self.combo_auxiliar.set("TODOS")

        self.combo_estructura.configure(values=opciones_est)
        if sel_est not in opciones_est:
            self.combo_estructura.set("TODOS")

        self._bloquear_eventos = False
        
        # 4. Refrescar los datos en la tabla
        self.actualizar_tabla_captura()

    def obtener_df_filtrado(self):
        """Filtra el dataframe activo según los valores seleccionados."""
        df = self.df_metas if self.combo_tipo.get() == "Metas" else self.df_indicadores
        if df is None or df.empty:
            return pd.DataFrame()

        cond = pd.Series(True, index=df.index)

        p = self.combo_programa.get()
        if p != "TODOS":
            cond &= (df['PROGRAMA'] == p)

        pr = self.combo_proyecto.get()
        if pr != "TODOS":
            cond &= (df['PROYECTO'] == pr)

        ax = self.combo_auxiliar.get()
        if ax != "TODOS":
            cond &= (df['AUX'] == ax)

        est = self.combo_estructura.get()
        if est != "TODOS":
            cond &= (df['ESTRUCTURA'] == est)

        return df[cond]

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
        """Renderiza dinámicamente según el modo activo."""
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
        trim_sel = self.combo_trimestre.get()
        df_filtrado = self.obtener_df_filtrado()

        if df_filtrado.empty:
            lbl_vacio = ctk.CTkLabel(self.scroll_frame, text="No se encontraron registros con los filtros seleccionados.")
            lbl_vacio.pack(pady=20)
            return

        col_progr = self.obtener_columna_programada(trim_sel)
        col_avance = self.obtener_columna_avance(trim_sel)
        col_desc = 'DESCRIPCIÓN DE LA META FÍSICA' if tipo_sel == "Metas" else 'DESCRIPCIÓN DEL INDICADOR'

        if tipo_sel == "Indicadores":
            headers = ["Estructura", "Área Encargada", "Nivel", "Descripción Indicador", "Variables", "U. Medida", "Progr.", "Avance Real", "Justificación"]
            widths = [110, 160, 90, 220, 150, 80, 50, 75, 150]
        else:
            headers = ["Estructura", "Área Encargada", "Descripción Meta Física", "U. Medida", "Progr.", "Avance Real", "Justificación"]
            widths = [120, 200, 280, 90, 50, 80, 180]

        for col_idx, (h_text, w) in enumerate(zip(headers, widths)):
            lbl_h = ctk.CTkLabel(
                self.scroll_frame, 
                text=h_text, 
                font=ctk.CTkFont(size=11, weight="bold"),
                width=w,
                anchor="w" if "Descripción" in h_text or h_text in ["Área Encargada", "Variables", "Justificación"] else "center"
            )
            lbl_h.grid(row=0, column=col_idx, padx=3, pady=5, sticky="w")

        for row_idx, (_, row_data) in enumerate(df_filtrado.iterrows(), start=1):
            clave_reg = str(row_data['CLAVE'])
            cod_est = str(row_data.get('ESTRUCTURA', ''))
            area_encargada = str(row_data.get('ÁREA ENCARGADA', row_data.get('DEPENDENCIA GENERAL', '')))
            desc_reg = str(row_data.get(col_desc, ''))
            u_medida = str(row_data.get('UNIDAD DE MEDIDA', ''))
            val_progr = row_data.get(col_progr, 0)
            val_avance_previo = row_data.get(col_avance, "")

            if pd.isna(val_avance_previo):
                val_avance_previo = ""

            curr_col = 0

            # 1. Estructura
            ctk.CTkLabel(self.scroll_frame, text=cod_est, width=widths[curr_col], anchor="center").grid(row=row_idx, column=curr_col, padx=3, pady=3)
            curr_col += 1

            # 2. Área Encargada
            area_fmt = area_encargada[:26] + "..." if len(area_encargada) > 26 else area_encargada
            ctk.CTkLabel(self.scroll_frame, text=area_fmt, width=widths[curr_col], anchor="w", justify="left").grid(row=row_idx, column=curr_col, padx=3, pady=3, sticky="w")
            curr_col += 1

            # 3. Solo para Indicadores: Nivel
            if tipo_sel == "Indicadores":
                val_nivel = str(row_data.get('NIVEL', ''))
                nivel_fmt = val_nivel[:15] + "..." if len(val_nivel) > 15 else val_nivel
                ctk.CTkLabel(self.scroll_frame, text=nivel_fmt, width=widths[curr_col], anchor="center").grid(row=row_idx, column=curr_col, padx=3, pady=3)
                curr_col += 1

            # 4. Descripción
            desc_limit = 35 if tipo_sel == "Indicadores" else 45
            desc_fmt = desc_reg[:desc_limit] + "..." if len(desc_reg) > desc_limit else desc_reg
            ctk.CTkLabel(self.scroll_frame, text=desc_fmt, width=widths[curr_col], anchor="w", justify="left").grid(row=row_idx, column=curr_col, padx=3, pady=3, sticky="w")
            curr_col += 1

            # 5. Solo para Indicadores: Variables
            if tipo_sel == "Indicadores":
                val_vars = str(row_data.get('VARIABLES', ''))
                vars_fmt = val_vars[:22] + "..." if len(val_vars) > 22 else val_vars
                ctk.CTkLabel(self.scroll_frame, text=vars_fmt, width=widths[curr_col], anchor="w", justify="left").grid(row=row_idx, column=curr_col, padx=3, pady=3, sticky="w")
                curr_col += 1

            # 6. U. Medida
            ctk.CTkLabel(self.scroll_frame, text=u_medida, width=widths[curr_col], anchor="center").grid(row=row_idx, column=curr_col, padx=3, pady=3)
            curr_col += 1

            # 7. Programado
            ctk.CTkLabel(self.scroll_frame, text=str(val_progr), font=ctk.CTkFont(weight="bold"), width=widths[curr_col], anchor="center").grid(row=row_idx, column=curr_col, padx=3, pady=3)
            curr_col += 1

            # 8. Input: Avance Real
            txt_avance = ctk.CTkEntry(self.scroll_frame, width=widths[curr_col], placeholder_text="0")
            txt_avance.insert(0, str(val_avance_previo))
            txt_avance.grid(row=row_idx, column=curr_col, padx=3, pady=3)
            curr_col += 1

            # 9. Input: Justificación
            txt_just = ctk.CTkEntry(self.scroll_frame, width=widths[curr_col], placeholder_text="Opcional")
            txt_just.grid(row=row_idx, column=curr_col, padx=3, pady=3)

            self.entradas_captura[clave_reg] = {
                "input_avance": txt_avance,
                "input_justificacion": txt_just,
                "tipo": tipo_sel
            }

    def guardar_captura(self):
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
            messagebox.showinfo("Sin Datos", "No hay valores de avance capturados para guardar.")
            return

        df_captura = pd.DataFrame(datos_capturados)
        trimestre_num = int(self.combo_trimestre.get()[0])

        if self.al_guardar_callback:
            self.al_guardar_callback(df_captura, trimestre_num, tipo_sel)
        else:
            messagebox.showinfo("Captura Registrada", f"Se registraron {len(df_captura)} avances de {tipo_sel} para el Trimestre {trimestre_num}.")
            self.cargar_datos_base()
            self.poblar_catalogos_filtros()