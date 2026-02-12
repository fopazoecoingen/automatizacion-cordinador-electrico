import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from core.descargar_archivos import (
    buscar_archivo_existente_tipo,
    descargar_y_descomprimir_zip_tipo,
    meses,
    TIPOS_ARCHIVO,
)
from core.leer_excel import (
    LectorBalance,
    leer_compra_venta_energia_gm_holdings,
    leer_ingresos_por_it,
    leer_total_ingresos_potencia_firme,
    leer_total_ingresos_sscc,
)
from core.plantilla_cliente import escribir_total_en_resultado


class InterfazInforme:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Generación de Informe Eléctrico")
        self.root.geometry("900x700")
        self.root.configure(bg="#E5E5E5")

        # Variables
        self.anyo_var = tk.IntVar(value=datetime.now().year)
        self.mes_combo = None
        self.barra_var = tk.StringVar(value="")
        self.procesando = False

        # Configurar estilo
        self.setup_styles()

        # Crear panel principal blanco
        self.create_main_panel()

        # Crear controles de ventana (simulación de los tres puntos)
        self.create_window_controls()

        # Crear sección de selección de período
        self.create_periodo_section()

        # Crear sección de configuración
        self.create_config_section()

        # Crear sección de selección de archivos (plantilla y destino)
        self.create_file_section()

        # Crear sección de progreso
        self.create_progress_section()

        # Crear botón de crear informe
        self.create_action_button()

    def setup_styles(self) -> None:
        """Configurar estilos personalizados."""
        style = ttk.Style()
        style.theme_use("clam")

        # Configurar estilo para el botón púrpura
        style.configure(
            "Purple.TButton",
            background="#7B2CBF",
            foreground="white",
            borderwidth=0,
            focuscolor="none",
            padding=10,
        )
        style.map("Purple.TButton", background=[("active", "#6A1B9A"), ("pressed", "#5A1A8A")])

    def create_window_controls(self) -> None:
        """Crear controles de ventana simulados (tres puntos)."""
        controls_frame = tk.Frame(self.root, bg="#E5E5E5")
        controls_frame.pack(fill=tk.X, padx=10, pady=5)

        for _ in range(3):
            dot = tk.Label(controls_frame, text="●", bg="#E5E5E5", fg="#999999", font=("Arial", 8))
            dot.pack(side=tk.LEFT, padx=2)

    def create_main_panel(self) -> None:
        """Crear panel principal blanco."""
        self.main_panel = tk.Frame(self.root, bg="white", relief=tk.FLAT)
        self.main_panel.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

    def create_periodo_section(self) -> None:
        """Crear sección de selección de año y mes."""
        periodo_frame = tk.Frame(self.main_panel, bg="white")
        periodo_frame.pack(fill=tk.X, padx=30, pady=20)

        # Título
        titulo_label = tk.Label(
            periodo_frame,
            text="Período",
            bg="white",
            font=("Arial", 10),
            anchor="w",
        )
        titulo_label.grid(row=0, column=0, padx=10, pady=(0, 5), sticky="w")

        seleccion_frame = tk.Frame(periodo_frame, bg="white")
        seleccion_frame.grid(row=1, column=0, padx=10, pady=5, sticky="ew")

        meses_lista = [f"{i:02d} - {meses[i]}" for i in range(1, 13)]

        año_label = tk.Label(
            seleccion_frame,
            text="Año:",
            bg="white",
            font=("Arial", 10, "bold"),
            anchor="w",
        )
        año_label.grid(row=0, column=0, padx=(0, 10), sticky="w")

        año_spinbox = tk.Spinbox(
            seleccion_frame,
            from_=2020,
            to=2030,
            textvariable=self.anyo_var,
            font=("Arial", 10),
            width=10,
        )
        año_spinbox.grid(row=0, column=1, padx=(0, 10), sticky="w")

        mes_label = tk.Label(
            seleccion_frame,
            text="Mes:",
            bg="white",
            font=("Arial", 10, "bold"),
            anchor="w",
        )
        mes_label.grid(row=0, column=2, padx=(0, 10), sticky="w")

        self.mes_combo = ttk.Combobox(
            seleccion_frame,
            values=meses_lista,
            state="readonly",
            font=("Arial", 10),
            width=18,
        )
        self.mes_combo.grid(row=0, column=3, sticky="w")
        self.mes_combo.current(datetime.now().month - 1)

        periodo_frame.grid_columnconfigure(0, weight=1)

    def create_config_section(self) -> None:
        """Crear sección de configuración con inputs de texto."""
        config_frame = tk.Frame(self.main_panel, bg="white")
        config_frame.pack(fill=tk.X, padx=30, pady=20)

        labels = ["Empresa", "Barra"]
        default_values = ["VIENTOS_DE_RENAICO", ""]

        self.entries = {}

        for i, (label, default) in enumerate(zip(labels, default_values)):
            # Label
            lbl = tk.Label(
                config_frame,
                text=label,
                bg="white",
                font=("Arial", 10),
                anchor="w",
            )
            lbl.grid(row=0, column=i, padx=10, pady=(0, 5), sticky="w")

            # Entry (campo de texto)
            entry = tk.Entry(config_frame, width=25, font=("Arial", 10))
            if default:
                entry.insert(0, default)
            entry.grid(row=1, column=i, padx=10, pady=5, sticky="ew")
            self.entries[label] = entry

        # Configurar pesos de columnas
        for i in range(2):
            config_frame.grid_columnconfigure(i, weight=1)

        # Nota sobre barra
        nota_label = tk.Label(
            config_frame,
            text="Nota: Deje 'Barra' vacío para procesar todas las barras",
            bg="white",
            font=("Arial", 8),
            fg="#666666",
            anchor="w",
        )
        nota_label.grid(row=2, column=0, columnspan=2, padx=10, pady=(5, 0), sticky="w")

    def create_file_section(self) -> None:
        """Crear sección de selección de archivos (plantilla y archivo de salida)."""
        file_frame = tk.Frame(self.main_panel, bg="white")
        file_frame.pack(fill=tk.X, padx=30, pady=20)

        # Plantilla base del cliente
        self.create_file_input(
            file_frame,
            "Plantilla base del cliente",
            "",
            0,
            modo="open",
        )

        # Destino del informe (copia basada en la plantilla)
        self.create_file_input(
            file_frame,
            "Ruta de destino del informe",
            "",
            1,
            modo="save",
        )

    def create_file_input(
        self,
        parent,
        label_text: str,
        default_path: str,
        row: int,
        modo: str = "open",
    ) -> None:
        """Crear un campo de entrada de archivo con botón de exploración."""
        # Label
        lbl = tk.Label(
            parent,
            text=label_text,
            bg="white",
            font=("Arial", 10),
            anchor="w",
        )
        lbl.grid(row=row * 2, column=0, padx=10, pady=(0, 5), sticky="w")

        # Frame para entrada y botón
        input_frame = tk.Frame(parent, bg="white")
        input_frame.grid(row=row * 2 + 1, column=0, padx=10, pady=5, sticky="ew")

        # Campo de texto
        entry = tk.Entry(input_frame, font=("Arial", 10), width=50)
        if default_path:
            entry.insert(0, default_path)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        # Guardar referencia al entry
        if label_text.startswith("Plantilla"):
            self.plantilla_entry = entry
        elif label_text.startswith("Ruta de destino"):
            self.destino_entry = entry

        # Botón de exploración
        browse_btn = tk.Button(
            input_frame,
            text="▼",
            font=("Arial", 8),
            bg="white",
            fg="#666666",
            relief=tk.FLAT,
            width=3,
            command=lambda: self.browse_file(entry, label_text, modo),
        )
        browse_btn.pack(side=tk.RIGHT)

        parent.grid_columnconfigure(0, weight=1)

    def browse_file(self, entry_widget, file_type: str, modo: str) -> None:
        """Abrir diálogo de selección de archivo."""
        if modo == "open":
            filename = filedialog.askopenfilename(
                title="Seleccionar plantilla",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            )
        else:
            filename = filedialog.asksaveasfilename(
                title="Seleccionar destino del informe",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            )

        if filename:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, filename)

    def create_progress_section(self) -> None:
        """Crear sección de progreso."""
        progress_frame = tk.Frame(self.main_panel, bg="white")
        progress_frame.pack(fill=tk.X, padx=30, pady=20)

        # Label de progreso
        self.progress_text_label = tk.Label(
            progress_frame,
            text="Esperando inicio del proceso...",
            bg="white",
            font=("Arial", 10),
            anchor="w",
        )
        self.progress_text_label.pack(fill=tk.X, pady=(0, 10))

        # Barra de progreso
        self.progress_var = tk.DoubleVar()
        self.progress_var.set(0)

        self.progress_bar = ttk.Progressbar(
            progress_frame,
            variable=self.progress_var,
            maximum=100,
            length=400,
            mode="determinate",
            style="TProgressbar",
        )
        self.progress_bar.pack(fill=tk.X)

        # Configurar color de la barra de progreso
        style = ttk.Style()
        style.configure(
            "TProgressbar",
            background="#20B2AA",
            troughcolor="#E0E0E0",
            borderwidth=0,
            lightcolor="#20B2AA",
            darkcolor="#20B2AA",
        )

    def create_action_button(self) -> None:
        """Crear botón de crear informe."""
        button_frame = tk.Frame(self.main_panel, bg="white")
        button_frame.pack(fill=tk.X, padx=30, pady=20)

        # Frame para alinear el botón a la derecha
        right_frame = tk.Frame(button_frame, bg="white")
        right_frame.pack(side=tk.RIGHT)

        # Botón púrpura
        self.create_btn = tk.Button(
            right_frame,
            text="Crear Informe",
            font=("Arial", 11, "bold"),
            bg="#7B2CBF",
            fg="white",
            relief=tk.FLAT,
            padx=30,
            pady=10,
            cursor="hand2",
            command=self.crear_informe,
        )
        self.create_btn.pack()

        # Efecto hover
        self.create_btn.bind("<Enter>", lambda e: self.create_btn.config(bg="#6A1B9A"))
        self.create_btn.bind("<Leave>", lambda e: self.create_btn.config(bg="#7B2CBF"))

    # --- Lógica de procesamiento -------------------------------------------------

    def crear_informe(self) -> None:
        """Función que se ejecuta al hacer clic en Crear Informe."""
        if self.procesando:
            messagebox.showwarning(
                "Proceso en curso",
                "Ya hay un proceso en ejecución. Por favor espere.",
            )
            return

        # Validar año y mes
        anyo = self.anyo_var.get()
        valor_mes = self.mes_combo.get()

        if not valor_mes:
            messagebox.showerror("Error", "Por favor seleccione un mes.")
            return

        try:
            mes = int(valor_mes.split(" - ")[0])
            if mes < 1 or mes > 12:
                messagebox.showerror(
                    "Error",
                    f"Mes inválido: {mes}. Debe estar entre 1 y 12.",
                )
                return
        except (ValueError, IndexError) as e:
            messagebox.showerror("Error", f"Error al obtener el mes: {e}")
            return

        if anyo < 2020 or anyo > 2030:
            messagebox.showerror(
                "Error",
                f"Año inválido: {anyo}. Debe estar entre 2020 y 2030.",
            )
            return

        print(f"[DEBUG] Procesando informe: {meses[mes]} {anyo}")

        ruta_plantilla = getattr(self, "plantilla_entry", None)
        ruta_destino_entry = getattr(self, "destino_entry", None)

        ruta_plantilla = ruta_plantilla.get().strip() if ruta_plantilla else ""
        ruta_destino = ruta_destino_entry.get().strip() if ruta_destino_entry else ""

        if not ruta_plantilla:
            messagebox.showerror("Error", "Por favor seleccione la plantilla base del cliente.")
            return

        if not ruta_destino:
            messagebox.showerror("Error", "Por favor seleccione una ruta de destino para el informe.")
            return

        # Tipos a descargar automáticamente por mes (Resultados, SSCC, Potencia — sin Antecedentes)
        tipos_a_descargar = ["energia_resultados", "sscc", "potencia"]

        nombre_barra = self.entries["Barra"].get().strip()
        nombre_empresa = self.entries["Empresa"].get().strip()

        # Iniciar proceso en hilo separado
        self.procesando = True
        self.create_btn.config(state=tk.DISABLED, text="Procesando...")
        self.progress_var.set(0)
        self.progress_text_label.config(text="Iniciando proceso...")

        thread = threading.Thread(
            target=self.procesar_informe_thread,
            args=(
                anyo,
                mes,
                ruta_plantilla,
                ruta_destino,
                nombre_barra,
                nombre_empresa,
                tipos_a_descargar,
            ),
            daemon=True,
        )
        thread.start()

    def procesar_informe_thread(
        self,
        anyo: int,
        mes: int,
        ruta_plantilla: str,
        ruta_destino: str,
        nombre_barra: str,
        nombre_empresa: str,
        tipos_seleccionados: list,
    ) -> None:
        """Procesar el informe en un hilo separado para un mes específico."""
        try:
            ruta_destino_path = Path(ruta_destino)
            ruta_destino_path.parent.mkdir(parents=True, exist_ok=True)
            from shutil import copyfile

            copyfile(ruta_plantilla, ruta_destino)
            print(f"[INFO] Plantilla copiada a destino: {ruta_destino}")

            # Procesar el mes
            self.procesar_mes(
                anyo,
                mes,
                ruta_destino,
                ruta_destino,
                nombre_barra,
                nombre_empresa,
                tipos_seleccionados,
                1,
                1,
            )

            # Mensaje final
            self.root.after(0, lambda: self.progress_var.set(100))
            self.root.after(
                0,
                lambda: self.progress_text_label.config(
                    text="✓ Proceso completado"
                ),
            )

            nombre_mes = meses[mes]
            mensaje = "Informe generado exitosamente.\n\n"
            mensaje += f"Período: {nombre_mes} {anyo}\n"
            if nombre_empresa:
                mensaje += f"Empresa: {nombre_empresa}\n"
            if nombre_barra:
                mensaje += f"Barra: {nombre_barra}\n"
            if not nombre_barra and not nombre_empresa:
                mensaje += "Barras: Todas (agrupadas)\n"
            mensaje += f"Archivo: {ruta_destino}"

            self.root.after(0, lambda: messagebox.showinfo("Éxito", mensaje))

        except Exception as e:
            error_msg = f"Error durante el procesamiento: {str(e)}"
            print(f"[ERROR] {error_msg}")
            import traceback

            traceback.print_exc()
            self.root.after(0, lambda: messagebox.showerror("Error", error_msg))
        finally:
            self.procesando = False
            self.root.after(
                0,
                lambda: self.create_btn.config(state=tk.NORMAL, text="Crear Informe"),
            )
            self.root.after(0, lambda: self.progress_var.set(0))

    def procesar_mes(
        self,
        anyo: int,
        mes: int,
        ruta_plantilla_destino: str,  # noqa: ARG002 - se usa como referencia lógica
        ruta_destino: str,
        nombre_barra: str,
        nombre_empresa: str,
        tipos_seleccionados: list,
        mes_actual: int,
        total_meses: int,
    ) -> None:
        """Procesar un mes individual del rango."""
        try:
            nombre_mes = meses[mes]
            print(f"\n[INFO] Procesando mes {mes_actual}/{total_meses}: {nombre_mes} {anyo}")

            # Calcular progreso base para este mes (distribuir 100% entre todos los meses)
            progreso_base = int((mes_actual - 1) * 100 / total_meses)
            progreso_por_mes = int(100 / total_meses)

            def calcular_progreso(porcentaje_mes: int) -> int:
                """Calcular el progreso global basado en el progreso de este mes."""
                return progreso_base + int(porcentaje_mes * progreso_por_mes / 100)

            # Paso 0: Validar estructura de carpetas (solo una vez)
            if mes_actual == 1:
                self.root.after(0, lambda: self.progress_var.set(calcular_progreso(2)))
                self.root.after(
                    0,
                    lambda: self.progress_text_label.config(
                        text="Validando estructura de carpetas..."
                    ),
                )

                carpeta_bd = Path("bd_data")
                carpeta_descomprimidos = carpeta_bd / "descomprimidos"

                if not carpeta_bd.exists():
                    carpeta_bd.mkdir(exist_ok=True)
                    self.root.after(
                        0,
                        lambda: self.progress_text_label.config(
                            text="✓ Carpeta 'bd_data' creada"
                        ),
                    )

                if not carpeta_descomprimidos.exists():
                    carpeta_descomprimidos.mkdir(parents=True, exist_ok=True)
                    self.root.after(
                        0,
                        lambda: self.progress_text_label.config(
                            text="✓ Carpeta 'descomprimidos' creada"
                        ),
                    )

            # Paso 1 y 2: Descargar/descomprimir cada tipo seleccionado (mismo proceso que ya existía)
            ruta_zip_energia = None  # Necesario para el informe
            codigo_error_energia = None
            total_tipos = len(tipos_seleccionados)

            for idx_tipo, tipo in enumerate(tipos_seleccionados):
                prog = calcular_progreso(5 + int(25 * idx_tipo / total_tipos))
                desc = TIPOS_ARCHIVO.get(tipo, tipo)
                self.root.after(0, lambda p=prog: self.progress_var.set(p))
                self.root.after(
                    0,
                    lambda d=desc, ma=mes_actual, tm=total_meses, nm=nombre_mes, a=anyo: self.progress_text_label.config(
                        text=f"[{ma}/{tm}] Descargando {d} para {nm} {a}..."
                    ),
                )

                archivo_existente = buscar_archivo_existente_tipo(anyo, mes, tipo)
                if archivo_existente:
                    self.root.after(
                        0,
                        lambda ma=mes_actual, tm=total_meses, ae=archivo_existente: self.progress_text_label.config(
                            text=f"[{ma}/{tm}] ✓ {ae.name} (ya existe)"
                        ),
                    )

                ruta_zip, ruta_descomprimida, codigo_error = descargar_y_descomprimir_zip_tipo(
                    anyo, mes, tipo, descomprimir=True, mostrar_progreso=False
                )

                if tipo == "energia_resultados":
                    ruta_zip_energia = ruta_zip
                    codigo_error_energia = codigo_error

                if ruta_zip:
                    self.root.after(
                        0,
                        lambda ma=mes_actual, tm=total_meses, rz=ruta_zip, d=desc: self.progress_text_label.config(
                            text=f"[{ma}/{tm}] ✓ {d}: {Path(rz).name}"
                        ),
                    )

            # Si energia_resultados falló, no podemos generar el informe para este mes
            if not ruta_zip_energia:
                ma, tm, nm, a = mes_actual, total_meses, nombre_mes, anyo
                if codigo_error_energia == 403:
                    self.root.after(
                        0,
                        lambda: self.progress_text_label.config(
                            text=f"[{ma}/{tm}] ✗ Error 403: Contenido no disponible para {nm} {a}"
                        ),
                    )
                    print(
                        f"[WARNING] Error 403 para {nombre_mes} {anyo}, "
                        "continuando con siguiente mes..."
                    )
                    return  # Continuar con el siguiente mes
                else:
                    self.root.after(
                        0,
                        lambda: self.progress_text_label.config(
                            text=f"[{ma}/{tm}] ✗ Error al descargar para {nm} {a}"
                        ),
                    )
                    print(
                        f"[WARNING] Error al descargar para {nombre_mes} {anyo}, "
                        "continuando con siguiente mes..."
                    )
                    return  # Continuar con el siguiente mes

            # Paso 3: Buscar archivo Balance
            self.root.after(0, lambda: self.progress_var.set(calcular_progreso(35)))
            self.root.after(
                0,
                lambda: self.progress_text_label.config(
                    text=f"[{mes_actual}/{total_meses}] Buscando archivo Balance..."
                ),
            )

            lector = LectorBalance(anyo, mes)

            self.root.after(0, lambda: self.progress_var.set(calcular_progreso(40)))
            self.root.after(
                0,
                lambda: self.progress_text_label.config(
                    text=(
                        f"[{mes_actual}/{total_meses}] ✓ Balance encontrado: "
                        f"{lector.ruta_archivo.name}"
                    )
                ),
            )

            # Paso 4: Leer Excel
            self.root.after(0, lambda: self.progress_var.set(calcular_progreso(45)))
            self.root.after(
                0,
                lambda: self.progress_text_label.config(
                    text=(
                        f"[{mes_actual}/{total_meses}] Leyendo hoja "
                        "'Balance Valorizado'..."
                    )
                ),
            )

            # Dejar que detecte automáticamente la fila de encabezados
            df_balance = lector.leer_balance_valorizado(header=None)

            self.root.after(0, lambda: self.progress_var.set(calcular_progreso(55)))
            self.root.after(
                0,
                lambda: self.progress_text_label.config(
                    text=(
                        f"[{mes_actual}/{total_meses}] ✓ Datos leídos: "
                        f"{len(df_balance)} filas"
                    )
                ),
            )

            # Paso 5: Calcular total monetario y escribir en plantilla del cliente
            self.root.after(0, lambda: self.progress_var.set(calcular_progreso(70)))

            # Preparar DataFrame filtrado como en guardar_en_plantilla
            columna_barra = None
            columna_monetario = None
            columna_empresa = None
            columna_fisico_kwh = None

            for col in df_balance.columns:
                col_lower = str(col).lower().replace(" ", "_")
                if col_lower == "barra":
                    columna_barra = col
                elif col_lower == "monetario":
                    columna_monetario = col
                elif "nombre_corto_empresa" in col_lower or "nombre corto empresa" in str(col).lower():
                    columna_empresa = col
                elif "fisico" in col_lower and "kwh" in col_lower:
                    columna_fisico_kwh = col

            # monetario: necesario para TOTAL INGRESOS POR ENERGIA CLP y fallback de POTENCIA FIRME
            if columna_monetario is None:
                print("[WARNING] No se encontró la columna 'monetario' en Balance Valorizado")

            if nombre_barra or nombre_empresa:
                df_guardar = df_balance.copy()

                if nombre_empresa:
                    if columna_empresa is None:
                        print("[ERROR] No se encontró la columna 'nombre_corto_empresa'")
                        return
                    df_guardar = df_guardar[
                        df_guardar[columna_empresa].astype(str).str.lower()
                        == nombre_empresa.lower()
                    ]

                if nombre_barra:
                    if columna_barra is None:
                        print("[ERROR] No se encontró la columna 'barra'")
                        return
                    df_guardar = df_guardar[
                        df_guardar[columna_barra].astype(str).str.lower()
                        == nombre_barra.lower()
                    ]

                mensaje_filtro = []
                if nombre_empresa:
                    mensaje_filtro.append(f"empresa: {nombre_empresa}")
                if nombre_barra:
                    mensaje_filtro.append(f"barra: {nombre_barra}")

                self.root.after(
                    0,
                    lambda: self.progress_text_label.config(
                        text=(
                            f"[{mes_actual}/{total_meses}] Filtrando datos por "
                            f"{', '.join(mensaje_filtro)}..."
                        )
                    ),
                )
            else:
                df_guardar = df_balance
                self.root.after(
                    0,
                    lambda: self.progress_text_label.config(
                        text=(
                            f"[{mes_actual}/{total_meses}] Calculando total monetario "
                            "para todas las barras..."
                        )
                    ),
                )

            self.root.after(0, lambda: self.progress_var.set(calcular_progreso(75)))

            # TOTAL INGRESOS POR POTENCIA FIRME CLP: leer desde Anexo 02.b Potencia
            # Tabla Datos: Empresa (B) | Potencia SEN (C) | TOTAL (D)
            total_monetario = leer_total_ingresos_potencia_firme(
                anyo, mes, nombre_empresa=nombre_empresa
            )
            if total_monetario is None:
                # Fallback: usar suma monetario del Balance Valorizado
                if columna_monetario is None:
                    print("[ERROR] No se encontró Anexo Potencia ni columna 'monetario' en Balance")
                    return
                total_monetario = (
                    df_guardar[columna_monetario].dropna().astype(float).sum()
                )
                print(
                    f"[INFO] Anexo Potencia no encontrado. Usando Balance Valorizado: "
                    f"{total_monetario:,.2f}"
                )
            else:
                print(f"[INFO] TOTAL INGRESOS POR POTENCIA FIRME CLP: {total_monetario:,.2f}")

            # Escribir TOTAL INGRESOS POR POTENCIA FIRME CLP en plantilla
            escribir_total_en_resultado(
                ruta_destino,
                anyo,
                mes,
                total_monetario,
                texto_concepto="TOTAL INGRESOS POR POTENCIA FIRME CLP",
            )

            # INGRESOS POR IT: Anexo 02.b Potencia, hoja 02.IT POTENCIA {Mes}-{YY} def
            total_it = leer_ingresos_por_it(anyo, mes, nombre_empresa=nombre_empresa)
            if total_it is not None:
                escribir_total_en_resultado(
                    ruta_destino,
                    anyo,
                    mes,
                    total_it,
                    texto_concepto="INGRESOS POR IT",
                )

            # TOTAL INGRESOS POR ENERGIA CLP: Balance Valorizado, columna monetario
            if columna_monetario is not None:
                total_energia = (
                    df_guardar[columna_monetario].dropna().astype(float).sum()
                )
                print(
                    f"[INFO] TOTAL INGRESOS POR ENERGIA CLP para {nombre_mes} {anyo}: "
                    f"{total_energia:,.2f} (Balance Valorizado, col monetario)"
                )
                escribir_total_en_resultado(
                    ruta_destino,
                    anyo,
                    mes,
                    total_energia,
                    texto_concepto="TOTAL INGRESOS POR ENERGIA CLP",
                )

            # TOTAL INGRESOS POR SSCC CLP: EXCEL 1_CUADROS_PAGO_SSCC, hoja CPI_
            # Filtra por Nemotecnico Deudor = empresa, suma columna Monto
            if nombre_empresa:
                total_sscc = leer_total_ingresos_sscc(anyo, mes, nombre_empresa)
                if total_sscc is not None:
                    escribir_total_en_resultado(
                        ruta_destino,
                        anyo,
                        mes,
                        total_sscc,
                        texto_concepto="TOTAL INGRESOS POR SSCC CLP",
                    )

            # Compra Venta Energia GM Holdings CLP: Balance, hoja Contratos, columna VENTA[CLP]
            total_gm_holdings = leer_compra_venta_energia_gm_holdings(
                anyo, mes,
                nombre_empresa=nombre_empresa,
                nombre_barra=nombre_barra,
            )
            if total_gm_holdings is not None:
                escribir_total_en_resultado(
                    ruta_destino,
                    anyo,
                    mes,
                    total_gm_holdings,
                    texto_concepto="Compra Venta Energia GM Holdings CLP",
                )

            # Calcular IMPORTACION MWh desde columna fisico_kwh (valor positivo, kWh -> MWh: /1000)
            if columna_fisico_kwh is not None:
                total_fisico_kwh = (
                    df_guardar[columna_fisico_kwh].dropna().astype(float).sum()
                )
                importacion_mwh = abs(total_fisico_kwh) / 1000.0
                print(
                    f"[INFO] IMPORTACION MWh para {nombre_mes} {anyo}: "
                    f"{importacion_mwh:,.2f} (desde {total_fisico_kwh:,.0f} kWh)"
                )
                print(f"  -> Dato obtenido (IMPORTACION MWh): {importacion_mwh:,.2f}")
                escribir_total_en_resultado(
                    ruta_destino,
                    anyo,
                    mes,
                    importacion_mwh,
                    texto_concepto="IMPORTACION MWh",
                )
            else:
                print("[WARNING] No se encontró la columna 'fisico_kwh' en Balance Valorizado")

            exito = True

            self.root.after(0, lambda: self.progress_var.set(calcular_progreso(85)))
            self.root.after(
                0,
                lambda: self.progress_text_label.config(
                    text=(
                        f"[{mes_actual}/{total_meses}] ✓ Guardado en hoja Resultado "
                        f"para {nombre_mes} {anyo}"
                    )
                ),
            )

            if exito:
                print(
                    f"[OK] Mes {mes_actual}/{total_meses} ({nombre_mes} {anyo}) "
                    "procesado exitosamente"
                )
            else:
                print(f"[ERROR] Error al guardar datos para {nombre_mes} {anyo}")

        except FileNotFoundError as e:
            print(
                f"[WARNING] Archivo no encontrado para {nombre_mes} {anyo}: {str(e)}"
            )
            self.root.after(
                0,
                lambda: self.progress_text_label.config(
                    text=(
                        f"[{mes_actual}/{total_meses}] ✗ Archivo no encontrado "
                        f"para {nombre_mes} {anyo}"
                    )
                ),
            )

        except Exception as e:
            print(f"[ERROR] Error procesando {nombre_mes} {anyo}: {str(e)}")
            import traceback

            traceback.print_exc()
            self.root.after(
                0,
                lambda: self.progress_text_label.config(
                    text=(
                        f"[{mes_actual}/{total_meses}] ✗ Error: {str(e)[:50]}..."
                    )
                ),
            )


def main() -> None:
    root = tk.Tk()
    app = InterfazInforme(root)
    root.mainloop()


__all__ = ["InterfazInforme", "main"]

