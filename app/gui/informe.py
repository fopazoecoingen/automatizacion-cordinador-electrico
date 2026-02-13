import json
import os
import sys
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

def _directorio_base_datos() -> Path:
    """
    Directorio base. Ejecutable: carpeta del .exe (ahí se crea bd_data).
    Desarrollo: carpeta actual.
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path.cwd()

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
    leer_ingresos_por_potencia,
    leer_total_ingresos_potencia_firme,
    leer_total_ingresos_sscc,
)
from core.plantilla_cliente import escribir_todos_en_resultado


# Paleta de colores profesional - Sector energético
COLORS = {
    "bg_main": "#F0F4F8",
    "bg_card": "#FFFFFF",
    "bg_header": "#1A365D",
    "accent": "#2B6CB0",
    "accent_hover": "#2C5282",
    "accent_light": "#EBF8FF",
    "text_primary": "#1A202C",
    "text_secondary": "#4A5568",
    "text_muted": "#718096",
    "border": "#E2E8F0",
    "success": "#38A169",
    "progress": "#3182CE",
}


class InterfazInforme:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Generación de Informe Eléctrico")
        self.root.geometry("920x760")
        self.root.minsize(700, 600)
        self.root.configure(bg=COLORS["bg_main"])

        # Variables
        self.anyo_var = tk.IntVar(value=datetime.now().year)
        self.mes_combo = None
        self.barra_var = tk.StringVar(value="")
        self.procesando = False

        # Configurar estilo
        self.setup_styles()

        # Crear header
        self.create_header()

        # Crear panel principal
        self.create_main_panel()

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

        # Cargar últimos datos guardados
        self._cargar_ultimos_datos()

    def _ruta_config(self) -> Path:
        """Ruta del archivo JSON con los últimos datos ingresados."""
        return _directorio_base_datos() / "config_ultimos_datos.json"

    def _guardar_ultimos_datos(
        self,
        anyo: int,
        mes: int,
        empresa: str,
        barra: str,
        nombre_medidor: str,
        plantilla: str,
        destino: str,
    ) -> None:
        """Guarda los últimos datos ingresados en un archivo JSON."""
        try:
            datos = {
                "anyo": anyo,
                "mes": mes,
                "empresa": empresa,
                "barra": barra,
                "nombre_medidor": nombre_medidor,
                "plantilla": plantilla,
                "destino": destino,
            }
            with open(self._ruta_config(), "w", encoding="utf-8") as f:
                json.dump(datos, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"[WARNING] No se pudieron guardar los últimos datos: {e}")

    def _cargar_ultimos_datos(self) -> None:
        """Carga los últimos datos guardados y los aplica a los inputs."""
        try:
            ruta = self._ruta_config()
            if not ruta.exists():
                return
            with open(ruta, "r", encoding="utf-8") as f:
                datos = json.load(f)
            if datos.get("anyo"):
                self.anyo_var.set(int(datos["anyo"]))
            if datos.get("mes") and self.mes_combo:
                mes = int(datos["mes"])
                if 1 <= mes <= 12:
                    self.mes_combo.current(mes - 1)
            if datos.get("empresa") is not None and "Empresa" in self.entries:
                self.entries["Empresa"].delete(0, tk.END)
                self.entries["Empresa"].insert(0, str(datos["empresa"]))
            if datos.get("barra") is not None and "Barra" in self.entries:
                self.entries["Barra"].delete(0, tk.END)
                self.entries["Barra"].insert(0, str(datos["barra"]))
            if datos.get("nombre_medidor") is not None and "Nombre Medidor" in self.entries:
                self.entries["Nombre Medidor"].delete(0, tk.END)
                self.entries["Nombre Medidor"].insert(0, str(datos["nombre_medidor"]))
            if datos.get("plantilla") and hasattr(self, "plantilla_entry"):
                self.plantilla_entry.delete(0, tk.END)
                self.plantilla_entry.insert(0, str(datos["plantilla"]))
            if datos.get("destino") and hasattr(self, "destino_entry"):
                self.destino_entry.delete(0, tk.END)
                self.destino_entry.insert(0, str(datos["destino"]))
        except Exception as e:
            print(f"[WARNING] No se pudieron cargar los últimos datos: {e}")

    def setup_styles(self) -> None:
        """Configurar estilos personalizados."""
        style = ttk.Style()
        style.theme_use("clam")

        style.configure(
            "Accent.TButton",
            background=COLORS["accent"],
            foreground="white",
            borderwidth=0,
            focuscolor="none",
            padding=(20, 12),
            font=("Segoe UI", 10, "bold"),
        )
        style.map(
            "Accent.TButton",
            background=[("active", COLORS["accent_hover"]), ("pressed", COLORS["accent_hover"])],
        )

    def create_header(self) -> None:
        """Crear cabecera de la aplicación."""
        header = tk.Frame(self.root, bg=COLORS["bg_header"], height=60)
        header.pack(fill=tk.X)
        header.pack_propagate(False)

        title = tk.Label(
            header,
            text="Generación de Informe Eléctrico",
            bg=COLORS["bg_header"],
            fg="white",
            font=("Segoe UI", 18, "bold"),
        )
        title.pack(side=tk.LEFT, padx=24, pady=18)

    def create_main_panel(self) -> None:
        """Crear panel principal con contenedor tipo card."""
        container = tk.Frame(self.root, bg=COLORS["bg_main"], padx=24, pady=20)
        container.pack(fill=tk.BOTH, expand=True)

        self.main_panel = tk.Frame(
            container,
            bg=COLORS["bg_card"],
            relief=tk.FLAT,
            highlightbackground=COLORS["border"],
            highlightthickness=1,
        )
        self.main_panel.pack(fill=tk.BOTH, expand=True)

    def create_periodo_section(self) -> None:
        """Crear sección de selección de año y mes."""
        periodo_frame = tk.Frame(self.main_panel, bg=COLORS["bg_card"])
        periodo_frame.pack(fill=tk.X, padx=32, pady=(24, 16))

        titulo_label = tk.Label(
            periodo_frame,
            text="Período",
            bg=COLORS["bg_card"],
            fg=COLORS["text_primary"],
            font=("Segoe UI", 11, "bold"),
            anchor="w",
        )
        titulo_label.grid(row=0, column=0, padx=0, pady=(0, 8), sticky="w")

        seleccion_frame = tk.Frame(periodo_frame, bg=COLORS["bg_card"])
        seleccion_frame.grid(row=1, column=0, padx=0, pady=0, sticky="ew")

        meses_lista = [f"{i:02d} - {meses[i]}" for i in range(1, 13)]

        año_label = tk.Label(
            seleccion_frame,
            text="Año",
            bg=COLORS["bg_card"],
            fg=COLORS["text_secondary"],
            font=("Segoe UI", 10),
            anchor="w",
        )
        año_label.grid(row=0, column=0, padx=(0, 8), sticky="w")

        año_spinbox = tk.Spinbox(
            seleccion_frame,
            from_=2020,
            to=2030,
            textvariable=self.anyo_var,
            font=("Segoe UI", 10),
            width=8,
            relief=tk.SOLID,
            bd=1,
            highlightthickness=0,
        )
        año_spinbox.grid(row=0, column=1, padx=(0, 20), sticky="w")

        mes_label = tk.Label(
            seleccion_frame,
            text="Mes",
            bg=COLORS["bg_card"],
            fg=COLORS["text_secondary"],
            font=("Segoe UI", 10),
            anchor="w",
        )
        mes_label.grid(row=0, column=2, padx=(0, 8), sticky="w")

        self.mes_combo = ttk.Combobox(
            seleccion_frame,
            values=meses_lista,
            state="readonly",
            font=("Segoe UI", 10),
            width=20,
        )
        self.mes_combo.grid(row=0, column=3, sticky="w")
        self.mes_combo.current(datetime.now().month - 1)

        periodo_frame.grid_columnconfigure(0, weight=1)

    def create_config_section(self) -> None:
        """Crear sección de configuración con inputs de texto."""
        config_frame = tk.Frame(self.main_panel, bg=COLORS["bg_card"])
        config_frame.pack(fill=tk.X, padx=32, pady=(8, 16))

        titulo_config = tk.Label(
            config_frame,
            text="Filtros de consulta",
            bg=COLORS["bg_card"],
            fg=COLORS["text_primary"],
            font=("Segoe UI", 11, "bold"),
            anchor="w",
        )
        titulo_config.grid(row=0, column=0, columnspan=3, padx=0, pady=(0, 8), sticky="w")

        labels = ["Empresa", "Barra", "Nombre Medidor"]
        default_values = ["VIENTOS_DE_RENAICO", "", ""]

        self.entries = {}

        for i, (label, default) in enumerate(zip(labels, default_values)):
            lbl = tk.Label(
                config_frame,
                text=label,
                bg=COLORS["bg_card"],
                fg=COLORS["text_secondary"],
                font=("Segoe UI", 10),
                anchor="w",
            )
            lbl.grid(row=1, column=i, padx=(0, 8), pady=(0, 4), sticky="w")

            entry = tk.Entry(
                config_frame,
                width=28,
                font=("Segoe UI", 10),
                relief=tk.SOLID,
                bd=1,
                fg=COLORS["text_primary"],
            )
            if default:
                entry.insert(0, default)
            entry.grid(row=2, column=i, padx=(0, 24 if i < 2 else 0), pady=(0, 0), sticky="ew")
            self.entries[label] = entry

        for i in range(3):
            config_frame.grid_columnconfigure(i, weight=1)

        nota_label = tk.Label(
            config_frame,
            text="Barra vacío = todas. Nombre Medidor aplica solo a IMPORTACION MWh y TOTAL INGRESOS POR ENERGIA CLP.",
            bg=COLORS["bg_card"],
            font=("Segoe UI", 9),
            fg=COLORS["text_muted"],
            anchor="w",
        )
        nota_label.grid(row=3, column=0, columnspan=3, padx=0, pady=(8, 0), sticky="w")

    def create_file_section(self) -> None:
        """Crear sección de selección de archivos (plantilla y archivo de salida)."""
        file_frame = tk.Frame(self.main_panel, bg=COLORS["bg_card"])
        file_frame.pack(fill=tk.X, padx=32, pady=(8, 16))

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
        lbl = tk.Label(
            parent,
            text=label_text,
            bg=COLORS["bg_card"],
            fg=COLORS["text_secondary"],
            font=("Segoe UI", 10),
            anchor="w",
        )
        lbl.grid(row=row * 2, column=0, padx=0, pady=(12, 4), sticky="w")

        input_frame = tk.Frame(parent, bg=COLORS["bg_card"])
        input_frame.grid(row=row * 2 + 1, column=0, padx=0, pady=(0, 4), sticky="ew")

        entry = tk.Entry(
            input_frame,
            font=("Segoe UI", 10),
            width=55,
            relief=tk.SOLID,
            bd=1,
            fg=COLORS["text_primary"],
        )
        if default_path:
            entry.insert(0, default_path)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))

        if label_text.startswith("Plantilla"):
            self.plantilla_entry = entry
        elif label_text.startswith("Ruta de destino"):
            self.destino_entry = entry

        browse_btn = tk.Button(
            input_frame,
            text="Examinar",
            font=("Segoe UI", 9),
            bg=COLORS["border"],
            fg=COLORS["text_primary"],
            relief=tk.FLAT,
            padx=14,
            pady=4,
            cursor="hand2",
            command=lambda: self.browse_file(entry, label_text, modo),
        )
        browse_btn.pack(side=tk.RIGHT)
        browse_btn.bind("<Enter>", lambda e: browse_btn.config(bg="#CBD5E0"))
        browse_btn.bind("<Leave>", lambda e: browse_btn.config(bg=COLORS["border"]))

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
        progress_frame = tk.Frame(self.main_panel, bg=COLORS["bg_card"])
        progress_frame.pack(fill=tk.X, padx=32, pady=(8, 16))

        self.progress_text_label = tk.Label(
            progress_frame,
            text="Esperando inicio del proceso...",
            bg=COLORS["bg_card"],
            fg=COLORS["text_secondary"],
            font=("Segoe UI", 10),
            anchor="w",
        )
        self.progress_text_label.pack(fill=tk.X, pady=(0, 8))

        self.progress_var = tk.DoubleVar()
        self.progress_var.set(0)

        style = ttk.Style()
        style.configure(
            "TProgressbar",
            background=COLORS["progress"],
            troughcolor=COLORS["border"],
            borderwidth=0,
            thickness=8,
        )
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            variable=self.progress_var,
            maximum=100,
            length=400,
            mode="determinate",
        )
        self.progress_bar.pack(fill=tk.X)

    def create_action_button(self) -> None:
        """Crear botón de crear informe."""
        button_frame = tk.Frame(self.main_panel, bg=COLORS["bg_card"])
        button_frame.pack(fill=tk.X, padx=32, pady=(16, 28))

        right_frame = tk.Frame(button_frame, bg=COLORS["bg_card"])
        right_frame.pack(side=tk.RIGHT)

        self.create_btn = tk.Button(
            right_frame,
            text="  Crear Informe  ",
            font=("Segoe UI", 11, "bold"),
            bg=COLORS["accent"],
            fg="white",
            relief=tk.FLAT,
            padx=28,
            pady=12,
            cursor="hand2",
            command=self.crear_informe,
            activebackground=COLORS["accent_hover"],
            activeforeground="white",
        )
        self.create_btn.pack()

        self.create_btn.bind("<Enter>", lambda e: self.create_btn.config(bg=COLORS["accent_hover"]))
        self.create_btn.bind("<Leave>", lambda e: self.create_btn.config(bg=COLORS["accent"]))

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
        nombre_medidor = self.entries["Nombre Medidor"].get().strip()

        # Guardar últimos datos para próxima ejecución
        self._guardar_ultimos_datos(
            anyo=anyo,
            mes=mes,
            empresa=nombre_empresa,
            barra=nombre_barra,
            nombre_medidor=nombre_medidor,
            plantilla=ruta_plantilla,
            destino=ruta_destino,
        )

        # Iniciar proceso en hilo separado
        self.procesando = True
        self.create_btn.config(state=tk.DISABLED, text="  Procesando...  ")
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
                nombre_medidor,
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
        nombre_medidor: str,
        tipos_seleccionados: list,
    ) -> None:
        """Procesar el informe en un hilo separado para un mes específico."""
        try:
            ruta_destino_path = Path(ruta_destino)
            ruta_destino_path.parent.mkdir(parents=True, exist_ok=True)

            # Procesar el mes (incluye validar descargas antes de copiar/ procesar)
            exito = self.procesar_mes(
                anyo,
                mes,
                ruta_plantilla,
                ruta_destino,
                nombre_barra,
                nombre_empresa,
                nombre_medidor,
                tipos_seleccionados,
                1,
                1,
            )

            if not exito:
                return  # Error 403 u otro: no mostrar mensaje de éxito ni hacer más nada

            # Mensaje final
            self.root.after(0, lambda: self.progress_var.set(100))
            self.root.after(
                0,
                lambda: self.progress_text_label.config(
                    text="[OK] Proceso completado"
                ),
            )

            nombre_mes = meses[mes]
            mensaje = "Informe generado exitosamente.\n\n"
            mensaje += f"Período: {nombre_mes} {anyo}\n"
            if nombre_empresa:
                mensaje += f"Empresa: {nombre_empresa}\n"
            if nombre_barra:
                mensaje += f"Barra: {nombre_barra}\n"
            if nombre_medidor:
                mensaje += f"Nombre Medidor: {nombre_medidor}\n"
            if not nombre_barra and not nombre_empresa and not nombre_medidor:
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
                lambda: self.create_btn.config(
                    state=tk.NORMAL,
                    text="  Crear Informe  ",
                    bg=COLORS["accent"],
                ),
            )
            self.root.after(0, lambda: self.progress_var.set(0))

    def procesar_mes(
        self,
        anyo: int,
        mes: int,
        ruta_plantilla: str,
        ruta_destino: str,
        nombre_barra: str,
        nombre_empresa: str,
        nombre_medidor: str,
        tipos_seleccionados: list,
        mes_actual: int,
        total_meses: int,
    ) -> bool:
        """Procesar un mes individual del rango. Retorna True si tuvo éxito, False si hubo error (ej. 403)."""
        try:
            nombre_mes = meses[mes]
            print(f"\n[INFO] Procesando mes {mes_actual}/{total_meses}: {nombre_mes} {anyo}")

            # Calcular progreso base para este mes (distribuir 100% entre todos los meses)
            progreso_base = int((mes_actual - 1) * 100 / total_meses)
            progreso_por_mes = int(100 / total_meses)

            def calcular_progreso(porcentaje_mes: int) -> int:
                """Calcular el progreso global basado en el progreso de este mes."""
                return progreso_base + int(porcentaje_mes * progreso_por_mes / 100)

            # Base de datos interna: ruta absoluta en AppData (el ejecutable la encuentra siempre)
            carpeta_bd = (_directorio_base_datos() / "bd_data").resolve()
            carpeta_descomprimidos = carpeta_bd / "descomprimidos"

            # Crear carpetas si no existen (necesario para que las descargas se guarden correctamente)
            carpeta_bd.mkdir(parents=True, exist_ok=True)
            carpeta_descomprimidos.mkdir(parents=True, exist_ok=True)

            if mes_actual == 1:
                self.root.after(0, lambda: self.progress_var.set(calcular_progreso(2)))
                self.root.after(
                    0,
                    lambda: self.progress_text_label.config(
                        text="[OK] Base de datos interna lista"
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

                archivo_existente = buscar_archivo_existente_tipo(
                    anyo, mes, tipo, carpeta_zip=str(carpeta_bd)
                )
                if archivo_existente:
                    self.root.after(
                        0,
                        lambda ma=mes_actual, tm=total_meses, ae=archivo_existente: self.progress_text_label.config(
                            text=f"[{ma}/{tm}] [OK] {ae.name} (ya existe)"
                        ),
                    )

                ruta_zip, ruta_descomprimida, codigo_error = descargar_y_descomprimir_zip_tipo(
                    anyo, mes, tipo,
                    carpeta_zip=str(carpeta_bd),
                    descomprimir=True, mostrar_progreso=False
                )

                if tipo == "energia_resultados":
                    ruta_zip_energia = ruta_zip
                    codigo_error_energia = codigo_error

                if ruta_zip:
                    self.root.after(
                        0,
                        lambda ma=mes_actual, tm=total_meses, rz=ruta_zip, d=desc: self.progress_text_label.config(
                            text=f"[{ma}/{tm}] [OK] {d}: {Path(rz).name}"
                        ),
                    )

            # Si energia_resultados falló, no podemos generar el informe: no copiar ni procesar
            if not ruta_zip_energia:
                ma, tm, nm, a = mes_actual, total_meses, nombre_mes, anyo
                if codigo_error_energia == 403:
                    self.root.after(
                        0,
                        lambda: self.progress_text_label.config(
                            text=f"[{ma}/{tm}] ✗ Error 403: Contenido no disponible para {nm} {a}"
                        ),
                    )
                    self.root.after(
                        50,
                        lambda n=nombre_mes, an=anyo: messagebox.showerror(
                            "Información no disponible",
                            f"No se encuentra la información disponible para la descarga para {n} {an}.",
                            parent=self.root,
                        ),
                    )
                    print(f"[WARNING] Error 403 para {nombre_mes} {anyo}: no se realiza ningún procesamiento.")
                    return False
                else:
                    self.root.after(
                        0,
                        lambda: self.progress_text_label.config(
                            text=f"[{ma}/{tm}] ✗ Error al descargar para {nm} {a}"
                        ),
                    )
                    self.root.after(
                        50,
                        lambda n=nombre_mes, an=anyo: messagebox.showerror(
                            "Información no disponible",
                            f"No se encuentra la información disponible para la descarga para {n} {an}.",
                            parent=self.root,
                        ),
                    )
                    print(
                        f"[WARNING] Error al descargar para {nombre_mes} {anyo}: no se realiza ningún procesamiento."
                    )
                    return False

            # Copiar plantilla solo después de confirmar que las descargas fueron exitosas
            from shutil import copyfile

            copyfile(ruta_plantilla, ruta_destino)
            print(f"[INFO] Plantilla copiada a destino: {ruta_destino}")

            # Paso 3: Buscar archivo Balance
            self.root.after(0, lambda: self.progress_var.set(calcular_progreso(35)))
            self.root.after(
                0,
                lambda: self.progress_text_label.config(
                    text=f"[{mes_actual}/{total_meses}] Buscando archivo Balance..."
                ),
            )

            lector = LectorBalance(anyo, mes, carpeta_base=str(carpeta_bd))

            self.root.after(0, lambda: self.progress_var.set(calcular_progreso(40)))
            self.root.after(
                0,
                lambda: self.progress_text_label.config(
                    text=(
                        f"[{mes_actual}/{total_meses}] [OK] Balance encontrado: "
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
                        f"[{mes_actual}/{total_meses}] [OK] Datos leídos: "
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
            columna_nombre_medidor = None

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
                elif "nombre_medidor" in col_lower or "nombre medidor" in str(col).lower():
                    columna_nombre_medidor = col

            # monetario: necesario para TOTAL INGRESOS POR ENERGIA CLP y fallback de POTENCIA FIRME
            if columna_monetario is None:
                print("[WARNING] No se encontró la columna 'monetario' en Balance Valorizado")

            if nombre_barra or nombre_empresa:
                df_guardar = df_balance.copy()

                if nombre_empresa:
                    if columna_empresa is None:
                        print("[ERROR] No se encontró la columna 'nombre_corto_empresa'")
                        return False
                    df_guardar = df_guardar[
                        df_guardar[columna_empresa].astype(str).str.lower()
                        == nombre_empresa.lower()
                    ]

                if nombre_barra:
                    if columna_barra is None:
                        print("[ERROR] No se encontró la columna 'barra'")
                        return False
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

            # Para IMPORTACION MWh y TOTAL INGRESOS POR ENERGIA CLP: aplicar filtro nombre_medidor
            df_para_energia_importacion = df_guardar.copy()
            if nombre_medidor:
                if columna_nombre_medidor is None:
                    print("[WARNING] Nombre Medidor especificado pero no se encontró columna 'nombre_medidor' en Balance Valorizado")
                else:
                    df_para_energia_importacion = df_para_energia_importacion[
                        df_para_energia_importacion[columna_nombre_medidor]
                        .astype(str).str.strip().str.upper()
                        == nombre_medidor.strip().upper()
                    ]
                    self.root.after(
                        0,
                        lambda: self.progress_text_label.config(
                            text=(
                                f"[{mes_actual}/{total_meses}] Filtrando por nombre_medidor: "
                                f"{nombre_medidor}..."
                            )
                        ),
                    )

            self.root.after(0, lambda: self.progress_var.set(calcular_progreso(75)))

            # Acumular datos encontrados para print resumen al final
            datos_encontrados = {}

            # TOTAL INGRESOS POR POTENCIA FIRME CLP: leer desde Anexo 02.b Potencia
            # Tabla Datos: Empresa (B) | Potencia SEN (C) | TOTAL (D)
            total_monetario = leer_total_ingresos_potencia_firme(
                anyo, mes, nombre_empresa=nombre_empresa, carpeta_base=str(carpeta_bd)
            )
            if total_monetario is None:
                # Fallback: usar suma monetario del Balance Valorizado
                if columna_monetario is None:
                    print("[ERROR] No se encontró Anexo Potencia ni columna 'monetario' en Balance")
                    return False
                total_monetario = (
                    df_guardar[columna_monetario].dropna().astype(float).sum()
                )
                print(
                    f"[INFO] Anexo Potencia no encontrado. Usando Balance Valorizado: "
                    f"{total_monetario:,.2f}"
                )
            else:
                print(f"[INFO] TOTAL INGRESOS POR POTENCIA FIRME CLP: {total_monetario:,.2f}")

            datos_encontrados["TOTAL INGRESOS POR POTENCIA FIRME CLP"] = total_monetario

            # INGRESOS POR IT POTENCIA: Anexo 02.b Potencia, hoja 02.IT POTENCIA {Mes}-{YY} def
            total_it = leer_ingresos_por_it(
                anyo, mes, nombre_empresa=nombre_empresa, carpeta_base=str(carpeta_bd)
            )
            datos_encontrados["INGRESOS POR IT POTENCIA"] = total_it

            # INGRESOS POR POTENCIA: Anexo 02.b Potencia, hoja 01.BALANCE POTENCIA {Mes}-{YY} def
            total_potencia = leer_ingresos_por_potencia(
                anyo, mes, nombre_empresa=nombre_empresa, carpeta_base=str(carpeta_bd)
            )
            datos_encontrados["INGRESOS POR POTENCIA"] = total_potencia

            # TOTAL INGRESOS POR ENERGIA CLP: Balance Valorizado, columna monetario
            # Usa filtro nombre_medidor si aplica
            if columna_monetario is not None:
                total_energia = (
                    df_para_energia_importacion[columna_monetario].dropna().astype(float).sum()
                )
                datos_encontrados["TOTAL INGRESOS POR ENERGIA CLP"] = total_energia
                print(
                    f"[INFO] TOTAL INGRESOS POR ENERGIA CLP para {nombre_mes} {anyo}: "
                    f"{total_energia:,.2f} (Balance Valorizado, col monetario)"
                )
            else:
                datos_encontrados["TOTAL INGRESOS POR ENERGIA CLP"] = None
                total_energia = None

            # TOTAL INGRESOS POR SSCC CLP: EXCEL 1_CUADROS_PAGO_SSCC, hoja CPI_
            # Filtra por Nemotecnico Deudor = empresa, suma columna Monto
            total_sscc = (
                leer_total_ingresos_sscc(anyo, mes, nombre_empresa, carpeta_base=str(carpeta_bd))
                if nombre_empresa else None
            )
            datos_encontrados["TOTAL INGRESOS POR SSCC CLP"] = total_sscc

            # Compra Venta Energia GM Holdings CLP: Balance, hoja Contratos, columna VENTA[CLP]
            total_gm_holdings = leer_compra_venta_energia_gm_holdings(
                anyo, mes,
                nombre_empresa=nombre_empresa,
                nombre_barra=nombre_barra,
                carpeta_base=str(carpeta_bd),
            )
            datos_encontrados["Compra Venta Energia GM Holdings CLP"] = total_gm_holdings

            # Calcular IMPORTACION MWh desde columna fisico_kwh (valor positivo, kWh -> MWh: /1000)
            # Usa filtro nombre_medidor si aplica
            importacion_mwh = None
            if columna_fisico_kwh is not None:
                total_fisico_kwh = (
                    df_para_energia_importacion[columna_fisico_kwh].dropna().astype(float).sum()
                )
                importacion_mwh = abs(total_fisico_kwh) / 1000.0
                print(
                    f"[INFO] IMPORTACION MWh para {nombre_mes} {anyo}: "
                    f"{importacion_mwh:,.2f} (desde {total_fisico_kwh:,.0f} kWh)"
                )
                print(f"  -> Dato obtenido (IMPORTACION MWh): {importacion_mwh:,.2f}")
            else:
                print("[WARNING] No se encontró la columna 'fisico_kwh' en Balance Valorizado")
            datos_encontrados["IMPORTACION MWh"] = importacion_mwh

            # Escribir todos los conceptos en una sola sesión Excel (evita errores COM en el ejecutable)
            pares_escribir = [
                ("TOTAL INGRESOS POR POTENCIA FIRME CLP", total_monetario),
            ]
            if total_it is not None:
                pares_escribir.append(("INGRESOS POR IT POTENCIA", total_it))
            if total_potencia is not None:
                pares_escribir.append(("INGRESOS POR POTENCIA", total_potencia))
            if total_energia is not None:
                pares_escribir.append(("TOTAL INGRESOS POR ENERGIA CLP", total_energia))
            if nombre_empresa and total_sscc is not None:
                pares_escribir.append(("TOTAL INGRESOS POR SSCC CLP", total_sscc))
            if total_gm_holdings is not None:
                pares_escribir.append(("Compra Venta Energia GM Holdings CLP", total_gm_holdings))
            if importacion_mwh is not None:
                pares_escribir.append(("IMPORTACION MWh", importacion_mwh))

            escribir_todos_en_resultado(ruta_destino, anyo, mes, pares_escribir)

            # Print resumen de todos los datos encontrados
            print("\n" + "=" * 60)
            print(f"RESUMEN DATOS ESCRITOS EN PLANTILLA - {nombre_mes} {anyo}")
            print("=" * 60)
            for concepto, valor in datos_encontrados.items():
                if valor is not None:
                    if isinstance(valor, float):
                        print(f"  {concepto}: {valor:,.2f}")
                    else:
                        print(f"  {concepto}: {valor}")
                else:
                    print(f"  {concepto}: (no encontrado)")
            print("=" * 60 + "\n")

            exito = True

            self.root.after(0, lambda: self.progress_var.set(calcular_progreso(85)))
            self.root.after(
                0,
                lambda: self.progress_text_label.config(
                    text=(
                        f"[{mes_actual}/{total_meses}] [OK] Guardado en hoja Resultado "
                        f"para {nombre_mes} {anyo}"
                    )
                ),
            )

            if exito:
                print(
                    f"[OK] Mes {mes_actual}/{total_meses} ({nombre_mes} {anyo}) "
                    "procesado exitosamente"
                )
                return True
            else:
                print(f"[ERROR] Error al guardar datos para {nombre_mes} {anyo}")
                return False

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
            return False

        except Exception as e:
            import traceback

            err_txt = str(e).strip() if str(e) else (repr(e) or type(e).__name__)
            err_txt = (err_txt[:80] + "…") if len(err_txt) > 80 else err_txt
            print(f"[ERROR] Error procesando {nombre_mes} {anyo}: {err_txt}")
            traceback.print_exc()

            def _mostrar_error():
                messagebox.showerror(
                    "Error al procesar",
                    f"Error al generar el informe para {nombre_mes} {anyo}:\n\n{err_txt}\n\n"
                    "Verifique que Excel esté instalado y que las rutas de plantilla y destino sean correctas.",
                    parent=self.root,
                )

            self.root.after(
                0,
                lambda ma=mes_actual, tm=total_meses, et=err_txt: self.progress_text_label.config(
                    text=f"[{ma}/{tm}] ✗ Error: {et}"
                ),
            )
            self.root.after(50, lambda: _mostrar_error())
            return False


def main() -> None:
    root = tk.Tk()
    app = InterfazInforme(root)
    root.mainloop()


__all__ = ["InterfazInforme", "main"]

