import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import messagebox, ttk

from core.descargar_archivos import (
    buscar_archivo_existente_tipo,
    descargar_y_descomprimir_zip_tipo,
    meses,
    TIPOS_ARCHIVO,
)


class InterfazDescarga:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Descarga de Archivos PLABACOM")
        self.root.geometry("520x580")
        self.root.configure(bg="#E5E5E5")

        # Variables
        self.anyo_var = tk.IntVar(value=datetime.now().year)
        self.mes_var = tk.IntVar(value=datetime.now().month)
        self.mes_combo = None  # Se asignará en create_widgets
        self.descargando = False

        # Variables para tipos de archivo (por defecto solo Resultados)
        self.tipo_vars = {
            "energia_resultados": tk.BooleanVar(value=True),
            "energia_antecedentes": tk.BooleanVar(value=False),
            "sscc": tk.BooleanVar(value=False),
            "potencia": tk.BooleanVar(value=False),
        }

        # Crear interfaz
        self.create_widgets()

    def create_widgets(self) -> None:
        """Crear los widgets de la interfaz."""
        # Frame principal
        main_frame = tk.Frame(self.root, bg="white", relief=tk.FLAT)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Título
        title_label = tk.Label(
            main_frame,
            text="Descargar Archivo PLABACOM",
            font=("Arial", 16, "bold"),
            bg="white",
        )
        title_label.pack(pady=(20, 30))

        # Frame para selección de año y mes
        selection_frame = tk.Frame(main_frame, bg="white")
        selection_frame.pack(pady=20, padx=30, fill=tk.X)

        # Año
        año_label = tk.Label(
            selection_frame,
            text="Año:",
            font=("Arial", 11),
            bg="white",
            anchor="w",
        )
        año_label.grid(row=0, column=0, sticky="w", pady=(0, 10))

        año_spinbox = tk.Spinbox(
            selection_frame,
            from_=2020,
            to=2030,
            textvariable=self.anyo_var,
            font=("Arial", 11),
            width=10,
            command=self.verificar_archivo_existente,
        )
        año_spinbox.grid(row=1, column=0, sticky="w", pady=(0, 20))
        año_spinbox.bind("<KeyRelease>", lambda e: self.verificar_archivo_existente())

        # Mes
        mes_label = tk.Label(
            selection_frame,
            text="Mes:",
            font=("Arial", 11),
            bg="white",
            anchor="w",
        )
        mes_label.grid(row=0, column=1, sticky="w", padx=(30, 0), pady=(0, 10))

        # Combobox para mes
        meses_lista = [f"{i:02d} - {meses[i]}" for i in range(1, 13)]
        self.mes_combo = ttk.Combobox(
            selection_frame,
            values=meses_lista,
            state="readonly",
            font=("Arial", 11),
            width=15,
        )
        self.mes_combo.grid(row=1, column=1, sticky="w", padx=(30, 0), pady=(0, 20))
        self.mes_combo.current(self.mes_var.get() - 1)  # Establecer mes actual
        self.mes_combo.bind(
            "<<ComboboxSelected>>", lambda e: self.verificar_archivo_existente()
        )

        # Frame para tipos de archivo
        tipos_frame = tk.Frame(main_frame, bg="white")
        tipos_frame.pack(pady=(10, 15), padx=30, fill=tk.X)

        tipos_label = tk.Label(
            tipos_frame,
            text="Tipos de archivo a descargar:",
            font=("Arial", 11),
            bg="white",
            anchor="w",
        )
        tipos_label.grid(row=0, column=0, sticky="w", pady=(0, 8))

        self.tipo_checkboxes = {}
        for i, (tipo_key, descripcion) in enumerate(TIPOS_ARCHIVO.items()):
            cb = tk.Checkbutton(
                tipos_frame,
                text=descripcion,
                variable=self.tipo_vars[tipo_key],
                font=("Arial", 10),
                bg="white",
                activebackground="white",
                anchor="w",
                command=self.verificar_archivo_existente,
            )
            row, col = (i // 2) + 1, (i % 2)
            cb.grid(row=row, column=col, sticky="w", padx=(0, 25), pady=2)
            self.tipo_checkboxes[tipo_key] = cb

        # Frame para información
        info_frame = tk.Frame(main_frame, bg="white")
        info_frame.pack(pady=20, padx=30, fill=tk.X)

        self.info_label = tk.Label(
            info_frame,
            text="Seleccione el año y mes a descargar",
            font=("Arial", 10),
            bg="white",
            fg="#666666",
            wraplength=400,
            justify=tk.LEFT,
        )
        self.info_label.pack(anchor="w")

        # Barra de progreso
        progress_frame = tk.Frame(main_frame, bg="white")
        progress_frame.pack(pady=20, padx=30, fill=tk.X)

        self.progress_label = tk.Label(
            progress_frame,
            text="",
            font=("Arial", 9),
            bg="white",
            fg="#666666",
        )
        self.progress_label.pack(anchor="w", pady=(0, 5))

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            variable=self.progress_var,
            maximum=100,
            length=400,
            mode="determinate",
        )
        self.progress_bar.pack(fill=tk.X)

        # Botón de descarga
        button_frame = tk.Frame(main_frame, bg="white")
        button_frame.pack(side=tk.BOTTOM, pady=30, padx=30, fill=tk.X)

        self.descargar_btn = tk.Button(
            button_frame,
            text="Descargar Archivo",
            font=("Arial", 12, "bold"),
            bg="#7B2CBF",
            fg="white",
            relief=tk.FLAT,
            padx=30,
            pady=12,
            cursor="hand2",
            command=self.iniciar_descarga,
        )
        self.descargar_btn.pack(expand=True)

        # Efecto hover
        self.descargar_btn.bind(
            "<Enter>", lambda e: self.descargar_btn.config(bg="#6A1B9A")
        )
        self.descargar_btn.bind(
            "<Leave>", lambda e: self.descargar_btn.config(bg="#7B2CBF")
        )

        # Verificar archivo existente al iniciar
        self.verificar_archivo_existente()

    def _obtener_tipos_seleccionados(self):
        """Retorna la lista de tipos de archivo que el usuario tiene marcados."""
        return [k for k, v in self.tipo_vars.items() if v.get()]

    def verificar_archivo_existente(self) -> None:
        """Verifica si los archivos ya existen y actualiza la información."""
        try:
            anyo = self.anyo_var.get()
            valor_mes = self.mes_combo.get()

            if not valor_mes:
                self.info_label.config(
                    text="Seleccione el año y mes a descargar", fg="#666666"
                )
                return

            mes = int(valor_mes.split(" - ")[0])
            nombre_mes = meses[mes]
            tipos_seleccionados = self._obtener_tipos_seleccionados()

            if not tipos_seleccionados:
                self.info_label.config(
                    text=(
                        f"Estado para {nombre_mes} {anyo}\n"
                        "Marque al menos un tipo de archivo a descargar"
                    ),
                    fg="#666666",
                )
                return

            # Verificar estado por cada tipo seleccionado
            lineas = [f"Estado para {nombre_mes} {anyo}:"]
            todos_existen = True
            for tipo in tipos_seleccionados:
                archivo = buscar_archivo_existente_tipo(anyo, mes, tipo)
                desc = TIPOS_ARCHIVO.get(tipo, tipo)
                if archivo:
                    tamaño = archivo.stat().st_size / (1024 * 1024)
                    lineas.append(f"  ✓ {desc}: {tamaño:.2f} MB")
                else:
                    lineas.append(f"  ✗ {desc}: no encontrado")
                    todos_existen = False

            self.info_label.config(
                text="\n".join(lineas),
                fg="#28a745" if todos_existen else "#666666",
            )
        except Exception:
            pass

    def iniciar_descarga(self) -> None:
        """Iniciar la descarga en un hilo separado."""
        if self.descargando:
            messagebox.showwarning(
                "Descarga en curso",
                "Ya hay una descarga en progreso. Por favor espere.",
            )
            return

        anyo = self.anyo_var.get()

        # Obtener el mes del combobox
        valor_mes = self.mes_combo.get()
        if not valor_mes:
            messagebox.showerror("Error", "Por favor seleccione un mes.")
            return

        try:
            mes = int(valor_mes.split(" - ")[0])
        except (ValueError, IndexError):
            messagebox.showerror("Error", "Error al obtener el mes seleccionado.")
            return

        # Validar
        if mes < 1 or mes > 12:
            messagebox.showerror("Error", "Por favor seleccione un mes válido.")
            return

        tipos_seleccionados = self._obtener_tipos_seleccionados()
        if not tipos_seleccionados:
            messagebox.showerror(
                "Error",
                "Seleccione al menos un tipo de archivo a descargar.",
            )
            return

        nombre_mes = meses[mes]

        # Actualizar interfaz
        self.descargando = True
        self.descargar_btn.config(state=tk.DISABLED, text="Descargando...")
        self.progress_var.set(0)
        self.progress_label.config(text="Iniciando descarga...")
        self.info_label.config(text=f"Descargando: {nombre_mes} {anyo}", fg="#666666")

        # Iniciar descarga en hilo separado
        thread = threading.Thread(
            target=self.descargar_archivo_thread,
            args=(anyo, mes, tipos_seleccionados),
            daemon=True,
        )
        thread.start()

    def descargar_archivo_thread(
        self,
        anyo: int,
        mes: int,
        tipos_seleccionados: list,
    ) -> None:
        """Descargar archivos en hilo separado para cada tipo seleccionado."""
        try:
            nombre_mes = meses[mes]
            total_tipos = len(tipos_seleccionados)
            resultados = []  # (tipo, ruta_zip, ruta_des, codigo_error)

            for idx, tipo in enumerate(tipos_seleccionados):
                progreso = int(10 + (idx / total_tipos) * 85)
                desc = TIPOS_ARCHIVO.get(tipo, tipo)
                self.root.after(0, lambda p=progreso: self.progress_var.set(p))
                self.root.after(
                    0,
                    lambda d=desc: self.progress_label.config(
                        text=f"Descargando {d}..."
                    ),
                )

                ruta_zip, ruta_des, codigo_error = descargar_y_descomprimir_zip_tipo(
                    anyo, mes, tipo, descomprimir=True, mostrar_progreso=False
                )
                resultados.append((tipo, ruta_zip, ruta_des, codigo_error))

            # Resumen de resultados
            exitosos = [(t, z, d, e) for t, z, d, e in resultados if z]
            fallidos = [(t, e) for t, z, d, e in resultados if not z]

            if exitosos:
                # Actualizar progreso
                self.root.after(0, lambda: self.progress_var.set(100))
                self.root.after(
                    0,
                    lambda: self.progress_label.config(
                        text="✓ Descarga y descompresión completadas"
                    ),
                )

                # Construir mensaje con todos los archivos descargados
                lineas_info = [f"✓ Archivos disponibles: {nombre_mes} {anyo}"]
                lineas_final = []
                for t, ruta_zip, ruta_des, _ in exitosos:
                    desc = TIPOS_ARCHIVO.get(t, t)
                    tamaño_zip = Path(ruta_zip).stat().st_size / (1024 * 1024)
                    lineas_info.append(f"  ✓ {desc}: {tamaño_zip:.2f} MB")
                    lineas_final.append(f"{desc}: {Path(ruta_zip).name}")
                    if ruta_des:
                        lineas_final.append(f"  Descomprimido: {Path(ruta_des).name}")

                self.root.after(
                    0,
                    lambda li=lineas_info: self.info_label.config(
                        text="\n".join(li),
                        fg="#28a745",
                    ),
                )

                titulo = "Descarga exitosa" if len(exitosos) == total_tipos else "Descarga parcial"
                mensaje_final = "Archivos descargados correctamente:\n\n" + "\n".join(lineas_final)
                if fallidos:
                    mensaje_final += "\n\nNo se pudieron descargar:\n"
                    for t, e in fallidos:
                        desc = TIPOS_ARCHIVO.get(t, t)
                        mensaje_final += f"  • {desc}"
                        if e == 403:
                            mensaje_final += " (no disponible en servidor)"
                        mensaje_final += "\n"

                self.root.after(
                    0,
                    lambda: messagebox.showinfo(titulo, mensaje_final),
                )

                # Actualizar la verificación para reflejar el estado actual
                self.root.after(0, self.verificar_archivo_existente)
            else:
                # Todos fallaron
                self.root.after(0, lambda: self.progress_var.set(0))
                hay_403 = any(e == 403 for _, e in fallidos)

                if hay_403:
                    # Error 403: Contenido no disponible
                    self.root.after(
                        0,
                        lambda: self.progress_label.config(
                            text="✗ Contenido no disponible"
                        ),
                    )
                    self.root.after(
                        0,
                        lambda: self.info_label.config(
                            text=(
                                f"✗ Contenido no disponible: {nombre_mes} {anyo}\n"
                                "El archivo no está disponible en el servidor.\n"
                                "Puede que aún no se haya publicado para este período."
                            ),
                            fg="#dc3545",
                        ),
                    )
                    self.root.after(
                        0,
                        lambda: messagebox.showerror(
                            "Contenido no disponible",
                            (
                                f"El archivo para {nombre_mes} {anyo} no está "
                                "disponible en el servidor (Error 403).\n\n"
                                "Esto puede deberse a:\n"
                                "• El archivo aún no se ha publicado para este período\n"
                                "• El año/mes seleccionado no existe en la base de datos\n"
                                "• El contenido no está disponible para descarga\n\n"
                                "Por favor, verifique que el año y mes sean correctos."
                            ),
                        ),
                    )
                else:
                    # Otro tipo de error
                    self.root.after(
                        0,
                        lambda: self.progress_label.config(
                            text="✗ Error en la descarga"
                        ),
                    )
                    self.root.after(
                        0,
                        lambda: self.info_label.config(
                            text=(
                                f"✗ Error al descargar: {nombre_mes} {anyo}\n"
                                "No se pudo descargar el archivo.\n"
                                "Verifique su conexión a internet."
                            ),
                            fg="#dc3545",
                        ),
                    )
                    self.root.after(
                        0,
                        lambda: messagebox.showerror(
                            "Error en la descarga",
                            (
                                f"No se pudo descargar el archivo para {nombre_mes} {anyo}.\n\n"
                                "Por favor, verifique:\n"
                                "• Su conexión a internet\n"
                                "• Que el servidor esté disponible\n"
                                "• Que el año y mes sean correctos"
                            ),
                        ),
                    )

        except Exception as e:
            self.root.after(0, lambda: self.progress_var.set(0))
            self.root.after(
                0,
                lambda: self.progress_label.config(text=f"✗ Error: {str(e)}"),
            )
            messagebox.showerror(
                "Error", f"Ocurrió un error durante la descarga:\n{str(e)}"
            )

        finally:
            # Restaurar botón
            self.descargando = False
            self.root.after(
                0,
                lambda: self.descargar_btn.config(
                    state=tk.NORMAL,
                    text="Descargar Archivo",
                ),
            )


def main() -> None:
    root = tk.Tk()
    app = InterfazDescarga(root)
    root.mainloop()


__all__ = ["InterfazDescarga", "main"]

