import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
from pathlib import Path
from v1.descargar_archivos import descargar_zip_si_no_existe, descargar_y_descomprimir_zip, meses, buscar_archivo_existente
import threading

class InterfazDescarga:
    def __init__(self, root):
        self.root = root
        self.root.title("Descarga de Archivos PLABACOM")
        self.root.geometry("500x500")
        self.root.configure(bg="#E5E5E5")
        
        # Variables
        self.anyo_var = tk.IntVar(value=datetime.now().year)
        self.mes_var = tk.IntVar(value=datetime.now().month)
        self.mes_combo = None  # Se asignará en create_widgets
        self.descargando = False
        
        # Crear interfaz
        self.create_widgets()
    
    def create_widgets(self):
        """Crear los widgets de la interfaz"""
        # Frame principal
        main_frame = tk.Frame(self.root, bg="white", relief=tk.FLAT)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Título
        title_label = tk.Label(main_frame, 
                              text="Descargar Archivo PLABACOM",
                              font=("Arial", 16, "bold"),
                              bg="white")
        title_label.pack(pady=(20, 30))
        
        # Frame para selección de año y mes
        selection_frame = tk.Frame(main_frame, bg="white")
        selection_frame.pack(pady=20, padx=30, fill=tk.X)
        
        # Año
        año_label = tk.Label(selection_frame, text="Año:", 
                            font=("Arial", 11), bg="white", anchor="w")
        año_label.grid(row=0, column=0, sticky="w", pady=(0, 10))
        
        año_spinbox = tk.Spinbox(selection_frame, 
                                 from_=2020, 
                                 to=2030, 
                                 textvariable=self.anyo_var,
                                 font=("Arial", 11),
                                 width=10,
                                 command=self.verificar_archivo_existente)
        año_spinbox.grid(row=1, column=0, sticky="w", pady=(0, 20))
        año_spinbox.bind("<KeyRelease>", lambda e: self.verificar_archivo_existente())
        
        # Mes
        mes_label = tk.Label(selection_frame, text="Mes:", 
                            font=("Arial", 11), bg="white", anchor="w")
        mes_label.grid(row=0, column=1, sticky="w", padx=(30, 0), pady=(0, 10))
        
        # Combobox para mes
        meses_lista = [f"{i:02d} - {meses[i]}" for i in range(1, 13)]
        self.mes_combo = ttk.Combobox(selection_frame,
                                     values=meses_lista,
                                     state="readonly",
                                     font=("Arial", 11),
                                     width=15)
        self.mes_combo.grid(row=1, column=1, sticky="w", padx=(30, 0), pady=(0, 20))
        self.mes_combo.current(self.mes_var.get() - 1)  # Establecer mes actual
        self.mes_combo.bind("<<ComboboxSelected>>", lambda e: self.verificar_archivo_existente())
        
        # Frame para información
        info_frame = tk.Frame(main_frame, bg="white")
        info_frame.pack(pady=20, padx=30, fill=tk.X)
        
        self.info_label = tk.Label(info_frame,
                                   text="Seleccione el año y mes a descargar",
                                   font=("Arial", 10),
                                   bg="white",
                                   fg="#666666",
                                   wraplength=400,
                                   justify=tk.LEFT)
        self.info_label.pack(anchor="w")
        
        # Barra de progreso
        progress_frame = tk.Frame(main_frame, bg="white")
        progress_frame.pack(pady=20, padx=30, fill=tk.X)
        
        self.progress_label = tk.Label(progress_frame,
                                      text="",
                                      font=("Arial", 9),
                                      bg="white",
                                      fg="#666666")
        self.progress_label.pack(anchor="w", pady=(0, 5))
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame,
                                           variable=self.progress_var,
                                           maximum=100,
                                           length=400,
                                           mode='determinate')
        self.progress_bar.pack(fill=tk.X)
        
        # Botón de descarga
        button_frame = tk.Frame(main_frame, bg="white")
        button_frame.pack(side=tk.BOTTOM, pady=30, padx=30, fill=tk.X)
        
        self.descargar_btn = tk.Button(button_frame,
                                      text="Descargar Archivo",
                                      font=("Arial", 12, "bold"),
                                      bg="#7B2CBF",
                                      fg="white",
                                      relief=tk.FLAT,
                                      padx=30,
                                      pady=12,
                                      cursor="hand2",
                                      command=self.iniciar_descarga)
        self.descargar_btn.pack(expand=True)
        
        # Efecto hover
        self.descargar_btn.bind("<Enter>", lambda e: self.descargar_btn.config(bg="#6A1B9A"))
        self.descargar_btn.bind("<Leave>", lambda e: self.descargar_btn.config(bg="#7B2CBF"))
        
        # Verificar archivo existente al iniciar
        self.verificar_archivo_existente()
    
    def verificar_archivo_existente(self):
        """Verifica si el archivo ya existe y actualiza la información"""
        try:
            anyo = self.anyo_var.get()
            valor_mes = self.mes_combo.get()
            
            if not valor_mes:
                self.info_label.config(text="Seleccione el año y mes a descargar", fg="#666666")
                return
            
            mes = int(valor_mes.split(" - ")[0])
            nombre_mes = meses[mes]
            
            # Buscar si el archivo ya existe
            archivo_existente = buscar_archivo_existente(anyo, mes)
            
            if archivo_existente:
                tamaño = archivo_existente.stat().st_size / (1024 * 1024)  # Tamaño en MB
                self.info_label.config(
                    text=f"✓ Archivo ya existe: {nombre_mes} {anyo}\nTamaño: {tamaño:.2f} MB\nUbicación: {archivo_existente.name}",
                    fg="#28a745"  # Verde
                )
            else:
                self.info_label.config(
                    text=f"Archivo no encontrado: {nombre_mes} {anyo}\nSe descargará al hacer clic en 'Descargar Archivo'",
                    fg="#666666"
                )
        except Exception:
            # Si hay algún error, simplemente no actualizar
            pass
    
    def iniciar_descarga(self):
        """Iniciar la descarga en un hilo separado"""
        if self.descargando:
            messagebox.showwarning("Descarga en curso", 
                                 "Ya hay una descarga en progreso. Por favor espere.")
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
        
        # Verificar si el archivo ya existe antes de descargar
        archivo_existente = buscar_archivo_existente(anyo, mes)
        nombre_mes = meses[mes]
        
        if archivo_existente:
            tamaño = archivo_existente.stat().st_size / (1024 * 1024)  # Tamaño en MB
            respuesta = messagebox.askyesno(
                "Archivo ya existe",
                f"El archivo para {nombre_mes} {anyo} ya existe:\n\n"
                f"Archivo: {archivo_existente.name}\n"
                f"Tamaño: {tamaño:.2f} MB\n\n"
                f"¿Desea descargarlo de nuevo?"
            )
            if not respuesta:
                return  # El usuario canceló
        
        # Actualizar interfaz
        self.descargando = True
        self.descargar_btn.config(state=tk.DISABLED, text="Descargando...")
        self.progress_var.set(0)
        self.progress_label.config(text="Iniciando descarga...")
        self.info_label.config(text=f"Descargando: {nombre_mes} {anyo}", fg="#666666")
        
        # Guardar si el archivo existía antes (para mostrar mensaje apropiado después)
        archivo_existia_antes = archivo_existente is not None
        
        # Iniciar descarga en hilo separado
        thread = threading.Thread(target=self.descargar_archivo_thread, 
                                 args=(anyo, mes, archivo_existia_antes), daemon=True)
        thread.start()
    
    def descargar_archivo_thread(self, anyo, mes, archivo_existia_antes):
        """Descargar archivo en hilo separado"""
        try:
            # Actualizar progreso inicial
            self.root.after(0, lambda: self.progress_var.set(10))
            self.root.after(0, lambda: self.progress_label.config(text="Verificando si el archivo existe..."))
            
            # Descargar y descomprimir (la función ya verifica si existe)
            ruta_zip, ruta_descomprimida, codigo_error = descargar_y_descomprimir_zip(anyo, mes, descomprimir=True)
            
            if ruta_zip:
                # Actualizar progreso
                self.root.after(0, lambda: self.progress_var.set(100))
                self.root.after(0, lambda: self.progress_label.config(text="✓ Descarga y descompresión completadas"))
                
                # Actualizar información
                tamaño_zip = Path(ruta_zip).stat().st_size / (1024 * 1024)
                nombre_mes = meses[mes]
                
                mensaje_info = f"✓ Archivo disponible: {nombre_mes} {anyo}\n"
                mensaje_info += f"ZIP: {tamaño_zip:.2f} MB - {Path(ruta_zip).name}\n"
                
                if ruta_descomprimida:
                    # Calcular tamaño de la carpeta descomprimida
                    tamaño_descomprimido = sum(f.stat().st_size for f in Path(ruta_descomprimida).rglob('*') if f.is_file()) / (1024 * 1024)
                    mensaje_info += f"Descomprimido: {tamaño_descomprimido:.2f} MB - {Path(ruta_descomprimida).name}"
                else:
                    mensaje_info += "Descompresión: No disponible"
                
                self.root.after(0, lambda: self.info_label.config(
                    text=mensaje_info,
                    fg="#28a745"))
                
                # Mostrar mensaje apropiado
                mensaje_final = ""
                if archivo_existia_antes:
                    mensaje_final = f"El archivo ya estaba disponible:\n\nZIP: {ruta_zip}"
                else:
                    mensaje_final = f"El archivo se ha descargado correctamente:\n\nZIP: {ruta_zip}"
                
                if ruta_descomprimida:
                    mensaje_final += f"\n\nDescomprimido en:\n{ruta_descomprimida}"
                
                if archivo_existia_antes:
                    self.root.after(0, lambda: messagebox.showinfo("Archivo encontrado", mensaje_final))
                else:
                    self.root.after(0, lambda: messagebox.showinfo("Descarga exitosa", mensaje_final))
                
                # Actualizar la verificación para reflejar el estado actual
                self.root.after(0, self.verificar_archivo_existente)
            else:
                # Verificar si fue error 403
                nombre_mes = meses[mes]
                self.root.after(0, lambda: self.progress_var.set(0))
                
                if codigo_error == 403:
                    # Error 403: Contenido no disponible
                    self.root.after(0, lambda: self.progress_label.config(text="✗ Contenido no disponible"))
                    self.root.after(0, lambda: self.info_label.config(
                        text=f"✗ Contenido no disponible: {nombre_mes} {anyo}\nEl archivo no está disponible en el servidor.\nPuede que aún no se haya publicado para este período.",
                        fg="#dc3545"))  # Rojo
                    self.root.after(0, lambda: messagebox.showerror("Contenido no disponible", 
                                       f"El archivo para {nombre_mes} {anyo} no está disponible en el servidor (Error 403).\n\n"
                                       f"Esto puede deberse a:\n"
                                       f"• El archivo aún no se ha publicado para este período\n"
                                       f"• El año/mes seleccionado no existe en la base de datos\n"
                                       f"• El contenido no está disponible para descarga\n\n"
                                       f"Por favor, verifique que el año y mes sean correctos."))
                else:
                    # Otro tipo de error
                    self.root.after(0, lambda: self.progress_label.config(text="✗ Error en la descarga"))
                    self.root.after(0, lambda: self.info_label.config(
                        text=f"✗ Error al descargar: {nombre_mes} {anyo}\nNo se pudo descargar el archivo.\nVerifique su conexión a internet.",
                        fg="#dc3545"))  # Rojo
                    self.root.after(0, lambda: messagebox.showerror("Error en la descarga", 
                                       f"No se pudo descargar el archivo para {nombre_mes} {anyo}.\n\n"
                                       f"Por favor, verifique:\n"
                                       f"• Su conexión a internet\n"
                                       f"• Que el servidor esté disponible\n"
                                       f"• Que el año y mes sean correctos"))
        
        except Exception as e:
            self.root.after(0, lambda: self.progress_var.set(0))
            self.root.after(0, lambda: self.progress_label.config(text=f"✗ Error: {str(e)}"))
            messagebox.showerror("Error", f"Ocurrió un error durante la descarga:\n{str(e)}")
        
        finally:
            # Restaurar botón
            self.descargando = False
            self.root.after(0, lambda: self.descargar_btn.config(state=tk.NORMAL, text="Descargar Archivo"))


def main():
    root = tk.Tk()
    app = InterfazDescarga(root)
    root.mainloop()


if __name__ == "__main__":
    main()
