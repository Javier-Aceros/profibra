import tkinter as tk
from tkinter import messagebox, ttk
import ttkbootstrap as ttk
from pathlib import Path
from modules.logger import Logger
from modules.file_processor import FileProcessor
from modules.consolidator import Consolidator
from modules.comparative_analyzer import ComparativeAnalyzer

class Aplicacion:
    def __init__(self, ventana):
        self.ventana = ventana
        self.configurar_interfaz()
        
        # Inicializar logger después de configurar la interfaz
        self.logger = Logger(self.log_frame)  # Pasamos el frame directamente
        
        # Inicializar los procesadores
        self.file_processor = FileProcessor(self.logger)
        self.consolidator = Consolidator(self.logger)
        self.analyzer = ComparativeAnalyzer(self.logger)
        
        # Configurar progreso
        self.progress = ttk.Progressbar(
            self.ventana, 
            orient="horizontal", 
            length=300, 
            mode="determinate"
        )
        self.progress.pack(pady=5)
        self.progress.pack_forget()  # Ocultar inicialmente
    
    def configurar_interfaz(self):
        self.ventana.title("Sistema de Comparación de Inventarios")
        self.ventana.geometry("800x600")  # Tamaño ajustado
        
        # Frame principal
        main_frame = ttk.Frame(self.ventana)
        main_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # Frame superior para controles
        top_frame = ttk.Frame(main_frame)
        top_frame.pack(fill="x", pady=(0, 10))  # Espacio solo abajo
        
        # Título
        ttk.Label(
            top_frame, 
            text="Comparador de Inventarios Automático", 
            font=("Segoe UI", 14, "bold")
        ).pack(pady=(0, 5))  # Espacio solo abajo

        # Instrucciones
        ttk.Label(
            top_frame, 
            text="La aplicación leerá automáticamente los archivos Excel de la carpeta 'inputs'",
            font=("Segoe UI", 10)
        ).pack(pady=(0, 2))  # Espacio solo abajo

        ttk.Label(
            top_frame,
            text="Archivos requeridos en 'inputs/': STIHL.xlsx, SUZUKI.xlsx, YAMAHA.xlsx, Valoración de inventarios.xlsx",
            font=("Segoe UI", 9)
        ).pack(pady=(0, 10))  # Espacio solo abajo

        # Botón de inicio
        ttk.Button(
            top_frame, 
            text="Iniciar Proceso", 
            command=self.iniciar_proceso,
            style="primary.TButton"
        ).pack(pady=(0, 10))  # Espacio solo abajo
        
        # Frame para el log (ocupará el espacio restante)
        self.log_frame = ttk.Frame(main_frame)
        self.log_frame.pack(fill="both", expand=True)
    
    def actualizar_progreso(self, valor, mensaje):
        """Actualiza la barra de progreso y muestra un mensaje"""
        self.progress["value"] = valor
        self.ventana.update_idletasks()
        self.logger.agregar_log(mensaje)
    
    def iniciar_proceso(self):
        try:
            self.progress.pack(pady=5)
            self.actualizar_progreso(10, "Iniciando proceso de consolidación...")
            
            # Paso 1: Crear consolidado
            self.actualizar_progreso(30, "Procesando archivos de inventario...")
            consolidado = self.consolidator.crear_consolidado()
            
            self.actualizar_progreso(60, f"Consolidado creado: {consolidado}")
            
            # Paso 2: Crear análisis comparativo
            self.actualizar_progreso(70, "Generando análisis comparativo...")
            resultado = self.analyzer.procesar_consolidado(consolidado)
            
            self.actualizar_progreso(100, "Proceso completado con éxito")
            
            messagebox.showinfo(
                "Éxito", 
                f"Proceso completado:\n"
                f"1. Consolidado: {Path(consolidado).name}\n"
                f"2. Análisis: {Path(resultado).name}\n\n"
                f"Los archivos se encuentran en la carpeta 'outputs'"
            )
            
            self.progress["value"] = 0
            self.progress.pack_forget()
                
        except Exception as e:
            self.logger.agregar_log(f"Error: {str(e)}", "error")
            messagebox.showerror(
                "Error", 
                f"Ocurrió un problema:\n{str(e)}\n\n"
                "Verifique que los archivos requeridos estén en la carpeta 'inputs' "
                "y que tengan el formato correcto."
            )
            self.progress["value"] = 0
            self.progress.pack_forget()

if __name__ == "__main__":
    ventana = ttk.Window(themename="flatly")
    app = Aplicacion(ventana)
    ventana.mainloop()
