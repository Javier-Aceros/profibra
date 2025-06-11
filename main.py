import tkinter as tk
from tkinter import messagebox, ttk
import ttkbootstrap as ttk
from modules.logger import Logger
from modules.file_processor import FileProcessor

class Aplicacion:
    def __init__(self, ventana):
        self.ventana = ventana
        self.configurar_interfaz()
        self.logger = Logger(ventana)
        self.procesador = FileProcessor(self.logger)
    
    def configurar_interfaz(self):
        self.ventana.title("Sistema de Comparación de Inventarios")
        self.ventana.geometry("800x600")
        
        # Título
        ttk.Label(self.ventana, text="Comparador de Inventarios Automático", 
                 font=("Segoe UI", 14, "bold")).pack(pady=10)

        # Instrucciones
        ttk.Label(self.ventana, text="La aplicación leerá automáticamente los archivos Excel de la carpeta 'inputs'",
                 font=("Segoe UI", 10)).pack(pady=5)

        ttk.Label(self.ventana, 
                 text="Archivos requeridos en 'inputs/': STIHL.xlsx, SUZUKI.xlsx, YAMAHA.xlsx, Valoración de inventarios.xlsx",
                 font=("Segoe UI", 9)).pack(pady=5)

        # Botón de inicio
        ttk.Button(self.ventana, text="Iniciar Proceso", 
                  command=self.iniciar_proceso).pack(pady=20)
    
    def iniciar_proceso(self):
        try:
            # Paso 1: Crear consolidado
            consolidado = self.procesador.crear_consolidado()
            
            # Paso 2: Crear análisis comparativo
            resultado = self.procesador.procesar_consolidado(consolidado)
            
            messagebox.showinfo("Éxito", 
                f"Proceso completado:\n"
                f"1. Consolidado: {consolidado}\n"
                f"2. Análisis: {resultado}")
                
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un problema:\n{str(e)}")

if __name__ == "__main__":
    ventana = ttk.Window(themename="flatly")
    app = Aplicacion(ventana)
    ventana.mainloop()