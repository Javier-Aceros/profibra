import tkinter as tk
from tkinter import ttk

class Logger:
    def __init__(self, ventana):
        self.ventana = ventana
        self.estado = ttk.Label(ventana, text="", font=("Segoe UI", 10))
        self.estado.pack(pady=10, padx=20, fill=tk.X)
        
        # Área de registro con scrollbar
        self.frame_registro = ttk.Frame(ventana)
        self.frame_registro.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)
        
        self.scrollbar = ttk.Scrollbar(self.frame_registro)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.registro = tk.Text(self.frame_registro, yscrollcommand=self.scrollbar.set,
                              height=10, wrap=tk.WORD, font=("Consolas", 9))
        self.registro.pack(fill=tk.BOTH, expand=True)
        self.scrollbar.config(command=self.registro.yview)
        
        # Configurar colores para diferentes tipos de mensajes
        self.registro.tag_config('info', foreground='blue')
        self.registro.tag_config('exito', foreground='green')
        self.registro.tag_config('error', foreground='red')
        self.registro.tag_config('advertencia', foreground='orange')

    def agregar_log(self, mensaje, tipo='info'):
        self.registro.insert(tk.END, f"{mensaje}\n", tipo)
        self.registro.see(tk.END)  # Auto-scroll al final
        self.ventana.update()
