import pandas as pd
from pathlib import Path
from openpyxl.styles import PatternFill
from openpyxl import load_workbook

class ComparativeAnalyzer:
    def __init__(self, logger):
        self.logger = logger

    def procesar_consolidado(self, ruta_consolidado):
        try:
            self.logger.agregar_log(f"\nProcesando archivo consolidado: {ruta_consolidado}...")

            # Leer datos consolidados
            df = pd.read_excel(ruta_consolidado, sheet_name="Consolidado")
            
            # Limpieza de datos
            df.columns = df.columns.str.strip()
            df["REFERENCIA"] = df["REFERENCIA"].astype(str).str.strip().replace("nan", "")
            
            # Separar referencias válidas vs inválidas
            con_ref = df[df["REFERENCIA"] != ""].copy()
            sin_ref = df[df["REFERENCIA"] == ""].copy()

            # Agrupar por referencia para consolidar orígenes
            grouped = con_ref.groupby("REFERENCIA").agg({
                'DESCRIPCION': 'first',
                'CANTIDAD': 'sum',
                'UBICACION': lambda x: ', '.join(str(v) for v in set(x) if v and str(v).strip() != ''),
                'CODIGO_SIIGO': 'first',
                'ORIGEN': lambda x: ', '.join(sorted(set(x)))
            }).reset_index()

            # Para referencias sin código, mantener los datos originales
            if not sin_ref.empty:
                sin_ref = sin_ref.groupby(['DESCRIPCION', 'ORIGEN']).agg({
                    'CANTIDAD': 'sum',
                    'UBICACION': lambda x: ', '.join(str(v) for v in set(x) if v and str(v).strip() != ''),
                    'CODIGO_SIIGO': 'first'
                }).reset_index()

            # Crear estructura final
            columnas_finales = [
                "REFERENCIA", "DESCRIPCION", "CANTIDAD", 
                "ORIGEN", "UBICACION", "CODIGO_SIIGO"
            ]

            tabla_final = pd.concat([
                grouped[columnas_finales],
                sin_ref[['DESCRIPCION', 'CANTIDAD', 'ORIGEN', 'UBICACION', 'CODIGO_SIIGO']]
            ], ignore_index=True)

            # Crear archivo de análisis
            analisis_file = Path("outputs") / "Analisis_Comparativo.xlsx"
            with pd.ExcelWriter(analisis_file, engine='openpyxl') as writer:
                tabla_final.to_excel(writer, sheet_name='Comparativo', index=False)
                
                # Asegurar visibilidad
                for sheet in writer.sheets.values():
                    sheet.sheet_state = 'visible'

            self.logger.agregar_log(f"Análisis comparativo creado: {analisis_file}", 'exito')
            return analisis_file

        except Exception as e:
            self.logger.agregar_log(f"Error durante el procesamiento: {str(e)}", 'error')
            raise
