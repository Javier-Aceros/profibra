import pandas as pd
from pathlib import Path
from openpyxl.styles import PatternFill, Font
from openpyxl import load_workbook

class ComparativeAnalyzer:
    def __init__(self, logger):
        self.logger = logger
        # Configurar estilos
        self.rojo = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")
        self.verde = PatternFill(start_color="99FF99", end_color="99FF99", fill_type="solid")
        self.fuente_negra = Font(color="000000")
        
    def procesar_consolidado(self, ruta_consolidado):
        try:
            self.logger.agregar_log(f"\nProcesando archivo consolidado: {ruta_consolidado}...")

            # Leer datos consolidados
            df = pd.read_excel(ruta_consolidado, sheet_name="Consolidado")
            
            # Limpieza inicial
            df.columns = df.columns.str.strip()
            df["REFERENCIA"] = df["REFERENCIA"].astype(str).str.strip().replace("nan", "")
            
            # Separar datos SIIGO y marcas
            df_siigo = df[df["ORIGEN"] == "SIIGO"].copy()
            df_marcas = df[df["ORIGEN"].isin(["STIHL", "SUZUKI", "YAMAHA"])].copy()

            # Función para limpiar ubicaciones
            def limpiar_ubicacion(x):
                if pd.isna(x) or str(x).strip().lower() in ["", "nan"]:
                    return ""
                return str(x).strip()

            # Agrupar datos de marcas
            grouped_marcas = df_marcas.groupby("REFERENCIA").agg({
                'DESCRIPCION': 'first',
                'CANTIDAD': 'sum',
                'UBICACION': lambda x: ', '.join(set(filter(None, [limpiar_ubicacion(v) for v in x]))),
                'ORIGEN': lambda x: ', '.join(sorted(set(x)))
            }).rename(columns={'UBICACION': 'UBICACION_MARCAS'}).reset_index()

            # Agrupar datos SIIGO
            grouped_siigo = df_siigo.groupby("REFERENCIA").agg({
                'DESCRIPCION': 'first',
                'CANTIDAD': 'sum',
                'UBICACION': lambda x: ', '.join(set(filter(None, [limpiar_ubicacion(v) for v in x]))),
                'CODIGO_SIIGO': 'first'
            }).rename(columns={'UBICACION': 'UBICACION_SIIGO'}).reset_index()

            # Combinar datos
            tabla_final = pd.merge(
                grouped_marcas,
                grouped_siigo,
                on="REFERENCIA",
                how="outer",
                suffixes=('', '_y')
            )

            # Limpiar y estandarizar columnas
            tabla_final['DESCRIPCION'] = tabla_final['DESCRIPCION'].combine_first(tabla_final['DESCRIPCION_y'])
            
            # Actualizar columna ORIGEN para incluir SIIGO cuando corresponda
            tabla_final['ORIGEN'] = tabla_final.apply(
                lambda x: ', '.join(
                    sorted(set(
                        filter(None, [
                            *str(x['ORIGEN']).split(', '), 
                            'SIIGO' if pd.notna(x['CANTIDAD_y']) else None
                        ])
                    ))
                ), 
                axis=1
            )
            
            tabla_final['INVENTARIO MANUAL'] = tabla_final['CANTIDAD'].fillna(0)
            tabla_final['SIIGO'] = tabla_final['CANTIDAD_y'].fillna(0)
            tabla_final['DIFERENCIA'] = tabla_final['INVENTARIO MANUAL'] - tabla_final['SIIGO']
            
            # Limpiar ubicaciones
            tabla_final['UBICACION_MARCAS'] = tabla_final['UBICACION_MARCAS'].replace('nan', '').fillna('')
            tabla_final['UBICACION_SIIGO'] = tabla_final['UBICACION_SIIGO'].replace('nan', '').fillna('')
            
            # Eliminar columnas temporales
            tabla_final.drop(columns=['DESCRIPCION_y', 'CANTIDAD', 'CANTIDAD_y'], inplace=True, errors='ignore')
            
            # Ordenar columnas
            columnas_finales = [
                "REFERENCIA", 
                "DESCRIPCION",
                "ORIGEN",
                "INVENTARIO MANUAL",
                "SIIGO",
                "DIFERENCIA",
                "UBICACION_MARCAS",
                "UBICACION_SIIGO",
                "CODIGO_SIIGO"
            ]
            
            # Asegurar que todas las columnas existan
            for col in columnas_finales:
                if col not in tabla_final.columns:
                    tabla_final[col] = ''
            
            tabla_final = tabla_final[columnas_finales]

            # Procesar ítems sin referencia
            sin_ref = df[df["REFERENCIA"] == ""].copy()
            if not sin_ref.empty:
                sin_ref['INVENTARIO MANUAL'] = sin_ref['CANTIDAD']
                sin_ref['SIIGO'] = 0
                sin_ref['DIFERENCIA'] = sin_ref['INVENTARIO MANUAL']
                sin_ref['UBICACION_MARCAS'] = sin_ref['UBICACION'].apply(limpiar_ubicacion)
                sin_ref['UBICACION_SIIGO'] = ''
                sin_ref = sin_ref[columnas_finales]
                tabla_final = pd.concat([tabla_final, sin_ref], ignore_index=True)

            # Crear archivo de análisis
            analisis_file = Path("outputs") / "Analisis_Comparativo.xlsx"
            with pd.ExcelWriter(analisis_file, engine='openpyxl') as writer:
                # Hoja Comparativo
                tabla_final.to_excel(writer, sheet_name='Comparativo', index=False)
                
                # Aplicar formatos
                ws = writer.sheets['Comparativo']
                
                # Formatear diferencias
                for row in ws.iter_rows(min_row=2, min_col=6, max_col=6):  # Columna DIFERENCIA
                    for cell in row:
                        if cell.value > 0:
                            cell.fill = self.verde
                        elif cell.value < 0:
                            cell.fill = self.rojo
                        cell.font = self.fuente_negra
                
                # Asegurar visibilidad
                for sheet in writer.sheets.values():
                    sheet.sheet_state = 'visible'

            self.logger.agregar_log(f"Análisis comparativo creado: {analisis_file}", 'exito')
            return analisis_file

        except Exception as e:
            self.logger.agregar_log(f"Error durante el procesamiento: {str(e)}", 'error')
            raise
