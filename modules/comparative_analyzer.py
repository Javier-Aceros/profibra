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

            df = pd.read_excel(ruta_consolidado, sheet_name="Consolidado")
            df.columns = df.columns.str.strip()
            df["REFERENCIA"] = df["REFERENCIA"].astype(str).str.strip().replace("nan", "")

            con_ref = df[df["REFERENCIA"] != ""].copy()
            sin_ref = df[df["REFERENCIA"] == ""].copy()

            pivot = con_ref.pivot_table(
                index="REFERENCIA",
                columns="ORIGEN",
                values="CANTIDAD",
                aggfunc="sum",
                fill_value=0
            ).reset_index()

            for marca in ['STIHL', 'SUZUKI', 'YAMAHA', 'SIIGO']:
                if marca not in pivot.columns:
                    pivot[marca] = 0

            info_adicional = con_ref.groupby("REFERENCIA").agg({
                'DESCRIPCION': 'first',
                'UBICACION': lambda x: ', '.join(str(v) for v in set(x) if pd.notna(v) and v != ''),
                'CODIGO_SIIGO': 'first'
            }).reset_index()

            pivot = pivot.merge(info_adicional, on="REFERENCIA", how="left")
            pivot["DIFERENCIA"] = pivot[['STIHL', 'SUZUKI', 'YAMAHA']].sum(axis=1) - pivot["SIIGO"]

            sin_ref["DIFERENCIA"] = sin_ref.apply(
                lambda x: x["CANTIDAD"] if x["ORIGEN"] in ['STIHL', 'SUZUKI', 'YAMAHA'] else -x["CANTIDAD"],
                axis=1
            )

            pivot["Inventario manual"] = pivot[["STIHL", "SUZUKI", "YAMAHA"]].sum(axis=1)
            sin_ref["Inventario manual"] = sin_ref[["STIHL", "SUZUKI", "YAMAHA"]].sum(axis=1)

            def obtener_marca(row):
                marcas = [m for m in ["STIHL", "SUZUKI", "YAMAHA"] if row[m] > 0]
                return marcas[0] if len(marcas) == 1 else ("VARIAS" if len(marcas) > 1 else "")

            pivot["MARCA"] = pivot.apply(obtener_marca, axis=1)
            sin_ref["MARCA"] = sin_ref.apply(obtener_marca, axis=1)

            columnas_finales = [
                "REFERENCIA", "STIHL", "SUZUKI", "YAMAHA", "Inventario manual", "SIIGO",
                "DESCRIPCION", "DIFERENCIA", "UBICACION", "CODIGO_SIIGO", "MARCA"
            ]

            tabla_final = pd.concat([
                pivot[columnas_finales],
                sin_ref[columnas_finales]
            ], ignore_index=True)

            analisis_file = Path("outputs") / "Analisis_Comparativo.xlsx"
            with pd.ExcelWriter(analisis_file, engine='openpyxl') as writer:
                tabla_final.to_excel(writer, sheet_name='Comparativo', index=False)

                ws = writer.sheets['Comparativo']
                rojo = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")
                for row in ws.iter_rows(min_row=2, min_col=7, max_col=7):
                    for cell in row:
                        if cell.value != 0:
                            cell.fill = rojo

                resumen = pd.DataFrame({
                    'Métrica': [
                        'Total ítems',
                        'Ítems con diferencias',
                        'Mayores en SIIGO',
                        'Mayores en Marcas',
                        'Diferencia total'
                    ],
                    'Valor': [
                        len(tabla_final),
                        sum(tabla_final['DIFERENCIA'] != 0),
                        sum(tabla_final['DIFERENCIA'] < 0),
                        sum(tabla_final['DIFERENCIA'] > 0),
                        tabla_final['DIFERENCIA'].sum()
                    ]
                })
                resumen.to_excel(writer, sheet_name='Resumen', index=False)

                for sheet in writer.sheets.values():
                    sheet.sheet_state = 'visible'

            self.logger.agregar_log(f"Análisis comparativo creado: {analisis_file}", 'exito')
            return analisis_file

        except Exception as e:
            self.logger.agregar_log(f"Error durante el procesamiento: {str(e)}", 'error')
            raise