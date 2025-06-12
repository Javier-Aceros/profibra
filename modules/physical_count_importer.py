import pandas as pd
from pathlib import Path

class PhysicalCountImporter:
    def __init__(self, logger):
        self.logger = logger

    def generar_importacion_conteo(self, ruta_analisis):
        try:
            self.logger.agregar_log(f"Generando archivo de importación desde: {ruta_analisis}")
            
            df = pd.read_excel(ruta_analisis, sheet_name="Comparativo")

            # Filtrar filas con un solo código no vacío
            df = df[df["CODIGO_SIIGO"].notna()]
            df = df[df["CODIGO_SIIGO"].astype(str).str.strip() != ""]
            df = df[~df["CODIGO_SIIGO"].astype(str).str.contains(",")]

            def construir_nombre(row):
                descripcion = str(row["DESCRIPCION"]).strip()
                ubicacion = ""
                if str(row["UBICACION_MARCAS"]).strip().lower() not in ["", "nan"]:
                    ubicacion = str(row["UBICACION_MARCAS"]).strip()
                elif str(row["UBICACION_SIIGO"]).strip().lower() not in ["", "nan"]:
                    ubicacion = str(row["UBICACION_SIIGO"]).strip()
                return f"{descripcion} ({ubicacion})" if ubicacion else descripcion

            # Create the base dataframe
            df_import = pd.DataFrame({
                "Código del producto \n(obligatorio) ": df["CODIGO_SIIGO"],
                "Nombre del producto / Servicio": df.apply(construir_nombre, axis=1),
                "Referencia de fábrica": df["REFERENCIA"],
                "Código de Bodega": "3-Almacén",
                "Existencias contadas \n(obligatorio)": df["INVENTARIO MANUAL"]
            })

            # Create additional rows for each item
            df_ecommerce = pd.DataFrame({
                "Código del producto \n(obligatorio) ": df["CODIGO_SIIGO"],
                "Nombre del producto / Servicio": df.apply(construir_nombre, axis=1),
                "Referencia de fábrica": df["REFERENCIA"],
                "Código de Bodega": "1-E-COMMERCE",
                "Existencias contadas \n(obligatorio)": 0
            })

            df_empty_warehouse = pd.DataFrame({
                "Código del producto \n(obligatorio) ": df["CODIGO_SIIGO"],
                "Nombre del producto / Servicio": df.apply(construir_nombre, axis=1),
                "Referencia de fábrica": df["REFERENCIA"],
                "Código de Bodega": "",
                "Existencias contadas \n(obligatorio)": 0
            })

            # Concatenate all dataframes
            df_final = pd.concat([df_import, df_ecommerce, df_empty_warehouse], ignore_index=True)

            # Sort by product code to keep related items together
            df_final.sort_values(by="Código del producto \n(obligatorio) ", inplace=True)

            output_path = Path("outputs") / "Importacion_conteo_fisico.xlsx"
            
            # Write to Excel with sheet name "Datos"
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                df_final.to_excel(writer, sheet_name="Datos", index=False)

            self.logger.agregar_log(f"Archivo de importación generado: {output_path}", "exito")
            return output_path
        
        except Exception as e:
            self.logger.agregar_log(f"Error generando importación: {str(e)}", "error")
            raise
