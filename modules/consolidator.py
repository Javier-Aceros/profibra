import pandas as pd
from pathlib import Path
from .file_processor import FileProcessor

class Consolidator:
    def __init__(self, logger):
        self.logger = logger
        self.file_processor = FileProcessor(logger)

    def crear_consolidado(self):
        try:
            self.logger.agregar_log("Iniciando proceso de consolidación...")

            inputs_dir = Path("inputs")
            if not inputs_dir.exists():
                raise FileNotFoundError("No se encontró la carpeta 'inputs'")

            archivos_requeridos = {
                'STIHL': None,
                'SUZUKI': None,
                'YAMAHA': None,
                'VALORACION': None
            }

            self.logger.agregar_log("Buscando archivos en la carpeta inputs...")
            for archivo in inputs_dir.glob("*.xlsx"):
                nombre = archivo.stem.upper()
                if 'STIHL' in nombre:
                    archivos_requeridos['STIHL'] = archivo
                elif 'SUZUKI' in nombre:
                    archivos_requeridos['SUZUKI'] = archivo
                elif 'YAMAHA' in nombre:
                    archivos_requeridos['YAMAHA'] = archivo
                elif 'VALORACION' in nombre or 'VALORACIÓN' in nombre:
                    archivos_requeridos['VALORACION'] = archivo

            for marca, archivo in archivos_requeridos.items():
                if archivo is None:
                    raise FileNotFoundError(f"No se encontró el archivo para {marca}")

            dfs = []
            for marca in ['STIHL', 'SUZUKI', 'YAMAHA']:
                df_marca = self.file_processor.leer_archivo_marca(archivos_requeridos[marca], marca)
                df_marca['STIHL'] = 0
                df_marca['SUZUKI'] = 0
                df_marca['YAMAHA'] = 0
                df_marca[marca] = df_marca['CANTIDAD']
                dfs.append(df_marca)

            df_siigo = self.file_processor.leer_archivo_siigo(archivos_requeridos['VALORACION'])
            df_siigo['STIHL'] = 0
            df_siigo['SUZUKI'] = 0
            df_siigo['YAMAHA'] = 0
            df_siigo['SIIGO'] = df_siigo['CANTIDAD']
            dfs.append(df_siigo)

            consolidado = pd.concat(dfs, ignore_index=True)

            output_dir = Path("outputs")
            output_dir.mkdir(exist_ok=True)

            consolidado_file = output_dir / "Consolidado_Inventarios.xlsx"
            with pd.ExcelWriter(consolidado_file, engine='openpyxl') as writer:
                consolidado.to_excel(writer, sheet_name='Consolidado', index=False)
                writer.sheets['Consolidado'].sheet_state = 'visible'

            self.logger.agregar_log(f"Archivo consolidado creado: {consolidado_file}", 'exito')
            return consolidado_file

        except Exception as e:
            self.logger.agregar_log(f"Error durante la consolidación: {str(e)}", 'error')
            raise