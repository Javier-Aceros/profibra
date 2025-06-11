import pandas as pd
from pathlib import Path

class FileProcessor:
    def __init__(self, logger):
        self.logger = logger

    def _buscar_fila_encabezados(self, df_raw, columnas_clave):
        """Busca la fila que contiene los encabezados requeridos."""
        for i, fila in df_raw.iterrows():
            valores = [str(cell).upper().strip() for cell in fila.values]
            if all(any(col in valores for col in cols) for cols in columnas_clave):
                return i
        return None

    def _mapear_columnas(self, df, mapeo_columnas):
        """Renombra columnas según el mapeo proporcionado."""
        columnas_renombradas = {}
        for col in df.columns:
            col_upper = str(col).upper().strip()
            for key, values in mapeo_columnas.items():
                if any(v == col_upper for v in values):
                    columnas_renombradas[col] = key
                    break
            else:
                columnas_renombradas[col] = col
        return df.rename(columns=columnas_renombradas)

    def _procesar_archivo_generico(self, archivo, mapeo_columnas, columnas_requeridas, origen):
        """Procesamiento genérico para archivos de inventario."""
        try:
            self.logger.agregar_log(f"Procesando archivo {archivo.name}...")

            df_raw = pd.read_excel(archivo, header=None)
            columnas_clave = [[col for col in cols] for cols in mapeo_columnas.values()]
            fila_encabezados = self._buscar_fila_encabezados(df_raw, columnas_clave)

            if fila_encabezados is None:
                raise ValueError("No se encontró fila con los encabezados requeridos")

            df = pd.read_excel(archivo, header=fila_encabezados)
            df.columns = df.columns.str.upper().str.strip()
            df = self._mapear_columnas(df, mapeo_columnas)

            for col in columnas_requeridas:
                if col not in df.columns:
                    raise ValueError(f"Falta columna requerida: {col}")

            # Limpieza de datos
            df = df.dropna(how='all')
            if 'REFERENCIA' in df.columns:
                df['REFERENCIA'] = df['REFERENCIA'].astype(str).str.strip()
            if 'DESCRIPCION' in df.columns:
                df['DESCRIPCION'] = df['DESCRIPCION'].astype(str).str.strip()
            if 'CANTIDAD' in df.columns:
                df['CANTIDAD'] = pd.to_numeric(df['CANTIDAD'], errors='coerce').fillna(0)

            # Agregar columnas adicionales si no existen
            if 'UBICACION' not in df.columns:
                df['UBICACION'] = ''
            if 'CODIGO_SIIGO' not in df.columns:
                df['CODIGO_SIIGO'] = ''

            df['ORIGEN'] = origen
            self.logger.agregar_log(f"Archivo {archivo.name} procesado correctamente", 'exito')
            return df

        except Exception as e:
            self.logger.agregar_log(f"Error procesando {archivo.name}: {str(e)}", 'error')
            raise

    def leer_archivo_marca(self, archivo, marca):
        """Procesa archivos de marcas específicas (STIHL, SUZUKI, YAMAHA)."""
        mapeo_columnas = {
            'REFERENCIA': ['REFERENCIA', 'CODIGO', 'SKU', 'MODELO', 'PART NUMBER'],
            'DESCRIPCION': ['NOMBRE', 'DESCRIPCION', 'DESCRIPCIÓN', 'DESCRIP', 'PRODUCTO'],
            'CANTIDAD': ['CANTIDAD', 'QTY', 'QUANTITY', 'STOCK'],
            'UBICACION': ['UBICACION', 'UBICACIÓN', 'LOCALIZACION', 'ALMACEN']
        }
        columnas_requeridas = ['REFERENCIA', 'DESCRIPCION', 'CANTIDAD']

        df = self._procesar_archivo_generico(archivo, mapeo_columnas, columnas_requeridas, marca)
        return df[['REFERENCIA', 'DESCRIPCION', 'CANTIDAD', 'ORIGEN', 'UBICACION', 'CODIGO_SIIGO']]

    def leer_archivo_siigo(self, archivo):
        try:
            self.logger.agregar_log(f"Procesando archivo de valoración: {archivo.name}...")

            df_raw = pd.read_excel(archivo, header=None)
            fila_encabezados = None
            for i, fila in df_raw.iterrows():
                valores = [str(cell).upper().strip() for cell in fila.values]
                tiene_codigo = any(col in valores for col in ['CÓDIGO PRODUCTO', 'CODIGO PRODUCTO', 'CODIGO'])
                tiene_referencia = any(col in valores for col in ['REFERENCIA FÁBRICA', 'REFERENCIA FABRICA', 'REFERENCIA'])
                tiene_saldo = any(col in valores for col in ['SALDO CANTIDADES', 'SALDO', 'CANTIDAD'])

                if tiene_codigo and tiene_referencia and tiene_saldo:
                    fila_encabezados = i
                    self.logger.agregar_log(f"Encabezados encontrados en la fila {fila_encabezados + 1}")
                    break

            if fila_encabezados is None:
                muestra = df_raw.head(5).applymap(lambda x: str(x).upper().strip())
                self.logger.agregar_log(f"Primeras filas del archivo:\n{muestra}")
                raise ValueError("No se encontró una fila con todos los encabezados requeridos")

            df = pd.read_excel(archivo, header=fila_encabezados)
            df.columns = df.columns.str.upper().str.strip()

            mapeo_columnas = {
                'CODIGO_PRODUCTO': ['CÓDIGO PRODUCTO', 'CODIGO PRODUCTO', 'CODIGO', 'CODIGO SIIGO'],
                'NOMBRE_PRODUCTO': ['NOMBRE PRODUCTO', 'DESCRIPCION', 'PRODUCTO', 'NOMBRE'],
                'REFERENCIA_FABRICA': ['REFERENCIA FÁBRICA', 'REFERENCIA FABRICA', 'REFERENCIA', 'SKU', 'MODELO'],
                'SALDO_CANTIDADES': ['SALDO CANTIDADES', 'SALDO', 'CANTIDAD', 'STOCK', 'EXISTENCIAS']
            }

            columnas_renombradas = {}
            for col in df.columns:
                col_upper = str(col).upper().strip()
                for key, values in mapeo_columnas.items():
                    if any(v == col_upper for v in values):
                        columnas_renombradas[col] = key
                        break
                else:
                    columnas_renombradas[col] = col

            df = df.rename(columns=columnas_renombradas)

            columnas_requeridas = ['CODIGO_PRODUCTO', 'REFERENCIA_FABRICA', 'SALDO_CANTIDADES']
            for col in columnas_requeridas:
                if col not in df.columns:
                    self.logger.agregar_log(f"Columnas disponibles: {list(df.columns)}")
                    raise ValueError(f"No se pudo identificar la columna '{col}'")

            if 'NOMBRE_PRODUCTO' not in df.columns:
                df['NOMBRE_PRODUCTO'] = ''

            df = df.dropna(how='all')
            df['REFERENCIA_FABRICA'] = df['REFERENCIA_FABRICA'].astype(str).str.strip()
            df['NOMBRE_PRODUCTO'] = df['NOMBRE_PRODUCTO'].astype(str).str.strip()
            df['SALDO_CANTIDADES'] = pd.to_numeric(df['SALDO_CANTIDADES'], errors='coerce').fillna(0)

            df['UBICACION'] = df['NOMBRE_PRODUCTO'].str.extract(r'\((.*?)\)$').fillna('')
            df['DESCRIPCION'] = df['NOMBRE_PRODUCTO'].str.replace(r'\(.*?\)$', '', regex=True).str.strip()

            df['ORIGEN'] = 'SIIGO'

            self.logger.agregar_log(f"Archivo de valoración procesado correctamente. Filas leídas: {len(df)}", 'exito')

            return df[['REFERENCIA_FABRICA', 'DESCRIPCION', 'SALDO_CANTIDADES', 'ORIGEN', 'UBICACION', 'CODIGO_PRODUCTO']].rename(columns={
                'REFERENCIA_FABRICA': 'REFERENCIA',
                'SALDO_CANTIDADES': 'CANTIDAD',
                'CODIGO_PRODUCTO': 'CODIGO_SIIGO'
            })

        except Exception as e:
            self.logger.agregar_log(f"Error procesando archivo de valoración: {str(e)}", 'error')
            raise