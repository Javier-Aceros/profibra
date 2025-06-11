from pathlib import Path
import pandas as pd

def verificar_archivos_requeridos(input_dir):
    """Verifica que existan los archivos requeridos en la carpeta inputs."""
    archivos_requeridos = {
        'STIHL': None,
        'SUZUKI': None,
        'YAMAHA': None,
        'VALORACION': None
    }
    
    for archivo in Path(input_dir).glob("*.xlsx"):
        nombre = archivo.stem.upper()
        if 'STIHL' in nombre:
            archivos_requeridos['STIHL'] = archivo
        elif 'SUZUKI' in nombre:
            archivos_requeridos['SUZUKI'] = archivo
        elif 'YAMAHA' in nombre:
            archivos_requeridos['YAMAHA'] = archivo
        elif 'VALORACION' in nombre or 'VALORACIÓN' in nombre:
            archivos_requeridos['VALORACION'] = archivo
    
    faltantes = [k for k, v in archivos_requeridos.items() if v is None]
    if faltantes:
        raise FileNotFoundError(f"Archivos faltantes: {', '.join(faltantes)}")
    
    return archivos_requeridos