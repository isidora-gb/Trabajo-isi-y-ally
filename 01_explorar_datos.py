# -*- coding: utf-8 -*-
"""PASO 1: SOLO EXPLORACION DEL EXCEL ORIGINAL.
No limpia, no guarda archivos, no modifica datos ni detecta atipicos.
Instala: python -m pip install pandas openpyxl
Ejecuta: python 01_explorar_datos.py
"""
from pathlib import Path
import sys
import pandas as pd

CARPETA = Path(__file__).resolve().parent
ARCHIVO = CARPETA / "CF-2023-I-204-1230_DAT.xlsx"
VARIABLES = [
    "AcelX1_SisGenEol", "AcelY1_SisGenEol", "AcelZ1_SisGenEol",
    "GiroX_SisGenEol", "GiroY_SisGenEol", "GiroZ_SisGenEol",
    "MagneX_SisGenEol", "MagneY_SisGenEol", "MagneZ_SisGenEol",
    "Ter_SisGenEol", "smp",
]

def main():
    archivo = Path(sys.argv[1]).expanduser().resolve() if len(sys.argv) > 1 else ARCHIVO
    if not archivo.is_file():
        raise FileNotFoundError(f"No se encuentra el Excel: {archivo}")

    print("\n" + "=" * 64)
    print("PASO 1 | LECTURA DEL EXCEL ORIGINAL (SIN LIMPIEZA)")
    print("=" * 64)
    datos = pd.read_excel(archivo, sheet_name=0, engine="openpyxl")
    disponibles = [col for col in VARIABLES if col in datos.columns]
    adicionales = [col for col in datos.columns if col not in VARIABLES]
    tabla = datos[disponibles]

    print(f"Archivo:                 {archivo.name}")
    print(f"Filas (registros):       {len(datos):,}")
    print(f"Columnas detectadas:     {datos.shape[1]}")
    print(f"Variables esperadas:     {len(disponibles)} de {len(VARIABLES)}")
    print(f"  Sensores:              {sum(col != 'smp' for col in disponibles)}")
    print(f"  Contador smp:          {'SI' if 'smp' in disponibles else 'NO'}")
    print(f"Columnas adicionales:    {len(adicionales)}")
    print(f"Celdas vacias en las {len(disponibles)} variables: {int(tabla.isna().sum().sum()):,}")

    print("\nTIPOS DE DATOS LEIDOS DESDE EL EXCEL:")
    filas = []
    for col in disponibles:
        filas.append({
            "Variable": col,
            "Tipo (pandas)": str(datos[col].dtype),
            "Valores presentes": int(datos[col].notna().sum()),
            "Valores vacios": int(datos[col].isna().sum()),
        })
    print(pd.DataFrame(filas).to_string(index=False))

    tipos = datos[disponibles].dtypes.astype(str)
    print("\nRESUMEN DE TIPOS (ANTES DE LIMPIAR):")
    for tipo, n in tipos.value_counts().items():
        print(f"  {tipo:<12}: {n} variable(s)")
    print("\nNOTA: 'smp' puede leerse como float64 si hay vacios o desalineaciones.")
    print("La limpieza posterior dejara 10 sensores float64 y smp int64.")
    print("Las columnas 'Unnamed' son posiciones extra del archivo, NO sensores nuevos.")
    print("\nFIN: no se ha guardado ni cambiado ningun archivo.\n")

if __name__ == "__main__":
    try:
        main()
    except (ValueError, FileNotFoundError) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        sys.exit(1)
