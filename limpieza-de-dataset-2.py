from pathlib import Path
import argparse
import sys

try:
    import numpy as np
    import pandas as pd
    import openpyxl  # Lector y escritor de archivos Excel.
except ModuleNotFoundError:
    print("Falta una biblioteca. Pega este comando en la terminal de VS Code:")
    prefijo = "& " if sys.platform == "win32" else ""
    print(f'{prefijo}"{sys.executable}" -m pip install pandas openpyxl')
    print("Cuando termine la instalación, vuelve a ejecutar el archivo.")
    raise SystemExit(1)


COLUMNAS = [
    "AcelX1_SisGenEol", "AcelY1_SisGenEol", "AcelZ1_SisGenEol",
    "GiroX_SisGenEol", "GiroY_SisGenEol", "GiroZ_SisGenEol",
    "MagneX_SisGenEol", "MagneY_SisGenEol", "MagneZ_SisGenEol",
    "Ter_SisGenEol", "smp",
]


def rellenar_vecinos(serie, cantidad=2):
    """Promedia los valores válidos de las filas más cercanas.

    Solo utiliza valores originalmente presentes. Una estimación nunca se
    convierte en vecino de otra estimación. Los empates se resuelven dando
    prioridad a la fila de menor número.
    """
    if cantidad < 1:
        raise ValueError("La cantidad de vecinos debe ser al menos 1.")

    valores = serie.to_numpy(dtype=float, copy=True)
    resultado = valores.copy()
    validos = np.flatnonzero(~np.isnan(valores))
    vacios = np.flatnonzero(np.isnan(valores))
    cambios = []

    if len(vacios) and not len(validos):
        raise ValueError(f"La columna {serie.name} está completamente vacía.")

    for fila in vacios:
        centro = int(np.searchsorted(validos, fila))
        candidatos = validos[
            max(0, centro - cantidad): min(len(validos), centro + cantidad)
        ]
        vecinos = sorted(
            candidatos, key=lambda i: (abs(int(i) - int(fila)), int(i))
        )[:cantidad]
        estimacion = float(np.mean(valores[vecinos]))
        resultado[fila] = estimacion
        cambios.append({
            "fila_excel": int(fila) + 2,
            "columna": serie.name,
            "valor_estimado": estimacion,
            "metodo": "Promedio de vecinos por posición de fila",
            "filas_vecinas_excel": ", ".join(str(int(i) + 2) for i in vecinos),
            "valores_vecinos": ", ".join(str(valores[i]) for i in vecinos),
        })

    return pd.Series(resultado, index=serie.index, name=serie.name), cambios


def limpiar(original, cantidad_vecinos=2):
    """Devuelve datos, registro de estimaciones y filas para revisar."""
    faltantes = [nombre for nombre in COLUMNAS if nombre not in original.columns]
    if faltantes:
        raise ValueError(f"No encuentro estos encabezados en el Excel: {faltantes}")

    adicionales = [nombre for nombre in original.columns if nombre not in COLUMNAS]
    con_extras = (
        original[adicionales].notna().any(axis=1)
        if adicionales else pd.Series(False, index=original.index)
    )
    revisar = original.loc[con_extras].copy()
    revisar.insert(0, "fila_excel_original", revisar.index.to_numpy() + 2)

    # Esta selección elimina las columnas adicionales, sin eliminar filas.
    datos = original[COLUMNAS].copy()
    datos = datos.replace(r"^\s*$", np.nan, regex=True)
    for columna in COLUMNAS:
        # Un texto inesperado produce un mensaje para revisarlo.
        datos[columna] = pd.to_numeric(datos[columna], errors="raise")

    cambios = []
    for columna in COLUMNAS[:-1]:
        datos[columna], registro = rellenar_vecinos(
            datos[columna], cantidad_vecinos
        )
        cambios.extend(registro)

    # smp parece ser un contador de muestras. No se promedia como un sensor.
    # En el original entregado solo falta en la última fila.
    vacios_smp = np.flatnonzero(datos["smp"].isna().to_numpy())
    for fila in vacios_smp:
        if fila != len(datos) - 1 or fila < 2:
            raise ValueError(
                "Hay un vacío en smp fuera del final del archivo. "
                "Revisa el contador antes de completar ese dato."
            )
        penultimo = datos["smp"].iloc[fila - 1]
        antepenultimo = datos["smp"].iloc[fila - 2]
        if (
            pd.isna(penultimo) or pd.isna(antepenultimo)
            or not float(penultimo).is_integer()
            or penultimo != antepenultimo + 1
        ):
            raise ValueError("No puedo inferir el último smp: su secuencia no es clara.")
        estimacion = float(penultimo + 1)
        datos.loc[datos.index[fila], "smp"] = estimacion
        cambios.append({
            "fila_excel": int(fila) + 2,
            "columna": "smp",
            "valor_estimado": estimacion,
            "metodo": "Contador inferido: número anterior + 1",
            "filas_vecinas_excel": f"{int(fila)}, {int(fila) + 1}",
            "valores_vecinos": f"{antepenultimo}, {penultimo}",
        })

    columnas_registro = [
        "fila_excel", "columna", "valor_estimado", "metodo",
        "filas_vecinas_excel", "valores_vecinos",
    ]
    registro = pd.DataFrame(cambios, columns=columnas_registro)
    assert len(datos) == len(original), "Cambió la cantidad de filas."
    assert list(datos.columns) == COLUMNAS, "Cambió el orden de las columnas."
    return datos, registro, revisar, adicionales


def main():
    argumentos = argparse.ArgumentParser(description="Limpieza inicial del Excel eólico")
    argumentos.add_argument("archivo", nargs="?", help="Ruta opcional del Excel original")
    argumentos.add_argument("--salida", help="Ruta opcional del resultado")
    argumentos.add_argument("--vecinos", type=int, default=2, help="Vecinos por vacío (2 por defecto)")
    opciones = argumentos.parse_args()

    carpeta = Path(__file__).resolve().parent
    if opciones.archivo:
        entrada = Path(opciones.archivo).resolve()
    else:
        nombres = ["CF-2023-I-204-1230_DAT.xlsx", "CF-2023-I-204-1230_DAT(1).xlsx"]
        entrada = next((carpeta / n for n in nombres if (carpeta / n).is_file()), None)
        if entrada is None:
            raise FileNotFoundError(
                "Coloca limpieza-de-dataset-2.py en la misma carpeta que tu Excel original. "
                "Nombre esperado: CF-2023-I-204-1230_DAT.xlsx"
            )

    salida = (
        Path(opciones.salida).resolve() if opciones.salida
        else entrada.parent / "Dataset_eolico_filtrado_2.xlsx"
    )
    if salida.resolve() == entrada.resolve():
        raise ValueError("El resultado debe tener un nombre distinto del original.")

    print("Leyendo el Excel original...", flush=True)
    original = pd.read_excel(
        entrada, sheet_name=0, engine="openpyxl", na_values=["NULL", "null", "None", " "]
    )
    datos, registro, revisar, adicionales = limpiar(original, opciones.vecinos)
    print("Guardando la nueva copia...", flush=True)
    salida.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(salida, engine="openpyxl") as escritor:
        datos.to_excel(escritor, sheet_name="Datos", index=False)
        registro.to_excel(escritor, sheet_name="Valores_estimados", index=False)
        revisar.to_excel(escritor, sheet_name="Filas_para_revisar", index=False)
        for hoja in escritor.book.worksheets:
            hoja.freeze_panes = "A2"

    print("\nLIMPIEZA INICIAL COMPLETADA")
    print(f"Filas conservadas: {len(datos)}")
    print(f"Columnas conservadas: {len(datos.columns)}")
    print(f"Columnas adicionales eliminadas de Datos: {len(adicionales)}")
    print(f"Vacíos completados: {len(registro)}")
    print(f"Vacíos restantes: {int(datos.isna().sum().sum())}")
    print(f"Filas con valores en columnas extra, para revisar: {len(revisar)}")
    print("Esas filas siguen en Datos. Revisa sus valores desplazados antes del PCA.")
    print(f"Archivo guardado en: {salida}")
    print("\nÚltimas filas:")
    print(datos.tail(3).to_string(index=False))


if __name__ == "__main__":
    try:
        main()
    except PermissionError:
        print("Cierra el Excel de salida si está abierto y ejecuta de nuevo el código.")
        raise SystemExit(1)
