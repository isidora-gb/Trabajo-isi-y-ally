from pathlib import Path

import pandas as pd


# ==============================================================
# CONFIGURACION
# ==============================================================

NOMBRE_ARCHIVO_ENTRADA = "CF-2023-I-204-1230_DAT.xlsx"
NOMBRE_HOJA_ENTRADA = "Hoja1"
NOMBRE_ARCHIVO_SALIDA = "Dataset_eolico_filtrado.xlsx"

# Estas son las 11 variables que se utilizaran en PCA y DBSCAN.
VARIABLES_UTILIZADAS = [
    "AcelX1_SisGenEol",
    "AcelY1_SisGenEol",
    "AcelZ1_SisGenEol",
    "GiroX_SisGenEol",
    "GiroY_SisGenEol",
    "GiroZ_SisGenEol",
    "MagneX_SisGenEol",
    "MagneY_SisGenEol",
    "MagneZ_SisGenEol",
    "Ter_SisGenEol",
    "smp",
]


def localizar_archivo(carpeta_programa: Path) -> Path:
    """Busca el Excel junto al programa o dentro de la carpeta dataset."""
    posibles_rutas = [
        carpeta_programa / NOMBRE_ARCHIVO_ENTRADA,
        carpeta_programa / "dataset" / NOMBRE_ARCHIVO_ENTRADA,
    ]

    for ruta in posibles_rutas:
        if ruta.is_file():
            return ruta

    raise FileNotFoundError(
        f"No se encontro el archivo '{NOMBRE_ARCHIVO_ENTRADA}'.\n"
        "Colocalo en la misma carpeta que este programa o dentro de "
        "una carpeta llamada 'dataset'."
    )


def filtrar_dataset(ruta_entrada: Path) -> tuple[pd.DataFrame, list[int]]:
    """Conserva las variables utilizadas y elimina registros incompletos."""
    dataset_completo = pd.read_excel(
        ruta_entrada,
        sheet_name=NOMBRE_HOJA_ENTRADA,
    )

    # Quita espacios accidentales de los nombres de las columnas.
    dataset_completo.columns = dataset_completo.columns.map(
        lambda nombre: nombre.strip() if isinstance(nombre, str) else nombre
    )

    variables_faltantes = [
        variable
        for variable in VARIABLES_UTILIZADAS
        if variable not in dataset_completo.columns
    ]

    if variables_faltantes:
        raise ValueError(
            "No se encontraron las siguientes variables en el Excel:\n- "
            + "\n- ".join(variables_faltantes)
        )

    # Al seleccionar solo estas columnas se eliminan las variables no usadas,
    # incluidas las columnas Unnamed: 11 hasta Unnamed: 20.
    dataset_filtrado = dataset_completo[VARIABLES_UTILIZADAS].copy()

    # Convierte celdas vacias o que contienen solo espacios en valores nulos.
    dataset_filtrado = dataset_filtrado.replace(r"^\s*$", pd.NA, regex=True)

    # PCA y DBSCAN requieren datos numericos. Un texto no convertible se
    # transforma en nulo y su fila tambien sera eliminada.
    for variable in VARIABLES_UTILIZADAS:
        dataset_filtrado[variable] = pd.to_numeric(
            dataset_filtrado[variable],
            errors="coerce",
        )

    filas_incompletas = dataset_filtrado.isna().any(axis=1)

    # Se suma 2 para informar el numero real de la fila en Excel:
    # una fila por el encabezado y otra porque Excel comienza en 1.
    filas_excel_eliminadas = (
        dataset_filtrado.index[filas_incompletas] + 2
    ).tolist()

    dataset_limpio = (
        dataset_filtrado.loc[~filas_incompletas]
        .reset_index(drop=True)
    )

    if dataset_limpio.empty:
        raise ValueError(
            "El filtro elimino todas las filas porque no habia registros "
            "completos en las 11 variables."
        )

    return dataset_limpio, filas_excel_eliminadas


def guardar_dataset(dataset_limpio: pd.DataFrame, ruta_salida: Path) -> None:
    """Guarda el nuevo dataset limpio en un archivo Excel."""
    dataset_limpio.to_excel(
        ruta_salida,
        sheet_name="Datos_limpios",
        index=False,
    )


def main() -> None:
    carpeta_programa = Path(__file__).resolve().parent
    ruta_entrada = localizar_archivo(carpeta_programa)
    ruta_salida = ruta_entrada.parent / NOMBRE_ARCHIVO_SALIDA

    try:
        dataset_limpio, filas_eliminadas = filtrar_dataset(ruta_entrada)
        guardar_dataset(dataset_limpio, ruta_salida)
    except PermissionError as error:
        raise PermissionError(
            "No se pudo abrir o guardar el archivo. Cierra el Excel si esta "
            "abierto y vuelve a ejecutar el programa."
        ) from error

    print("=" * 70)
    print("FILTRO COMPLETADO")
    print("=" * 70)
    print(f"Archivo original: {ruta_entrada}")
    print(f"Filas originales: {len(dataset_limpio) + len(filas_eliminadas)}")
    print(f"Filas eliminadas: {len(filas_eliminadas)}")
    print(f"Filas conservadas: {len(dataset_limpio)}")
    print(f"Variables conservadas: {len(dataset_limpio.columns)}")
    print(f"Valores nulos restantes: {int(dataset_limpio.isna().sum().sum())}")

    if filas_eliminadas:
        print(f"Filas de Excel eliminadas: {filas_eliminadas}")
    else:
        print("No fue necesario eliminar filas.")

    print(f"\nNuevo dataset guardado en:\n{ruta_salida}")


if __name__ == "__main__":
    main()