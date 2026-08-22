import pandas as pd
from pathlib import Path
from pandas.api.types import (
    is_integer_dtype,
    is_float_dtype,
    is_bool_dtype,
    is_datetime64_any_dtype
)

# Mostrar todas las columnas en la terminal
pd.set_option("display.max_columns", None)
pd.set_option("display.width", 200)

# ==================================================
# 1. LOCALIZAR EL ARCHIVO
# ==================================================

carpeta_actual = Path(__file__).resolve().parent
nombre_archivo = "CF-2023-I-204-1230_DAT.xlsx"

# Buscarlo en la carpeta del programa o en una carpeta llamada dataset
posibles_rutas = [
    carpeta_actual / nombre_archivo,
    carpeta_actual / "dataset" / nombre_archivo
]

ruta_archivo = None

for ruta in posibles_rutas:
    if ruta.exists():
        ruta_archivo = ruta
        break

if ruta_archivo is None:
    raise FileNotFoundError(
        "No se encontró el archivo Excel.\n"
        "Colócalo en la misma carpeta que Taller-IA.py "
        "o dentro de una carpeta llamada dataset."
    )

print("=" * 70)
print("ARCHIVO ENCONTRADO")
print("=" * 70)
print(ruta_archivo)

# ==================================================
# 2. LEER EL ARCHIVO EXCEL
# ==================================================

dataset_completo = pd.read_excel(
    ruta_archivo,
    sheet_name="Hoja1"
)

print("\n" + "=" * 70)
print("DIMENSIONES DEL DATASET")
print("=" * 70)

print(f"Cantidad de filas: {dataset_completo.shape[0]}")
print(f"Cantidad total de columnas: {dataset_completo.shape[1]}")

# ==================================================
# 3. SEPARAR COLUMNAS PRINCIPALES Y ADICIONALES
# ==================================================

# Variables que tienen un nombre definido
columnas_principales = [
    columna for columna in dataset_completo.columns
    if not str(columna).startswith("Unnamed")
]

# Columnas adicionales sin nombre
columnas_adicionales = [
    columna for columna in dataset_completo.columns
    if str(columna).startswith("Unnamed")
]

# Dataset con las 11 variables principales
dataset_principal = dataset_completo[
    columnas_principales
].copy()

print("\n" + "=" * 70)
print("VARIABLES PRINCIPALES")
print("=" * 70)

for numero, columna in enumerate(columnas_principales, start=1):
    print(f"{numero}. {columna}")

print(f"\nCantidad de variables principales: {len(columnas_principales)}")

print("\n" + "=" * 70)
print("COLUMNAS ADICIONALES SIN NOMBRE")
print("=" * 70)

if len(columnas_adicionales) == 0:
    print("No se encontraron columnas adicionales.")
else:
    for numero, columna in enumerate(columnas_adicionales, start=1):
        print(f"{numero}. {columna}")

    print(
        f"\nCantidad de columnas adicionales: "
        f"{len(columnas_adicionales)}"
    )

# ==================================================
# 4. MOSTRAR LAS PRIMERAS 5 FILAS
# ==================================================

print("\n" + "=" * 70)
print("PRIMERAS 5 FILAS DE LAS VARIABLES PRINCIPALES")
print("=" * 70)

print(dataset_principal.head(5).to_string())

# ==================================================
# 5. IDENTIFICAR LOS TIPOS DE DATOS
# ==================================================

def identificar_tipo(serie):
    if is_bool_dtype(serie):
        return "Booleano"
    elif is_integer_dtype(serie):
        return "Entero"
    elif is_float_dtype(serie):
        return "Float"
    elif is_datetime64_any_dtype(serie):
        return "Fecha u hora"
    else:
        return "Texto u otro"


tabla_tipos = pd.DataFrame({
    "Variable": dataset_completo.columns,
    "Tipo en pandas": dataset_completo.dtypes.astype(str).values,
    "Clasificación": [
        identificar_tipo(dataset_completo[columna])
        for columna in dataset_completo.columns
    ]
})

print("\n" + "=" * 70)
print("TIPO DE DATO DE CADA COLUMNA")
print("=" * 70)

print(tabla_tipos.to_string(index=False))

# ==================================================
# 6. CONTAR LOS VALORES NULOS DE TODAS LAS COLUMNAS
# ==================================================

resumen_nulos = pd.DataFrame({
    "Columna": dataset_completo.columns,
    "Valores nulos": dataset_completo.isnull().sum().values,
    "Valores no nulos": dataset_completo.notnull().sum().values,
    "Porcentaje de nulos": (
        dataset_completo.isnull().mean().values * 100
    )
})

print("\n" + "=" * 70)
print("RESUMEN DE VALORES NULOS")
print("=" * 70)

print(
    resumen_nulos.to_string(
        index=False,
        formatters={
            "Porcentaje de nulos": lambda valor: f"{valor:.4f}%"
        }
    )
)

# ==================================================
# 7. NULOS EN LAS 11 VARIABLES PRINCIPALES
# ==================================================

nulos_principales = pd.DataFrame({
    "Variable": columnas_principales,
    "Cantidad de nulos": [
        dataset_principal[columna].isnull().sum()
        for columna in columnas_principales
    ],
    "Porcentaje de nulos": [
        dataset_principal[columna].isnull().mean() * 100
        for columna in columnas_principales
    ]
})

variables_con_nulos = nulos_principales[
    nulos_principales["Cantidad de nulos"] > 0
]

print("\n" + "=" * 70)
print("VARIABLES PRINCIPALES QUE TIENEN NULOS")
print("=" * 70)

if variables_con_nulos.empty:
    print("Las variables principales no contienen valores nulos.")
else:
    print(
        variables_con_nulos.to_string(
            index=False,
            formatters={
                "Porcentaje de nulos": lambda valor: f"{valor:.6f}%"
            }
        )
    )

# ==================================================
# 8. UBICACIÓN DE LOS NULOS PRINCIPALES
# ==================================================

print("\n" + "=" * 70)
print("UBICACIÓN DE LOS VALORES NULOS")
print("=" * 70)

hay_nulos = False

for columna in columnas_principales:

    cantidad_nulos = dataset_principal[columna].isnull().sum()

    if cantidad_nulos > 0:
        hay_nulos = True

        # Se suma 2 porque Excel comienza en 1 y tiene encabezado
        filas_excel = (
            dataset_principal.index[
                dataset_principal[columna].isnull()
            ] + 2
        ).tolist()

        print(
            f"{columna}: {cantidad_nulos} nulo(s). "
            f"Fila(s) de Excel: {filas_excel}"
        )

if not hay_nulos:
    print("No se encontraron valores nulos.")

# ==================================================
# 9. MOSTRAR LAS FILAS CON NULOS PRINCIPALES
# ==================================================

filas_con_nulos = dataset_principal[
    dataset_principal.isnull().any(axis=1)
].copy()

filas_con_nulos.insert(
    0,
    "Fila de Excel",
    filas_con_nulos.index + 2
)

print("\n" + "=" * 70)
print("FILAS QUE CONTIENEN NULOS")
print("=" * 70)

if filas_con_nulos.empty:
    print("No existen filas con valores nulos.")
else:
    print(filas_con_nulos.to_string(index=False))

# ==================================================
# 10. ANALIZAR LOS VALORES ADICIONALES
# ==================================================

print("\n" + "=" * 70)
print("ANÁLISIS DE COLUMNAS ADICIONALES")
print("=" * 70)

if len(columnas_adicionales) == 0:
    print("No existen columnas adicionales.")
else:
    cantidad_por_columna = (
        dataset_completo[columnas_adicionales]
        .notnull()
        .sum()
    )

    print("Cantidad de valores presentes en cada columna adicional:")
    print(cantidad_por_columna.to_string())

    total_valores_adicionales = int(
        dataset_completo[columnas_adicionales]
        .notnull()
        .sum()
        .sum()
    )

    mascara_adicionales = (
        dataset_completo[columnas_adicionales]
        .notnull()
        .any(axis=1)
    )

    cantidad_filas_adicionales = int(
        mascara_adicionales.sum()
    )

    print(
        f"\nCantidad total de valores adicionales: "
        f"{total_valores_adicionales}"
    )

    print(
        f"Cantidad de filas con valores adicionales: "
        f"{cantidad_filas_adicionales}"
    )

    # Obtener filas que contienen datos en columnas Unnamed
    filas_adicionales = dataset_completo[
        mascara_adicionales
    ].copy()

    # Agregar el número real de la fila en Excel
    filas_adicionales.insert(
        0,
        "Fila de Excel",
        filas_adicionales.index + 2
    )

    print("\nPrimeras 5 filas que contienen valores adicionales:")
    print(filas_adicionales.head(5).to_string(index=False))

# ==================================================
# 11. INFORMACIÓN GENERAL DE PANDAS
# ==================================================

print("\n" + "=" * 70)
print("INFORMACIÓN GENERAL ENTREGADA POR PANDAS")
print("=" * 70)

dataset_completo.info()

# ==================================================
# 12. RESUMEN FINAL
# ==================================================

print("\n" + "=" * 70)
print("RESUMEN FINAL")
print("=" * 70)

print(f"Filas totales: {dataset_completo.shape[0]}")
print(f"Columnas totales: {dataset_completo.shape[1]}")
print(f"Variables principales: {len(columnas_principales)}")
print(f"Columnas adicionales: {len(columnas_adicionales)}")

print(
    "Nulos en las variables principales:",
    int(dataset_principal.isnull().sum().sum())
)

if len(columnas_adicionales) > 0:
    print(
        "Valores en columnas adicionales:",
        int(
            dataset_completo[columnas_adicionales]
            .notnull()
            .sum()
            .sum()
        )
    )

    print(
        "\nADVERTENCIA: las columnas adicionales contienen datos. "
        "No deben eliminarse hasta comprobar por qué algunas filas "
        "tienen valores desplazados o concatenados."
    )