import pandas as pd
from pathlib import Path

# Mostrar todas las columnas en la terminal.
pd.set_option("display.max_columns", None)
pd.set_option("display.width", 200)

# ==================================================
# 1. BUSCAR EL EXCEL ORIGINAL
# ==================================================

# Buscar junto a este código o dentro de "dataset".
carpeta = Path(__file__).resolve().parent
archivos_encontrados = []

for ubicacion in [carpeta, carpeta / "dataset"]:
    if ubicacion.exists():
        archivos_encontrados.extend(
            ubicacion.glob("CF-2023-I-204-1230_DAT*.xlsx")
        )

archivos_encontrados = sorted({
    archivo.resolve()
    for archivo in archivos_encontrados
    if not archivo.name.startswith("~$")
})

if not archivos_encontrados:
    raise FileNotFoundError(
        "\nNo se encontró el Excel original.\n"
        "Guarda el archivo CF-2023-I-204-1230_DAT.xlsx "
        "junto a Taller-IA.py o dentro de la carpeta dataset.\n"
        "También se aceptan nombres con (1), (2), etc."
    )

if len(archivos_encontrados) == 1:
    ruta_archivo = archivos_encontrados[0]

else:
    print("\nSe encontraron varios archivos:")

    for numero, archivo in enumerate(archivos_encontrados, start=1):
        print(f"{numero}. {archivo.name}")

    seleccion = int(
        input("\nEscribe el número del archivo original que quieres leer: ")
    )

    if not 1 <= seleccion <= len(archivos_encontrados):
        raise ValueError("El número seleccionado no es válido.")

    ruta_archivo = archivos_encontrados[seleccion - 1]

print("\nARCHIVO SELECCIONADO")
print(ruta_archivo)

# ==================================================
# 2. LEER LA PRIMERA HOJA
# ==================================================

with pd.ExcelFile(ruta_archivo, engine="openpyxl") as excel:
    nombre_hoja = excel.sheet_names[0]
    datos = pd.read_excel(excel, sheet_name=nombre_hoja)

print(f"\nHoja leída: {nombre_hoja}")

# ==================================================
# 3. IDENTIFICAR COLUMNAS
# ==================================================

columnas_principales = [
    columna
    for columna in datos.columns
    if not str(columna).startswith("Unnamed")
]

columnas_adicionales = [
    columna
    for columna in datos.columns
    if str(columna).startswith("Unnamed")
]

datos_principales = datos[columnas_principales].copy()

print("\nDIMENSIONES DEL ARCHIVO")
print(f"Cantidad de filas de datos: {datos.shape[0]}")
print(f"Cantidad total de columnas: {datos.shape[1]}")
print(f"Columnas principales con nombre: {len(columnas_principales)}")
print(f"Columnas adicionales sin nombre: {len(columnas_adicionales)}")

print("\nLISTADO DE COLUMNAS PRINCIPALES")

for numero, columna in enumerate(columnas_principales, start=1):
    print(f"{numero}. {columna}")

# ==================================================
# 4. MOSTRAR LAS PRIMERAS FILAS
# ==================================================

print("\nPRIMERAS 5 FILAS")
print(datos_principales.head(5).to_string(index=False))

# ==================================================
# 5. IDENTIFICAR TIPOS DE DATOS
# ==================================================

tabla_tipos = pd.DataFrame({
    "Columna": datos.columns,
    "Tipo de dato": datos.dtypes.astype(str).values
})

print("\nTIPOS DE DATOS")
print(tabla_tipos.to_string(index=False))

# ==================================================
# 6. CONTAR NULOS
# ==================================================

resumen_nulos = pd.DataFrame({
    "Columna": datos.columns,
    "Cantidad de nulos": datos.isna().sum().values,
    "Cantidad de no nulos": datos.notna().sum().values,
    "Porcentaje de nulos": datos.isna().mean().values * 100
})

print("\nVALORES NULOS POR COLUMNA")
print(
    resumen_nulos.to_string(
        index=False,
        float_format=lambda valor: f"{valor:.6f}"
    )
)

# ==================================================
# 7. UBICAR NULOS EN COLUMNAS PRINCIPALES
# ==================================================

print("\nUBICACIÓN DE NULOS EN COLUMNAS PRINCIPALES")

hay_nulos = False

for columna in columnas_principales:
    mascara_nulos = datos_principales[columna].isna()

    if mascara_nulos.any():
        hay_nulos = True

        # +2: encabezado en fila 1 y datos desde fila 2.
        filas_excel = (
            datos_principales.index[mascara_nulos] + 2
        ).tolist()

        print(
            f"{columna}: {int(mascara_nulos.sum())} nulo(s). "
            f"Filas de Excel: {filas_excel}"
        )

if not hay_nulos:
    print("No hay nulos en las columnas principales.")

# ==================================================
# 8. MOSTRAR FILAS CON NULOS
# ==================================================

filas_con_nulos = datos_principales[
    datos_principales.isna().any(axis=1)
].copy()

print("\nFILAS CON NULOS EN COLUMNAS PRINCIPALES")

if filas_con_nulos.empty:
    print("No se encontraron filas con nulos.")

else:
    filas_con_nulos.insert(
        0,
        "Fila de Excel",
        filas_con_nulos.index + 2
    )

    print(filas_con_nulos.head(20).to_string(index=False))

    if len(filas_con_nulos) > 20:
        print(
            f"Se muestran las primeras 20 de "
            f"{len(filas_con_nulos)} filas."
        )

# ==================================================
# 9. REVISAR COLUMNAS SIN NOMBRE
# ==================================================

print("\nREVISIÓN DE COLUMNAS ADICIONALES")

cantidad_filas_adicionales = 0

if not columnas_adicionales:
    print("No existen columnas adicionales sin nombre.")

else:
    conteos = datos[columnas_adicionales].notna().sum()

    print("\nCantidad de valores presentes por columna:")
    print(conteos.to_string())

    mascara_adicionales = (
        datos[columnas_adicionales].notna().any(axis=1)
    )

    cantidad_filas_adicionales = int(
        mascara_adicionales.sum()
    )

    print(
        f"\nFilas con valores adicionales: "
        f"{cantidad_filas_adicionales}"
    )

    print(
        f"Valores adicionales totales: "
        f"{int(conteos.sum())}"
    )

    if cantidad_filas_adicionales > 0:
        ejemplos = datos.loc[mascara_adicionales].head(5).copy()

        ejemplos.insert(
            0,
            "Fila de Excel",
            ejemplos.index + 2
        )

        print("\nPrimeras 5 filas con valores adicionales:")
        print(ejemplos.to_string(index=False))

        print(
            "\nEstas filas requieren revisión de estructura "
            "antes del análisis."
        )

# ==================================================
# 10. RESUMEN FINAL
# ==================================================

print("\n" + "=" * 60)
print("LECTURA Y DIAGNÓSTICO COMPLETADOS")
print("=" * 60)

print(f"Filas de datos: {len(datos)}")
print(f"Columnas principales: {len(columnas_principales)}")
print(f"Columnas adicionales: {len(columnas_adicionales)}")

print(
    "Nulos en columnas principales:",
    int(datos_principales.isna().sum().sum())
)

print(
    "Filas con valores adicionales:",
    cantidad_filas_adicionales
)

print("\nNo se eliminaron filas ni columnas.")
print("No se reemplazaron valores.")
print("El Excel original no fue modificado.")