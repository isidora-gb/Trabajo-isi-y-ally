"""
TALLER IA - DEPURACION DEL DATASET EOLICO MEDIANTE 6 DESVIACIONES ESTANDAR

Entradas:
    Dataset_eolico_filtrado_2.xlsx (o una version con sufijo, p. ej. (4))
    Hoja: Datos

Salidas:
    Dataset_eolico_depurado_6sigma.xlsx   -> datos para los graficos y modelos
    Resultados_6sigma.xlsx                -> estadisticas y auditoria

Regla: borrar toda la fila si cualquiera de las variables fisicas evaluadas
esta ESTRICTAMENTE por debajo de media - 6*DS o por encima de media + 6*DS.
El contador 'smp' se conserva y se describe, pero no se evalua como magnitud
fisica. No se utiliza RIC/IQR ni se realizan eliminaciones iterativas.
"""

from pathlib import Path
import numpy as np
import pandas as pd

# ---------------- 1. CONFIGURACION ----------------
CARPETA = Path(__file__).resolve().parent
ARCHIVO_PREFERIDO = CARPETA / "Dataset_eolico_filtrado_2.xlsx"
PATRON_ALTERNATIVO = "Dataset_eolico_filtrado_2*.xlsx"
HOJA_DATOS = "Datos"
NUM_DESVIACIONES = 6
COLUMNAS_NO_FISICAS = {"smp"}  # Contador/identificador, no una medicion del proceso.

ARCHIVO_LIMPIO = CARPETA / "Dataset_eolico_depurado_6sigma.xlsx"
ARCHIVO_REPORTE = CARPETA / "Resultados_6sigma.xlsx"

# ---------------- 2. LEER DATOS YA FILTRADOS ----------------
if ARCHIVO_PREFERIDO.exists():
    archivo_entrada = ARCHIVO_PREFERIDO
else:
    candidatos = [
        p for p in CARPETA.glob(PATRON_ALTERNATIVO)
        if not p.name.startswith("~$")
    ]
    if not candidatos:
        raise FileNotFoundError(
            "No se encontro Dataset_eolico_filtrado_2.xlsx. "
            "Pon el Excel filtrado en la misma carpeta que este .py."
        )
    archivo_entrada = max(candidatos, key=lambda p: p.stat().st_mtime)

print(f"Archivo de entrada: {archivo_entrada.name}")
datos = pd.read_excel(archivo_entrada, sheet_name=HOJA_DATOS, engine="openpyxl")
if datos.empty:
    raise ValueError("La hoja Datos no contiene registros.")

# Validacion: no ocultar errores de lectura convirtiendo valores incorrectos en ceros.
for variable in datos.columns:
    columna_numerica = pd.to_numeric(datos[variable], errors="coerce")
    if columna_numerica.isna().any():
        n_invalidos = int(columna_numerica.isna().sum())
        raise ValueError(
            f"La columna '{variable}' contiene {n_invalidos} valores vacios/no numericos. "
            "Revisa primero la limpieza inicial, sin borrar registros silenciosamente."
        )
    datos[variable] = columna_numerica

columnas = datos.columns.tolist()
columnas_fisicas = [c for c in columnas if c not in COLUMNAS_NO_FISICAS]
if not columnas_fisicas:
    raise ValueError("No se encontraron variables fisicas para evaluar.")
if not np.isfinite(datos.to_numpy(dtype=float)).all():
    raise ValueError("Se detectaron valores infinitos o no validos en el archivo.")

print(f"Filas iniciales: {len(datos):,} | Variables: {len(columnas)}")
print("Variables evaluadas:", ", ".join(columnas_fisicas))
print("Variables no fisicas excluidas de la decision:", ", ".join(sorted(COLUMNAS_NO_FISICAS)))

# ---------------- 3. ESTADISTICA ORIGINAL (UNA SOLA VEZ) ----------------
medias = datos.mean(axis=0)
desviaciones = datos.std(axis=0, ddof=1)  # Desviacion estandar MUESTRAL.
limite_inferior = medias - NUM_DESVIACIONES * desviaciones
limite_superior = medias + NUM_DESVIACIONES * desviaciones

# Para cada columna numerica: verificar si cada dato sobrepasa ±6 DS.
# COMPARACION ESTRICTA: si esta exactamente en el limite NO se elimina.
fuera_de_limites = datos.lt(limite_inferior) | datos.gt(limite_superior)
# Si una columna tiene desviacion estandar 0, todos sus valores son iguales;
# no puede generar observaciones fuera de ±6 DS.

# Separar estadistica descriptiva de la regla de eliminacion:
# 'smp' tiene estadistica, pero no decide si se elimina una fila.
filas_a_eliminar = fuera_de_limites[columnas_fisicas].any(axis=1)

z_scores = (datos - medias).div(desviaciones.replace(0, np.nan)).fillna(0.0)

# ---------------- 4. RESPALDO Y ELIMINACION DE FILAS COMPLETAS ----------------
datos_eliminados = datos.loc[filas_a_eliminar].copy()
datos_eliminados.insert(0, "Fila_Excel_original", datos_eliminados.index + 2)
datos_eliminados["Variables_que_superan_6DS"] = [
    ", ".join(fuera_de_limites.columns[fuera_de_limites.loc[i] &
                                     fuera_de_limites.columns.isin(columnas_fisicas)])
    for i in datos.index[filas_a_eliminar]
]
datos_eliminados["Mayor_z_absoluto"] = (
    z_scores.loc[filas_a_eliminar, columnas_fisicas].abs().max(axis=1).values
)

# Comprobacion independiente del umbral por cada valor retirado.
detalles = []
for i in datos.index[filas_a_eliminar]:
    for variable in columnas_fisicas:
        if fuera_de_limites.at[i, variable]:
            valor = float(datos.at[i, variable])
            media = float(medias[variable])
            ds = float(desviaciones[variable])
            z = float(z_scores.at[i, variable])
            assert ds > 0 and abs(z) > NUM_DESVIACIONES, (
                f"ERROR de verificacion en fila {i + 2}, variable {variable}."
            )
            detalles.append({
                "Fila_Excel_original": int(i + 2),
                "Variable": variable,
                "Valor_observado": valor,
                "Media": media,
                "Desviacion_estandar": ds,
                "Limite_inferior": float(limite_inferior[variable]),
                "Limite_superior": float(limite_superior[variable]),
                "Z_con_signo": z,
                "Z_absoluto": abs(z),
                "Verificacion": "FUERA DE ±6 DS",
            })

datos_limpios = datos.loc[~filas_a_eliminar].reset_index(drop=True)
assert len(datos_limpios) + len(datos_eliminados) == len(datos)
assert not fuera_de_limites.loc[~filas_a_eliminar, columnas_fisicas].any().any()

# ---------------- 5. ESTADISTICAS ANTES Y DESPUES ----------------
estadisticas_antes = pd.DataFrame({
    "Variable": columnas,
    "Cantidad": len(datos),
    "Media": medias.to_numpy(),
    "Desviacion_estandar": desviaciones.to_numpy(),
    "Limite_inferior_6DS": limite_inferior.to_numpy(),
    "Limite_superior_6DS": limite_superior.to_numpy(),
    "Minimo": datos.min().to_numpy(),
    "Maximo": datos.max().to_numpy(),
    "Valores_fuera_6DS": fuera_de_limites.sum().to_numpy(),
    "Usada_para_eliminar": ["SI" if c in columnas_fisicas else "NO - contador" for c in columnas],
})

estadisticas_despues = pd.DataFrame({
    "Variable": columnas,
    "Cantidad": len(datos_limpios),
    "Media": datos_limpios.mean().to_numpy(),
    "Desviacion_estandar": datos_limpios.std(ddof=1).to_numpy(),
    "Minimo": datos_limpios.min().to_numpy(),
    "Maximo": datos_limpios.max().to_numpy(),
})

resumen = pd.DataFrame({
    "Indicador": [
        "Archivo de origen", "Regla", "Metodo de dispersion", "Se uso RIC/IQR",
        "Filas iniciales", "Filas eliminadas", "Filas finales",
        "Porcentaje eliminado", "Iteraciones de la regla", "Columnas excluidas de la eliminacion",
    ],
    "Resultado": [
        archivo_entrada.name, "|z| > 6: borrar la fila completa",
        "Desviacion estandar muestral (ddof=1)", "NO",
        len(datos), len(datos_eliminados), len(datos_limpios),
        100 * len(datos_eliminados) / len(datos),
        1, ", ".join(sorted(COLUMNAS_NO_FISICAS)),
    ],
})

# ---------------- 6. GUARDAR LOS DOS ARCHIVOS ----------------
try:
    datos_limpios.to_excel(ARCHIVO_LIMPIO, sheet_name="Datos", index=False, engine="openpyxl")
    with pd.ExcelWriter(ARCHIVO_REPORTE, engine="openpyxl") as writer:
        resumen.to_excel(writer, sheet_name="Resumen", index=False)
        estadisticas_antes.to_excel(writer, sheet_name="Estadisticas_antes", index=False)
        estadisticas_despues.to_excel(writer, sheet_name="Estadisticas_despues", index=False)
        datos_eliminados.to_excel(writer, sheet_name="Filas_eliminadas", index=False)
        pd.DataFrame(detalles).to_excel(writer, sheet_name="Verificacion_6DS", index=False)
except PermissionError as exc:
    raise PermissionError(
        "Cierra los Excel de salida si estan abiertos y vuelve a ejecutar el programa."
    ) from exc

# ---------------- 7. MOSTRAR RESULTADOS EN LA TERMINAL ----------------
print("\n====== RESULTADOS DE LA REGLA 6 DS ======")
for _, fila in estadisticas_antes.iterrows():
    print(
        f"{fila['Variable']:<22} DS={fila['Desviacion_estandar']:.8f} "
        f"LI={fila['Limite_inferior_6DS']:.8f} "
        f"LS={fila['Limite_superior_6DS']:.8f} "
        f"Fuera={int(fila['Valores_fuera_6DS'])}"
    )
print(f"\nFilas antes:     {len(datos):,}")
print(f"Filas borradas:  {len(datos_eliminados):,}")
print(f"Filas despues:  {len(datos_limpios):,}")
print("\nArchivo para GRAFICOS:", ARCHIVO_LIMPIO.name)
print("Archivo de VERIFICACION:", ARCHIVO_REPORTE.name)
print("Proceso terminado: NO se utilizo RIC/IQR.")
