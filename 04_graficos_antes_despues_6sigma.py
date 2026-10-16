# -*- coding: utf-8 -*-
"""Gráficos triaxiales antes/después de ±6σ, basados en 03_tratamiento_6sigma.py.

EJECUCIÓN:
    python 04_graficos_antes_despues_6sigma.py

ENTRADAS (rutas reales que utiliza/genera 03_tratamiento_6sigma.py):
    Resultados_Limpieza_Inicial/CSV/datos_limpios.csv
    Resultados_6Sigma/CSV/datos_tratados_6sigma.csv

SALIDA: 6 imágenes PNG (300 dpi), sin archivos PDF.
    Resultados_6Sigma/Graficos_Triaxiales/Antes_6Sigma/
    Resultados_6Sigma/Graficos_Triaxiales/Despues_6Sigma/
    Resultados_6Sigma/Graficos_Triaxiales/Revision_SMP/
       revision_smp_discontinuidades_antes.csv
       revision_smp_discontinuidades_despues.csv
       resumen_continuidad_smp.csv

La etapa 6σ retira FILAS COMPLETAS. Por ello se conserva en los gráficos
el número de registro original de cada fila superviviente; los huecos no
se comprimen. El eje NO se etiqueta como tiempo en segundos.
"""
from __future__ import annotations

from collections import defaultdict, deque
from pathlib import Path
import csv
import sys

try:
    import numpy as np
    import matplotlib
    matplotlib.use("Agg", force=True)
    import matplotlib.pyplot as plt
    from matplotlib.ticker import MaxNLocator
except ImportError as exc:
    raise SystemExit(
        f"Falta la biblioteca {exc.name}. Instala con: "
        f'"{sys.executable}" -m pip install numpy matplotlib'
    ) from exc

BASE = Path(__file__).resolve().parent
ANTES = BASE / "Resultados_Limpieza_Inicial" / "CSV" / "datos_limpios.csv"
DESPUES = BASE / "Resultados_6Sigma" / "CSV" / "datos_tratados_6sigma.csv"
SALIDA = BASE / "Resultados_6Sigma" / "Graficos_Triaxiales"

# Usa True para comparar con idénticos límites verticales.
# False permite ver más detalle DESPUÉS, pero las escalas ya no coincidirán.
MISMA_ESCALA_Y = True

SENSORES = {
    "Acelerometro": (
        "AcelX1_SisGenEol", "AcelY1_SisGenEol", "AcelZ1_SisGenEol"
    ),
    "Giroscopio": (
        "GiroX_SisGenEol", "GiroY_SisGenEol", "GiroZ_SisGenEol"
    ),
    "Magnetometro": (
        "MagneX_SisGenEol", "MagneY_SisGenEol", "MagneZ_SisGenEol"
    ),
}
NOMBRES = {
    "Acelerometro": "Acelerómetro",
    "Giroscopio": "Giroscopio",
    "Magnetometro": "Magnetómetro",
}
COLORES = {"X": "#2676A8", "Y": "#D95C58", "Z": "#238D78"}


def leer_datos(ruta: Path):
    if not ruta.is_file():
        raise FileNotFoundError(f"No se encontró:\n  {ruta}")
    with ruta.open("r", encoding="utf-8-sig", newline="") as archivo:
        columnas = next(csv.reader(archivo))
    datos = np.loadtxt(
        ruta, delimiter=",", skiprows=1, encoding="utf-8-sig",
        dtype=np.float64, ndmin=2
    )
    if datos.size == 0:
        raise ValueError(f"El archivo está vacío: {ruta}")
    if datos.shape[1] != len(columnas):
        raise ValueError(f"Número de columnas inconsistente en {ruta}")
    if not np.isfinite(datos).all():
        raise ValueError(f"Hay valores vacíos, no numéricos o infinitos en {ruta}")
    return columnas, datos


def alinear_posiciones(antes: np.ndarray, despues: np.ndarray, indice_smp: int) -> np.ndarray:
    """Recupera posiciones ORIGINALES usando smp y verificando cada fila.

    03_tratamiento_6sigma.py no cambia valores ni orden, solo elimina filas.
    """
    a = antes[:, indice_smp]
    b = despues[:, indice_smp]
    if len(np.unique(a)) == len(a):
        # Caso normal: contador smp único en cada registro.
        mapa = {valor: i for i, valor in enumerate(a)}
        try:
            indices = np.array([mapa[valor] for valor in b], dtype=np.int64)
        except KeyError as exc:
            raise ValueError("El archivo tratado contiene smp ausentes en el limpio.") from exc
    else:
        # Si hay smp repetidos, usamos orden de aparición y validamos filas.
        candidatos = defaultdict(deque)
        for i, valor in enumerate(a):
            candidatos[valor].append(i)
        indices = []
        for fila in despues:
            opciones = candidatos.get(fila[indice_smp])
            if not opciones:
                raise ValueError("No se pudo ubicar una fila tratada en la base limpia.")
            hallado = None
            while opciones:
                pos = opciones.popleft()
                if np.allclose(antes[pos], fila, atol=1e-12, rtol=1e-12):
                    hallado = pos
                    break
            if hallado is None:
                raise ValueError("No se pudieron emparejar registros con smp repetido.")
            indices.append(hallado)
        indices = np.asarray(indices, dtype=np.int64)

    if not np.array_equal(indices, np.sort(indices)):
        raise ValueError("La secuencia tratada no conserva el orden de las filas originales.")
    if not np.allclose(antes[indices], despues, atol=1e-12, rtol=1e-12):
        raise ValueError(
            "El CSV tratado tiene valores distintos de los de la base limpia; "
            "no corresponde a una simple eliminación de filas por 6σ."
        )
    return indices + 1  # Índice de registro empezando en 1.


def limites_y(serie: np.ndarray) -> tuple[float, float]:
    mi = float(np.min(serie))
    ma = float(np.max(serie))
    margen = max((ma - mi) * .05, abs(mi) * .005, 1e-8)
    return mi - margen, ma + margen


def crear_figura(datos, x, indices_columnas, sensor, etapa, destino, n_original, y_limits=None):
    fig, ax = plt.subplots(figsize=(13.5, 5.0), facecolor="white")
    fig.subplots_adjust(left=.085, right=.985, bottom=.17, top=.77)
    for eje, indice in zip("XYZ", indices_columnas):
        ax.plot(
            x, datos[:, indice], lw=.45, alpha=.75, color=COLORES[eje],
            label=f"{eje}", rasterized=True
        )
    nombre = NOMBRES[sensor]
    fig.suptitle(f"{nombre} — componentes X, Y y Z", x=.085, y=.97,
                 ha="left", fontsize=14, fontweight="bold", color="#202b33")
    fig.text(.085, .895,
             f"{etapa} · {len(datos):,} registros · datos sin suavizar · "
             f"posición original (base: {n_original:,} registros)",
             color="#637381", fontsize=9.5)
    ax.set_xlabel("Número de registro original (no representa segundos)", fontsize=10)
    ax.set_ylabel("Valor registrado (unidad por confirmar)", fontsize=10)
    ax.set_xlim(1, n_original)
    if y_limits is not None:
        ax.set_ylim(*y_limits)
    ax.grid(axis="y", color="#DCE1E7", linewidth=.7, alpha=.7)
    ax.grid(axis="x", color="#EBEEF1", linewidth=.5, alpha=.65)
    ax.spines[["right", "top"]].set_visible(False)
    ax.spines[["left", "bottom"]].set_color("#AEB8C2")
    ax.tick_params(labelsize=9, colors="#53606D")
    ax.xaxis.set_major_locator(MaxNLocator(nbins=8))
    ax.legend(loc="upper center", bbox_to_anchor=(.5, 1.16),
              frameon=False, ncol=3, fontsize=10)
    destino.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(destino, dpi=300, bbox_inches="tight", pad_inches=.15,
                facecolor="white")
    plt.close(fig)
    print(f"  ✓ {destino.relative_to(BASE)}", flush=True)


def clasificar_paso(diferencia: float) -> str:
    """Clasifica incrementos; jamás reordena la secuencia."""
    if np.isclose(diferencia, 1.0, rtol=0, atol=1e-8):
        return "Correcto (+1)"
    if np.isclose(diferencia, 0.0, rtol=0, atol=1e-8):
        return "Repetido (0)"
    if diferencia < 0:
        return "Retroceso o reinicio (<0)"
    if diferencia > 1:
        return "Salto hacia adelante (>1)"
    return "Incremento positivo distinto de 1"


def guardar_tabla_csv(ruta: Path, encabezados: list[str], filas: list[dict]) -> None:
    """CSV legible en Excel (UTF-8 con BOM), incluso sin irregularidades."""
    ruta.parent.mkdir(parents=True, exist_ok=True)
    with ruta.open("w", newline="", encoding="utf-8-sig") as salida:
        writer = csv.DictWriter(salida, fieldnames=encabezados)
        writer.writeheader()
        writer.writerows(filas)


def revisar_smp_detallado(originales, tratados, columnas, posiciones_despues):
    """Revisa smp antes/después e identifica saltos atribuibles a filas retiradas.

    - Antes: pares consecutivos reales en la base limpia.
    - Después: pares consecutivos de la base tratada, preservando su posición
      de origen. Elimina una fila => diferencia de posición > 1.
    - Diferencia de smp distinta de +1 NO prueba una pérdida de adquisición.
    """
    indice = columnas.index("smp")
    smp_antes = originales[:, indice]
    smp_despues = tratados[:, indice]
    n_original = len(originales)
    pos_antes = np.arange(1, n_original + 1, dtype=np.int64)
    pos_despues = posiciones_despues.astype(np.int64)
    directorio = SALIDA / "Revision_SMP"

    encabezados = [
        "Etapa", "Fila_CSV_anterior", "Fila_CSV_actual",
        "Posicion_original_anterior", "Posicion_original_actual",
        "smp_anterior", "smp_actual", "diferencia_smp",
        "tipo_diferencia", "filas_retiradas_entre_registros",
        "origen_probable", "nota",
    ]
    resumen_encabezados = [
        "Etapa", "Registros", "Transiciones", "Pasos_igual_1",
        "Saltos_mayores_1", "Repetidos", "Retrocesos_reinicios",
        "Otros_incrementos", "Registros_con_smp_no_entero",
        "Transiciones_con_filas_retiradas", "Irregularidades_reportadas",
    ]
    resumen = []

    for etapa, smp, posiciones in (
        ("ANTES_6SIGMA", smp_antes, pos_antes),
        ("DESPUES_6SIGMA", smp_despues, pos_despues),
    ):
        conteos = {
            "Correcto (+1)": 0,
            "Salto hacia adelante (>1)": 0,
            "Repetido (0)": 0,
            "Retroceso o reinicio (<0)": 0,
            "Incremento positivo distinto de 1": 0,
        }
        filas = []
        tramos_con_retiros = 0
        # Número de pasos atípicos acumulado entre posiciones originales.
        pasos_originales = np.diff(smp_antes)
        pasos_anormales_originales = ~np.isclose(pasos_originales, 1.0, atol=1e-8, rtol=0)
        prefijo_anomalias = np.concatenate(([0], np.cumsum(pasos_anormales_originales)))

        for j in range(1, len(smp)):
            paso = float(smp[j] - smp[j - 1])
            tipo = clasificar_paso(paso)
            conteos[tipo] += 1
            pos_ant = int(posiciones[j - 1])
            pos_act = int(posiciones[j])
            retiradas = pos_act - pos_ant - 1
            if retiradas > 0:
                tramos_con_retiros += 1

            if etapa == "ANTES_6SIGMA":
                if tipo == "Correcto (+1)":
                    continue
                origen = "Presente antes de 6sigma"
                nota = "Revisar origen del contador; no implica necesariamente pérdida de muestras."
            else:
                anomalias_en_tramo = int(prefijo_anomalias[pos_act - 1] -
                                         prefijo_anomalias[pos_ant - 1])
                if retiradas > 0 and anomalias_en_tramo == 0:
                    origen = "Explicable por filas retiradas en 6sigma"
                    nota = ("Todas las transiciones originales en este tramo eran +1; "
                            "la diferencia incluye registros eliminados.")
                elif retiradas > 0 and anomalias_en_tramo > 0:
                    origen = "Filas retiradas + discontinuidad preexistente"
                    nota = ("El tramo también contiene transiciones smp distintas "
                            "de +1 en la base limpia.")
                else:
                    origen = "Presente antes de 6sigma"
                    nota = "Los registros eran adyacentes también en la base limpia."
                # Incluso con diferencia smp +1, guardar el hueco provocado
                # por retirada de filas para mantener una auditoría transparente.
                if tipo == "Correcto (+1)" and retiradas == 0:
                    continue

            filas.append({
                "Etapa": etapa,
                "Fila_CSV_anterior": j + 1,    # encabezado en fila 1
                "Fila_CSV_actual": j + 2,
                "Posicion_original_anterior": pos_ant,
                "Posicion_original_actual": pos_act,
                "smp_anterior": f"{smp[j - 1]:.15g}",
                "smp_actual": f"{smp[j]:.15g}",
                "diferencia_smp": f"{paso:.15g}",
                "tipo_diferencia": tipo,
                "filas_retiradas_entre_registros": retiradas,
                "origen_probable": origen,
                "nota": nota,
            })

        no_enteros = int(np.sum(~np.isclose(smp, np.rint(smp), atol=1e-8, rtol=0)))
        resumen.append({
            "Etapa": etapa,
            "Registros": len(smp),
            "Transiciones": max(0, len(smp) - 1),
            "Pasos_igual_1": conteos["Correcto (+1)"],
            "Saltos_mayores_1": conteos["Salto hacia adelante (>1)"],
            "Repetidos": conteos["Repetido (0)"],
            "Retrocesos_reinicios": conteos["Retroceso o reinicio (<0)"],
            "Otros_incrementos": conteos["Incremento positivo distinto de 1"],
            "Registros_con_smp_no_entero": no_enteros,
            "Transiciones_con_filas_retiradas": tramos_con_retiros,
            "Irregularidades_reportadas": len(filas),
        })

        nombre = ("revision_smp_discontinuidades_antes.csv" if etapa == "ANTES_6SIGMA"
                  else "revision_smp_discontinuidades_despues.csv")
        ruta_csv = directorio / nombre
        guardar_tabla_csv(ruta_csv, encabezados, filas)
        print(f"\n=== CONTINUIDAD SMP: {etapa} ===")
        print(f"Registros: {len(smp):,} | inicio={smp[0]:.15g} | final={smp[-1]:.15g}")
        print(f"Pasos +1: {conteos['Correcto (+1)']:,}; "
              f"saltos >1: {conteos['Salto hacia adelante (>1)']:,}; "
              f"repeticiones: {conteos['Repetido (0)']:,}; "
              f"retrocesos: {conteos['Retroceso o reinicio (<0)']:,}; "
              f"otros: {conteos['Incremento positivo distinto de 1']:,}.")
        print(f"Transiciones con filas retiradas entre muestras: {tramos_con_retiros:,}")
        if no_enteros:
            print(f"AVISO: {no_enteros:,} valores smp no son enteros; revisar significado del campo.")
        print(f"Registros de auditoría guardados: {len(filas):,} -> {ruta_csv.relative_to(BASE)}")
        for fila in filas[:5]:
            print("  Posiciones originales "
                  f"{fila['Posicion_original_anterior']} -> {fila['Posicion_original_actual']}: "
                  f"smp {fila['smp_anterior']} -> {fila['smp_actual']} "
                  f"(Δ={fila['diferencia_smp']}); {fila['origen_probable']}")

    ruta_resumen = directorio / "resumen_continuidad_smp.csv"
    guardar_tabla_csv(ruta_resumen, resumen_encabezados, resumen)
    print(f"\nResumen comparativo guardado: {ruta_resumen.relative_to(BASE)}")
    print("Interpretación: un salto smp no prueba por sí solo pérdida de datos; "
          "tras 6σ puede deberse a eliminación de filas.")


def main():
    print("=== GRÁFICOS ANTES Y DESPUÉS DE 6σ ===", flush=True)
    columnas_antes, originales = leer_datos(ANTES)
    columnas_despues, tratados = leer_datos(DESPUES)
    if columnas_antes != columnas_despues:
        raise ValueError("Los encabezados de ambos CSV son diferentes.")
    if "smp" not in columnas_antes:
        raise ValueError("No se encontró la columna 'smp'.")
    for grupo in SENSORES.values():
        for nombre in grupo:
            if nombre not in columnas_antes:
                raise ValueError(f"Falta la columna {nombre!r} en {ANTES.name}.")
    posiciones_despues = alinear_posiciones(
        originales, tratados, columnas_antes.index("smp")
    )
    posiciones_antes = np.arange(1, len(originales) + 1)
    print(f"Base limpia: {len(originales):,} registros")
    print(f"Después de 6σ: {len(tratados):,} registros")
    print(f"Filas retiradas por 6σ: {len(originales) - len(tratados):,}")
    revisar_smp_detallado(originales, tratados, columnas_antes, posiciones_despues)
    print("\nGuardando únicamente PNG a 300 dpi:", flush=True)

    for n, (sensor, nombres_columnas) in enumerate(SENSORES.items(), start=1):
        indices = [columnas_antes.index(c) for c in nombres_columnas]
        limites = limites_y(originales[:, indices]) if MISMA_ESCALA_Y else None
        nombre = f"{n:02d}_{sensor.lower()}_xyz.png"
        crear_figura(
            originales, posiciones_antes, indices, sensor,
            "ANTES de 6σ (base limpia)", SALIDA / "Antes_6Sigma" / nombre,
            len(originales), limites
        )
        crear_figura(
            tratados, posiciones_despues, indices, sensor,
            "DESPUÉS de 6σ (base tratada)", SALIDA / "Despues_6Sigma" / nombre,
            len(originales), limites
        )
    print("\nLISTO: 6 imágenes PNG + 3 CSV de continuidad smp; sin archivos PDF.")
    print(f"Carpeta: {SALIDA}")


if __name__ == "__main__":
    try:
        main()
    except (FileNotFoundError, ValueError, OSError) as exc:
        print(f"\nERROR: {exc}", file=sys.stderr)
        print("Ejecuta primero 03_tratamiento_6sigma.py si aún no existe el CSV tratado.",
              file=sys.stderr)
        sys.exit(1)
