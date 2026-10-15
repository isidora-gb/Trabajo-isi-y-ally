# -*- coding: utf-8 -*-
"""
PASO 02 - LIMPIEZA INICIAL + DESCRIPTIVOS + GRAFICOS (SIN TRATAR ATIPICOS)

Ejecutar:
    python -m pip install pandas numpy openpyxl matplotlib
    python 02_limpiar_datos.py

Deja junto a este archivo el Excel original CF-2023-I-204-1230_DAT.xlsx.
Tambien puedes indicar la ruta como argumento.

Salidas organizadas en Resultados_Limpieza_Inicial/: 
    CSV/: 5 tablas CSV
    Excel/Resultados_limpieza.xlsx: 9 hojas simples
    Graficos_Paper_Profesional_Sin_Zoom/: 32 PNG individuales con estilo de paper (300 dpi)
    RESUMEN_TERMINAL.txt: el mismo resumen que se imprime en la terminal.

Este script corrige estructuras verificables e imputa valores faltantes; conserva
las filas y TODOS los valores observados. No calcula, marca ni filtra atipicos
por 6 sigma ni por IQR: esa etapa se realiza con otro script.

Notas: smp es un contador de muestras, no representa segundos; las unidades
fisicas y limites admisibles de los sensores no estan documentados aqui.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

try:
    import numpy as np
    import pandas as pd
    from openpyxl.utils import get_column_letter
    import matplotlib
    matplotlib.use("Agg", force=True)  # Sin ventanas emergentes
    import matplotlib.pyplot as plt
    from matplotlib.ticker import MaxNLocator
except ImportError:
    print("Faltan paquetes. Instala con: python -m pip install pandas numpy openpyxl matplotlib")
    raise SystemExit(1)

SENSORES = [
    "AcelX1_SisGenEol", "AcelY1_SisGenEol", "AcelZ1_SisGenEol",
    "GiroX_SisGenEol", "GiroY_SisGenEol", "GiroZ_SisGenEol",
    "MagneX_SisGenEol", "MagneY_SisGenEol", "MagneZ_SisGenEol",
    "Ter_SisGenEol",
]
COLUMNAS = SENSORES + ["smp"]
MODULO_SMP = 65536

REG_CAMPOS = [
    "tipo", "fila_excel", "columna", "valor_anterior", "valor_nuevo",
    "metodo", "detalle",
]
REAL_CAMPOS = [
    "fila_excel", "celdas_presentes", "columna_inicio_bloque",
    "columna_fin_bloque", "smp_anterior", "smp_recuperado",
    "smp_siguiente", "paso_anterior", "paso_siguiente",
    "fragmentos_conservados",
]
FRAG_CAMPOS = ["fila_excel", "celda_original", "valor_original", "observacion"]
SALTOS_CAMPOS = [
    "fila_excel_anterior", "fila_excel_actual", "smp_anterior",
    "smp_actual", "muestras_no_registradas",
]
REV_CAMPOS = ["fila_excel", "situacion", "estado"]


def buscar_entrada(carpeta: Path, nombre: str | None) -> Path:
    if nombre:
        archivo = Path(nombre).expanduser().resolve()
        if not archivo.is_file():
            raise FileNotFoundError(f"No existe el Excel indicado: {archivo}")
        return archivo
    posibles = sorted(
        p for p in carpeta.glob("CF-2023-I-204-1230_DAT*.xlsx")
        if not p.name.startswith("~$")
    )
    if len(posibles) != 1:
        raise FileNotFoundError(
            "Deja UN solo Excel original CF-2023-I-204-1230_DAT*.xlsx "
            "junto al script o indica su ruta: "
            'python 02_limpiar_datos.py "ruta/al/archivo.xlsx". '
            f"Encontrados: {[p.name for p in posibles]}"
        )
    return posibles[0]


def es_smp(v: float) -> bool:
    return bool(np.isfinite(v) and v.is_integer() and 0 <= v < MODULO_SMP)


def formato_valor(v: object) -> float | int | None:
    if v is None or pd.isna(v):
        return None
    return float(v)


def normalizar_original(original: pd.DataFrame) -> pd.DataFrame:
    if list(original.columns[:11]) != COLUMNAS:
        raise ValueError(
            "Los primeros 11 encabezados no coinciden con los esperados.\n"
            f"Esperados: {COLUMNAS}\n"
            f"Recibidos: {list(original.columns[:11])}"
        )
    extras = list(original.columns[11:])
    if any(not str(c).startswith("Unnamed:") for c in extras):
        raise ValueError(
            f"Hay columnas extras con nombre propio ({extras}); revisalas antes de limpiar."
        )

    limpio = original.replace(r"^\s*$", np.nan, regex=True).copy()
    for col in limpio.columns:
        try:
            limpio[col] = pd.to_numeric(limpio[col], errors="raise")
        except (ValueError, TypeError) as exc:
            raise ValueError(f"La columna {col!r} incluye texto que no es numerico.") from exc
    if np.isinf(limpio.to_numpy(dtype=float)).any():
        raise ValueError("El original contiene valores infinitos; revisalos manualmente.")
    return limpio


def recuperar_filas(
    tabla: pd.DataFrame,
) -> tuple[pd.DataFrame, list[dict], list[dict], list[dict], list[dict], list[dict]]:
    """Realinea bloques de 11 campos completos validados por contador smp."""
    matriz = tabla.to_numpy(dtype=float)
    anchos = np.isfinite(matriz).sum(axis=1)
    datos = tabla[COLUMNAS].copy()
    correcciones: list[dict] = []
    realineaciones: list[dict] = []
    fragmentos: list[dict] = []
    revisiones: list[dict] = []
    columnas = len(COLUMNAS)

    for fila in np.flatnonzero(anchos > columnas):
        ancho = int(anchos[fila])
        valores = matriz[fila]
        fila_excel = int(fila + 2)
        if not (np.isfinite(valores[:ancho]).all() and np.isnan(valores[ancho:]).all()):
            raise ValueError(f"Fila {fila_excel}: hay huecos entre los datos desalineados.")
        if not (12 <= ancho <= 21):
            raise ValueError(f"Fila {fila_excel}: {ancho} celdas, patron no contemplado.")
        if (
            fila == 0 or fila == len(matriz) - 1
            or anchos[fila - 1] != 11 or anchos[fila + 1] != 11
        ):
            raise ValueError(
                f"Fila {fila_excel}: no tiene dos filas vecinas completas para validarla."
            )

        inicio = ancho - columnas
        bloque = valores[inicio:ancho].copy()
        previo = float(matriz[fila - 1, 10])
        contador = float(bloque[-1])
        siguiente = float(matriz[fila + 1, 10])
        if not all(es_smp(v) for v in (previo, contador, siguiente)):
            raise ValueError(f"Fila {fila_excel}: no se puede verificar smp.")
        salto_izq = int((contador - previo) % MODULO_SMP)
        salto_der = int((siguiente - contador) % MODULO_SMP)
        if salto_izq not in (1, 2) or salto_der != 1:
            raise ValueError(
                f"Fila {fila_excel}: bloque desplazado no concuerda con smp vecino."
            )

        for j, columna in enumerate(COLUMNAS):
            antes = valores[j]
            nuevo = bloque[j]
            correcciones.append({
                "tipo": "realineacion",
                "fila_excel": fila_excel,
                "columna": columna,
                "valor_anterior": formato_valor(antes),
                "valor_nuevo": formato_valor(nuevo),
                "metodo": "Bloque completo de 11 campos validado por smp",
                "detalle": f"Origen: {get_column_letter(inicio + j + 1)}{fila_excel}",
            })
        datos.iloc[fila, :] = bloque
        realineaciones.append({
            "fila_excel": fila_excel,
            "celdas_presentes": ancho,
            "columna_inicio_bloque": get_column_letter(inicio + 1),
            "columna_fin_bloque": get_column_letter(ancho),
            "smp_anterior": int(previo),
            "smp_recuperado": int(contador),
            "smp_siguiente": int(siguiente),
            "paso_anterior": salto_izq,
            "paso_siguiente": salto_der,
            "fragmentos_conservados": inicio,
        })
        for j in range(inicio):
            fragmentos.append({
                "fila_excel": fila_excel,
                "celda_original": f"{get_column_letter(j + 1)}{fila_excel}",
                "valor_original": float(valores[j]),
                "observacion": "Fragmento previo al bloque verificado; no se uso para reconstruir",
            })
        revisiones.append({
            "fila_excel": fila_excel,
            "situacion": f"Fila con {ancho} valores en el original",
            "estado": "Realineada y validada; fragmentos auditados",
        })

    return datos, correcciones, realineaciones, fragmentos, revisiones, []


def completar_sensores(
    datos: pd.DataFrame, registro: list[dict], vecinos: int
) -> pd.DataFrame:
    for columna in SENSORES:
        # Tomar una COPIA original: una estimacion no alimenta otra estimacion.
        valores = datos[columna].to_numpy(dtype=float, copy=True)
        validos = np.flatnonzero(np.isfinite(valores))
        vacios = np.flatnonzero(~np.isfinite(valores))
        if len(vacios) and not len(validos):
            raise ValueError(f"{columna}: no hay ningun dato real para estimar faltantes.")
        for pos in vacios:
            # Solo hacen falta 'vecinos' a la izquierda y a la derecha para
            # encontrar los n datos de menor distancia global.
            punto = int(np.searchsorted(validos, pos))
            candidatos = validos[
                max(0, punto - vecinos):min(len(validos), punto + vecinos)
            ]
            elegidos = sorted(candidatos, key=lambda j: (abs(int(j) - int(pos)), int(j)))[:vecinos]
            estimado = float(np.mean(valores[elegidos]))
            datos.iat[int(pos), datos.columns.get_loc(columna)] = estimado
            registro.append({
                "tipo": "imputacion_sensor",
                "fila_excel": int(pos + 2),
                "columna": columna,
                "valor_anterior": None,
                "valor_nuevo": estimado,
                "metodo": f"Promedio de {len(elegidos)} vecinos observados mas cercanos",
                "detalle": "Filas Excel: " + ", ".join(str(int(j) + 2) for j in elegidos),
            })
    return datos


def validar_y_completar_smp(
    datos: pd.DataFrame, registro: list[dict]
) -> tuple[pd.DataFrame, list[dict], list[int]]:
    smp = datos["smp"].to_numpy(dtype=float, copy=True)
    indices_vacios = np.flatnonzero(np.isnan(smp))
    if len(indices_vacios):
        ultimo = len(smp) - 1
        if len(indices_vacios) != 1 or int(indices_vacios[0]) != ultimo or ultimo < 2:
            raise ValueError(
                "smp tiene uno o mas vacios fuera de la ultima fila. "
                "No se estima el contador sin verificacion."
            )
        a, b = smp[-3], smp[-2]
        if not (es_smp(float(a)) and es_smp(float(b)) and int((b - a) % MODULO_SMP) == 1):
            raise ValueError("Ultimo smp vacio: los dos anteriores no permiten inferirlo.")
        nuevo = int((b + 1) % MODULO_SMP)
        smp[-1] = nuevo
        datos.iat[ultimo, datos.columns.get_loc("smp")] = nuevo
        registro.append({
            "tipo": "inferencia_contador",
            "fila_excel": ultimo + 2,
            "columna": "smp",
            "valor_anterior": None,
            "valor_nuevo": nuevo,
            "metodo": "Contador anterior + 1 (modulo 65536)",
            "detalle": f"smp anteriores: {int(a)}, {int(b)}",
        })

    if not all(es_smp(float(v)) for v in smp):
        raise ValueError("smp contiene fracciones o valores fuera de 0..65535.")

    saltos = []
    reinicios = []
    for i in range(1, len(smp)):
        anterior, actual = int(smp[i - 1]), int(smp[i])
        salto = (actual - anterior) % MODULO_SMP
        if salto not in (1, 2):
            raise ValueError(
                f"Fila Excel {i + 2}: smp tiene un salto inesperado {anterior} -> {actual}."
            )
        if actual < anterior:
            if anterior != MODULO_SMP - 1 or actual != 0:
                raise ValueError(f"Fila Excel {i + 2}: reinicio de smp no explicable.")
            reinicios.append(i + 2)
        if salto == 2:
            saltos.append({
                "fila_excel_anterior": i + 1,
                "fila_excel_actual": i + 2,
                "smp_anterior": anterior,
                "smp_actual": actual,
                "muestras_no_registradas": 1,
            })
    datos["smp"] = smp.astype("int64")
    return datos, saltos, reinicios




def calcular_estadisticos(datos: pd.DataFrame) -> pd.DataFrame:
    """Media, DS muestral y extremos por variable, ignorando solo los NaN.

    La tabla 'antes' se calcula DESPUES de realinear filas (para comparar
    las mismas variables) pero ANTES de imputar celdas vacias.
    """
    filas = []
    for columna in COLUMNAS:
        serie = pd.to_numeric(datos[columna], errors="raise")
        conteo = int(serie.notna().sum())
        faltantes = int(serie.isna().sum())
        filas.append({
            "Variable": columna,
            "Tipo": str(datos[columna].dtype),
            "Cantidad": int(len(serie)),
            "Validos": conteo,
            "Nulos": faltantes,
            "Nulos_pct": round(100 * faltantes / len(serie), 6) if len(serie) else np.nan,
            "Media": float(serie.mean()) if conteo else np.nan,
            "Mediana": float(serie.median()) if conteo else np.nan,
            "DS_muestral": float(serie.std(ddof=1)) if conteo > 1 else np.nan,
            "Minimo": float(serie.min()) if conteo else np.nan,
            "Maximo": float(serie.max()) if conteo else np.nan,
            "Q1": float(serie.quantile(0.25)) if conteo else np.nan,
            "Q3": float(serie.quantile(0.75)) if conteo else np.nan,
        })
    return pd.DataFrame(filas)


def mostrar_tabla_estadisticos(titulo: str, estadisticos: pd.DataFrame) -> None:
    """Mostrar en terminal una tabla compacta con datos pedidos por la pauta."""
    print("\n" + titulo)
    campos = ["Variable", "Validos", "Nulos", "Media", "DS_muestral", "Minimo", "Maximo"]
    vista = estadisticos[campos].copy()
    print(vista.to_string(index=False, float_format=lambda x: f"{x:,.6g}"))


def escribir_excel(destino: Path, hojas: dict[str, pd.DataFrame]) -> None:
    """Tablas sencillas: encabezados y datos, sin graficos ni disenos complejos."""
    if any(len(df) >= 1_048_576 for df in hojas.values()):
        raise ValueError("El Excel no admite tantas filas; usa los CSV.")
    with pd.ExcelWriter(destino, engine="openpyxl") as escritor:
        for nombre, tabla in hojas.items():
            tabla.to_excel(escritor, sheet_name=nombre, index=False)
            hoja = escritor.sheets[nombre]
            hoja.freeze_panes = "A2"
            hoja.auto_filter.ref = hoja.dimensions
            # Anchos fijos sencillos; no se recorren cientos de miles de celdas.
            for i, encabezado in enumerate(tabla.columns, start=1):
                hoja.column_dimensions[get_column_letter(i)].width = min(32, max(16, len(str(encabezado)) + 2))





# ------------------ FIGURAS CIENTIFICAS: SIN ZOOM ------------------
# Estilo sobrio de paper: sin letras A/B/C, fondo blanco, ejes simples,
# colores diferenciados y líneas de referencia claramente visibles.
from matplotlib.ticker import MaxNLocator
from matplotlib.colors import to_rgb, to_hex

PALETA = {
    "AcelX1_SisGenEol": "#2F65B0",
    "AcelY1_SisGenEol": "#1598A6",
    "AcelZ1_SisGenEol": "#7A59B5",
    "GiroX_SisGenEol": "#C76B2C",
    "GiroY1_SisGenEol": "#2D8A5F",
    "GiroY_SisGenEol": "#2D8A5F",
    "GiroZ_SisGenEol": "#B14662",
    "MagneX_SisGenEol": "#A68A22",
    "MagneY_SisGenEol": "#2D80BF",
    "MagneZ_SisGenEol": "#9B4F93",
    "Ter_SisGenEol": "#D55A4E",
}
FONDO = "#FFFFFF"
TEXTO = "#20252B"
SUAVE = "#6E7781"
REJILLA = "#E7EAEE"
COLOR_MEDIA = "#1F2933"   # gris oscuro casi negro
COLOR_MEDIANA = "#C43C39" # rojo oscuro para diferenciarla de la media


def nombre_corto(nombre: str) -> str:
    return nombre.replace("_SisGenEol", "")


def numero_breve(valor: float) -> str:
    if not np.isfinite(valor):
        return "N/D"
    if valor == 0:
        return "0"
    if abs(valor) < 1e-3 or abs(valor) >= 1e4:
        return f"{valor:.3e}"
    return f"{valor:.5g}"


def aclarar_color(color: str, mezcla: float = 0.55) -> str:
    """Mezcla el color con blanco para suavizar la señal principal."""
    base = np.array(to_rgb(color))
    blanco = np.array([1.0, 1.0, 1.0])
    nuevo = (1 - mezcla) * base + mezcla * blanco
    return to_hex(np.clip(nuevo, 0, 1))


def configurar_estilo() -> None:
    plt.ioff()
    plt.rcParams.update({
        "font.family": "DejaVu Sans",
        "font.size": 10,
        "axes.titlesize": 12,
        "axes.titleweight": "bold",
        "axes.labelsize": 10,
        "xtick.labelsize": 9,
        "ytick.labelsize": 9,
        "legend.fontsize": 9,
        "axes.linewidth": 0.8,
        "figure.facecolor": FONDO,
        "axes.facecolor": FONDO,
        "savefig.facecolor": FONDO,
        "savefig.dpi": 300,
        "pdf.fonttype": 42,
        "ps.fonttype": 42,
        "axes.spines.top": False,
        "axes.spines.right": False,
        "axes.formatter.useoffset": False,
        "axes.unicode_minus": True,
        "agg.path.chunksize": 20000,
    })


def estilo_ejes(ax, eje_grid="y") -> None:
    ax.set_facecolor(FONDO)
    ax.spines["left"].set_color("#AEB7C1")
    ax.spines["bottom"].set_color("#AEB7C1")
    ax.tick_params(axis="both", colors=SUAVE, width=0.8, length=3.5)
    ax.set_axisbelow(True)
    ax.grid(axis=eje_grid, color=REJILLA, linewidth=0.6)
    ax.xaxis.set_major_locator(MaxNLocator(nbins=6))
    ax.yaxis.set_major_locator(MaxNLocator(nbins=6))


def figura_paper(titulo: str, subtitulo: str):
    fig = plt.figure(figsize=(7.1, 5.0), facecolor=FONDO)
    ax = fig.add_axes([0.14, 0.21, 0.79, 0.66])
    fig.text(0.14, 0.935, titulo, fontsize=13, fontweight="bold", color=TEXTO)
    fig.text(0.14, 0.895, subtitulo, fontsize=9.5, color=SUAVE)
    return fig, ax


def pie_figura(fig, texto: str) -> None:
    fig.text(0.14, 0.08, texto, fontsize=8.4, color=SUAVE)


def guardar_figura(fig, carpeta: Path, nombre: str) -> Path:
    carpeta.mkdir(parents=True, exist_ok=True)
    ruta = carpeta / nombre
    fig.savefig(ruta, dpi=300, bbox_inches="tight", pad_inches=0.08, facecolor=FONDO)
    plt.close(fig)
    return ruta


def _subtitulo_resumen(serie: np.ndarray) -> str:
    return (
        f"n = {len(serie):,}   Media = {numero_breve(float(np.mean(serie)))}   "
        f"DE = {numero_breve(float(np.std(serie, ddof=1)))}   "
        f"Mediana = {numero_breve(float(np.median(serie)))}"
    )


def grafico_histograma(serie, variable, carpeta, indice):
    del indice
    color = PALETA[variable]
    color_barras = aclarar_color(color, 0.18)
    fig, ax = figura_paper(
        f"Distribución de {nombre_corto(variable)}",
        _subtitulo_resumen(serie),
    )
    minimo, maximo = float(np.min(serie)), float(np.max(serie))
    if minimo == maximo:
        margen = max(abs(minimo) * 0.05, 1e-6)
        bordes = np.linspace(minimo - margen, maximo + margen, 30)
    else:
        bordes = np.linspace(minimo, maximo, 50)
    ax.hist(
        serie, bins=bordes, color=color_barras, alpha=0.92,
        edgecolor="white", linewidth=0.45, label="Frecuencia"
    )
    ax.axvline(float(np.mean(serie)), color=COLOR_MEDIA, linestyle="--", linewidth=1.45,
               label="Media", zorder=4)
    ax.axvline(float(np.median(serie)), color=COLOR_MEDIANA, linestyle="-", linewidth=1.65,
               label="Mediana", zorder=5)
    ax.set_xlabel("Valor medido")
    ax.set_ylabel("Frecuencia")
    ax.legend(frameon=False, loc="upper right")
    estilo_ejes(ax, "y")
    pie_figura(fig, "Histograma sin recortes ni zoom; se muestra el rango completo de los datos.")
    return guardar_figura(fig, carpeta, "01_Histograma.png")


def grafico_caja(serie, variable, carpeta, indice):
    del indice
    color = PALETA[variable]
    fig, ax = figura_paper(
        f"Resumen de dispersión de {nombre_corto(variable)}",
        _subtitulo_resumen(serie),
    )
    q1 = float(np.quantile(serie, 0.25))
    q3 = float(np.quantile(serie, 0.75))
    iqr = q3 - q1
    limite_inf = q1 - 1.5 * iqr
    limite_sup = q3 + 1.5 * iqr
    atipicos = serie[(serie < limite_inf) | (serie > limite_sup)]

    ax.boxplot(
        [serie], vert=True, widths=0.26, patch_artist=True, showfliers=False, showmeans=True,
        boxprops={"facecolor": color, "alpha": 0.23, "edgecolor": color, "linewidth": 1.6},
        medianprops={"color": COLOR_MEDIANA, "linewidth": 1.9},
        meanprops={"marker": "D", "markerfacecolor": COLOR_MEDIA, "markeredgecolor": COLOR_MEDIA,
                   "markersize": 4.2},
        whiskerprops={"color": color, "linewidth": 1.2},
        capprops={"color": color, "linewidth": 1.2},
    )

    if len(atipicos):
        rng = np.random.default_rng(12345)
        jitter = rng.uniform(-0.10, 0.10, size=len(atipicos))
        ax.scatter(
            1.0 + jitter, atipicos,
            s=8, color=color, alpha=0.32, edgecolors='none', rasterized=True,
            label='Puntos fuera de 1.5×RIC', zorder=3,
        )

    minimo, maximo = float(np.min(serie)), float(np.max(serie))
    margen = max((maximo - minimo) * 0.05, abs(minimo) * 1e-6, 1e-12)
    ax.set_ylim(minimo - margen, maximo + margen)
    ax.set_xlim(0.72, 1.28)
    ax.set_xticks([1], [nombre_corto(variable)])
    ax.set_ylabel("Valor medido")
    if len(atipicos):
        ax.legend(frameon=False, loc='upper right')
    estilo_ejes(ax, "y")
    pie_figura(fig, f"Sin zoom: se muestra el rango completo min–max. Los puntos se dispersan horizontalmente para que se distingan (n fuera de 1.5×RIC = {len(atipicos):,}).")
    return guardar_figura(fig, carpeta, "02_Caja.png")


def grafico_evolucion(serie, variable, carpeta, indice):
    del indice
    color = PALETA[variable]
    color_senal = aclarar_color(color, 0.22)
    fig, ax = figura_paper(
        f"Comportamiento por muestra de {nombre_corto(variable)}",
        _subtitulo_resumen(serie),
    )
    orden = np.arange(1, len(serie) + 1)
    ax.plot(orden, serie, color=color_senal, linewidth=0.30, alpha=0.75,
            rasterized=True, label="Señal registrada", zorder=1)
    ax.axhline(float(np.mean(serie)), color=COLOR_MEDIA, linestyle="--", linewidth=1.55,
               label="Media", zorder=4)
    ax.axhline(float(np.median(serie)), color=COLOR_MEDIANA, linestyle="-", linewidth=1.7,
               label="Mediana", zorder=5)
    ax.set_xlim(1, len(serie))
    ax.set_xlabel("Orden de muestra")
    ax.set_ylabel("Valor medido")
    ax.ticklabel_format(axis="x", style="sci", scilimits=(4, 4))
    ax.legend(frameon=False, loc="upper right")
    estilo_ejes(ax, "y")
    pie_figura(fig, "Serie completa sin suavizado, sin diezmado y sin recortes; smp representa orden de muestra.")
    return guardar_figura(fig, carpeta, "03_Evolucion.png")


def grafico_correlaciones(datos, carpeta):
    matriz = datos[SENSORES].corr(method="pearson").to_numpy(dtype=float)
    fig = plt.figure(figsize=(7.8, 6.7), facecolor=FONDO)
    ax = fig.add_axes([0.17, 0.18, 0.62, 0.68])
    fig.text(0.17, 0.935, "Matriz de correlación de Pearson", fontsize=13, fontweight="bold", color=TEXTO)
    fig.text(0.17, 0.895, "Relación lineal entre sensores limpios; smp se excluye por ser contador.", fontsize=9.5, color=SUAVE)
    imagen = ax.imshow(matriz, cmap="RdBu_r", vmin=-1, vmax=1, interpolation="nearest")
    etiquetas = [nombre_corto(c) for c in SENSORES]
    ax.set_xticks(range(len(SENSORES)), labels=etiquetas, rotation=45, ha="right")
    ax.set_yticks(range(len(SENSORES)), labels=etiquetas)
    ax.tick_params(length=0, colors=SUAVE)
    for i in range(len(SENSORES)):
        for j in range(len(SENSORES)):
            valor = matriz[i, j]
            ax.text(j, i, f"{valor:.2f}", ha="center", va="center", fontsize=7.7,
                    color="white" if abs(valor) >= 0.55 else TEXTO)
    cax = fig.add_axes([0.82, 0.23, 0.025, 0.58])
    cb = fig.colorbar(imagen, cax=cax)
    cb.set_label("r", color=TEXTO, fontsize=10)
    cb.outline.set_visible(False)
    pie_figura(fig, "La correlación cuantifica asociación lineal y no implica causalidad.")
    return guardar_figura(fig, carpeta, "01_Correlaciones.png")


def grafico_nulos(antes, despues, carpeta):
    previos = antes.set_index("Variable").loc[COLUMNAS, "Nulos"].astype(int).to_numpy()
    posteriores = despues.set_index("Variable").loc[COLUMNAS, "Nulos"].astype(int).to_numpy()
    pos = np.arange(len(COLUMNAS))
    fig, ax = figura_paper(
        "Valores faltantes por variable",
        "Comparación antes y después de la limpieza inicial.",
    )
    ax.bar(pos - 0.18, previos, width=0.36, color="#8C8C8C", label="Antes")
    ax.bar(pos + 0.18, posteriores, width=0.36, color="#2F65B0", label="Después")
    ax.set_xticks(pos, [nombre_corto(c) for c in COLUMNAS], rotation=45, ha="right")
    ax.set_ylabel("Cantidad de valores faltantes")
    ax.legend(frameon=False, loc="upper right")
    estilo_ejes(ax, "y")
    pie_figura(fig, "La limpieza imputa faltantes verificables; no elimina registros.")
    return guardar_figura(fig, carpeta, "02_Nulos_antes_despues.png")


def generar_graficos(datos: pd.DataFrame, antes: pd.DataFrame,
                     despues: pd.DataFrame, correlaciones: pd.DataFrame,
                     carpeta: Path) -> list[Path]:
    """Genera 30 figuras individuales + 2 generales, sin zoom y sin 6 sigma."""
    del correlaciones
    configurar_estilo()
    carpeta.mkdir(parents=True, exist_ok=True)
    rutas = []
    indice = 0
    for variable in SENSORES:
        carpeta_variable = carpeta / variable
        serie = datos[variable].to_numpy(dtype=float)
        if not np.isfinite(serie).all():
            raise ValueError(f"No se pueden graficar valores no finitos en {variable}")
        indice += 1
        rutas.append(grafico_histograma(serie, variable, carpeta_variable, indice))
        indice += 1
        rutas.append(grafico_caja(serie, variable, carpeta_variable, indice))
        indice += 1
        rutas.append(grafico_evolucion(serie, variable, carpeta_variable, indice))
        print(f"   [OK] {nombre_corto(variable):8s} · 3 graficos · mejor contraste", flush=True)
    general = carpeta / "00_Generales"
    rutas.append(grafico_correlaciones(datos, general))
    rutas.append(grafico_nulos(antes, despues, general))
    print("   [OK] Correlaciones y nulos · 2 graficos", flush=True)
    indice_tabla = pd.DataFrame([{
        "Variable": p.parent.name,
        "Tipo": p.stem.split("_", 1)[-1],
        "Archivo": str(p.relative_to(carpeta)).replace("\\", "/"),
        "Resolucion_dpi": 300,
        "Zoom": "NO",
        "Estilo": "Paper mejorado",
    } for p in rutas])
    indice_tabla.to_csv(carpeta / "INDICE_GRAFICOS.csv", index=False, encoding="utf-8-sig")
    return rutas

def main() -> None:
    p = argparse.ArgumentParser(description="Limpieza inicial del Excel eolico + estadisticos + figuras; NO 6sigma")
    p.add_argument("archivo", nargs="?", help="Excel original; por defecto se detecta junto al .py")
    p.add_argument("--salida", help="Carpeta para resultados, por defecto Resultados_Limpieza_Inicial")
    p.add_argument("--vecinos", type=int, default=2, help="Numero de vecinos para estimar sensores")
    p.add_argument("--solo-graficos", action="store_true", help="Usa el CSV limpio existente para rehacer los graficos, sin leer el Excel original")
    args = p.parse_args()
    if args.vecinos < 1:
        raise ValueError("--vecinos debe ser al menos 1")

    carpeta = Path(__file__).resolve().parent
    raiz = Path(args.salida).expanduser().resolve() if args.salida else carpeta / "Resultados_Limpieza_Inicial"
    directorio_csv, directorio_excel = raiz / "CSV", raiz / "Excel"
    directorio_graficos = raiz / "Graficos_Paper_Profesional_Sin_Zoom"
    directorio_csv.mkdir(parents=True, exist_ok=True)
    directorio_excel.mkdir(parents=True, exist_ok=True)
    directorio_graficos.mkdir(parents=True, exist_ok=True)

    if args.solo_graficos:
        origen_csv = directorio_csv / "datos_limpios.csv"
        if not origen_csv.exists():
            raise FileNotFoundError(f"Falta {origen_csv}. Ejecuta primero la limpieza completa.")
        print(f"Leyendo datos ya limpios: {origen_csv}",flush=True)
        limpios = pd.read_csv(origen_csv,encoding="utf-8-sig")
        est_desp = calcular_estadisticos(limpios)
        archivo_est_antes = directorio_csv / "estadisticos_antes.csv"
        est_antes = (pd.read_csv(archivo_est_antes,encoding="utf-8-sig")
                     if archivo_est_antes.exists() else est_desp)
        rutas = generar_graficos(limpios,est_antes,est_desp,
                                limpios[SENSORES].corr().reset_index(),directorio_graficos)
        print(f"Listo: {len(rutas)} figuras en {directorio_graficos}",flush=True)
        return

    entrada = buscar_entrada(carpeta,args.archivo)
    print(f"Leyendo archivo original: {entrada.name}", flush=True)
    original = pd.read_excel(
        entrada, sheet_name=0, engine="openpyxl",
        na_values=["NULL", "null", "None", " "],
    )
    if original.empty:
        raise ValueError("El Excel original esta vacio.")
    num = normalizar_original(original)
    datos, correcciones, realineaciones, fragmentos, revisiones, _ = recuperar_filas(num)
    # Estadisticos 'antes' sobre las columnas YA correctamente realineadas.
    # Asi no se calculan medias equivocadas por el desplazamiento del origen.
    estadisticos_antes = calcular_estadisticos(datos)
    faltantes_antes = datos[SENSORES].isna().sum()
    smp_faltantes_antes = int(datos["smp"].isna().sum())
    datos = completar_sensores(datos, correcciones, args.vecinos)
    datos, saltos, reinicios = validar_y_completar_smp(datos, correcciones)

    for col in SENSORES:
        datos[col] = datos[col].astype("float64")
    datos = datos[COLUMNAS].reset_index(drop=True)
    if len(datos) != len(original):
        raise AssertionError("La limpieza altero la cantidad de filas.")
    if datos.isna().any().any():
        raise AssertionError("Quedaron celdas vacias en los datos finales.")
    if not np.isfinite(datos.to_numpy(dtype=float)).all():
        raise AssertionError("Los datos finales contienen numeros no finitos.")

    estadisticos_despues = calcular_estadisticos(datos)
    correlaciones = datos[SENSORES].corr(method="pearson")
    correlaciones.index.name = "Variable"
    correlaciones = correlaciones.reset_index()

    registro = pd.DataFrame(correcciones, columns=REG_CAMPOS)
    realineadas = pd.DataFrame(realineaciones, columns=REAL_CAMPOS)
    frag_df = pd.DataFrame(fragmentos, columns=FRAG_CAMPOS)
    saltos_df = pd.DataFrame(saltos, columns=SALTOS_CAMPOS)
    revision_df = pd.DataFrame(revisiones, columns=REV_CAMPOS)
    imputadas = int((registro["tipo"] == "imputacion_sensor").sum())
    contador_estimado = int((registro["tipo"] == "inferencia_contador").sum())
    calidad = pd.DataFrame([
        {
            "Variable": col,
            "Tipo final": "int64" if col == "smp" else "float64",
            "Vacios antes": (smp_faltantes_antes if col == "smp" else int(faltantes_antes[col])),
            "Vacios despues": int(datos[col].isna().sum()),
            "Valores imputados": int(((registro["tipo"] == "imputacion_sensor") & (registro["columna"] == col)).sum()),
        }
        for col in COLUMNAS
    ])

    resumen = pd.DataFrame([
        ("Archivo original", entrada.name),
        ("Filas originales", len(original)),
        ("Filas limpias", len(datos)),
        ("Filas eliminadas", 0),
        ("Variables reales", len(datos.columns)),
        ("Sensores float64", len(SENSORES)),
        ("Contadores int64", 1),
        ("Columnas extra en archivo", original.shape[1] - len(datos.columns)),
        ("Filas realineadas", len(realineadas)),
        ("Celdas auditadas por realineacion", int((registro["tipo"] == "realineacion").sum())),
        ("Celdas de sensores imputadas", imputadas),
        ("Contadores smp inferidos", contador_estimado),
        ("Fragmentos originales conservados", len(frag_df)),
        ("Saltos smp de 2 posiciones", len(saltos_df)),
        ("Reinicios smp 65535 a 0", len(reinicios)),
        ("Celdas vacias finales", int(datos.isna().sum().sum())),
        ("Tratamiento de atipicos", "NO; se analizara en otro programa"),
        ("Referencia de datos ANTES", "Tras realinear columnas, antes de imputar nulos"),
        ("Validez fisica de extremos", "No verificable sin unidades y rangos de sensores"),
        ("Eje temporal", "No hay segundos: smp es contador de muestras"),
    ], columns=["Indicador", "Resultado"])

    def guardar_csv(df: pd.DataFrame, nombre: str) -> None:
        df.to_csv(
            directorio_csv / nombre, index=False, encoding="utf-8-sig",
            float_format="%.15g",
        )

    guardar_csv(datos, "datos_limpios.csv")
    guardar_csv(registro, "registro_correcciones.csv")
    guardar_csv(resumen, "resumen_limpieza.csv")
    guardar_csv(estadisticos_antes, "estadisticos_antes.csv")
    guardar_csv(estadisticos_despues, "estadisticos_despues.csv")

    archivo_excel = directorio_excel / "Resultados_limpieza.xlsx"
    print("Preparando Excel con tablas simples y auditoria...", flush=True)
    escribir_excel(archivo_excel, {
        "Resumen": resumen,
        "Variables": calidad,
        "Estadisticos_antes": estadisticos_antes,
        "Estadisticos_despues": estadisticos_despues,
        "Correlaciones": correlaciones,
        "Datos_limpios": datos,
        "Correcciones": registro,
        "Filas_realineadas": realineadas,
        "Fragmentos_auditados": frag_df,
    })

    print("\nGenerando 32 graficos cientificos separados (sin zoom ni ventanas)...", flush=True)
    rutas_graficos = generar_graficos(datos, estadisticos_antes,
                                     estadisticos_despues, correlaciones,
                                     directorio_graficos)

    resumen_texto = [
        "="*76,
        "INFORME DE LIMPIEZA INICIAL DEL DATASET EOLICO",
        "="*76,
        f"Archivo leido: {entrada.name}",
        f"Registros antes: {len(original):,} | Despues: {len(datos):,} | Eliminados: 0",
        f"Variables: {len(COLUMNAS)} (10 sensores float64 y contador smp int64)",
        f"Columnas extra originales (sin usar en senales): {original.shape[1]-len(COLUMNAS)}",
        "",
        "1. CORRECCIONES EFECTUADAS",
        f"   Filas realineadas y validadas con smp: {len(realineadas):,}",
        f"   Celdas de sensores estimadas con {args.vecinos} vecinos: {imputadas:,}",
        f"   Valores de smp inferidos: {contador_estimado:,}",
        f"   Fragmentos originales guardados para auditoria: {len(frag_df):,}",
        f"   Saltos del contador de 2 posiciones: {len(saltos_df):,}",
        f"   Reinicios del contador modulo 65536: {len(reinicios):,}",
        "",
        "2. VALORES FALTANTES POR VARIABLE (antes -> despues)",
    ]
    for _, f in calidad.iterrows():
        resumen_texto.append(
            f"   {f['Variable']:<20} {int(f['Vacios antes']):>5} -> "
            f"{int(f['Vacios despues']):<5}   Tipo final: {f['Tipo final']}")

    campos = ["Variable", "Validos", "Nulos", "Media", "DS_muestral", "Minimo", "Maximo"]
    def estadistico_para_texto(tabla):
        return tabla[campos].to_string(index=False, float_format=lambda x: f"{x:,.6g}")
    resumen_texto += [
        "", "3. ESTADISTICOS ANTES (filas realineadas, ANTES de imputar)",
        estadistico_para_texto(estadisticos_antes),
        "", "4. ESTADISTICOS DESPUES DE LA LIMPIEZA",
        estadistico_para_texto(estadisticos_despues),
        "", "5. INTERPRETACION Y LIMITACIONES",
        "   Se conservaron todas las filas y los valores registrados.",
        "   No se aplico NI se calculo 6 sigma: corresponde a la siguiente etapa.",
        "   La presencia de valores extremos no confirma defectos fisicos sin rangos/unidades.",
        "   smp es orden de muestra, no tiempo real en segundos.",
        f"   Graficos PNG: {len(rutas_graficos)} figuras guardadas; NINGUNA ventana abierta.",
        "", "6. RUTAS DE SALIDA",
        f"   CSV: {directorio_csv}",
        f"   Excel: {archivo_excel}",
        f"   Figuras: {directorio_graficos}",
        "="*76,
    ]
    informe="\n".join(resumen_texto)+"\n"
    print("\n"+informe,flush=True)
    (raiz/"RESUMEN_TERMINAL.txt").write_text(informe,encoding="utf-8")


if __name__ == "__main__":
    try:
        main()
    except (ValueError, FileNotFoundError, PermissionError, AssertionError) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        sys.exit(1)
