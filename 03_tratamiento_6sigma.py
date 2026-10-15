# -*- coding: utf-8 -*-
"""ETAPA 03 - DETECCION Y TRATAMIENTO DE ATIPICOS POR 6 DESVIACIONES ESTANDAR.

REQUISITOS: python -m pip install numpy matplotlib openpyxl
EJECUTAR:   python 03_tratamiento_6sigma.py

Entrada por defecto:
    Resultados_Limpieza_Inicial/CSV/datos_limpios.csv

Salida:
    Resultados_6Sigma/CSV/*.csv (incluye datos_tratados_6sigma.csv)
    Resultados_6Sigma/Excel/Resultados_6Sigma.xlsx (tablas simples)
    Resultados_6Sigma/Graficos_Paper/ (33 PNG, sin zoom; incluye correlaciones)
    Resultados_6Sigma/RESUMEN_TERMINAL.txt

Regla: |x - media| / DS_muestral >= 6. Estadisticos calculados UNA sola
vez usando todas las filas de la entrada, sin iteraciones. Se evalúan SOLO
las 10 columnas de sensores, NO el contador 'smp'.

Modo por defecto: elimina de una NUEVA COPIA toda fila que supere 6 DS en
al menos un sensor, y la conserva en un registro de auditoria. NO modifica
el CSV de la etapa de limpieza ni el Excel original. La deteccion 6 DS es un
criterio estadistico; NO prueba que los valores sean errores fisicos.

--solo-diagnostico: detecta y documenta, pero NO elimina ninguna fila.
--sin-graficos: acelera la ejecucion si solo se requieren tablas.

Las abscisas en los graficos de evolucion indican POSICION ORIGINAL DE
MUESTRA (no tiempo real en segundos). Las comparaciones usan siempre los
mismos limites graficos completos min/max del dataset antes del tratamiento.
"""

from __future__ import annotations

import argparse
import csv
from pathlib import Path
import sys

try:
    import numpy as np
    import matplotlib
    matplotlib.use("Agg", force=True)  # No abrir ventanas.
    import matplotlib.pyplot as plt
    from matplotlib.lines import Line2D
    from matplotlib.patches import Patch
    from matplotlib.ticker import MaxNLocator
    from matplotlib.colors import to_rgb
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
except ImportError as exc:
    print(f"Falta la biblioteca: {exc.name}")
    print(f'Instalala en el mismo Python: "{sys.executable}" -m pip install numpy matplotlib openpyxl')
    raise SystemExit(1) from exc

SENSORES = [
    "AcelX1_SisGenEol", "AcelY1_SisGenEol", "AcelZ1_SisGenEol",
    "GiroX_SisGenEol", "GiroY_SisGenEol", "GiroZ_SisGenEol",
    "MagneX_SisGenEol", "MagneY_SisGenEol", "MagneZ_SisGenEol",
    "Ter_SisGenEol",
]
COLUMNAS = SENSORES + ["smp"]
DS_UMBRAL = 6.0
COLORES = {
    "AcelX1_SisGenEol": "#2F65B0", "AcelY1_SisGenEol": "#1598A6",
    "AcelZ1_SisGenEol": "#7A59B5", "GiroX_SisGenEol": "#C76B2C",
    "GiroY_SisGenEol": "#2D8A5F", "GiroZ_SisGenEol": "#B14662",
    "MagneX_SisGenEol": "#A68A22", "MagneY_SisGenEol": "#2D80BF",
    "MagneZ_SisGenEol": "#9B4F93", "Ter_SisGenEol": "#D55A4E",
}
NEGRO = "#293139"
GRIS = "#7D8792"
ROJO = "#C23C3C"


def nombre_corto(nombre: str) -> str:
    return nombre.replace("_SisGenEol", "")


def numero(x: float) -> str:
    return f"{x:.6g}" if np.isfinite(x) else "N/D"


def aclarar(color: str, fraccion: float = .72) -> tuple[float, float, float]:
    return tuple((1-fraccion)*np.array(to_rgb(color)) + fraccion)


def leer_csv(ruta: Path) -> np.ndarray:
    if not ruta.is_file():
        raise FileNotFoundError(
            f"No se encontro {ruta}. Ejecuta primero 02_limpiar_datos.py "
            "o usa --entrada para indicar tu datos_limpios.csv."
        )
    with ruta.open("r", encoding="utf-8-sig", newline="") as f:
        encabezados = next(csv.reader(f))
    if encabezados != COLUMNAS:
        raise ValueError(f"Encabezados inesperados en {ruta.name}: {encabezados}")
    datos = np.loadtxt(ruta, dtype=np.float64, delimiter=",", skiprows=1, encoding="utf-8-sig")
    if datos.ndim == 1:
        datos = datos.reshape(1, -1)
    if datos.shape[1] != len(COLUMNAS) or len(datos) < 2:
        raise ValueError("Se esperan al menos dos filas y las 11 columnas del dataset limpio.")
    if not np.isfinite(datos).all():
        raise ValueError("Hay NaN o valores infinitos: revisa primero la limpieza inicial.")
    if not np.all(np.isclose(datos[:, -1], np.round(datos[:, -1]), atol=1e-9, rtol=0)):
        raise ValueError("El contador smp deberia contener enteros.")
    return datos


def deteccion(datos: np.ndarray):
    sensores = datos[:, :10]
    medias = sensores.mean(axis=0)
    ds = sensores.std(axis=0, ddof=1)
    z = np.zeros_like(sensores)
    validos = ds > 0
    z[:, validos] = (sensores[:, validos] - medias[validos]) / ds[validos]
    marcas = np.abs(z) >= DS_UMBRAL
    filas_marcadas = marcas.any(axis=1)
    return medias, ds, z, marcas, filas_marcadas


def estadisticos(datos: np.ndarray) -> list[dict]:
    filas = []
    for j, nombre in enumerate(COLUMNAS):
        a = datos[:, j]
        filas.append({
            "Variable": nombre,
            "Cantidad": len(a),
            "Tipo": "int64" if nombre == "smp" else "float64",
            "Media": float(a.mean()) if len(a) else None,
            "Mediana": float(np.median(a)) if len(a) else None,
            "DS_muestral": float(a.std(ddof=1)) if len(a) > 1 else None,
            "Minimo": float(a.min()) if len(a) else None,
            "Q1": float(np.quantile(a, .25)) if len(a) else None,
            "Q3": float(np.quantile(a, .75)) if len(a) else None,
            "Maximo": float(a.max()) if len(a) else None,
        })
    return filas


def escribir_csv_tabla(destino: Path, filas: list[dict], encabezados: list[str]) -> None:
    destino.parent.mkdir(parents=True, exist_ok=True)
    with destino.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=encabezados)
        writer.writeheader()
        writer.writerows(filas)


def escribir_csv_datos(destino: Path, datos: np.ndarray) -> None:
    destino.parent.mkdir(parents=True, exist_ok=True)
    np.savetxt(destino, datos, delimiter=",", header=",".join(COLUMNAS),
               comments="", encoding="utf-8", fmt=["%.15g"]*10+["%.0f"])


def construir_auditoria(datos, medias, ds, z, marcas, filas_marcadas):
    detalles = []
    retiradas = []
    idx = np.flatnonzero(filas_marcadas)
    for i in idx:
        fuera = np.flatnonzero(marcas[i])
        retiradas.append({
            "Fila_CSV": int(i) + 2,
            "Posicion_muestra_original": int(i) + 1,
            "smp": int(datos[i, 10]),
            "Variables_fuera_6DS": ", ".join(SENSORES[j] for j in fuera),
            "Z_max_absoluto": float(np.max(np.abs(z[i]))),
            **{COLUMNAS[j]: int(datos[i, j]) if j == 10 else float(datos[i, j])
               for j in range(len(COLUMNAS))},
        })
        for j in fuera:
            detalles.append({
                "Fila_CSV": int(i) + 2,
                "Posicion_muestra_original": int(i) + 1,
                "smp": int(datos[i, 10]),
                "Variable": SENSORES[j],
                "Valor_observado": float(datos[i, j]),
                "Media_origen": float(medias[j]),
                "DS_muestral_origen": float(ds[j]),
                "Limite_inferior_6DS": float(medias[j] - DS_UMBRAL*ds[j]),
                "Limite_superior_6DS": float(medias[j] + DS_UMBRAL*ds[j]),
                "Z_con_signo": float(z[i, j]),
                "Z_absoluto": float(abs(z[i, j])),
                "Criterio": "|z| >= 6 (una pasada)",
                "Estado": "Candidato a revision fisica",
            })
    return retiradas, detalles


def generar_excel(destino: Path, tablas: dict[str, list[dict]]):
    """Genera un Excel de tablas simples con openpyxl (sin artifact_tool).

    Las observaciones completas quedan tambien disponibles en CSV. En Excel
    se incluyen resumen, estadisticos y auditoria para que sea facil revisar.
    """
    libro = Workbook()
    primera = True
    for nombre_hoja, filas in tablas.items():
        if primera:
            hoja = libro.active
            hoja.title = nombre_hoja
            primera = False
        else:
            hoja = libro.create_sheet(title=nombre_hoja)

        if filas:
            encabezados = list(filas[0].keys())
            hoja.append(encabezados)
            for fila in filas:
                hoja.append([
                    valor.item() if isinstance(valor, np.generic) else valor
                    for valor in (fila.get(col) for col in encabezados)
                ])
        else:
            encabezados = ["Observacion"]
            hoja.append(encabezados)
            hoja.append(["No hay filas detectadas con este criterio"])

        for celda in hoja[1]:
            celda.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            celda.fill = PatternFill(fill_type="solid", fgColor="344054")
            celda.alignment = Alignment(vertical="center")
        hoja.row_dimensions[1].height = 22
        hoja.freeze_panes = "A2"
        hoja.auto_filter.ref = hoja.dimensions

        for indice, titulo in enumerate(encabezados, start=1):
            # Tablas legibles sin columnas exageradamente anchas.
            ancho = min(34, max(15, len(str(titulo)) + 3))
            hoja.column_dimensions[get_column_letter(indice)].width = ancho

    destino.parent.mkdir(parents=True, exist_ok=True)
    libro.save(destino)


def estilo_paper():
    plt.ioff()
    plt.rcParams.update({
        "font.family": "DejaVu Sans", "font.size": 9,
        "axes.titlesize": 10, "axes.labelsize": 9,
        "xtick.labelsize": 8, "ytick.labelsize": 8,
        "axes.spines.top": False, "axes.spines.right": False,
        "axes.linewidth": .75, "pdf.fonttype": 42,
        "savefig.dpi": 300, "figure.facecolor": "white",
        "axes.facecolor": "white", "axes.formatter.useoffset": False,
        "agg.path.chunksize": 20000,
    })


def ejes(ax, rejilla="y"):
    ax.spines["left"].set_color("#B9BFC6")
    ax.spines["bottom"].set_color("#B9BFC6")
    ax.tick_params(colors="#6B737C", length=3, width=.8)
    ax.set_axisbelow(True)
    ax.grid(axis=rejilla, color="#E9EAED", lw=.65)
    ax.yaxis.set_major_locator(MaxNLocator(nbins=5))


def figura_comparativa(titulo, subtitulo, pie, comparador="h"):
    if comparador == "v":
        fig, ax = plt.subplots(2, 1, figsize=(8.6, 6.9), sharex=True, sharey=True)
        fig.subplots_adjust(left=.12, right=.97, top=.78, bottom=.17, hspace=.26)
    else:
        fig, ax = plt.subplots(1, 2, figsize=(10.2, 4.6), sharex=True, sharey=True)
        fig.subplots_adjust(left=.085, right=.98, top=.76, bottom=.23, wspace=.18)
    fig.suptitle(titulo, x=.085 if comparador=="h" else .12, y=.96,
                 ha="left", fontweight="bold", fontsize=12, color=NEGRO)
    fig.text(.085 if comparador=="h" else .12, .898, subtitulo, fontsize=9, color="#64707A")
    fig.text(.085 if comparador=="h" else .12, .055, pie, fontsize=8, color="#64707A")
    return fig, ax


def guardar_png(fig, ruta: Path):
    ruta.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(ruta, dpi=300, facecolor="white", bbox_inches="tight", pad_inches=.12)
    plt.close(fig)


def limites_completos(valores: np.ndarray):
    a, b = float(np.min(valores)), float(np.max(valores))
    margen = max(.025*(b-a), .000000001, abs(a)*1e-7)
    return a-margen, b+margen


def graficar_histogramas(original, tratado, variable, j, salida):
    color = COLORES[variable]
    fig, axs = figura_comparativa(
        f"Distribución de {nombre_corto(variable)} · comparación 6σ",
        "Mismos límites para ambas muestras · Sin zoom",
        "La escala horizontal abarca el mínimo y máximo originales; sin recortes por percentiles.")
    serie_antes, serie_despues = original[:,j], tratado[:,j]
    mi, ma = float(np.min(serie_antes)), float(np.max(serie_antes))
    bins = np.linspace(mi, ma, 61) if ma>mi else np.linspace(mi-.001,ma+.001,61)
    max_frec = 1.0
    for ax, serie, nombre, color_barras in zip(
        axs, [serie_antes, serie_despues], ["Antes", "Después"], ["#B5BBC3", color]
    ):
        frec, _, _ = ax.hist(serie, bins=bins, color=color_barras,
                             edgecolor="white", linewidth=.3, alpha=.9)
        max_frec = max(max_frec, max(frec))
        ax.axvline(serie.mean(), ls="--", c=NEGRO, lw=1.15, zorder=4)
        ax.axvline(np.median(serie), ls="-", c=ROJO, lw=1.25, zorder=5)
        ax.set_title(f"{nombre} (n = {len(serie):,})", loc="left", fontweight="bold")
        ax.set_xlabel("Valor medido")
        ejes(ax)
    axs[0].set_ylabel("Frecuencia")
    axs[0].set_xlim(*limites_completos(serie_antes))
    axs[0].set_ylim(0, max_frec*1.09)
    fig.legend(handles=[Patch(facecolor="#B5BBC3",label="Antes"),
                        Patch(facecolor=color,label="Después"),
                        Line2D([0],[0],color=NEGRO,ls="--",label="Media"),
                        Line2D([0],[0],color=ROJO,ls="-",label="Mediana")],
               frameon=False, ncol=4, loc="upper center", bbox_to_anchor=(.55,.865))
    guardar_png(fig, salida)


def graficar_cajas(original, tratado, variable, j, salida):
    color = COLORES[variable]
    serie_antes, serie_despues = original[:,j], tratado[:,j]
    fig, ax = plt.subplots(figsize=(7.6, 5.0))
    fig.subplots_adjust(left=.15, right=.92, top=.79, bottom=.22)
    fig.suptitle(f"Diagrama de caja de {nombre_corto(variable)} · antes y después",
                 x=.15, ha="left", y=.96, fontweight="bold", fontsize=11.5)
    fig.text(.15,.89,"Escala completa del mínimo al máximo original · Puntos externos visibles",fontsize=9,color="#64707A")
    fig.text(.15,.06,"Bigotes: 1.5 × RIC solo como representación gráfica; tratamiento aplicado: ±6 DE.",fontsize=8,color="#64707A")
    rng = np.random.default_rng(983)
    for lugar, serie, color_base in [(1,serie_antes,"#909AA5"),(2,serie_despues,color)]:
        ax.boxplot([serie], positions=[lugar], widths=.26, patch_artist=True,
                   showmeans=True, showfliers=False,
                   boxprops=dict(facecolor=color_base,alpha=.3,edgecolor=color_base,linewidth=1.4),
                   medianprops=dict(color=ROJO,linewidth=1.6),
                   meanprops=dict(marker="D",markerfacecolor=NEGRO,markeredgecolor=NEGRO,markersize=4),
                   whiskerprops=dict(color=color_base,linewidth=1.25),
                   capprops=dict(color=color_base,linewidth=1.25))
        q1,q3=np.quantile(serie,[.25,.75])
        atipicos = serie[(serie<q1-1.5*(q3-q1))|(serie>q3+1.5*(q3-q1))]
        if len(atipicos):
            xx = lugar+rng.uniform(-.16,.16,size=len(atipicos))
            ax.scatter(xx,atipicos,s=7,c=[color_base],alpha=.32,edgecolors="none",
                       rasterized=True,zorder=3)
    ax.set_xlim(.55,2.45)
    ax.set_ylim(*limites_completos(serie_antes))
    ax.set_xticks([1,2], [f"Antes\n(n={len(serie_antes):,})", f"Después\n(n={len(serie_despues):,})"])
    ax.set_ylabel("Valor medido")
    ejes(ax)
    guardar_png(fig,salida)


def graficar_evolucion(original, tratado, retenidas, variable, j, salida):
    color=COLORES[variable]
    ya,yb=original[:,j],tratado[:,j]
    fig,axs=figura_comparativa(
        f"Comportamiento por muestra · {nombre_corto(variable)}",
        "Posiciones originales · Mismo rango vertical y horizontal · Sin zoom",
        "No se suavizan ni recortan muestras. El eje horizontal no equivale a segundos.","v")
    n=len(ya)
    for ax, ypos, vals, texto in [
        (axs[0], np.arange(1,n+1),ya,f"Antes  (n = {len(ya):,})"),
        (axs[1], np.flatnonzero(retenidas)+1,yb,f"Después  (n = {len(yb):,})")
    ]:
        # Los datos completos pueden verse densos al representar 143 mil muestras.
        ax.plot(ypos,vals,c=aclarar(color,.53),lw=.26,alpha=.58,rasterized=True,zorder=1)
        ax.axhline(vals.mean(),c=NEGRO,ls="--",lw=1.3,zorder=4)
        ax.axhline(np.median(vals),c=ROJO,ls="-",lw=1.5,zorder=5)
        ax.set_title(texto,loc="left",fontweight="bold")
        ax.set_ylabel("Valor medido")
        ax.set_xlim(1,n)
        ax.set_ylim(*limites_completos(ya))
        ax.ticklabel_format(axis="x",style="sci",scilimits=(4,4))
        ejes(ax)
    axs[1].set_xlabel("Posición original de muestra")
    fig.legend(handles=[Line2D([0],[0],color=aclarar(color,.53),lw=1,label="Señal"),
                        Line2D([0],[0],color=NEGRO,ls="--",label="Media"),
                        Line2D([0],[0],color=ROJO,label="Mediana")],
               loc="upper right",bbox_to_anchor=(.96,.86),ncol=3,frameon=False)
    guardar_png(fig,salida)


def graficar_conteos(marcas, carpeta):
    n=marcas.sum(axis=0)
    fig,ax=plt.subplots(figsize=(8.0,4.7))
    fig.subplots_adjust(left=.11,right=.97,bottom=.28,top=.82)
    fig.suptitle("Observaciones fuera de ±6 desviaciones estándar",
                 x=.11,ha="left",y=.96,fontweight="bold",fontsize=12)
    fig.text(.11,.89,"Se excluye smp; criterio aplicado una sola vez",fontsize=9,color="#64707A")
    pos=np.arange(10)
    ax.bar(pos,n,color=[COLORES[k] for k in SENSORES],width=.67)
    ax.set_xticks(pos,[nombre_corto(s) for s in SENSORES],rotation=35,ha="right")
    ax.set_ylabel("Cantidad de valores")
    ax.yaxis.set_major_locator(MaxNLocator(integer=True))
    for x,y in zip(pos,n):
        ax.text(x,y+.06,str(int(y)),ha="center",fontsize=8)
    ax.set_ylim(0,max(1.3,float(n.max())*1.18))
    ejes(ax)
    guardar_png(fig,carpeta/"Conteo_atipicos_6sigma.png")




def matriz_correlaciones(datos: np.ndarray) -> np.ndarray:
    """Coeficientes r de Pearson de los 10 sensores; smp no participa."""
    return np.corrcoef(datos[:, :len(SENSORES)], rowvar=False)


def filas_matriz_correlacion(matriz: np.ndarray) -> list[dict]:
    return [
        {"Variable": fila, **{col: float(matriz[i, j]) if np.isfinite(matriz[i, j]) else None
                               for j, col in enumerate(SENSORES)}}
        for i, fila in enumerate(SENSORES)
    ]


def filas_cambios_correlacion(antes: np.ndarray, despues: np.ndarray) -> list[dict]:
    # Eliminar diagonal y duplicados: 45 parejas únicas entre 10 sensores.
    return [
        {"Variable_1": SENSORES[i], "Variable_2": SENSORES[j],
         "Pearson_antes": float(antes[i, j]),
         "Pearson_despues": float(despues[i, j]),
         "Cambio_r": float(despues[i, j] - antes[i, j])}
        for i in range(len(SENSORES)) for j in range(i + 1, len(SENSORES))
    ]


def _decorar_matriz(ax, matriz: np.ndarray, titulo: str):
    imagen = ax.imshow(matriz, cmap="RdBu_r", vmin=-1, vmax=1,
                       interpolation="nearest", aspect="equal")
    etiquetas = [nombre_corto(s) for s in SENSORES]
    ax.set_xticks(np.arange(len(SENSORES)), etiquetas,
                  rotation=45, ha="right", rotation_mode="anchor", fontsize=8)
    ax.set_yticks(np.arange(len(SENSORES)), etiquetas, fontsize=8)
    ax.tick_params(length=0, colors="#414750")
    for borde in ax.spines.values():
        borde.set_visible(False)
    for i in range(len(SENSORES)):
        for j in range(len(SENSORES)):
            valor = matriz[i, j]
            if np.isfinite(valor):
                ax.text(j, i, f"{valor:.2f}", ha="center", va="center",
                        fontsize=7.1, color=("white" if abs(valor) >= 0.65 else "#26313A"))
    ax.set_title(titulo, fontsize=11, fontweight="semibold", pad=12, color="#25303A")
    return imagen


def graficar_correlaciones_despues(matriz: np.ndarray, n: int, destino: Path):
    """Mapa de calor posterior a 6σ, con escala Pearson completa [-1,1]."""
    fig, ax = plt.subplots(figsize=(8.2, 7.2), facecolor="white")
    fig.subplots_adjust(left=.22, right=.84, bottom=.20, top=.85)
    imagen = _decorar_matriz(ax, matriz, "Correlación de Pearson · después de 6σ")
    cbar = fig.colorbar(imagen, ax=ax, fraction=.05, pad=.035, shrink=.87)
    cbar.set_label("Coeficiente de correlación (r)", fontsize=9)
    cbar.ax.tick_params(labelsize=8)
    fig.text(.22, .91, f"n = {n:,} muestras · 10 sensores · smp excluido", fontsize=9, color="#64707A")
    fig.text(.22, .035, "Se muestran todas las variables. Escala fija de r = −1 a +1; sin zoom.",
             fontsize=8.4, color="#64707A")
    guardar_png(fig, destino)


def graficar_correlaciones_comparadas(antes: np.ndarray, despues: np.ndarray,
                                      n_antes: int, n_despues: int, destino: Path):
    """Comparación directa: exactamente la misma escala de color para ambos mapas."""
    fig, axs = plt.subplots(1, 2, figsize=(13.4, 6.5), facecolor="white")
    fig.subplots_adjust(left=.12, right=.90, bottom=.23, top=.82, wspace=.34)
    imagen = _decorar_matriz(axs[0], antes, f"Antes de 6σ  (n = {n_antes:,})")
    _decorar_matriz(axs[1], despues, f"Después de 6σ  (n = {n_despues:,})")
    fig.suptitle("Correlación de Pearson · antes y después de 6σ",
                 x=.12, ha="left", fontsize=13, fontweight="bold", color="#25303A", y=.96)
    cbar = fig.colorbar(imagen, ax=axs, fraction=.022, pad=.04, shrink=.8)
    cbar.set_label("Coeficiente r", fontsize=9)
    cbar.ax.tick_params(labelsize=8)
    fig.text(.12, .055, "Escala de color idéntica (−1 a +1); los cambios pequeños son esperables al retirar pocas filas.",
             fontsize=8.5, color="#64707A")
    guardar_png(fig, destino)

def generar_graficos(original, tratado, retenidas, marcas, carpeta):
    estilo_paper()
    rutas=[]
    for j,variable in enumerate(SENSORES):
        ruta=carpeta/variable
        graficar_histogramas(original,tratado,variable,j,ruta/"01_Histogramas_antes_despues.png")
        graficar_cajas(original,tratado,variable,j,ruta/"02_Cajas_antes_despues.png")
        graficar_evolucion(original,tratado,retenidas,variable,j,ruta/"03_Evolucion_antes_despues.png")
        rutas+=list(sorted(ruta.glob("*.png")))
        print(f"  ✓ {nombre_corto(variable):7s} | 3 graficos comparativos",flush=True)
    carpeta_general=carpeta/"00_Generales"
    carpeta_general.mkdir(parents=True,exist_ok=True)
    graficar_conteos(marcas,carpeta_general)
    rutas.append(carpeta_general/"Conteo_atipicos_6sigma.png")
    corr_antes = matriz_correlaciones(original)
    corr_despues = matriz_correlaciones(tratado)
    ruta_actual = carpeta_general / "Correlacion_Despues_6Sigma.png"
    ruta_comparacion = carpeta_general / "Correlaciones_Antes_Despues_6Sigma.png"
    graficar_correlaciones_despues(corr_despues, len(tratado), ruta_actual)
    graficar_correlaciones_comparadas(corr_antes, corr_despues, len(original), len(tratado), ruta_comparacion)
    rutas.extend([ruta_actual, ruta_comparacion])
    print("  ✓ Correlaciones después y antes/después | 2 gráficos adicionales", flush=True)
    return rutas


def main():
    parser = argparse.ArgumentParser(description="Tratamiento 6 sigma independiente de la limpieza inicial")
    parser.add_argument("--entrada",type=Path,help="Ruta de datos_limpios.csv")
    parser.add_argument("--salida",type=Path,help="Carpeta resultados")
    parser.add_argument("--solo-diagnostico",action="store_true",help="No elimina filas; solo las marca para revision")
    parser.add_argument("--sin-graficos",action="store_true",help="No genera figuras")
    args=parser.parse_args()

    base=Path(__file__).resolve().parent
    entrada=(args.entrada or base/"Resultados_Limpieza_Inicial"/"CSV"/"datos_limpios.csv").expanduser().resolve()
    raiz=(args.salida or base/"Resultados_6Sigma").expanduser().resolve()
    csv_dir, excel_dir,fig_dir=raiz/"CSV",raiz/"Excel",raiz/"Graficos_Paper"
    csv_dir.mkdir(parents=True,exist_ok=True)
    excel_dir.mkdir(parents=True,exist_ok=True)
    print(f"Leyendo CSV limpio: {entrada}",flush=True)
    original=leer_csv(entrada)
    medias,ds,z,marcas,filas_marcadas=deteccion(original)
    idx=np.flatnonzero(filas_marcadas)
    retenidas=(~filas_marcadas) if not args.solo_diagnostico else np.ones(len(original),dtype=bool)
    tratado=original[retenidas]
    retiradas,detalles=construir_auditoria(original,medias,ds,z,marcas,filas_marcadas)
    reg_cols=["Fila_CSV","Posicion_muestra_original","smp","Variable","Valor_observado",
              "Media_origen","DS_muestral_origen","Limite_inferior_6DS","Limite_superior_6DS",
              "Z_con_signo","Z_absoluto","Criterio","Estado"]
    filas_cols=["Fila_CSV","Posicion_muestra_original","smp","Variables_fuera_6DS","Z_max_absoluto"]+COLUMNAS
    limites=[{
        "Variable":nombre,"Media_origen":float(medias[j]),"DS_muestral_origen":float(ds[j]),
        "Limite_inferior_6DS":float(medias[j]-DS_UMBRAL*ds[j]),
        "Limite_superior_6DS":float(medias[j]+DS_UMBRAL*ds[j]),
        "Observaciones_fuera":int(marcas[:,j].sum()),"Excluida": "NO",
    } for j,nombre in enumerate(SENSORES)]
    limites.append({"Variable":"smp","Media_origen":None,"DS_muestral_origen":None,
                    "Limite_inferior_6DS":None,"Limite_superior_6DS":None,
                    "Observaciones_fuera":None,"Excluida":"SI - contador de muestras"})
    antes=estadisticos(original)
    despues=estadisticos(tratado)
    comparacion=[]
    for a,b in zip(antes,despues):
        comparacion.append({
            "Variable":a["Variable"],
            "Media_antes":a["Media"],"Media_despues":b["Media"],
            "DS_antes":a["DS_muestral"],"DS_despues":b["DS_muestral"],
            "Min_antes":a["Minimo"],"Min_despues":b["Minimo"],
            "Max_antes":a["Maximo"],"Max_despues":b["Maximo"],
        })
    modalidad="Solo diagnostico: NO se eliminaron filas" if args.solo_diagnostico else "Se retiraron filas de una copia"
    resumen=[
        {"Indicador":"Archivo de entrada","Resultado":entrada.name},
        {"Indicador":"Criterio","Resultado":"|z| >= 6, una sola pasada"},
        {"Indicador":"DS utilizada","Resultado":"Desviacion estandar muestral (ddof=1)"},
        {"Indicador":"Variables evaluadas","Resultado":"10 sensores; smp excluido"},
        {"Indicador":"Filas de entrada","Resultado":int(len(original))},
        {"Indicador":"Filas detectadas por 6DS","Resultado":int(len(idx))},
        {"Indicador":"Valores individuales detectados","Resultado":int(marcas.sum())},
        {"Indicador":"Filas retiradas de la copia","Resultado":int(np.sum(~retenidas))},
        {"Indicador":"Filas de salida","Resultado":int(len(tratado))},
        {"Indicador":"Porcentaje retirado","Resultado":float(100*np.mean(~retenidas))},
        {"Indicador":"Modalidad","Resultado":modalidad},
        {"Indicador":"¿Se modifico el CSV original?","Resultado":"NO"},
        {"Indicador":"¿Se imputaron los valores 6DS?","Resultado":"NO"},
        {"Indicador":"Limitacion metodologica","Resultado":"6DS identifica valores extremos, no valida errores fisicos"},
        {"Indicador":"Eje de evolucion","Resultado":"Orden original de muestra, no segundos"},
    ]
    escribir_csv_datos(csv_dir/"datos_tratados_6sigma.csv",tratado)
    escribir_csv_tabla(csv_dir/"registro_atipicos_6sigma.csv",detalles,reg_cols)
    escribir_csv_tabla(csv_dir/"filas_detectadas_6sigma.csv",retiradas,filas_cols)
    for nombre,tabla in [
        ("estadisticos_antes_6sigma.csv",antes),
        ("estadisticos_despues_6sigma.csv",despues),
        ("limites_6sigma.csv",limites),
        ("comparacion_estadisticos.csv",comparacion),
        ("resumen_6sigma.csv",resumen),
    ]:
        escribir_csv_tabla(csv_dir/nombre,tabla,list(tabla[0].keys()))
    # Correlación posterior al tratamiento y comparación con la entrada limpia.
    corr_antes = matriz_correlaciones(original)
    corr_despues = matriz_correlaciones(tratado)
    tabla_corr_antes = filas_matriz_correlacion(corr_antes)
    tabla_corr_despues = filas_matriz_correlacion(corr_despues)
    tabla_cambios_corr = filas_cambios_correlacion(corr_antes, corr_despues)
    for nombre, tabla in [
        ("correlacion_antes_6sigma.csv", tabla_corr_antes),
        ("correlacion_despues_6sigma.csv", tabla_corr_despues),
        ("cambios_correlacion_6sigma.csv", tabla_cambios_corr),
    ]:
        escribir_csv_tabla(csv_dir / nombre, tabla, list(tabla[0].keys()))
    excel=excel_dir/"Resultados_6Sigma.xlsx"
    generar_excel(excel,{
        "Resumen":resumen,
        "Limites_6Sigma":limites,
        "Estadisticos_antes":antes,
        "Estadisticos_despues":despues,
        "Comparacion":comparacion,
        "Filas_detectadas":retiradas,
        "Detalle_atipicos":detalles,
        "Correlacion_antes":tabla_corr_antes,
        "Correlacion_despues":tabla_corr_despues,
        "Cambio_correlacion":tabla_cambios_corr,
    })
    rutas=[]
    if not args.sin_graficos:
        print("Generando graficos sin zoom y sin ventanas (PNG 300 dpi)...",flush=True)
        rutas=generar_graficos(original,tratado,retenidas,marcas,fig_dir)
    resumen_txt=[
        "="*74,
        "ETAPA 03 | TRATAMIENTO DE ATIPICOS POR 6 DESVIACIONES ESTANDAR",
        "="*74,
        f"Archivo fuente: {entrada}",
        f"Muestras originales: {len(original):,}",
        f"Regla de deteccion: |(x - media)/DS_muestral| >= 6, una sola pasada",
        f"Filas detectadas: {len(idx):,}",
        f"Observaciones marcadas: {int(marcas.sum()):,}",
        f"Filas retiradas en copia: {int(np.sum(~retenidas)):,}",
        f"Muestras tratadas: {len(tratado):,}",
        f"smp NO es una variable fisica y NO se usa para decidir eliminaciones.",
        "",
        "VARIABLE                 MEDIA          DS       LIMITE INF.     LIMITE SUP.   FUERA",
    ]
    for j,n in enumerate(SENSORES):
        resumen_txt.append(f"{n:<21} {medias[j]:>11.5g} {ds[j]:>11.5g} "
                           f"{(medias[j]-6*ds[j]):>16.5g} {(medias[j]+6*ds[j]):>16.5g} "
                           f"{int(marcas[:,j].sum()):>5d}")
    resumen_txt += ["", "FILAS DETECTADAS (corresponden a fila del CSV limpio + encabezado):"]
    for f in detalles:
        resumen_txt.append(f" Fila CSV {f['Fila_CSV']}: smp={f['smp']}; "
                           f"{f['Variable']}={f['Valor_observado']:.6g}; "
                           f"z={f['Z_con_signo']:.3f}")
    mayor_cambio = max(tabla_cambios_corr, key=lambda f: abs(f["Cambio_r"]))
    resumen_txt += [
        "", "CORRELACIONES DE PEARSON (10 SENSORES, smp EXCLUIDO):",
        f"* Matriz despues: {csv_dir / 'correlacion_despues_6sigma.csv'}",
        f"* Mayor cambio: {nombre_corto(mayor_cambio['Variable_1'])} / {nombre_corto(mayor_cambio['Variable_2'])}",
        f"  r antes = {mayor_cambio['Pearson_antes']:.6f} | r despues = {mayor_cambio['Pearson_despues']:.6f}",
        "", "IMPORTANTE:",
        "* El criterio ±6 DS es estadistico y NO demuestra fallo fisico del sensor.",
        "* Revisa el contexto de estas 6 observaciones antes de justificar su exclusion.",
        "* El CSV inicial sigue intacto. El eliminado solo se retira de la copia tratada.",
        "* Las filas retiradas crean huecos en la secuencia temporal de smp.",
        "* No se sustituye por promedios ni se recalculan iterativamente los limites.",
        f"* Se guardaron {len(rutas)} graficos comparativos (sin zoom).",
        "",f"Excel: {excel}",f"CSV tratados: {csv_dir/'datos_tratados_6sigma.csv'}",
        f"Auditoria: {csv_dir/'registro_atipicos_6sigma.csv'}",
        "="*74,
    ]
    texto="\n".join(resumen_txt)+"\n"
    print("\n"+texto,flush=True)
    (raiz/"RESUMEN_TERMINAL.txt").write_text(texto,encoding="utf-8")
    (raiz/"LEEME.txt").write_text(
        "PASO 03: 6 desviaciones estandar sobre datos ya limpios.\n"
        "Regla: se retira la fila completa de una COPIA cuando cualquier sensor "
        "tiene |z| >= 6; una unica pasada, sin iteracion.\n"
        "No modifica la entrada original ni reemplaza valores por promedios.\n"
        "El Excel resume resultados; el dataset completo se encuentra en "
        "CSV/datos_tratados_6sigma.csv.\n"
        "Los puntos fuera de 1.5*RIC de los boxplots son elementos "
        "DESCRIPTIVOS; no indican que se haya aplicado RIC al tratamiento.\n",
        encoding="utf-8",
    )


if __name__=="__main__":
    try:
        main()
    except (FileNotFoundError,ValueError,OSError) as err:
        print(f"ERROR: {err}",file=sys.stderr)
        sys.exit(1)
