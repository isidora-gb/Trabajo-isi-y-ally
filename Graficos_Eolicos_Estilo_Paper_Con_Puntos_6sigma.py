# -*- coding: utf-8 -*-
"""
TALLER IA — GRÁFICOS CIENTÍFICOS, SIN ZOOM (VERSIÓN CORREGIDA)
================================================================
Lee un Excel ANTES y otro DESPUÉS de aplicar |z| >= 6.
Genera figuras independientes listas para incorporarse a un informe:
    01_Histogramas           : 10 PNG (2 paneles por variable)
    02_Cajas_min_max        : 10 PNG (bigotes mínimo-máximo, puntos |z|≥6)
    03_Series_completas     : 10 PNG (todos los registros; sin submuestreo)
    04_Barras_estadisticas  : 40 PNG (una métrica/variable por archivo)
    05_Barras_6sigma        :  1 PNG (cantidad por variable)
TOTAL: 71 FIGURAS + README y resumen CSV.

No aplica nueva limpieza, no modifica los Excel, no usa RIC para descartar
registros. No hace zoom, ni recorta ejes, ni usa ejes rotos ni insets.
Los límites de los ejes incluyen los mínimos y máximos originales.
Los ejes de los gráficos antes/después mantienen una ESCALA COMÚN.

Instala dependencias: python -m pip install numpy pandas matplotlib openpyxl
Ejecuta: python Graficos_Eolicos_Estilo_Paper_Sin_Zoom.py
Prueba corta: python Graficos_Eolicos_Estilo_Paper_Sin_Zoom.py --solo GiroZ
Para generar además PDF vectoriales, añade --pdf
"""
from __future__ import annotations

import argparse
from pathlib import Path
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import MaxNLocator, ScalarFormatter, FuncFormatter

BASE = Path(__file__).resolve().parent
NOMBRE_DEPURADO = "Dataset_eolico_depurado_6sigma.xlsx"
CARPETA_RESULTADOS = "Graficos_Eolicos_PAPER_CON_PUNTOS_6SIGMA"
UMBRAL = 6.0

# Nombre en el Excel: (tipo de medición, dirección, etiqueta corta)
SENALES = {
    "AcelX1_SisGenEol": ("Aceleración", "X", "AcelX1"),
    "AcelY1_SisGenEol": ("Aceleración", "Y", "AcelY1"),
    "AcelZ1_SisGenEol": ("Aceleración", "Z", "AcelZ1"),
    "GiroX_SisGenEol": ("Giro", "X", "GiroX"),
    "GiroY_SisGenEol": ("Giro", "Y", "GiroY"),
    "GiroZ_SisGenEol": ("Giro", "Z", "GiroZ"),
    "MagneX_SisGenEol": ("Campo magnético", "X", "MagneX"),
    "MagneY_SisGenEol": ("Campo magnético", "Y", "MagneY"),
    "MagneZ_SisGenEol": ("Campo magnético", "Z", "MagneZ"),
    "Ter_SisGenEol": ("Temperatura", None, "Temperatura"),
}
SUBCARPETAS = {
    "hist": "01_Histogramas",
    "box": "02_Cajas_min_max",
    "serie": "03_Series_completas",
    "barras": "04_Barras_estadisticas",
    "conteo": "05_Barras_6sigma",
}

# Paleta sobria, compatible con publicación y legibilidad en impresión.
ANTES = "#344F71"       # azul pizarra
DESPUES = "#148A87"     # verde petróleo
DESTACADO = "#A75B35"   # terracota: registros con |z|>=6
TEXTO = "#222A31"
CUADRICULA = "#DCE1E5"

plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "font.size": 9,
    "axes.titlesize": 10,
    "axes.titleweight": "normal",
    "axes.labelsize": 9,
    "xtick.labelsize": 8,
    "ytick.labelsize": 8,
    "axes.linewidth": 0.8,
    "axes.edgecolor": TEXTO,
    "axes.labelcolor": TEXTO,
    "xtick.color": TEXTO,
    "ytick.color": TEXTO,
    "text.color": TEXTO,
    "figure.facecolor": "white",
    "axes.facecolor": "white",
    "savefig.facecolor": "white",
    "legend.fontsize": 8,
    "legend.frameon": False,
    "pdf.fonttype": 42,
    "svg.fonttype": "none",
})


def entrada_filtrada(carpeta):
    preferido = carpeta / "Dataset_eolico_filtrado_2.xlsx"
    if preferido.exists():
        return preferido
    candidatos = [p for p in carpeta.glob("Dataset_eolico_filtrado_2*.xlsx")
                  if p.is_file() and not p.name.startswith("~$")]
    if not candidatos:
        raise FileNotFoundError("Falta Dataset_eolico_filtrado_2.xlsx (o versión con sufijo).")
    return max(candidatos, key=lambda p: p.stat().st_mtime)


def cargar(ruta):
    tabla = pd.read_excel(ruta, sheet_name="Datos", engine="openpyxl")
    faltan = [c for c in SENALES if c not in tabla.columns]
    if faltan:
        raise ValueError(f"En {ruta.name} faltan las columnas: {faltan}")
    tabla = tabla[list(SENALES)].copy()
    for c in tabla.columns:
        tabla[c] = pd.to_numeric(tabla[c], errors="raise")
    if tabla.empty or tabla.isna().any().any() or not np.isfinite(tabla.to_numpy()).all():
        raise ValueError(f"{ruta.name} tiene datos vacíos, no numéricos o infinitos.")
    return tabla


def etiqueta(var):
    tipo, eje, corto = SENALES[var]
    return f"{tipo}, eje {eje} ({corto})" if eje is not None else f"{tipo} ({corto})"


def limites_todos(a, b, extra=0.025):
    """Mínimo y máximo de AMBAS muestras. Padding sólo exterior; jamás zoom."""
    menor = float(min(np.min(a), np.min(b)))
    mayor = float(max(np.max(a), np.max(b)))
    amplitud = mayor - menor
    margen = (extra * amplitud) if amplitud > 0 else max(abs(menor)*0.03, 1e-9)
    return menor-margen, mayor+margen


def estilo(ax, eje_rejilla="y"):
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(axis=eje_rejilla, color=CUADRICULA, lw=0.48)
    ax.set_axisbelow(True)
    ax.tick_params(which="both", direction="out", length=3, width=0.65, pad=3)


def eje_continuo(ax, tipo="x", max_ticks=6):
    fmt = ScalarFormatter(useMathText=True, useOffset=False)
    fmt.set_powerlimits((-3, 4))
    axis = ax.xaxis if tipo == "x" else ax.yaxis
    axis.set_major_formatter(fmt)
    axis.set_major_locator(MaxNLocator(nbins=max_ticks, min_n_ticks=3))


def guardar(fig, ruta, dpi, pdf=False):
    ruta.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(ruta, dpi=dpi, facecolor="white", bbox_inches="tight", pad_inches=0.11)
    if pdf:
        fig.savefig(ruta.with_suffix(".pdf"), facecolor="white",
                    bbox_inches="tight", pad_inches=0.11)
    plt.close(fig)


def histograma(var, a, b, ruta, dpi, pdf):
    """Todos los datos, histograma con bins/axes idénticos entre paneles."""
    x = a[var].to_numpy()
    y = b[var].to_numpy()
    lo, hi = limites_todos(x, y, extra=0.0)
    if lo == hi:
        lo, hi = limites_todos(x, y)
    # Rango exacto con los verdaderos extremos; NO quitar colas.
    bins = np.linspace(lo, hi, 91)
    if len(np.unique(bins)) < 91:
        bins = np.linspace(*limites_todos(x, y), 91)
    h1, _ = np.histogram(x, bins=bins)
    h2, _ = np.histogram(y, bins=bins)

    fig, axes = plt.subplots(1, 2, figsize=(7.2, 3.15), sharex=True, sharey=True,
                             layout="constrained")
    for ax, values, nombre, color in zip(
        axes, (x, y), ("(a) Antes de 6σ", "(b) Después de 6σ"), (ANTES, DESPUES)
    ):
        ax.hist(values, bins=bins, color=color, linewidth=0, alpha=0.92)
        ax.set_title(f"{SENALES[var][2]} · {nombre}", loc="left", pad=7)
        ax.set_xlabel("Valor registrado (unidad no informada)")
        ax.set_xlim(*limites_todos(x, y))
        ax.set_ylim(0, max(h1.max(), h2.max())*1.06 if max(h1.max(), h2.max()) else 1)
        estilo(ax, "y")
        eje_continuo(ax)
    axes[0].set_ylabel("Frecuencia (registros)")
    guardar(fig, ruta, dpi, pdf)


def caja_minmax(var, a, b, zabs, ruta, dpi, pdf):
    """Cajas min–máx con puntos superpuestos sólo para |z| >= 6.

    Los cuartiles se utilizan únicamente para describir la distribución,
    NO para identificar valores extremos ni eliminar registros.
    El umbral de 6 sigma se calcula sobre el archivo ANTES y es fijo.
    """
    x = a[var].to_numpy()
    y = b[var].to_numpy()
    fig, ax = plt.subplots(figsize=(7.0, 3.10), layout="constrained")
    result = ax.boxplot([x, y], positions=[2, 1], vert=False, widths=0.44,
                        patch_artist=True, showfliers=False, whis=(0, 100),
                        medianprops={"color":"white", "linewidth":1.15},
                        whiskerprops={"color":TEXTO, "linewidth":0.88},
                        capprops={"color":TEXTO, "linewidth":0.88},
                        boxprops={"linewidth":0.85})
    for patch, color in zip(result["boxes"], [ANTES, DESPUES]):
        patch.set_facecolor(color)
        patch.set_edgecolor(color)
        patch.set_alpha(0.94)

    # Dibujar explícitamente los registros extremos según la regla del
    # trabajo (|z| >= 6 respecto de la media y DS del archivo anterior).
    # Se colocan ligeramente por encima del bigote del grupo 'Antes'
    # para que sean visibles aunque ese bigote abarque mínimo y máximo.
    indices = np.flatnonzero(zabs[var].to_numpy() >= UMBRAL)
    if len(indices):
        # Desplazamiento vertical fijo (NO altera el valor en el eje X).
        alturas = 2.33 + np.linspace(-0.045, 0.045, len(indices))
        ax.scatter(x[indices], alturas, s=37, marker='o', color=DESTACADO,
                   edgecolors='white', linewidths=0.65, zorder=6,
                   label=f'Registros |z| ≥ 6: {len(indices)}')
        ax.legend(loc='lower right', fontsize=8, handlelength=1.1,
                  borderpad=0.3, labelspacing=0.35)

    ax.set_yticks([2, 1], labels=['Antes', 'Después'])
    ax.set_xlim(*limites_todos(x, y))  # Rango completo, sin zoom.
    ax.set_xlabel('Valor registrado (unidad no informada)')
    ax.set_title(f'{SENALES[var][2]} · caja y extremos de 6σ',
                 loc='left', pad=9, fontsize=9)
    ax.set_ylim(0.4, 2.65)
    estilo(ax, 'x')
    eje_continuo(ax)
    guardar(fig, ruta, dpi, pdf)


def serie_completa(var, antes, despues, posiciones, zabs, ruta, dpi, pdf):
    """Traza TODOS los registros sin downsampling. Posiciones originales.
    NaN en huecos eliminados para evitar conectar puntos separados.
    """
    x = antes[var].to_numpy()
    y_original = np.full(len(x), np.nan, dtype=float)
    y_original[posiciones] = despues[var].to_numpy()
    fig, axes = plt.subplots(2, 1, figsize=(7.25,4.4), sharex=True,
                             sharey=True, layout="constrained")
    rango = limites_todos(x,despues[var].to_numpy())
    numeros = np.arange(1, len(x)+1)
    for ax, vals, color, nombre in zip(
        axes, (x, y_original), (ANTES, DESPUES),
        ("(a) Antes de 6σ", "(b) Después de 6σ")
    ):
        ax.plot(numeros, vals, color=color, linewidth=0.44,
                solid_capstyle="butt", rasterized=True)
        ax.set_title(f"{SENALES[var][2]} · {nombre}", loc="left",pad=7)
        ax.set_xlim(1, len(x))
        ax.set_ylim(*rango)
        ax.set_ylabel("Valor registrado\n(unidad no informada)")
        estilo(ax,"y")
        eje_continuo(ax,"y",5)
    extremos = np.flatnonzero(zabs[var].to_numpy() >= UMBRAL)
    if len(extremos):
        axes[0].scatter(extremos+1,x[extremos],s=11,color=DESTACADO,
                        edgecolors="white",linewidths=0.35,zorder=5,
                        label=r"$|z|\geq6$")
        axes[0].legend(loc="upper right")
    axes[1].set_xlabel("Posición en el archivo filtrado original (registro)")
    axes[1].xaxis.set_major_locator(MaxNLocator(5,integer=True))
    guardar(fig,ruta,dpi,pdf)


def barritas(var, a, b, ruta_dir, dpi, pdf):
    """Una figura por métrica (NO una lámina 2x2). Eje incluye el cero."""
    x, y = a[var],b[var]
    metricas = {
        "Media": (float(x.mean()),float(y.mean())),
        "Desviacion_estandar": (float(x.std(ddof=1)),float(y.std(ddof=1))),
        "Minimo": (float(x.min()),float(y.min())),
        "Maximo": (float(x.max()),float(y.max())),
    }
    for nom, (v1, v2) in metricas.items():
        fig, ax = plt.subplots(figsize=(5.6,2.4),layout="constrained")
        ax.barh([1,0],[v1,v2],color=[ANTES,DESPUES],height=0.44)
        ax.set_yticks([1,0],labels=["Antes","Después"])
        # Comparar valores absolutos desde cero, incluso si son negativos.
        minimo = min(0,v1,v2)
        maximo = max(0,v1,v2)
        anchura = maximo-minimo
        if anchura == 0:
            anchura = max(abs(maximo),1e-9)
        margen = anchura * 0.23
        ax.set_xlim(minimo-margen,maximo+margen)
        ax.set_ylim(-0.55,1.55)
        for pos,val in ((1,v1),(0,v2)):
            offset=anchura*0.025
            if val>=0:
                ax.text(val+offset,pos, f"{val:.6g}",ha="left",va="center",fontsize=8)
            else:
                ax.text(val-offset,pos, f"{val:.6g}",ha="right",va="center",fontsize=8)
        ax.axvline(0,color=TEXTO,lw=0.75)
        ax.set_xlabel("Valor registrado (unidad no informada)")
        ax.set_title(f"{nom.replace('_',' ').capitalize()} — {SENALES[var][2]}",loc="left",pad=7)
        estilo(ax,"x")
        eje_continuo(ax)
        guardar(fig,ruta_dir/f"{nom}_{SENALES[var][2]}.png",dpi,pdf)


def barra_extremos(zabs, n_eliminadas, variables, ruta, dpi, pdf):
    conteos=np.array([int((zabs[v]>=UMBRAL).sum()) for v in variables])
    nombres=[SENALES[v][2] for v in variables]
    fig,ax=plt.subplots(figsize=(6.9,3.65), layout="constrained")
    ys=np.arange(len(nombres))
    barras=ax.barh(ys,conteos,color=[DESTACADO if c else "#B8C3CB" for c in conteos],height=.63)
    ax.set_yticks(ys,labels=nombres)
    ax.invert_yaxis()
    ax.set_xlim(0,max(1,conteos.max())*1.24)
    ax.xaxis.set_major_locator(MaxNLocator(nbins=6,integer=True))
    ax.set_xlabel("Observaciones con |z| ≥ 6 (por variable)")
    ax.set_title(f"Observaciones extremas por señal · {n_eliminadas} filas únicas excluidas",
                 loc="left",pad=7)
    estilo(ax,"x")
    for rect,num in zip(barras,conteos):
        ax.text(rect.get_width()+max(0.07,float(conteos.max())*0.022),
                rect.get_y()+rect.get_height()/2,str(num),ha="left",va="center",fontsize=8)
    guardar(fig,ruta,dpi,pdf)


def validar_registros(a,b):
    mu=a.mean()
    ds=a.std(ddof=1).replace(0,np.nan)
    zabs=((a-mu)/ds).abs().fillna(0)
    excluidos=zabs.ge(UMBRAL).any(axis=1)
    esperado=a.loc[~excluidos].reset_index(drop=True)
    if esperado.shape!=b.shape or not np.allclose(esperado.to_numpy(),b.to_numpy(),rtol=1e-9,atol=1e-11):
        raise ValueError("El Excel depurado no coincide con el filtrado al aplicar |z|>=6. "
                         "Comprueba que ambos pertenecen a la misma limpieza.")
    posiciones=np.flatnonzero(~excluidos.to_numpy())
    return zabs,excluidos,posiciones


def main():
    parser=argparse.ArgumentParser(description="Figuras científicas completas del dataset eólico")
    parser.add_argument("--antes",type=Path,default=None,help="Ruta Excel antes de 6sigma")
    parser.add_argument("--despues",type=Path,default=None,help="Ruta Excel depurado")
    parser.add_argument("--salida",type=Path,default=None,help="Carpeta de resultados")
    parser.add_argument("--dpi",type=int,default=300,help="Resolución PNG (por defecto 300)")
    parser.add_argument("--pdf",action="store_true",help="Exportar también PDF vectorial")
    parser.add_argument("--solo",nargs="+",help="Probar una o más variables: GiroZ AcelX1...")
    args=parser.parse_args()
    if args.dpi<150:
        raise ValueError("Utiliza al menos 150 ppp; se recomiendan 300 para publicación.")
    an=args.antes or entrada_filtrada(BASE)
    de=args.despues or BASE/NOMBRE_DEPURADO
    raiz=args.salida or BASE/CARPETA_RESULTADOS
    print(f"ANTES   : {an}",flush=True)
    print(f"DESPUÉS : {de}",flush=True)
    a=cargar(an)
    b=cargar(de)
    zabs,excluidos,posiciones=validar_registros(a,b)
    vs=list(SENALES)
    if args.solo:
        codigos=set(args.solo)
        vs=[v for v in vs if v in codigos or SENALES[v][2] in codigos]
        if not vs or len(vs)!=len(codigos):
            raise ValueError("Variable desconocida en --solo (prueba: --solo GiroZ)")
    for sub in SUBCARPETAS.values():
        (raiz/sub).mkdir(parents=True,exist_ok=True)
    print(f"Registros: {len(a)} antes, {len(b)} después; {int(excluidos.sum())} excluidos.",flush=True)
    for idx,var in enumerate(vs,1):
        nombre=SENALES[var][2]
        print(f"[{idx}/{len(vs)}] {nombre}",flush=True)
        histograma(var,a,b,raiz/SUBCARPETAS['hist']/f"Histograma_{nombre}.png",args.dpi,args.pdf)
        caja_minmax(var,a,b,zabs,raiz/SUBCARPETAS['box']/f"Caja_{nombre}.png",args.dpi,args.pdf)
        serie_completa(var,a,b,posiciones,zabs,raiz/SUBCARPETAS['serie']/f"Serie_{nombre}.png",args.dpi,args.pdf)
        barritas(var,a,b,raiz/SUBCARPETAS['barras'],args.dpi,args.pdf)
    barra_extremos(zabs,int(excluidos.sum()),vs,
                   raiz/SUBCARPETAS['conteo']/"Barras_Registros_6sigma.png",args.dpi,args.pdf)
    resumen = pd.DataFrame({
        "Variable":vs,
        "minimo_antes":[a[v].min() for v in vs],
        "maximo_antes":[a[v].max() for v in vs],
        "minimo_despues":[b[v].min() for v in vs],
        "maximo_despues":[b[v].max() for v in vs],
        "z_ge_6":[int((zabs[v]>=UMBRAL).sum()) for v in vs],
    })
    resumen.to_csv(raiz/"00_Resumen_rangos.csv",index=False,encoding="utf-8-sig")
    texto=f"""FIGURAS EÓLICAS — ESTILO CIENTÍFICO (SIN ZOOM)
===============================================
Excel antes: {an.name}
Excel después: {de.name}
Registros antes/después: {len(a)} / {len(b)}
Filas excluidas: {int(excluidos.sum())}
Variables físicas: {len(vs)} (smp no se grafica)
Figuras PNG: {len(vs)*7+1}

Todos los valores de ambos Excel se incluyen en los histogramas y los gráficos
muestran sus rangos COMPLETOS, sin ampliación, sin recortes y sin ejes rotos.
Las series dibujan TODOS los puntos en el orden del archivo original.
El índice de registro NO equivale a tiempo porque no se conoce el intervalo
real entre mediciones. La serie depurada deja huecos en las posiciones excluidas.
Los diagramas de caja muestran bigotes en el MÍNIMO y MÁXIMO (whis=0,100).
Los puntos terracota se dibujan explícitamente cuando |z|>=6, con valores
de media y desviación estándar del dataset ANTES; no son atípicos según RIC.
La caja usa Q1, mediana y Q3 sólo como descriptores, nunca para depuración.

Se usa la misma escala en ambos paneles de cada figura. Se respeta el cero
al comparar métricas mediante barras, incluso para valores negativos.
Las unidades originales de los sensores NO están confirmadas; no se inventan.
La detección |z|>=6 se reproduce exclusivamente para verificar que los dos
Excel corresponden a la misma limpieza; NO modifica los archivos.
La presencia de registros extremos NO demuestra por sí misma una falla física.
"""
    (raiz/"00_LEEME.txt").write_text(texto,encoding="utf-8")
    print(f"FINALIZADO: {len(vs)*7+1} PNG en {raiz}",flush=True)


if __name__=="__main__":
    main()
