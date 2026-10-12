# -*- coding: utf-8 -*-
"""Genera 30 gráficos separados ANTES de la limpieza:
- 10 histogramas
- 10 diagramas de caja
- 10 gráficos de comportamiento por muestra

Usa el Excel ORIGINAL y no modifica los datos.
"""
from pathlib import Path
import argparse
import sys
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.ticker import MaxNLocator

SENSORES = [
    'AcelX1_SisGenEol', 'AcelY1_SisGenEol', 'AcelZ1_SisGenEol',
    'GiroX_SisGenEol', 'GiroY_SisGenEol', 'GiroZ_SisGenEol',
    'MagneX_SisGenEol', 'MagneY_SisGenEol', 'MagneZ_SisGenEol',
    'Ter_SisGenEol',
]

ROSA = '#CE5A91'
ROSA_FUERTE = '#93345F'
ROSA_CLARO = '#F4B6D2'
ROSA_BANDA = '#F8CADC'
TEXTO = '#2F2F2F'
SECUNDARIO = '#5E5E5E'
EJES = '#D9D9D9'
GRILLA = '#ECECEC'


def nombre_corto(nombre):
    return nombre.replace('_SisGenEol', '')


def formato_num(v):
    if not np.isfinite(v):
        return 'N/D'
    if v == 0:
        return '0'
    if abs(v) >= 1e4 or abs(v) < 1e-3:
        return f'{v:.3g}'
    return f'{v:.6g}'


def zoom_visual(x):
    a, b = np.percentile(x, [0.5, 99.5])
    if not np.isfinite(a) or not np.isfinite(b):
        a, b = float(np.min(x)), float(np.max(x))
    if a == b:
        a, b = float(np.min(x)), float(np.max(x))
    if a == b:
        margen = 0.05 * abs(a) if a else 1.0
        return a - margen, b + margen
    margen = .06 * (b - a)
    return a - margen, b + margen


def fondo_y_ejes(ax, grid='y'):
    ax.set_facecolor('white')
    for spine in ax.spines.values():
        spine.set_color(EJES)
    ax.tick_params(axis='both', colors=SECUNDARIO, labelsize=10)
    ax.set_axisbelow(True)
    if grid == 'x':
        ax.grid(axis='x', color=GRILLA, lw=0.8)
    elif grid == 'both':
        ax.grid(color=GRILLA, lw=0.8)
    else:
        ax.grid(axis='y', color=GRILLA, lw=0.8)
    ax.xaxis.set_major_locator(MaxNLocator(nbins=6))


def estadisticas_base(s):
    x = s.dropna().to_numpy(dtype=float)
    q1, q3 = np.quantile(x, [.25, .75])
    iqr = q3 - q1
    li = q1 - 1.5 * iqr
    ls = q3 + 1.5 * iqr
    atipicos = int(((x < li) | (x > ls)).sum())
    return {
        'x': x,
        'media': float(x.mean()),
        'ds': float(x.std(ddof=1)),
        'mediana': float(np.median(x)),
        'q1': float(q1),
        'q3': float(q3),
        'iqr': float(iqr),
        'li': float(li),
        'ls': float(ls),
        'atipicos': atipicos,
    }


def encabezado(fig, sensor, tipo, total, stats_line):
    """Cabecera con cuatro niveles y distancias fijas: evita superposiciones."""
    fig.text(.085, .961, 'ANÁLISIS EXPLORATORIO  ·  ANTES DE LA LIMPIEZA',
             ha='left', va='top', fontsize=10, weight='bold', color=ROSA_FUERTE)
    fig.text(.085, .902, f'{tipo} — {sensor}',
             ha='left', va='top', fontsize=19, weight='bold', color=TEXTO)
    fig.text(.085, .817, f'Datos originales  ·  {total:,} registros'.replace(',', '.'),
             ha='left', va='top', fontsize=11, color=SECUNDARIO)
    fig.text(.085, .765, stats_line,
             ha='left', va='top', fontsize=10, color=SECUNDARIO)
    # Separador tenue para distinguir inequívocamente cabecera y datos.
    fig.add_artist(plt.Line2D([.085, .965], [.713, .713],
                  transform=fig.transFigure, color=GRILLA, lw=1))


def pie(fig, texto):
    fig.text(.085, .052, texto, fontsize=9.1, color=SECUNDARIO)


def grafico_histograma(sensor, serie, total, destino):
    st = estadisticas_base(serie)
    x = st['x']
    a, b = zoom_visual(x)
    visibles = x[(x >= a) & (x <= b)]

    fig, ax = plt.subplots(figsize=(11.2, 6.6), facecolor='white')
    fig.subplots_adjust(left=.085, right=.965, top=.677, bottom=.16)
    encabezado(fig, sensor, 'Histograma', total,
               f"Media = {formato_num(st['media'])}   |   Mediana = {formato_num(st['mediana'])}   |   DS = {formato_num(st['ds'])}")
    ax.hist(visibles, bins=50, color=ROSA, edgecolor='white', linewidth=.35)
    ax.axvline(st['mediana'], color=ROSA_FUERTE, lw=1.8, ls='--', label='Mediana')
    ax.set_xlim(a, b)
    ax.set_xlabel('Valor registrado (unidad por confirmar)', color=SECUNDARIO, fontsize=10)
    ax.set_ylabel('Frecuencia', color=SECUNDARIO, fontsize=10)
    fondo_y_ejes(ax, 'y')
    ax.legend(loc='upper left', frameon=False, fontsize=9, labelcolor=SECUNDARIO)
    ax.text(.98, .95, f"Fuera de vista: {len(x)-len(visibles):,}".replace(',', '.'),
            ha='right', va='top', transform=ax.transAxes, fontsize=9.5, color=SECUNDARIO)
    pie(fig, 'Datos originales sin corregir. La vista usa zoom visual P0,5–P99,5; los valores no se eliminan.')
    archivo = destino / f'Histograma_{sensor}_ANTES_LIMPIEZA.png'
    fig.savefig(archivo, dpi=180, facecolor='white')
    plt.close(fig)


def grafico_caja(sensor, serie, total, destino):
    st = estadisticas_base(serie)
    x = st['x']
    a, b = zoom_visual(x)

    fig, ax = plt.subplots(figsize=(11.2, 6.6), facecolor='white')
    fig.subplots_adjust(left=.085, right=.965, top=.677, bottom=.17)
    encabezado(fig, sensor, 'Diagrama de caja', total,
               f"Media = {formato_num(st['media'])}   |   Mediana = {formato_num(st['mediana'])}   |   DS = {formato_num(st['ds'])}")
    ax.boxplot(x, vert=False, widths=.32, patch_artist=True,
               boxprops={'facecolor': ROSA_CLARO, 'edgecolor': ROSA_FUERTE, 'linewidth': 1.3},
               medianprops={'color': ROSA_FUERTE, 'linewidth': 2.2},
               whiskerprops={'color': ROSA, 'linewidth': 1.25},
               capprops={'color': ROSA, 'linewidth': 1.25},
               flierprops={'marker': '.', 'markersize': 2,
                           'markeredgecolor': ROSA, 'alpha': .38})
    ax.set_xlim(a, b)
    ax.set_ylim(.65, 1.35)
    ax.set_yticks([])
    ax.set_xlabel('Valor registrado (unidad por confirmar)', color=SECUNDARIO, fontsize=10)
    fondo_y_ejes(ax, 'x')
    ax.text(.98, .93, f"Atípicos IQR: {st['atipicos']:,}".replace(',', '.'),
            ha='right', va='top', transform=ax.transAxes, fontsize=10, color=SECUNDARIO)
    ax.text(.98, .83, f"Q1 = {formato_num(st['q1'])}   |   Q3 = {formato_num(st['q3'])}   |   IQR = {formato_num(st['iqr'])}",
            ha='right', va='top', transform=ax.transAxes, fontsize=9.4, color=SECUNDARIO)
    ax.text(.98, .73, f"Límites IQR: {formato_num(st['li'])} a {formato_num(st['ls'])}",
            ha='right', va='top', transform=ax.transAxes, fontsize=9.4, color=SECUNDARIO)
    pie(fig, 'Datos originales sin corregir. Algunos registros pueden estar desalineados. El zoom es solo visual.')
    archivo = destino / f'Caja_{sensor}_ANTES_LIMPIEZA.png'
    fig.savefig(archivo, dpi=180, facecolor='white')
    plt.close(fig)


def grafico_evolucion(sensor, serie, total, destino):
    st = estadisticas_base(serie)
    datos = serie.to_numpy(dtype=float)
    x = st['x']
    a, b = zoom_visual(x)
    tam_bloque = max(1, int(np.ceil(total / 550)))
    puntos_x, medianas, p10s, p90s = [], [], [], []
    for ini in range(0, total, tam_bloque):
        vals = datos[ini:ini + tam_bloque]
        vals = vals[np.isfinite(vals)]
        if not len(vals):
            continue
        puntos_x.append(ini + (min(ini + tam_bloque, total) - ini) / 2)
        medianas.append(float(np.median(vals)))
        p10s.append(float(np.percentile(vals, 10)))
        p90s.append(float(np.percentile(vals, 90)))
    puntos_x = np.asarray(puntos_x)
    medianas = np.asarray(medianas)
    p10s = np.asarray(p10s)
    p90s = np.asarray(p90s)

    fig, ax = plt.subplots(figsize=(11.2, 6.6), facecolor='white')
    fig.subplots_adjust(left=.085, right=.965, top=.677, bottom=.16)
    encabezado(fig, sensor, 'Comportamiento por muestra', total,
               f"Media = {formato_num(st['media'])}   |   Mediana = {formato_num(st['mediana'])}   |   DS = {formato_num(st['ds'])}")
    ax.fill_between(puntos_x, p10s, p90s, color=ROSA_BANDA, alpha=.9, linewidth=0, label='P10–P90 por bloque')
    ax.plot(puntos_x, medianas, color=ROSA_FUERTE, lw=1.35, label='Mediana por bloque')
    ax.set_xlim(1, total)
    ax.set_ylim(a, b)
    ax.set_xlabel('Orden de muestra (no equivale a segundos)', color=SECUNDARIO, fontsize=10)
    ax.set_ylabel('Valor registrado', color=SECUNDARIO, fontsize=10)
    fondo_y_ejes(ax, 'y')
    ax.legend(loc='upper left', frameon=False, fontsize=9, labelcolor=SECUNDARIO)
    ax.text(.98, .95, f"Fuera del zoom: {np.count_nonzero((x < a) | (x > b)):,}".replace(',', '.'),
            ha='right', va='top', transform=ax.transAxes, fontsize=9.5, color=SECUNDARIO)
    pie(fig, 'Resumen gráfico en bloques sucesivos. Eje X = orden de muestra; no equivale a tiempo real en segundos.')
    archivo = destino / f'Evolucion_{sensor}_ANTES_LIMPIEZA.png'
    fig.savefig(archivo, dpi=180, facecolor='white')
    plt.close(fig)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('archivo', nargs='?', help='Ruta opcional del Excel ORIGINAL')
    args = parser.parse_args()
    carpeta = Path(__file__).resolve().parent
    if args.archivo:
        entrada = Path(args.archivo).resolve()
    else:
        candidatos = sorted(p for p in carpeta.glob('CF-2023-I-204-1230_DAT*.xlsx') if not p.name.startswith('~$'))
        if len(candidatos) != 1:
            raise ValueError('Deja un único Excel original CF-2023-I-204-1230_DAT*.xlsx en esta carpeta o indica la ruta.')
        entrada = candidatos[0]
    if not entrada.is_file():
        raise FileNotFoundError(f'No existe el archivo: {entrada}')

    print(f'Leyendo Excel original: {entrada.name}', flush=True)
    df = pd.read_excel(entrada, sheet_name=0, engine='openpyxl')
    faltantes = [n for n in SENSORES if n not in df.columns]
    if faltantes:
        raise ValueError(f'Faltan sensores: {faltantes}')
    sensores = df[SENSORES].apply(pd.to_numeric, errors='raise').replace([np.inf, -np.inf], np.nan)
    total = len(sensores)

    raiz = carpeta / 'Graficos_ANTES_Limpieza_Separados_Titulos_Corregidos'
    carpetas = {
        'hist': raiz / 'Histogramas',
        'caja': raiz / 'Cajas',
        'evo': raiz / 'Evolucion',
    }
    for c in carpetas.values():
        c.mkdir(parents=True, exist_ok=True)

    for nombre in SENSORES:
        corto = nombre_corto(nombre)
        serie = sensores[nombre]
        grafico_histograma(corto, serie, total, carpetas['hist'])
        grafico_caja(corto, serie, total, carpetas['caja'])
        grafico_evolucion(corto, serie, total, carpetas['evo'])
        print(f'Listo: {corto}', flush=True)

    (raiz / 'LEEME.txt').write_text(
        'Gráficos separados - ANTES de la limpieza\n\n'
        'Se generaron 30 gráficos: 10 histogramas, 10 diagramas de caja y 10 gráficos de comportamiento por muestra.\n'
        'Todos usan el Excel original sin corregir. No se trataron nulos, atípicos ni desalineaciones.\n'
        'El diseño usa fondo blanco y paleta rosada solo en los elementos del gráfico.\n',
        encoding='utf-8'
    )
    print(f'Se guardaron 30 gráficos en {raiz}')


if __name__ == '__main__':
    try:
        main()
    except (FileNotFoundError, ValueError, PermissionError) as err:
        print(f'ERROR: {err}')
        sys.exit(1)
