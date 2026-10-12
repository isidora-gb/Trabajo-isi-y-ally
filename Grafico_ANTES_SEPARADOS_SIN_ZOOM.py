from pathlib import Path
import sys
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.ticker import MaxNLocator

ETAPA_LARGA = 'antes de la limpieza'
PREFIJO = 'ANTES'
HOJA = 0
ARCHIVOS_PREFERIDOS = ['CF-2023-I-204-1230_DAT.xlsx','CF-2023-I-204-1230_DAT(3).xlsx','CF-2023-I-204-1230_DAT(2).xlsx']
PATRON = 'CF-2023-I-204-1230_DAT*.xlsx'
SENSORES = [
    'AcelX1_SisGenEol', 'AcelY1_SisGenEol', 'AcelZ1_SisGenEol',
    'GiroX_SisGenEol', 'GiroY_SisGenEol', 'GiroZ_SisGenEol',
    'MagneX_SisGenEol', 'MagneY_SisGenEol', 'MagneZ_SisGenEol',
    'Ter_SisGenEol'
]

ROSA = '#CE5A91'
ROSA_FUERTE = '#93345F'
ROSA_CLARO = '#F4B6D2'
ROSA_BANDA = '#F8CADC'
TEXTO = '#292929'
SECUNDARIO = '#575757'
REJILLA = '#E9E9E9'


def corto(nombre):
    return nombre.replace('_SisGenEol', '')


def formatear(v):
    if not np.isfinite(v):
        return 'N/D'
    if v == 0:
        return '0'
    return f'{v:.5g}' if abs(v) >= .001 else f'{v:.3g}'


def rango_completo(s):
    a, b = np.nanmin(s), np.nanmax(s)
    if a == b:
        margen = 0.05 * abs(a) if a else 1
        return a - margen, b + margen
    dif = b - a
    margen = .02 * dif
    return a - margen, b + margen


def estilo(ax, caja=False):
    ax.set_facecolor('white')
    ax.set_axisbelow(True)
    ax.grid(axis='x' if caja else 'y', color=REJILLA, linewidth=.8)
    for sp in ax.spines.values():
        sp.set_color('#D7D7D7')
    ax.tick_params(labelsize=9, colors=SECUNDARIO)
    ax.xaxis.set_major_locator(MaxNLocator(5))
    if not caja:
        ax.yaxis.set_major_locator(MaxNLocator(5))
    ax.xaxis.label.set_color(SECUNDARIO)
    ax.yaxis.label.set_color(SECUNDARIO)


def estadisticas(s):
    x = s.dropna().to_numpy(dtype=float)
    q1, q3 = np.quantile(x, [.25, .75])
    iqr = q3 - q1
    atip = np.count_nonzero((x < q1 - 1.5 * iqr) | (x > q3 + 1.5 * iqr))
    return x, np.mean(x), np.std(x, ddof=1), np.median(x), int(atip)


def crear_figura_base(sensor, total, subtitulo):
    fig, ax = plt.subplots(figsize=(10.5, 6.0), facecolor='white')
    fig.text(.08, .965, f'ANÁLISIS EXPLORATORIO  ·  {ETAPA_LARGA.upper()}',
             ha='left', va='top', color=ROSA_FUERTE, fontsize=10.5, weight='bold')
    fig.text(.08, .92, subtitulo, ha='left', va='top', color=TEXTO, fontsize=18, weight='bold')
    fig.text(.08, .875, f'{corto(sensor)}  ·  {total:,} registros'.replace(',', '.'),
             ha='left', va='top', color=SECUNDARIO, fontsize=11)
    fig.add_artist(plt.Line2D([.08, .965], [.845, .845], transform=fig.transFigure,
                              color=REJILLA, lw=1))
    fig.subplots_adjust(left=.10, right=.95, top=.78, bottom=.16)
    return fig, ax


def grafico_hist(sensor, serie, destino, total):
    x, mu, ds, mediana, _ = estadisticas(serie)
    low, high = rango_completo(x)
    fig, ax = crear_figura_base(sensor, total, 'Distribución de los registros')
    ax.set_title(f'Media {formatear(mu)}   ·   DS {formatear(ds)}', loc='left', pad=12,
                 color=TEXTO, fontsize=10.5, weight='medium')
    ax.hist(x, bins=48, color=ROSA, edgecolor='white', linewidth=.35)
    ax.axvline(mediana, color=ROSA_FUERTE, linestyle='--', linewidth=1.6)
    ax.set_xlim(low, high)
    ax.set_xlabel('Valor registrado', fontsize=9)
    ax.set_ylabel('Frecuencia', fontsize=9)
    estilo(ax)
    fig.text(.08, .06, 'Histograma sin zoom: se muestra el rango completo de valores. Las estadísticas se calcularon con todos los valores no nulos.',
             fontsize=9.1, color=SECUNDARIO, ha='left')
    ruta = destino / f'Histograma_{corto(sensor)}.png'
    fig.savefig(ruta, dpi=170, facecolor='white')
    plt.close(fig)


def grafico_caja(sensor, serie, destino, total):
    x, mu, ds, mediana, atip = estadisticas(serie)
    low, high = rango_completo(x)
    fig, ax = crear_figura_base(sensor, total, 'Dispersión y valores atípicos')
    ax.set_title(f'Media {formatear(mu)}   ·   DS {formatear(ds)}', loc='left', pad=12,
                 color=TEXTO, fontsize=10.5, weight='medium')
    ax.boxplot(x, vert=False, widths=.33, patch_artist=True,
               boxprops={'facecolor': ROSA_CLARO, 'edgecolor': ROSA_FUERTE, 'linewidth': 1.3},
               medianprops={'color': ROSA_FUERTE, 'linewidth': 2},
               whiskerprops={'color': ROSA, 'linewidth': 1.2},
               capprops={'color': ROSA, 'linewidth': 1.2},
               flierprops={'marker': '.', 'markersize': 1.8, 'markeredgecolor': ROSA, 'alpha': .4})
    ax.set_xlim(low, high)
    ax.set_yticks([])
    ax.set_xlabel('Valor registrado', fontsize=9)
    estilo(ax, caja=True)
    ax.text(.98, .95, f'Atípicos IQR: {atip:,}'.replace(',', '.'), transform=ax.transAxes,
            ha='right', va='top', fontsize=8.8, color=SECUNDARIO)
    fig.text(.08, .06, 'Atípicos por criterio IQR (1,5×IQR). Se muestra el rango completo; no se eliminaron observaciones.',
             fontsize=9.1, color=SECUNDARIO, ha='left')
    ruta = destino / f'Caja_{corto(sensor)}.png'
    fig.savefig(ruta, dpi=170, facecolor='white')
    plt.close(fig)


def grafico_evol(sensor, serie, destino, total):
    x, mu, ds, mediana, _ = estadisticas(serie)
    low, high = rango_completo(x)
    fig, ax = crear_figura_base(sensor, total, 'Comportamiento por orden de muestra')
    ax.set_title(f'Media {formatear(mu)}   ·   DS {formatear(ds)}', loc='left', pad=12,
                 color=TEXTO, fontsize=10.5, weight='medium')
    datos = serie.to_numpy(dtype=float)
    tam = max(1, int(np.ceil(len(datos) / 550)))
    xp, central, lowp, highp = [], [], [], []
    for j in range(0, len(datos), tam):
        bloque = datos[j:j+tam]
        bloque = bloque[np.isfinite(bloque)]
        if len(bloque) == 0:
            continue
        xp.append(j + len(datos[j:j+tam]) / 2)
        central.append(float(np.median(bloque)))
        lowp.append(float(np.percentile(bloque, 10)))
        highp.append(float(np.percentile(bloque, 90)))
    ax.fill_between(xp, lowp, highp, color=ROSA_BANDA, alpha=.88, label='P10–P90 por bloque')
    ax.plot(xp, central, color=ROSA_FUERTE, lw=1.3, label='Mediana por bloque')
    ax.set_xlim(1, len(datos))
    ax.set_ylim(low, high)
    ax.set_xlabel('Orden de muestra (no segundos)', fontsize=9)
    ax.set_ylabel('Valor registrado', fontsize=9)
    estilo(ax)
    ax.legend(loc='upper left', frameon=False, fontsize=8)
    fig.text(.08, .06, 'Bandas P10–P90 y mediana por bloques. El eje horizontal indica orden, no tiempo real.',
             fontsize=9.1, color=SECUNDARIO, ha='left')
    ruta = destino / f'Evolucion_{corto(sensor)}.png'
    fig.savefig(ruta, dpi=170, facecolor='white')
    plt.close(fig)


def main():
    carpeta = Path(__file__).resolve().parent
    import argparse
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('archivo', nargs='?', help='Ruta al Excel de esta etapa (opcional)')
    args = parser.parse_args()
    if args.archivo:
        entrada = Path(args.archivo).resolve()
    else:
        entrada = next((carpeta / n for n in ARCHIVOS_PREFERIDOS if (carpeta / n).is_file()), None)
        if entrada is None:
            candidatos = sorted(p for p in carpeta.glob(PATRON) if not p.name.startswith('~$'))
            if len(candidatos) != 1:
                raise FileNotFoundError('Coloca el Excel correspondiente en esta carpeta o indica su ruta.')
            entrada = candidatos[0]
    print(f'Leyendo: {entrada.name}', flush=True)
    tabla = pd.read_excel(entrada, sheet_name=HOJA, engine='openpyxl')
    faltantes = [x for x in SENSORES if x not in tabla.columns]
    if faltantes:
        raise ValueError(f'No se encuentran estas columnas: {faltantes}')
    tabla = tabla[SENSORES].apply(pd.to_numeric, errors='raise').replace([np.inf, -np.inf], np.nan)
    total = len(tabla)
    raiz = carpeta / 'Graficos_ANTES_Separados_SIN_ZOOM'
    hist_dir = raiz / 'Histogramas'
    cajas_dir = raiz / 'Cajas'
    evol_dir = raiz / 'Evolucion'
    for d in (hist_dir, cajas_dir, evol_dir):
        d.mkdir(parents=True, exist_ok=True)
    for s in SENSORES:
        grafico_hist(s, tabla[s], hist_dir, total)
        grafico_caja(s, tabla[s], cajas_dir, total)
        grafico_evol(s, tabla[s], evol_dir, total)
        print('Generado:', corto(s), flush=True)
    print(f'LISTO: 30 figuras guardadas en {raiz}', flush=True)


if __name__ == '__main__':
    try:
        main()
    except (FileNotFoundError, ValueError, PermissionError) as e:
        print(f'ERROR: {e}')
        sys.exit(1)
