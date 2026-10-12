# -*- coding: utf-8 -*-
"""Tarea 2: gráficos agrupados, cuatro subgráficos como máximo por imagen.
Etapa: antes de la limpieza.

Ejecutar en VS Code: python Grafico_ANTES_Agrupados_Corregido.py
El código busca el Excel correspondiente en su misma carpeta y nunca lo modifica.
Requisitos: pandas, numpy, openpyxl, matplotlib.
"""
from pathlib import Path
import sys
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.ticker import MaxNLocator

ETAPA = 'antes'
ETAPA_LARGA = 'antes de la limpieza'
PREFIJO = 'ANTES'
HOJA = 0
ARCHIVOS_PREFERIDOS = ['CF-2023-I-204-1230_DAT.xlsx','CF-2023-I-204-1230_DAT(3).xlsx','CF-2023-I-204-1230_DAT(2).xlsx']
PATRON = 'CF-2023-I-204-1230_DAT*.xlsx'

GRUPOS = [
    ('Acelerómetros', ['AcelX1_SisGenEol', 'AcelY1_SisGenEol', 'AcelZ1_SisGenEol']),
    ('Giroscopios', ['GiroX_SisGenEol', 'GiroY_SisGenEol', 'GiroZ_SisGenEol']),
    ('Magnetómetros y temperatura', ['MagneX_SisGenEol', 'MagneY_SisGenEol',
                                    'MagneZ_SisGenEol', 'Ter_SisGenEol']),
]
SENSORES = [sensor for _, grupo in GRUPOS for sensor in grupo]

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


def zoom(s):
    a, b = np.nanpercentile(s, [0.5, 99.5])
    if a == b:
        a, b = np.nanmin(s), np.nanmax(s)
    dif = b - a
    margen = .055 * dif if dif > 0 else (0.05 * abs(a) if a else 1)
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
    iqr = q3-q1
    atip = np.count_nonzero((x < q1-1.5*iqr) | (x > q3+1.5*iqr))
    return x, np.mean(x), np.std(x, ddof=1), np.median(x), int(atip)


def lamina(tipo, familia, sensores, tabla, destino):
    nombres_tipos = {
        'hist': 'Distribución de los registros',
        'caja': 'Dispersión y valores atípicos',
        'evol': 'Comportamiento por orden de muestra',
    }
    fig, axs = plt.subplots(2, 2, figsize=(14, 9.3), facecolor='white')
    axs = list(axs.flat)
    fig.text(.07, .972, f'ANÁLISIS EXPLORATORIO  ·  {ETAPA_LARGA.upper()}',
             ha='left', va='top', color=ROSA_FUERTE, fontsize=10.5, weight='bold')
    fig.text(.07, .925, nombres_tipos[tipo],
             ha='left', va='top', color=TEXTO, fontsize=19, weight='bold')
    fig.text(.07, .869, f'{familia}  ·  {len(tabla):,} registros  ·  {len(sensores)} sensores'.replace(',', '.'),
             ha='left', va='top', color=SECUNDARIO, fontsize=11)
    fig.add_artist(plt.Line2D([.07, .965], [.831, .831], transform=fig.transFigure,
                              color=REJILLA, lw=1))
    # Reserva superior suficiente para cabecera y títulos de los subgráficos.
    fig.subplots_adjust(left=.075, right=.975, top=.757, bottom=.123,
                        hspace=.52, wspace=.28)

    for idx, nombre in enumerate(sensores):
        ax = axs[idx]
        serie = tabla[nombre]
        x, mu, ds, mediana, atip = estadisticas(serie)
        low, high = zoom(x)
        statsline = f'Media {formatear(mu)}   ·   DS {formatear(ds)}'
        ax.set_title(f'{corto(nombre)}   |   {statsline}', loc='left', pad=13,
                     color=TEXTO, fontsize=10.5, weight='medium')

        if tipo == 'hist':
            visible = x[(x >= low) & (x <= high)]
            ax.hist(visible, bins=48, color=ROSA, edgecolor='white', linewidth=.35)
            ax.axvline(mediana, color=ROSA_FUERTE, linestyle='--', linewidth=1.6)
            ax.set_xlim(low, high)
            ax.set_xlabel('Valor registrado', fontsize=9)
            ax.set_ylabel('Frecuencia', fontsize=9)
            estilo(ax)
            ax.text(.98, .95, f'Fuera de vista: {len(x)-len(visible):,}'.replace(',', '.'),
                    transform=ax.transAxes, ha='right', va='top', fontsize=8.8,
                    color=SECUNDARIO)
        elif tipo == 'caja':
            ax.boxplot(x, vert=False, widths=.33, patch_artist=True,
                       boxprops={'facecolor': ROSA_CLARO, 'edgecolor': ROSA_FUERTE,
                                 'linewidth': 1.3},
                       medianprops={'color': ROSA_FUERTE, 'linewidth': 2},
                       whiskerprops={'color': ROSA, 'linewidth': 1.2},
                       capprops={'color': ROSA, 'linewidth': 1.2},
                       flierprops={'marker': '.', 'markersize': 1.8,
                                   'markeredgecolor': ROSA, 'alpha': .4})
            ax.set_xlim(low, high)
            ax.set_yticks([])
            ax.set_xlabel('Valor registrado', fontsize=9)
            estilo(ax, caja=True)
            ax.text(.98, .96, f'Atípicos IQR: {atip:,}'.replace(',', '.'),
                    transform=ax.transAxes, ha='right', va='top', fontsize=8.8,
                    color=SECUNDARIO)
        else:
            datos = serie.to_numpy(dtype=float)
            tam = max(1, int(np.ceil(len(datos)/550)))
            xp, central, lowp, highp = [], [], [], []
            for j in range(0, len(datos), tam):
                bloque = datos[j:j+tam]
                bloque = bloque[np.isfinite(bloque)]
                if len(bloque) == 0:
                    continue
                xp.append(j + len(datos[j:j+tam])/2)
                central.append(float(np.median(bloque)))
                lowp.append(float(np.percentile(bloque, 10)))
                highp.append(float(np.percentile(bloque, 90)))
            ax.fill_between(xp, lowp, highp, color=ROSA_BANDA, alpha=.88,
                            label='P10–P90 por bloque')
            ax.plot(xp, central, color=ROSA_FUERTE, lw=1.3, label='Mediana por bloque')
            ax.set_xlim(1, len(datos))
            ax.set_ylim(low, high)
            ax.set_xlabel('Orden de muestra (no segundos)', fontsize=9)
            ax.set_ylabel('Valor registrado', fontsize=9)
            estilo(ax)
            ax.legend(loc='upper left', frameon=False, fontsize=8)

    for ax in axs[len(sensores):]:
        ax.set_visible(False)

    pie = {
        'hist': 'Histogramas: zoom visual P0,5–P99,5. Las estadísticas se calcularon con todos los valores no nulos.',
        'caja': 'Atípicos por criterio IQR (1,5×IQR). El zoom es solo visual; no se eliminaron observaciones.',
        'evol': 'Bandas P10–P90 y mediana por bloques. El eje horizontal indica orden, no tiempo real.',
    }[tipo]
    fig.text(.07, .065, pie, fontsize=9.3, color=SECUNDARIO, ha='left')
    fig.text(.07, .042,
             'Datos originales; algunas filas pueden estar desalineadas.', fontsize=9.0, color=SECUNDARIO, ha='left')

    slug = {'hist':'01_Histogramas', 'caja':'02_Cajas', 'evol':'03_Evolucion'}[tipo]
    familia_slug = familia.lower().replace('ó','o').replace('í','i').replace(' ','_')
    path = destino / f'{slug}_{familia_slug}.png'
    fig.savefig(path, dpi=170, facecolor='white')
    plt.close(fig)
    return path


def main():
    carpeta = Path(__file__).resolve().parent
    import argparse
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('archivo', nargs='?', help='Ruta al Excel de esta etapa (opcional)')
    args = parser.parse_args()
    if args.archivo:
        entrada = Path(args.archivo).resolve()
    else:
        entrada = next((carpeta / nombre for nombre in ARCHIVOS_PREFERIDOS
                        if (carpeta / nombre).is_file()), None)
        if entrada is None:
            candidatos = sorted(p for p in carpeta.glob(PATRON) if not p.name.startswith('~$'))
            if len(candidatos) != 1:
                raise FileNotFoundError('Coloca el Excel correspondiente en esta carpeta o indica su ruta al ejecutar el programa.')
            entrada = candidatos[0]
    if not entrada.is_file():
        raise FileNotFoundError(f'No existe: {entrada}')
    print(f'Leyendo: {entrada.name}', flush=True)
    tabla = pd.read_excel(entrada, sheet_name=HOJA, engine='openpyxl')
    faltantes = [x for x in SENSORES if x not in tabla.columns]
    if faltantes:
        raise ValueError(f'No se encuentran estas columnas: {faltantes}')
    tabla = tabla[SENSORES].apply(pd.to_numeric, errors='raise').replace([np.inf,-np.inf], np.nan)
    salida = carpeta / f'Graficos_{PREFIJO}_Agrupados_Titulos_Corregidos'
    salida.mkdir(exist_ok=True)
    for tipo in ('hist','caja','evol'):
        for familia, nombres in GRUPOS:
            path = lamina(tipo, familia, nombres, tabla, salida)
            print('Generado:', path.name, flush=True)
    (salida / 'LEEME.txt').write_text(
        f'Etapa: {ETAPA_LARGA}\n9 figuras (3 tipos x 3 familias).\n'
        'Los gráficos no modifican ni eliminan registros.\n'
        'Datos originales; algunas filas pueden estar desalineadas.\n', encoding='utf-8')
    print(f'LISTO: 9 figuras guardadas en {salida}', flush=True)


if __name__ == '__main__':
    try:
        main()
    except (FileNotFoundError, ValueError, PermissionError) as e:
        print(f'ERROR: {e}')
        sys.exit(1)
