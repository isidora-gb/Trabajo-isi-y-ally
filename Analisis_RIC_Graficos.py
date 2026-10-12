# -*- coding: utf-8 -*-
"""Tarea 2 — Cuartiles, RIC, atípicos leves/extremos y gráficos COMPLETOS.

Entradas: Excel original CF-2023-I-204-1230_DAT*.xlsx y Excel limpio
Dataset_eolico_filtrado_2*.xlsx (hoja Datos, que es la primera hoja).

Salidas:
  Tabla_RIC_Atipicos_Leves_Extremos.xlsx (tablas simples, sin filtros),
  Graficos_RIC_Completos/ANTES/... y Graficos_RIC_Completos/DESPUES/...
  30 PNG por etapa, separados por Histogramas / Cajas / Evolucion.

No aplica zoom, NO elimina registros ni modifica las bases. Las señales de
Evolución dibujan TODAS las muestras reales (no agregados ni submuestreo).
El IQR / RIC se calcula por separado para cada etapa y sensor.

Requisitos en Python: numpy, matplotlib, openpyxl.
Instalación: python -m pip install numpy matplotlib openpyxl
Ejecutar: python Analisis_RIC_y_Graficos_Completos.py
Opcional: python Analisis_RIC_y_Graficos_Completos.py --original ruta.xlsx --limpio ruta.xlsx

La lectura numérica de Excel se hace con zipfile/XML estándar.
El Excel de resumen se exporta con openpyxl (sin artifact_tool).
"""

from pathlib import Path
from zipfile import ZipFile
import argparse
import xml.etree.ElementTree as ET
import re
import math
import numpy as np

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

NOMBRES = [
    'AcelX1_SisGenEol', 'AcelY1_SisGenEol', 'AcelZ1_SisGenEol',
    'GiroX_SisGenEol', 'GiroY_SisGenEol', 'GiroZ_SisGenEol',
    'MagneX_SisGenEol', 'MagneY_SisGenEol', 'MagneZ_SisGenEol',
    'Ter_SisGenEol',
]
BASE = Path(__file__).resolve().parent
# Paleta de alto contraste para diferenciar atípicos leves y extremos
BASE_GRAFICO = '#B0B7C3'
BASE_CAJA = '#D9DEE7'
BASE_CLARO = '#EEF1F5'
LINEA_CENTRAL = '#404040'
LEVE = '#F39C12'
EXTREMO = '#1F4E79'
TEXTO = '#242424'
SECUNDARIO = '#5C5C5C'
REJILLA = '#E9E9E9'


def encontrar(patron, preferidos):
    for nombre in preferidos:
        if (BASE / nombre).is_file():
            return BASE / nombre
    candidatos = sorted(p for p in BASE.glob(patron) if not p.name.startswith('~$'))
    if len(candidatos) == 1:
        return candidatos[0]
    raise ValueError(f'Se necesita un Excel de {patron}; encontrados: {[p.name for p in candidatos]}. Indica su ruta con --original o --limpio.')


def indice_columna(coordenada):
    n = 0
    for letra in re.match(r'[A-Z]+', coordenada).group(0):
        n = n * 26 + ord(letra) - 64
    return n - 1


def leer_diez_sensores_excel(ruta):
    """Extrae las 10 columnas de sensores de la primera hoja XLSX.

    Recorre los registros por XML sin construir el DataFrame completo. Respeta
    celdas realmente vacías y NO recupera los fragmentos desalineados del bruto.
    """
    ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    matriz = np.full((145000, 10), np.nan, dtype=np.float64)
    filas = 0
    with ZipFile(ruta, 'r') as z:
        with z.open('xl/worksheets/sheet1.xml') as hoja:
            # Leer por filas; limpiar cada elemento para ahorrar memoria.
            contexto = ET.iterparse(hoja, events=('start', 'end'))
            _, root = next(contexto)
            for evento, elemento in contexto:
                if evento == 'end' and elemento.tag == ns + 'row':
                    num_excel = int(elemento.attrib['r'])
                    if num_excel > 1:
                        pos = num_excel - 2
                        if pos >= matriz.shape[0]:
                            matriz = np.vstack([matriz, np.full((50000, 10), np.nan)])
                        for celda in elemento.iter(ns + 'c'):
                            ref = celda.get('r', '')
                            if not ref:
                                continue
                            j = indice_columna(ref)
                            if j >= 10:
                                continue
                            v = celda.find(ns + 'v')
                            if v is not None and v.text:
                                try:
                                    matriz[pos, j] = float(v.text)
                                except ValueError:
                                    pass
                        filas = max(filas, pos + 1)
                    elemento.clear()
                    root.clear()
    return matriz[:filas]


def formato(valor):
    if not np.isfinite(valor):
        return 'N/D'
    return f'{valor:.6g}'


def stats_ric(x):
    valores = x[np.isfinite(x)]
    if len(valores) < 2:
        raise ValueError('No hay suficientes valores numéricos para el sensor.')
    q1, q3 = np.percentile(valores, [25, 75])
    ric = q3 - q1
    li15 = q1 - 1.5 * ric
    ls15 = q3 + 1.5 * ric
    li3 = q1 - 3 * ric
    ls3 = q3 + 3 * ric
    leves_inf = (valores < li15) & (valores >= li3)
    leves_sup = (valores > ls15) & (valores <= ls3)
    extremos_inf = valores < li3
    extremos_sup = valores > ls3
    normales = (valores >= li15) & (valores <= ls15)
    resumen = {
        'N': int(len(valores)), 'Nulos': int(len(x) - len(valores)),
        'media': float(valores.mean()), 'DS': float(valores.std(ddof=1)),
        'min': float(valores.min()), 'max': float(valores.max()),
        'q1': float(q1), 'q3': float(q3), 'ric': float(ric),
        'li15': float(li15), 'ls15': float(ls15),
        'li3': float(li3), 'ls3': float(ls3),
        'habituales': int(normales.sum()),
        'leves_inf': int(leves_inf.sum()), 'leves_sup': int(leves_sup.sum()),
        'leves': int(leves_inf.sum() + leves_sup.sum()),
        'extremos_inf': int(extremos_inf.sum()), 'extremos_sup': int(extremos_sup.sum()),
        'extremos': int(extremos_inf.sum() + extremos_sup.sum()),
    }
    resumen['total_atipicos'] = resumen['leves'] + resumen['extremos']
    resumen['pct_atipicos'] = resumen['total_atipicos'] / resumen['N']
    return resumen


def identificadores(x, st):
    m = np.isfinite(x)
    leve = m & (((x < st['li15']) & (x >= st['li3'])) | ((x > st['ls15']) & (x <= st['ls3'])))
    extremo = m & ((x < st['li3']) | (x > st['ls3']))
    return leve, extremo


def escribir_excel(antes, despues, resumen_antes, resumen_despues, destino):
    """Escribe un Excel simple usando openpyxl (sin artifact_tool).

    Las cifras derivadas se calculan en Python y se guardan como valores,
    para que sean visibles sin recalcular fórmulas al abrir el libro.
    """
    wb = Workbook()
    wb.remove(wb.active)

    def hoja(nombre, titulo, descripcion, encabezados, filas, ancho_max=29):
        ws = wb.create_sheet(nombre)
        ws.append([titulo])
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(encabezados))
        ws.cell(1,1).font = Font(name='Calibri', bold=True, size=14, color='292929')
        ws.row_dimensions[1].height = 27
        ws.append([descripcion])
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(encabezados))
        ws.cell(2,1).font = Font(name='Calibri', size=10, color='555555')
        ws.row_dimensions[2].height = 25
        ws.append([])
        ws.append(encabezados)
        for c in ws[4]:
            c.font = Font(name='Calibri', bold=True, size=10, color='292929')
            c.fill = PatternFill('solid', fgColor='EAEAEA')
            c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws.row_dimensions[4].height=35
        for valores in filas:
            ws.append(valores)
        for j, header in enumerate(encabezados, 1):
            # Tablas simples, sin filtros ni filas inmovilizadas.
            ws.column_dimensions[get_column_letter(j)].width = min(ancho_max, max(14, min(26,len(str(header))+4)))
            for fila in range(5, ws.max_row + 1):
                c = ws.cell(fila,j)
                c.font = Font(name='Calibri', size=10)
                if isinstance(c.value, float):
                    c.number_format = '0.000000' if ' (%)' not in header else '0.00%'
                elif isinstance(c.value, int):
                    c.number_format = '#,##0'
                c.alignment = Alignment(vertical='center')
        return ws

    hoja('Criterios', 'Criterios para identificar atípicos',
         'RIC = Q3 - Q1; leves entre 1,5 y 3 RIC, extremos más allá de 3 RIC.',
         ['Concepto', 'Criterio', 'Interpretación'], [
             ['RIC', 'Q3 - Q1', 'Rango intercuartílico'],
             ['Límite inferior 1,5', 'Q1 - 1,5 × RIC', 'Inicio de atípicos inferiores'],
             ['Límite superior 1,5', 'Q3 + 1,5 × RIC', 'Inicio de atípicos superiores'],
             ['Límite inferior 3', 'Q1 - 3 × RIC', 'Extremos inferiores'],
             ['Límite superior 3', 'Q3 + 3 × RIC', 'Extremos superiores'],
             ['Habitual', 'LI 1,5 ≤ valor ≤ LS 1,5', 'No atípico según IQR'],
             ['Atípico leve', 'Entre límites de 1,5 y 3 RIC', 'No incluye los extremos'],
             ['Atípico extremo', 'Por fuera de 3 RIC', 'Valores más alejados'],
         ])
    ws = wb['Criterios']
    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 36
    ws.column_dimensions['C'].width = 44
    for row in ws.iter_rows(min_row=5):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical='center')
        ws.row_dimensions[row[0].row].height = 25

    cabecera_lim = ['Sensor', 'N válidos', 'Q1', 'Q3', 'RIC = Q3-Q1',
                    'Límite inferior 1,5', 'Límite superior 1,5',
                    'Límite inferior 3', 'Límite superior 3']
    cabecera_at = ['Sensor','N válidos','Habituales','Leves inferiores','Leves superiores',
                   'Leves total','Extremos inferiores','Extremos superiores','Extremos total',
                   'Atípicos total','Atípicos (%)']
    for etapa, resumen in [('ANTES', resumen_antes), ('DESPUES', resumen_despues)]:
        lim, ati = [], []
        for nombre, st in zip(NOMBRES, resumen):
            nombre = nombre.replace('_SisGenEol', '')
            lim.append([nombre,st['N'],st['q1'],st['q3'],st['ric'],st['li15'],st['ls15'],st['li3'],st['ls3']])
            ati.append([nombre,st['N'],st['habituales'],st['leves_inf'],st['leves_sup'],
                        st['leves'],st['extremos_inf'],st['extremos_sup'],st['extremos'],
                        st['total_atipicos'],st['pct_atipicos']])
        hoja('Limites '+etapa, 'Límites IQR/RIC · '+etapa,
             'Rangos estadísticos por sensor, no límites físicos del instrumento.',
             cabecera_lim,lim)
        hoja('Atipicos '+etapa, 'Atípicos leves y extremos · '+etapa,
             'Leves y extremos se cuentan sin duplicarlos ni eliminar datos.',
             cabecera_at,ati)
        wb['Atipicos '+etapa].column_dimensions['K'].width=18

    comparar = []
    for nombre,a,d in zip(NOMBRES,resumen_antes,resumen_despues):
        comparar.append([nombre.replace('_SisGenEol',''),a['leves'],d['leves'],
                         a['extremos'],d['extremos'],a['total_atipicos'],d['total_atipicos'],
                         d['total_atipicos']-a['total_atipicos'],a['Nulos'],d['Nulos']])
    hoja('Comparacion','Comparación antes y después de la limpieza inicial',
         'Cada conjunto tiene sus propios cuartiles y límites IQR; no se eliminaron atípicos.',
         ['Sensor','Leves antes','Leves después','Extremos antes','Extremos después',
          'Total antes','Total después','Variación total','Nulos antes','Nulos después'],comparar)
    notas = [
        ['Dataset original', 'CF-2023-I-204-1230_DAT*.xlsx, primera hoja.'],
        ['Dataset limpio', 'Dataset_eolico_filtrado_2*.xlsx, hoja Datos.'],
        ['Columnas', 'Diez sensores. El contador smp no se analiza como sensor.'],
        ['Clasificación', 'Leves fuera de 1,5 y dentro de 3 RIC; extremos fuera de 3.'],
        ['Datos originales', 'Algunas filas están desalineadas; sus extremos pueden ser errores estructurales.'],
        ['Limpieza inicial', 'Desalineaciones corregidas y nulos tratados; atípicos conservados.'],
        ['Sin zoom', 'Se incluyen todos los valores en histogramas y diagramas de caja.'],
        ['Validez física', 'Pendiente de confirmar unidades y especificaciones de sensores.'],
        ['Formato', 'Celdas comunes, sin filtros, sin tablas dinámicas ni paneles congelados.'],
    ]
    hoja('Notas','Notas metodológicas',
         'Descripción del tratamiento y precauciones de interpretación.',
         ['Aspecto','Detalle'],notas,ancho_max=50)
    wb['Notas'].column_dimensions['A'].width = 30
    wb['Notas'].column_dimensions['B'].width = 100
    for row in wb['Notas'].iter_rows(min_row=5):
        row[1].alignment = Alignment(wrap_text=True,vertical='center')
        wb['Notas'].row_dimensions[row[0].row].height=28

    wb.save(destino)
    print('Excel guardado en:',destino,flush=True)


def estilo_fig(ax):
    ax.set_facecolor('white')
    ax.set_axisbelow(True)
    ax.grid(color=REJILLA, linewidth=0.8)
    for sp in ax.spines.values():
        sp.set_color('#D5D5D5')
    ax.tick_params(labelsize=9.8, colors=SECUNDARIO)
    ax.xaxis.label.set_color(SECUNDARIO)
    ax.yaxis.label.set_color(SECUNDARIO)


def fig_base(sensor, etapa, tipo, st, mostrar_minmax=False):
    fig, ax=plt.subplots(figsize=(11.2, 6.5), facecolor='white')
    fig.text(0.065,0.970,f'ANÁLISIS EXPLORATORIO  ·  {etapa}',ha='left',va='top',
             color=LINEA_CENTRAL,fontsize=10.8,fontweight='bold')
    fig.text(0.065,0.918,f'{tipo} — {sensor}',ha='left',va='top',
             color=TEXTO,fontsize=18,fontweight='bold')
    fig.text(0.065,0.862,
             f"N = {st['N']:,}   |   Media = {formato(st['media'])}   |   DS = {formato(st['DS'])}"
             .replace(',', '.'), ha='left',va='top',fontsize=10.4,color=SECUNDARIO)
    if mostrar_minmax:
        # Segunda línea de estadísticas, sin poner textos encima del histograma.
        fig.text(0.065, 0.821,
                 f"Mínimo = {formato(st['min'])}   |   Máximo = {formato(st['max'])}",
                 ha='left', va='top', fontsize=10.4, color=SECUNDARIO)
        fig.add_artist(Line2D([.065,.955],[.782,.782],transform=fig.transFigure,
                              color=REJILLA,lw=1))
        fig.subplots_adjust(left=.085,right=.96,top=.728,bottom=.16)
    else:
        fig.add_artist(Line2D([.065,.955],[.821,.821],transform=fig.transFigure,
                              color=REJILLA,lw=1))
        fig.subplots_adjust(left=.085,right=.96,top=.765,bottom=.16)
    estilo_fig(ax)
    return fig,ax


def leyenda_atipicos(ax):
    # Etiquetas fijas: no depende de si hay 0 puntos de una clase.
    handles=[
        Line2D([],[],marker='o',markersize=5,linestyle='',color=LEVE,label='Leves (1,5–3 RIC)'),
        Line2D([],[],marker='o',markersize=5,linestyle='',color=EXTREMO,label='Extremos (>3 RIC)')
    ]
    ax.legend(handles=handles,loc='lower left',frameon=False,fontsize=9)


def graficar_sensor(col, sensor, st, etapa, carpetas, solo_histogramas=False):
    x=col[np.isfinite(col)]
    orden=np.arange(1,len(col)+1)
    leve,extremo=identificadores(col,st)
    # 1: distribución. Se incluyen todos los valores en el histograma (sin filtrado).
    fig, ax=fig_base(sensor,etapa,'Histograma completo',st,mostrar_minmax=True)
    ax.hist(x,bins=65,color=BASE_GRAFICO,edgecolor='white',linewidth=.35)
    ax.axvline(st['media'],color=LINEA_CENTRAL,lw=1.8,ls='--',label='Media')
    ax.set_xlim(st['min'],st['max']) if st['min']<st['max'] else None
    ax.set_xlabel('Valor registrado (unidad por confirmar)')
    ax.set_ylabel('Frecuencia')
    ax.legend(frameon=False,loc='upper right',fontsize=9)
    nota = ('Rango completo, sin zoom ni exclusiones. Datos originales: pueden existir desalineaciones.'
            if etapa.startswith('ANTES') else
            'Rango completo, sin zoom ni exclusiones. Datos después de corregir estructura y nulos.')
    fig.text(.065,.055,nota, fontsize=9.1,color=SECUNDARIO)
    fig.savefig(carpetas['Histogramas']/f'Histograma_{sensor}.png',dpi=155,facecolor='white')
    plt.close(fig)
    if solo_histogramas:
        return

    # 2: caja / bigotes IQR 1,5. Atípicos leves y extremos diferenciados.
    fig, ax=fig_base(sensor,etapa,'Diagrama de caja completo',st)
    ax.boxplot(x,vert=False,positions=[0],widths=.34,showfliers=False,patch_artist=True,
               boxprops={'facecolor':BASE_CAJA,'edgecolor':LINEA_CENTRAL,'linewidth':1.4},
               whiskerprops={'color':'#7A7A7A','linewidth':1.3},
               capprops={'color':'#7A7A7A','linewidth':1.3},
               medianprops={'color':LINEA_CENTRAL,'linewidth':2.2})
    rng=np.random.default_rng(12)
    for mask,color,sz in [(leve,LEVE,14),(extremo,EXTREMO,22)]:
        out=col[mask]
        if len(out):
            y=rng.normal(loc=0,scale=.036,size=len(out))
            ax.scatter(out,y,s=sz,color=color,alpha=.48,rasterized=True,zorder=4,linewidths=0)
    ax.set_ylim(-.48,.48)
    ax.set_yticks([])
    if st['min']<st['max']:
        margen=(st['max']-st['min'])*.025
        ax.set_xlim(st['min']-margen,st['max']+margen)
    ax.set_xlabel('Valor registrado (unidad por confirmar)')
    ax.text(.98,.94,
            f"Leves: {st['leves']:,}  |  Extremos: {st['extremos']:,}"
            .replace(',','.'),transform=ax.transAxes,
            ha='right',va='top',fontsize=10,color=SECUNDARIO)
    leyenda_atipicos(ax)
    fig.text(.065,.055,'Bigotes: 1,5 RIC. Se grafican todos los atípicos; puntos superpuestos pueden parecer uno solo. Sin zoom.',
             fontsize=9.1,color=SECUNDARIO)
    fig.savefig(carpetas['Cajas']/f'Caja_{sensor}.png',dpi=155,facecolor='white')
    plt.close(fig)

    # 3: evolución completa, sin reducir 143561 muestras a bloques. Todas se dibujan.
    fig, ax=fig_base(sensor,etapa,'Comportamiento por muestra',st)
    ax.plot(orden,col,color=BASE_GRAFICO,lw=.45,alpha=.70,rasterized=True,label='Todos los registros')
    if leve.any():
        ax.scatter(orden[leve],col[leve],s=10,color=LEVE,alpha=.9,rasterized=True,zorder=3,label='Leves')
    if extremo.any():
        ax.scatter(orden[extremo],col[extremo],s=22,color=EXTREMO,alpha=.95,rasterized=True,zorder=4,label='Extremos')
    ax.set_xlim(1,len(col))
    if st['min']<st['max']:
        m=(st['max']-st['min'])*.025
        ax.set_ylim(st['min']-m,st['max']+m)
    ax.set_xlabel('Orden de muestra (no equivale a segundos)')
    ax.set_ylabel('Valor registrado')
    ax.legend(frameon=False,loc='upper right',fontsize=8.5)
    fig.text(.065,.055,'Cada muestra está incluida (sin promedios por bloques y sin zoom); la superposición puede formar bandas densas.',
             fontsize=9.1,color=SECUNDARIO)
    fig.savefig(carpetas['Evolucion']/f'Evolucion_{sensor}.png',dpi=155,facecolor='white')
    plt.close(fig)


def escribir_graficos(etapa, datos, stats, salida, solo_histogramas=False):
    carpetas={n:salida/n for n in ['Histogramas','Cajas','Evolucion']}
    for d in carpetas.values():
        d.mkdir(parents=True,exist_ok=True)
    for j,nombre in enumerate(NOMBRES):
        graficar_sensor(datos[:,j], nombre.replace('_SisGenEol',''), stats[j], etapa, carpetas, solo_histogramas)
        print('  Gráficos:',etapa,nombre.replace('_SisGenEol',''),flush=True)
    (salida/'LEEME.txt').write_text(
        f'Etapa: {etapa}\nTotal: {10 if solo_histogramas else 30} PNG.\n'
        'Histogramas con mínimo y máximo visibles en el encabezado; rango completo sin zoom.\n'
        'La evolución representa las muestras individuales, no agregados.\n'
        'IQR/RIC: leves fuera de 1,5 y dentro de 3; extremos fuera de 3.\n'
        'No se eliminan datos. Los originales pueden contener registros desalineados.\n',
        encoding='utf-8')


def main():
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('--original',type=Path,help='Excel original')
    p.add_argument('--limpio',type=Path,help='Excel después de la limpieza inicial')
    p.add_argument('--solo-excel',action='store_true',help='Generar únicamente el Excel')
    p.add_argument('--solo-graficos',action='store_true',help='Generar únicamente gráficos')
    p.add_argument('--solo-histogramas',action='store_true',help='Actualizar solo los 20 histogramas')
    a=p.parse_args()
    orig=a.original or encontrar('CF-2023-I-204-1230_DAT*.xlsx',
                         ['CF-2023-I-204-1230_DAT(3).xlsx','CF-2023-I-204-1230_DAT.xlsx'])
    limpio=a.limpio or encontrar('Dataset_eolico_filtrado_2*.xlsx',
                         ['Dataset_eolico_filtrado_2(3).xlsx','Dataset_eolico_filtrado_2.xlsx'])
    if not orig.is_file() or not limpio.is_file():
        raise FileNotFoundError('No se encontraron ambos Excel.')
    print('Leyendo archivo original...',flush=True)
    antes=leer_diez_sensores_excel(orig)
    print('Leyendo archivo limpio...',flush=True)
    despues=leer_diez_sensores_excel(limpio)
    if len(antes)!=len(despues):
        raise ValueError('El número de filas difiere entre los archivos.')
    print('Filas por etapa:',len(antes),'; sensores:',len(NOMBRES),flush=True)
    res_antes=[stats_ric(antes[:,i]) for i in range(10)]
    res_despues=[stats_ric(despues[:,i]) for i in range(10)]
    if not a.solo_graficos:
        escribir_excel(antes,despues,res_antes,res_despues,BASE/'Tabla_RIC_Atipicos_Leves_Extremos.xlsx')
    if not a.solo_excel:
        base=BASE/'Graficos_RIC_Completos_CON_MIN_MAX'
        escribir_graficos('ANTES DE LA LIMPIEZA',antes,res_antes,base/'ANTES', a.solo_histogramas)
        escribir_graficos('DESPUÉS DE LA LIMPIEZA INICIAL',despues,res_despues,base/'DESPUES', a.solo_histogramas)
    print('FINALIZADO.',flush=True)

if __name__=='__main__':
    main()
