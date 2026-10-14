"""Matriz de correlación científica para el dataset eólico.

Entrada: Dataset_eolico_depurado_6sigma.xlsx (hoja Datos)
Salida: 06_Correlaciones/Matriz_correlacion_Pearson_PAPER.png
Únicamente PNG (600 dpi); no crea PDF, CSV ni otros archivos.

Instalar dependencias si falta alguna:
    python -m pip install pandas numpy matplotlib openpyxl
Ejecutar:
    python Correlacion_Eolica_Pearson_Estilo_Paper_Final.py
"""

from pathlib import Path
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.colors import LinearSegmentedColormap, TwoSlopeNorm

BASE = Path(__file__).resolve().parent
ENTRADA = BASE / "Dataset_eolico_depurado_6sigma.xlsx"
SALIDA = BASE / "06_Correlaciones" / "Matriz_correlacion_Pearson_PAPER.png"

# Se respetan la selección y el orden de las diez variables físicas.
VARIABLES = {
    "AcelX1_SisGenEol": "AcelX",
    "AcelY1_SisGenEol": "AcelY",
    "AcelZ1_SisGenEol": "AcelZ",
    "GiroX_SisGenEol": "GiroX",
    "GiroY_SisGenEol": "GiroY",
    "GiroZ_SisGenEol": "GiroZ",
    "MagneX_SisGenEol": "MagneX",
    "MagneY_SisGenEol": "MagneY",
    "MagneZ_SisGenEol": "MagneZ",
    "Ter_SisGenEol": "Temperatura",
}

if not ENTRADA.is_file():
    archivos = [a for a in BASE.glob("Dataset_eolico_depurado_6sigma*.xlsx")
                if not a.name.startswith("~$")]
    if not archivos:
        raise FileNotFoundError("Coloca el Excel depurado junto al archivo .py")
    ENTRADA = max(archivos, key=lambda p: p.stat().st_mtime)

raw = pd.read_excel(ENTRADA, sheet_name="Datos", engine="openpyxl")
faltantes = [nombre for nombre in VARIABLES if nombre not in raw]
if faltantes:
    raise ValueError(f"Faltan variables en el Excel: {faltantes}")
medidas = raw[list(VARIABLES)].apply(pd.to_numeric, errors="raise")
if not np.isfinite(medidas.to_numpy(dtype=float)).all():
    raise ValueError("El Excel contiene mediciones vacías o no finitas")
if (medidas.nunique() < 2).any():
    raise ValueError("Una variable es constante; Pearson no puede calcularse")

matriz = medidas.corr(method="pearson").to_numpy()
etiquetas = list(VARIABLES.values())
n = len(etiquetas)

plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "font.size": 9,
    "axes.unicode_minus": True,
    "figure.facecolor": "white",
    "axes.facecolor": "white",
    "savefig.facecolor": "white",
})

# Escala simétrica fija, necesaria para interpretar +1 y -1 sin distorsión.
paleta = LinearSegmentedColormap.from_list(
    "cientifica", ["#af6c57", "#f7f7f6", "#21667c"], N=256
)
fig = plt.figure(figsize=(8.2, 7.45), dpi=160)
# Se reservan zonas independientes para título, matriz, etiquetas y leyenda.
ax = fig.add_axes([0.155, 0.165, 0.69, 0.725])
cbax = fig.add_axes([0.877, 0.275, 0.020, 0.52])
im = ax.imshow(matriz, cmap=paleta,
               norm=TwoSlopeNorm(vmin=-1, vcenter=0, vmax=1),
               interpolation="nearest", aspect="equal")

ax.set_xticks(np.arange(n))
ax.set_xticklabels(etiquetas, rotation=45, ha="right", rotation_mode="anchor")
ax.set_yticks(np.arange(n))
ax.set_yticklabels(etiquetas)
ax.tick_params(axis="both", which="major", length=0, pad=7, labelsize=8.9)
ax.set_xticks(np.arange(-0.5, n, 1), minor=True)
ax.set_yticks(np.arange(-0.5, n, 1), minor=True)
ax.grid(which="minor", color="white", linewidth=0.85)
ax.tick_params(which="minor", left=False, bottom=False)
for spine in ax.spines.values():
    spine.set_visible(False)

for i in range(n):
    for j in range(n):
        v = matriz[i, j]
        if abs(v) < 0.005:  # evita -0.00 y +0.00
            etiqueta = "0.00"
        elif i == j:
            etiqueta = "1.00"
        else:
            etiqueta = f"{v:+.2f}"
        ax.text(j, i, etiqueta, ha="center", va="center",
                fontsize=8.0, color="white" if abs(v) > 0.66 else "#293741")

cbar = fig.colorbar(im, cax=cbax, ticks=[-1, -0.5, 0, 0.5, 1])
cbar.ax.tick_params(labelsize=8.4, length=0, pad=6)
cbar.outline.set_visible(False)
cbar.set_label("Coeficiente de Pearson (r)", fontsize=9, labelpad=9)

# Un encabezado discreto y ningún texto encima del área de los datos.
fig.text(0.155, 0.951, "Matriz de correlación de Pearson",
         ha="left", va="top", fontsize=11.5, color="#263844", fontweight="semibold")
fig.text(0.845, 0.950, f"n = {len(medidas):,}".replace(",", "."),
         ha="right", va="top", fontsize=8.8, color="#556571")

SALIDA.parent.mkdir(parents=True, exist_ok=True)
# La figura tiene márgenes propios: NO usar bbox_inches='tight' ni tight_layout.
fig.savefig(SALIDA, dpi=600, facecolor="white")
plt.close(fig)
print(f"Figura generada: {SALIDA}")
print(f"Registros: {len(medidas):,} | variables: {n}")
print("Única salida: PNG (600 ppp).")
