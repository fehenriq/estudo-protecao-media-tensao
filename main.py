import matplotlib.pyplot as plt
from matplotlib.ticker import ScalarFormatter
from matplotlib.ticker import FormatStrFormatter
import numpy as np

x_fase_montante = np.array([
    202, 240, 260, 280, 300, 400, 600, 800, 1000, 1200, 1400, 1600, 1800, 2000, 2200, 2400, 2600, 2800, 3000,
    3200, 3400, 3600, 3800, 4000, 4200, 4400, 4500, 4500, 4500, 4500, 4500, 4500, 4500, 4500, 4500, 4500, 4500,
    4500, 4500, 4500, 4500, 4500, 4500, 4500, 4500, 4500, 4500, 4500, 4500
])

x_neutro_montante = np.array([
    60.6, 72, 78, 84, 90, 120, 180, 240, 300, 360, 420, 480, 540, 600, 660, 720, 780, 840, 900, 960, 1020, 1080,
    1140, 1200, 1200, 1200, 1200, 1200, 1200, 1200, 1200, 1200, 1200, 1200, 1200, 1200, 1200, 1200, 1200, 1200,
    1200, 1200, 1200, 1200, 1200, 1200, 1200, 1200, 1200
])

x_fase_jusante = np.array([
    29.9, 35.5, 38.5, 41.4, 44.4, 59.2, 88.8, 118.4, 148, 177.6, 207.2, 236.8, 266.4, 296, 325.6, 340, 340, 340,
    340, 340, 340, 340, 340, 340, 340, 340, 340, 340, 340, 340, 340, 340, 340, 340, 340, 340, 340, 340, 340, 340,
    340, 340, 340, 340, 340, 340, 340, 340, 340
])

x_neutro_jusante = np.array([
    9.1, 10.8, 11.7, 12.6, 13.5, 18, 27, 36, 45, 52, 52, 52, 52, 52, 52, 52, 52, 52, 52, 52, 52, 52, 52, 52, 52,
    52, 52, 52, 52, 52, 52, 52, 52, 52, 52, 52, 52, 52, 52, 52, 52, 52, 52, 52, 52, 52, 52, 52, 52
])

y_montante = np.array([
    140.68,	7.66,	5.32,	4.15,	3.44,	2.01,	1.26,	1,	0.86,	0.77,	0.71,	0.66,	0.62,	0.59,	0.57,	0.55,	0.53,	0.52,
    0.5,	0.49,	0.48,	0.47,	0.46,	0.45,	0.45,	0.44,	0.43,	0.43,	0.42,	0.42,	0.41,	0.41,	0.4,	0.4,	0.39,	0.38,
    0.38,	0.37,	0.37,	0.36,	0.35,	0.35,	0.34,	0.33,	0.33,	0.32,	0.31,	0.29,	0.01
])

y_jusante = np.array([
    135, 6.75, 4.5, 3.38, 2.7, 1.35, 0.68, 0.45, 0.34, 0.27, 0.23, 0.19, 0.17, 0.15, 0.14, 0.12, 0.11, 0.1,
    0.1, 0.09, 0.09, 0.08, 0.08, 0.07, 0.07, 0.06, 0.06, 0.06, 0.06, 0.05, 0.05, 0.05, 0.05, 0.05, 0.04, 0.04,
    0.04, 0.04, 0.03, 0.03, 0.03, 0.03, 0.03, 0.02, 0.02, 0.02, 0.02, 0.01, 0.01
])

x_carga = np.array([
    22.74, 22.74
])

y_carga = np.array([
    1, 1.05
])

x_ICC3F = np.array([
    10000, 10100
])

y_ICC3F = np.array([
    3, 3.1
])

x_ANSI = np.array([
    1.35, 11.35
])

y_ANSI = np.array([
    5, 5.105
])

x_ICC1F = np.array([
    4000, 4100
])

y_ICC1F = np.array([
    3, 3.1
])

x_IMAG = np.array([
    323.9, 313.9
])

y_IMAG = np.array([
    0.1, 0.095
])

x_51_GS_MONTANTE = np.array([
    12, 36
])

y_51_GS_MONTANTE = np.array([
    6, 6
])

x_51_GS_JUSANTE = np.array([
    8, 24
])

y_51_GS_JUSANTE = np.array([
    4, 4
])

x_elo = np.array([
    23.69, 23.69, 24, 24.48, 24.63, 24.8, 25.04, 25.45, 25.78, 26.3, 27.24, 28.71, 30, 31.41, 31.84, 32.73,
    33.57, 34.74, 35.89, 37.31, 39.99, 45.08, 48.98, 54.71, 56.88, 58.75, 61.52, 65, 69.5, 77.9, 89.56, 110.22,
    127.95, 161.08, 171.79, 183.99, 201.84, 215.77, 245.47, 289.72, 353.09, 486.77, 600, 664.56
])

y_elo = np.array([
    300, 200, 150, 100, 90, 80, 70, 60, 50, 40, 30, 20, 15, 10, 9, 8, 7, 6, 5, 4, 3, 2, 1.5, 1, 0.9, 0.804, 0.7,
    0.6, 0.5, 0.4, 0.3, 0.2, 0.15, 0.1, 0.09, 0.08, 0.068, 0.059, 0.05, 0.04, 0.03, 0.02, 0.015, 0.013
])

fig, ax = plt.subplots(figsize=(10, 15))
ax.loglog(x_fase_montante, y_montante,
          label="FASE MONTANTE", linestyle="-", color="red")
ax.loglog(x_neutro_montante, y_montante, label="NEUTRO MONTANTE",
          linestyle="--", color="dodgerblue")
ax.loglog(x_fase_jusante, y_jusante, label="FASE JUSANTE",
          linestyle="-", color="black")
ax.loglog(x_neutro_jusante, y_jusante, label="NEUTRO JUSANTE",
          linestyle="--", color="limegreen")
ax.loglog(x_carga, y_carga, marker="o", label="I CARGA",
          linestyle=":", color="orchid")
ax.loglog(x_ICC3F, y_ICC3F, marker="o", label="ICC3F",
          linestyle=":", color="darkgreen")
ax.loglog(x_ANSI, y_ANSI, label="ANSI", linestyle=":", color="royalblue")
ax.loglog(x_ICC1F, y_ICC1F, marker="o", label="ICC1F",
          linestyle=":", color="darkorange")
ax.loglog(x_IMAG, y_IMAG, marker="o", label="IMAG",
          linestyle=":", color="fuchsia")
ax.loglog(x_51_GS_MONTANTE, y_51_GS_MONTANTE,
          label="51 GS MONTANTE", linestyle=":", color="darkred")
ax.loglog(x_51_GS_JUSANTE, y_51_GS_JUSANTE,
          label="51 GS JUSANTE", linestyle=":", color="orange")
ax.loglog(x_elo, y_elo, label="ELO", linestyle="-", color="brown")

ax.set(
    xlabel="Corrente (A)",
    ylabel="Tempo (s)",
    xlim=(0.1, 10000),
    ylim=(0.01, 1000),
    title="COORDENOGRAMA FASES E NEUTRO\nDISJUNTOR GERAL DA CABINE X CONCESSIONÁRIA"
)

ax.xaxis.set_major_formatter(ScalarFormatter())
ax.yaxis.set_major_formatter(FormatStrFormatter('%.2f'))
box = ax.get_position()
ax.set_position([box.x0, box.y0, box.width * 0.8, box.height])
ax.legend(loc='center left', bbox_to_anchor=(1, 0.5))
ax.grid()
fig.savefig("grafico.png")
plt.show()
