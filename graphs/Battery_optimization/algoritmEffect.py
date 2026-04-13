import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter

# ДЛЯ ГРАФИКА ЭФФЕКТИНВОСТЬ АДАПТИВНОГО АЛГОРИТМА ДЛЯ КАЖДОЙ УСТАНОВЛЕННОЙ МОЩНОСТИ ВЭУ
print("Вставь данные (LCOE, LPSP, LOLP через таб/пробел, с запятой). Пустая строка — завершение:")

lcoe, lpsp, lolp = [], [], []

while True:
    line = input()
    if not line.strip():
        break

    parts = line.split()
    if len(parts) != 3:
        print("Ошибка: нужно 3 значения")
        continue

    try:
        lcoe.append(float(parts[0].replace(",", ".")))
        lpsp.append(float(parts[1].replace(",", ".")))
        lolp.append(float(parts[2].replace(",", ".")))
    except ValueError:
        print("Ошибка преобразования числа")
        continue

# X = 25, 50, 75, ...
x = [25 * (i + 1) for i in range(len(lcoe))]

# Размер ~ половины A5 + увеличение ширины на ~4%
fig_width = (130 / 25.4) * 1.04
fig_height = 95 / 25.4

plt.rcParams["font.family"] = "DejaVu Sans"

fig, ax1 = plt.subplots(figsize=(fig_width, fig_height), dpi=300)

label_fs = 12
tick_fs = 11
legend_fs = 8

# Форматирование
def comma_formatter(x, pos):
    val = round(x, 1)
    if val.is_integer():
        return f"{int(val)}"
    return str(val).replace(".", ",")

formatter = FuncFormatter(comma_formatter)

# Левая ось
line1, = ax1.plot(
    x, lcoe,
    color="black",
    linestyle="-",
    marker="o",
    linewidth=1.0,
    label="LCOE"
)
ax1.set_xlabel("Установленная мощность ВЭУ, %", fontsize=label_fs)
ax1.set_ylabel("Эффективность LCOE, %", fontsize=label_fs)
ax1.tick_params(axis='both', labelsize=tick_fs)
ax1.yaxis.set_major_formatter(formatter)
ax1.grid(True, linestyle="--", linewidth=0.5, alpha=0.4)

# Правая ось
ax2 = ax1.twinx()

line2, = ax2.plot(
    x, lpsp,
    color="black",
    linestyle="--",
    marker="s",
    linewidth=1.0,
    label="LPSP"
)

line3, = ax2.plot(
    x, lolp,
    color="black",
    linestyle=":",
    marker="^",
    linewidth=1.0,
    label="LOLP"
)

ax2.set_ylabel("Эффективность LPSP и LOLP, %", fontsize=label_fs)
ax2.tick_params(axis='y', labelsize=tick_fs)
ax2.yaxis.set_major_formatter(formatter)

# Легенда
lines = [line1, line2, line3]
labels = [l.get_label() for l in lines]
ax1.legend(
    lines, labels,
    loc="lower center",
    bbox_to_anchor=(0.4, 0.08),
    ncol=1,
    fontsize=legend_fs,
    frameon=True,
    edgecolor="black",
    handlelength=1.2,
    handletextpad=0.4,
    columnspacing=0.8,
    borderpad=0.3
)

plt.tight_layout()
plt.savefig("plot_bw_a5_half.png", dpi=300, bbox_inches="tight")
plt.show()