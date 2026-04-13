import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter

# ДЛЯ ГРАФИКА ОПТИМАЛЬНАЯ ЕМКОСТЬ ДЛЯ КАЖДОГО ЗНАЧЕНИЯ МОЩНОСТИ ВЭУ С УЧЕТОМ ТОКОВ РАЗРЯДА

print("Вставь данные для 3 кривых (1C, 2C, 3C через таб/пробел, с запятой). Пустая строка — завершение:")

y1, y2, y3 = [], [], []

while True:
    line = input()
    if not line.strip():
        break

    parts = line.split()
    if len(parts) != 3:
        print("Ошибка: нужно 3 значения")
        continue

    try:
        y1.append(float(parts[0].replace(",", ".")))
        y2.append(float(parts[1].replace(",", ".")))
        y3.append(float(parts[2].replace(",", ".")))
    except ValueError:
        print("Ошибка преобразования числа")
        continue

# X = 25, 50, 75, ...
x = [25 * (i + 1) for i in range(len(y1))]

# Размеры как в твоем примере
fig_width = (130 / 25.4) * 1.04
fig_height = 95 / 25.4

plt.rcParams["font.family"] = "DejaVu Sans"

fig, ax = plt.subplots(figsize=(fig_width, fig_height), dpi=300)

label_fs = 12
tick_fs = 11
legend_fs = 8


# Форматирование чисел с запятой
def comma_formatter(x, pos):
    val = round(x, 1)
    if float(val).is_integer():
        return f"{int(val)}"
    return str(val).replace(".", ",")


formatter = FuncFormatter(comma_formatter)

# Линии
line1, = ax.plot(
    x, y1,
    color="black",
    linestyle="-",
    marker="o",
    linewidth=1.0,
    markersize=4,
    label="1C"
)

line2, = ax.plot(
    x, y2,
    color="black",
    linestyle="--",
    marker="^",
    linewidth=1.0,
    markersize=4,
    label="2C"
)

line3, = ax.plot(
    x, y3,
    color="black",
    linestyle=":",
    marker="s",
    linewidth=1.0,
    markersize=4,
    label="3C"
)

# Подписи осей
ax.set_xlabel("Установленная мощность ВЭУ, %", fontsize=label_fs)
ax.set_ylabel("Оптимальная емкость СНЭ, %", fontsize=label_fs)

# Настройка тиков
ax.tick_params(axis="both", labelsize=tick_fs)
ax.xaxis.set_major_formatter(formatter)
ax.yaxis.set_major_formatter(formatter)

# Сетка
ax.grid(True, linestyle="--", linewidth=0.5, alpha=0.4)

# Легенда — параметры взяты по образцу
lines = [line1, line2, line3]
labels = [l.get_label() for l in lines]
ax.legend(
    lines, labels,
    loc="lower center",
    bbox_to_anchor=(0.5, 0.05),
    ncol=1,
    fontsize=legend_fs,
    frameon=True,
    edgecolor="black",
    handlelength=2,
    handletextpad=0.4,
    columnspacing=0.8,
    borderpad=0.3
)

plt.tight_layout()
plt.savefig("plot_bw_single_axis.png", dpi=300, bbox_inches="tight")
plt.show()
