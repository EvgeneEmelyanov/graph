import math
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter

print("Вставь данные для 2 кривых (LOLP и LCOE через таб/пробел, с запятой). Пустая строка — завершение:")

y_lolh, y_lpsp = [], []

while True:
    line = input()
    if not line.strip():
        break

    parts = line.split()
    if len(parts) != 2:
        print("Ошибка: нужно 2 значения")
        continue

    try:
        y_lolh.append(float(parts[0].replace(",", ".")))
        y_lpsp.append(float(parts[1].replace(",", ".")))
    except ValueError:
        print("Ошибка преобразования числа")
        continue

if not y_lolh or not y_lpsp:
    raise ValueError("Нет данных для построения графика.")

# X = 25, 50, 75, ...
x = [25 * (i + 1) for i in range(len(y_lolh))]

# Размеры как в твоем примере
fig_width = (130 / 25.4) * 1.04
fig_height = 95 / 25.4

plt.rcParams["font.family"] = "DejaVu Sans"

fig, ax1 = plt.subplots(figsize=(fig_width, fig_height), dpi=300)
ax2 = ax1.twinx()

label_fs = 12
tick_fs = 11
legend_fs = 8


def comma_formatter(x, pos):
    val = round(x, 1)
    if float(val).is_integer():
        return f"{int(val)}"
    return str(val).replace(".", ",")


def nice_upper_limit(max_val: float) -> float:
    """
    Округляет максимум вверх до 'красивого' значения.
    Примеры:
    2.88 -> 5
    4.03 -> 5
    11.29 -> 20
    40.06 -> 50
    """
    if max_val <= 0:
        return 1.0

    exponent = math.floor(math.log10(max_val))
    fraction = max_val / (10 ** exponent)

    if fraction <= 1:
        nice_fraction = 1
    elif fraction <= 2:
        nice_fraction = 2
    elif fraction <= 5:
        nice_fraction = 5
    else:
        nice_fraction = 10

    return nice_fraction * (10 ** exponent)


formatter = FuncFormatter(comma_formatter)

# Линия 1 — LOLH
line1, = ax1.plot(
    x, y_lolh,
    color="black",
    linestyle="-",
    marker="o",
    linewidth=1.0,
    markersize=4,
    label="LOLP"
)

# Линия 2 — LPSP
line2, = ax2.plot(
    x, y_lpsp,
    color="black",
    linestyle="--",
    marker="s",
    linewidth=1.0,
    markersize=4,
    label="LCOE"
)

# Подписи осей
ax1.set_xlabel("Установленная мощность ВЭУ, %", fontsize=label_fs)
ax1.set_ylabel("Эффективность LOLP, %", fontsize=label_fs)
ax2.set_ylabel("Эффективность LCOE, %", fontsize=label_fs)

# Автомасштаб по максимумам
y1_max = nice_upper_limit(max(y_lolh))
y2_max = nice_upper_limit(max(y_lpsp))

ax1.set_ylim(0, y1_max)
ax2.set_ylim(0, y2_max)

# Одинаковое количество делений для визуального совпадения сетки
tick_count = 6
ax1.set_yticks([i * y1_max / (tick_count - 1) for i in range(tick_count)])
ax2.set_yticks([i * y2_max / (tick_count - 1) for i in range(tick_count)])

# Настройка тиков
ax1.tick_params(axis="both", labelsize=tick_fs)
ax2.tick_params(axis="y", labelsize=tick_fs)

ax1.xaxis.set_major_formatter(formatter)
ax1.yaxis.set_major_formatter(formatter)
ax2.yaxis.set_major_formatter(formatter)

# Сетка по левой оси
ax1.grid(True, linestyle="--", linewidth=0.5, alpha=0.4)

# Общая легенда
lines = [line1, line2]
labels = [l.get_label() for l in lines]
ax1.legend(
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
plt.savefig("plot_bw_dual_axis_auto_scaled.png", dpi=300, bbox_inches="tight")
plt.show()