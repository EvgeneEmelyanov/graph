import matplotlib.pyplot as plt
import numpy as np

# =========================
# ГЛОБАЛЬНЫЕ НАСТРОЙКИ ШРИФТА (+2)
# =========================
plt.rcParams.update({
    "font.size": 16,
    "axes.labelsize": 18,
    "xtick.labelsize": 15,
    "ytick.labelsize": 15,
    "legend.fontsize": 15
})

# =========================
# ДАННЫЕ
# =========================
regions = [
    "Республика Саха (Якутия)",
    "Камчатский край",
    "Республика Коми",
    "Магаданская область",
    "НАО",
    "Мурманская область",
    "Республика Карелия",
    "Томская область",
    "Республика Тыва",
    "Красноярский край"
]

costs = [43, 35, 34, 33, 32, 31, 30, 27, 25, 24]
subsidies = [30, 19, 5, 7, 13, 23, 11, 18, 18, 14]

# =========================
# СОРТИРОВКА
# =========================
data = list(zip(regions, costs, subsidies))
data_sorted = sorted(data, key=lambda x: x[1], reverse=True)

regions_top = [x[0] for x in data_sorted]
costs_top = [x[1] for x in data_sorted]
subs_top = [x[2] for x in data_sorted]

# =========================
# ГРАФИК
# =========================
y = np.arange(len(regions_top))

plt.figure(figsize=(11, 7))

color_cost = "#426f91"
color_subs = "#c95f46"

bar_width = 0.4

bars1 = plt.barh(y - bar_width/2, costs_top, height=bar_width,
                 color=color_cost, label="Удельные расходы")

bars2 = plt.barh(y + bar_width/2, subs_top, height=bar_width,
                 color=color_subs, label="Субсидии")

plt.yticks(y, regions_top)
plt.xlabel("руб/кВт·ч")

# Сетка
plt.grid(axis='x', linestyle='--', alpha=0.6)

# =========================
# ПОДПИСИ ЗНАЧЕНИЙ
# =========================
for i, (c, s) in enumerate(zip(costs_top, subs_top)):
    plt.text(c + 0.5, i - bar_width/2, f"{c}", va='center')
    plt.text(s + 0.5, i + bar_width/2, f"{s}", va='center')

plt.legend()
plt.gca().invert_yaxis()

plt.tight_layout()
plt.show()