from pathlib import Path
from calendar import monthrange

import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import MaxNLocator


# ============================================================
# НАСТРОЙКИ
# ============================================================
FILE = Path(r"D:\08_ModelingData\02_Wind.txt")
START_DATE = "2001-01-01 00:00:00"
OUTPUT_PNG = r"D:\wind.png"
FIG_DPI = 300

plt.rcParams.update({
    "font.size": 16,
    "axes.labelsize": 18,
    "xtick.labelsize": 15,
    "ytick.labelsize": 15
})


# ============================================================
# ЧТЕНИЕ ФАЙЛА (десятичная запятая)
# ============================================================
def load_wind_series(file_path: Path) -> pd.Series:
    values = []
    with open(file_path, "r", encoding="utf-8") as f:
        for line in f:
            s = line.strip()
            if not s:
                continue
            s = s.replace(",", ".")
            values.append(float(s))
    return pd.Series(values, name="wind")


wind = load_wind_series(FILE)


# ============================================================
# ВРЕМЯ (с учётом високосных лет)
# ============================================================
time_index = pd.date_range(start=START_DATE, periods=len(wind), freq="h")

df = pd.DataFrame({
    "time": time_index,
    "wind": wind
})

df["year"] = df["time"].dt.year
df["month"] = df["time"].dt.month


# ============================================================
# СРЕДНЕМЕСЯЧНЫЕ ЗНАЧЕНИЯ ПО ГОДАМ
# ============================================================
monthly_stats = (
    df.groupby(["year", "month"])
      .agg(mean_wind=("wind", "mean"),
           hours=("wind", "size"))
      .reset_index()
)

def expected_hours(year, month):
    return monthrange(year, month)[1] * 24

monthly_stats["expected_hours"] = monthly_stats.apply(
    lambda row: expected_hours(int(row["year"]), int(row["month"])),
    axis=1
)

# оставляем только полные месяцы
monthly_stats = monthly_stats[
    monthly_stats["hours"] == monthly_stats["expected_hours"]
]

data_by_month = [
    monthly_stats.loc[monthly_stats["month"] == m, "mean_wind"].values
    for m in range(1, 13)
]

months = ["Янв", "Фев", "Мар", "Апр", "Май", "Июн",
          "Июл", "Авг", "Сен", "Окт", "Ноя", "Дек"]


# ============================================================
# ГРАФИК
# ============================================================
fig, ax = plt.subplots(figsize=(11, 6))

bp = ax.boxplot(
    data_by_month,
    patch_artist=True,
    widths=0.6,
    showfliers=True,
    whis=1.5
)

# цвет
for box in bp["boxes"]:
    box.set(facecolor="#426f91", alpha=0.75)

for median in bp["medians"]:
    median.set(color="black", linewidth=2)

# оси
ax.set_xticks(range(1, 13))
ax.set_xticklabels(months)
ax.set_xlabel("Месяц")
ax.set_ylabel("Среднемесячная скорость ветра, м/с")

# сетка
ax.grid(axis="y", linestyle="--", alpha=0.6)

# ТОЛЬКО ЦЕЛЫЕ ЧИСЛА по оси Y
ax.yaxis.set_major_locator(MaxNLocator(integer=True))


# ============================================================
# СОХРАНЕНИЕ
# ============================================================
plt.tight_layout()
plt.savefig(OUTPUT_PNG, dpi=FIG_DPI, bbox_inches="tight")
plt.show()