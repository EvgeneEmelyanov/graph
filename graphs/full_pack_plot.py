import numpy as np
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from matplotlib.colors import LinearSegmentedColormap, Normalize

# ============================================================
# 0) НАСТРОЙКИ EXCEL (ОДИН ФАЙЛ)
# ============================================================
# FILE  = r"D:\10_results\02_Scopus\5.1.xlsx"
FILE  = r"D:\10_results\def.xlsx"
SHEET = "SWEEP_2"

OUT_PNG = r"C:\Users\Balt_\Desktop\4_surfaces.png"

# ============================================================
# 1) ВЫБОР НАБОРА ДАННЫХ (1/2/3) — как в твоём примере
# ============================================================
DATASET = 2  # <-- меняй: 1, 2 или 3

X_ROW = 2
TRI_LIMIT = 1.0000001  # не используется в rect, оставлен для совместимости

if DATASET == 1:
    Y_START_ROW, Y_END_ROW = 3, 9
    X_START_COL, X_END_COL = 2, 12
    # 4 блока Z (строки в Excel): [Fuel, Hours, ENS, Costs] — порядок как в твоём 2x2
    Z_BLOCKS = [(3, 9), (14, 20), (25, 31), (36, 42)]
    SHAPE_MODE = "rect"
elif DATASET == 2:
    Y_START_ROW, Y_END_ROW = 3, 12
    X_START_COL, X_END_COL = 2, 12
    Z_BLOCKS = [(3, 12), (17, 26), (31, 40), (45, 54)]
    SHAPE_MODE = "rect"
elif DATASET == 3:
    Y_START_ROW, Y_END_ROW = 3, 13
    X_START_COL, X_END_COL = 2, 12
    Z_BLOCKS = [(3, 13), (18, 28), (33, 43), (48, 58)]
    SHAPE_MODE = "triangle"  # если вдруг нужно, оставлено как в твоём коде
else:
    raise ValueError("DATASET должен быть 1, 2 или 3")

# ============================================================
# 2) УГЛЫ ОБЗОРА (ОСТАВЛЯЕМ КАК СЕЙЧАС)
# ============================================================
VIEW_ANGLES = {
    "fuel":  dict(elev=22, azim=65),
    "hours": dict(elev=22, azim=35),
    "ens":   dict(elev=22, azim=35),
    "costs": dict(elev=22, azim=35),
}

# ============================================================
# 3) ЦВЕТА (ОСТАВЛЯЕМ "ЭТАЛОННЫЙ" СТИЛЬ)
#    - facecolors по cmap(norm(Z))
#    - shade=False (цвета строго по Z, без освещения)
# ============================================================
colors = ["limegreen", "yellowgreen", "yellow", "orange", "red", "maroon"]
cmap = LinearSegmentedColormap.from_list("custom_scale", colors, N=256)

# ============================================================
# 4) ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ЧТЕНИЯ ИЗ EXCEL
# ============================================================
def to_float(v):
    """Аккуратный парсер чисел из Excel (учёт строк, пробелов, запятых, ###)."""
    if v is None:
        return np.nan
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        s = v.strip()
        if not s or s == "###":
            return np.nan
        s = s.replace(" ", "").replace("\u00A0", "").replace(",", ".")
        try:
            return float(s)
        except ValueError:
            return np.nan
    return np.nan

def _cell_to_label(v):
    if v is None:
        return ""
    if isinstance(v, str):
        return v.strip()
    return str(v).strip()

def read_axis_labels(ws):
    """
    Как в твоём большом коде:
    X-лейбл из (row=1, col=2), Y-лейбл из (row=1, col=3).
    """
    xlab = _cell_to_label(ws.cell(row=1, column=2).value)
    ylab = _cell_to_label(ws.cell(row=1, column=3).value)
    return (xlab or "Емкость СНЭ,\nкВт*ч"), (ylab or "Ток разряда, С")

def get_z_label(ws, z_start_row):
    """
    Как в твоём большом коде:
    - если блок начинается с Y_START_ROW -> Z из A1
    - иначе Z из (row=z_start_row-2, col=X_START_COL-1)
    """
    if z_start_row == Y_START_ROW:
        return _cell_to_label(ws.cell(row=1, column=1).value)
    return _cell_to_label(ws.cell(row=z_start_row - 2, column=X_START_COL - 1).value)

def read_block(ws, z_start_row, z_end_row):
    """
    Возвращает: x_vals, y_vals, Z (2D), z_label
    """
    zlab = get_z_label(ws, z_start_row)

    x = np.array(
        [to_float(ws.cell(row=X_ROW, column=col).value) for col in range(X_START_COL, X_END_COL + 1)],
        dtype=float
    )
    y = np.array(
        [to_float(ws.cell(row=row, column=1).value) for row in range(Y_START_ROW, Y_END_ROW + 1)],
        dtype=float
    )

    expected_rows = len(y)
    actual_rows = z_end_row - z_start_row + 1
    if actual_rows != expected_rows:
        raise RuntimeError(
            f"Диапазон Z по строкам не совпадает с Y: Y={expected_rows} строк, Z={actual_rows} строк.\n"
            f"Проверь Z_BLOCKS и Y_START_ROW/Y_END_ROW для DATASET={DATASET}."
        )

    Z = np.full((len(y), len(x)), np.nan, dtype=float)
    for i, row in enumerate(range(z_start_row, z_end_row + 1)):
        for j, col in enumerate(range(X_START_COL, X_END_COL + 1)):
            Z[i, j] = to_float(ws.cell(row=row, column=col).value)

    # (на случай triangle-режима)
    if SHAPE_MODE == "triangle":
        Xg, Yg = np.meshgrid(x, y)
        mask = np.isnan(Z) | ((Xg + Yg) > TRI_LIMIT)
        Z = Z.copy()
        Z[mask] = np.nan

    if np.all(np.isnan(x)) or np.all(np.isnan(y)):
        raise RuntimeError("X или Y не считались (всё NaN). Проверь лист/диапазоны.")
    if np.all(np.isnan(Z)):
        raise RuntimeError(
            "Блок Z полностью NaN. Частая причина: в xlsx нет сохранённого кэша формул "
            "(openpyxl data_only=True читает кэш)."
        )

    return x, y, Z, zlab

# ============================================================
# 5) ЧТЕНИЕ ИЗ EXCEL (ОДИН РАЗ)
# ============================================================
wb = load_workbook(FILE, data_only=True)
ws = wb[SHEET]

X_AXIS_TITLE, Y_AXIS_TITLE = read_axis_labels(ws)

# Читаем 4 блока в порядке: Fuel, Hours, ENS, Costs
(x_vals, y_vals, Z_fuel,  zlab_fuel)  = read_block(ws, *Z_BLOCKS[0])
(_,      _,      Z_hours, zlab_hours) = read_block(ws, *Z_BLOCKS[1])
(_,      _,      Z_ens,   zlab_ens)   = read_block(ws, *Z_BLOCKS[2])
(_,      _,      Z_costs, zlab_costs) = read_block(ws, *Z_BLOCKS[3])

# Сетка для текущего набора
X, Y = np.meshgrid(x_vals, y_vals)

# ============================================================
# 6) "ЭТАЛОННАЯ" ОТРИСОВКА (ОФОРМЛЕНИЕ КАК СЕЙЧАС)
# ============================================================
def plot_surface_flatcolors(ax, Z, title, zlabel, view):
    Z = np.asarray(Z, dtype=float)
    norm = Normalize(vmin=np.nanmin(Z), vmax=np.nanmax(Z))

    ax.plot_surface(
        X, Y, Z,
        facecolors=cmap(norm(Z)),
        rstride=1, cstride=1,
        linewidth=0.3,
        edgecolor="black",
        antialiased=True,
        shade=False,  # <-- фиксирует стиль (без освещения)
    )

    ax.set_title(title, fontsize=12, y=1)
    ax.set_xlabel(X_AXIS_TITLE, fontsize=10)
    ax.set_ylabel(Y_AXIS_TITLE, fontsize=10)
    ax.set_zlabel(zlabel, fontsize=10, rotation=180)
    ax.view_init(**view)

# ============================================================
# 7) ФИГУРА 2×2 (РАСПОЛОЖЕНИЕ КАК СЕЙЧАС)
# ============================================================
fig = plt.figure(figsize=(13, 11), dpi=300)
gs = fig.add_gridspec(2, 2, wspace=0.0, hspace=0.0)

ax1 = fig.add_subplot(gs[0, 0], projection="3d")
ax2 = fig.add_subplot(gs[0, 1], projection="3d")
ax3 = fig.add_subplot(gs[1, 0], projection="3d")
ax4 = fig.add_subplot(gs[1, 1], projection="3d")

# Заголовки оставлены как в твоём 2×2 примере; zlabel берём из Excel (если пусто — ставим дефолт)
plot_surface_flatcolors(ax1, Z_fuel,  "Экономические затраты",        (zlab_fuel  or "Z"), VIEW_ANGLES["fuel"])
plot_surface_flatcolors(ax2, Z_hours, "Недоотпуск энергии",          (zlab_hours or "Z"), VIEW_ANGLES["hours"])
plot_surface_flatcolors(ax3, Z_ens,   "Расход топлива",                   (zlab_ens   or "Z"), VIEW_ANGLES["ens"])
plot_surface_flatcolors(ax4, Z_costs, "Моточасы ДГУ", (zlab_costs or "Z"), VIEW_ANGLES["costs"])

fig.patch.set_facecolor("white")
for ax in (ax1, ax2, ax3, ax4):
    ax.set_facecolor("white")

plt.savefig(OUT_PNG, dpi=300, facecolor=fig.get_facecolor())
print("Готово:", OUT_PNG)
print(f"FILE={FILE} | SHEET={SHEET} | DATASET={DATASET} | SHAPE_MODE={SHAPE_MODE}")
