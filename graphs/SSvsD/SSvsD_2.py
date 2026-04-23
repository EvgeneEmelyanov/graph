import os
import re
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.colors import LinearSegmentedColormap
from matplotlib.ticker import FuncFormatter, ScalarFormatter
from openpyxl import load_workbook

# =========================================================
# НАСТРОЙКИ
# =========================================================
EXCEL_PATH = r"D:\2comb_results.xlsx"
SHEET_NAME = "SWEEP_2"
OUTPUT_DIR = r"D:\results"

SCHEME_1_NAME = "Секционированная"
SCHEME_2_NAME = "Двойная"

PEAK_LOAD_KW = 1346.0
WT_COUNT = 2

X_TRIM = (0, 0)
Y_TRIM = (4, 0)

METRICS = [
    ("LCOE", "LCOE", "руб/кВт·ч"),
    ("LPSP", "LPSP", ""),
    ("LOLP", "LOLP", ""),
    ("ENS_evtN", "ENS событий", "шт"),
    ("ENS_evtMaxH", "Макс. ENS", "ч"),
]

HEADER_ALIASES = {
    "LCOE": [
        "LCOE",
        "LCOE,руб/кВт∙ч",
        "LCOE, руб/кВт∙ч",
        "LCOE,руб/кВт·ч",
        "LCOE, руб/кВт·ч",
        "LCOE,руб/кВт*ч",
        "LCOE, руб/кВт*ч",
    ],
    "LPSP": [
        "LPSP",
    ],
    "LOLP": [
        "LOLP",
    ],
    "ENS_evtN": [
        "ENS_evtN",
    ],
    "ENS_evtMaxH": [
        "ENS_evtMaxH",
        "ENS_evt_MaxH",
        "ENS_evtMaxH ",
        "ENS_evt_MaxH ",
    ],
}

# =========================================================
# РАЗМЕРЫ ШРИФТОВ
# =========================================================
AXIS_LABEL_SIZE = 14
TICK_SIZE = 12
CB_LABEL_SIZE = 13
CB_TICK_SIZE = 12
HEADER_SIZE = 16
METRIC_LABEL_SIZE = 15

# =========================================================
# ПАЛИТРЫ
# =========================================================
MATTE_METRIC_CMAP = LinearSegmentedColormap.from_list(
    "metric",
    ["#f2f2f2", "#cfd8dc", "#8ea9b8", "#5f86a2", "#426f91"]
)

MATTE_DIFF_CMAP = LinearSegmentedColormap.from_list(
    "diff",
    ["#4d7a99", "#8fa8b7", "#e9e9e9", "#ddb48c", "#c98d5b"]
)

# =========================================================
# ФОРМАТТЕРЫ
# =========================================================
def comma_formatter(decimals=2):
    def _fmt(x, pos):
        if not np.isfinite(x):
            return ""
        s = f"{x:.{decimals}f}".rstrip("0").rstrip(".")
        return s.replace(".", ",")
    return FuncFormatter(_fmt)


class CommaScalarFormatter(ScalarFormatter):
    def __call__(self, x, pos=None):
        s = super().__call__(x, pos)
        return s.replace(".", ",")


# =========================================================
# ВСПОМОГАТЕЛЬНЫЕ
# =========================================================
def normalize_text(x):
    if x is None:
        return ""
    return (
        str(x)
        .strip()
        .replace(" ", "")
        .replace("·", "∙")
        .replace("*", "∙")
    )


def to_float(v):
    if v is None:
        return np.nan
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", "."))
    except Exception:
        return np.nan


def find_metric(ws, key):
    aliases = HEADER_ALIASES.get(key, [key])
    aliases = [normalize_text(a) for a in aliases]

    found = []

    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            val = normalize_text(ws.cell(r, c).value)
            if val in aliases:
                found.append((r, c))

    found = sorted(found, key=lambda rc: (rc[0], rc[1]))

    if len(found) < 2:
        raise ValueError(
            f"Для метрики '{key}' найдено блоков: {len(found)}. "
            f"Проверь подписи в Excel или добавь вариант в HEADER_ALIASES."
        )

    return found[:2]


def read_block(ws, r, c):
    xrow = r + 1
    yrow = r + 2
    xcol = c + 1

    xcols = []
    cc = xcol
    while cc <= ws.max_column and ws.cell(xrow, cc).value not in [None, ""]:
        xcols.append(cc)
        cc += 1

    yrows = []
    rr = yrow
    while rr <= ws.max_row and ws.cell(rr, c).value not in [None, ""]:
        yrows.append(rr)
        rr += 1

    if not xcols:
        raise ValueError(f"Не удалось прочитать ось X для блока в ячейке ({r}, {c})")
    if not yrows:
        raise ValueError(f"Не удалось прочитать ось Y для блока в ячейке ({r}, {c})")

    x = np.array([to_float(ws.cell(xrow, k).value) for k in xcols], dtype=float)
    y = np.array([to_float(ws.cell(k, c).value) for k in yrows], dtype=float)

    z = np.array(
        [[to_float(ws.cell(rr, cc).value) for cc in xcols] for rr in yrows],
        dtype=float
    )

    return x, y, z


def trim(x, y, z):
    xl, xr = X_TRIM
    yt, yb = Y_TRIM

    xe = len(x) - xr if xr > 0 else len(x)
    ye = len(y) - yb if yb > 0 else len(y)

    return x[xl:xe], y[yt:ye], z[yt:ye, xl:xe]


def diff(a, b):
    out = np.full_like(a, np.nan, dtype=float)
    m = np.abs(a) > 1e-12
    out[m] = (a[m] - b[m]) / a[m] * 100
    return out


def xlabels(x):
    labels = []
    for v in x:
        if np.isfinite(v):
            labels.append(str(int(round(v))))
        else:
            labels.append("")
    return labels


def ylabels(y):
    labels = []
    for v in y:
        if np.isfinite(v):
            labels.append(str(int(round(v * WT_COUNT / PEAK_LOAD_KW * 100))))
        else:
            labels.append("")
    return labels


def style_axis(ax, x, y, show_x=False, show_y=False):
    ax.set_xticks(np.arange(len(x)))
    ax.set_xticklabels(xlabels(x), rotation=90, fontsize=TICK_SIZE)

    ax.set_yticks(np.arange(len(y)))
    ax.set_yticklabels(ylabels(y), fontsize=TICK_SIZE)

    if show_x:
        ax.set_xlabel("Мощность ДГУ, кВт", fontsize=AXIS_LABEL_SIZE)

    if show_y:
        ax.set_ylabel("Мощность ВЭУ, %", fontsize=AXIS_LABEL_SIZE)


def setup_metric_colorbar(cb, key, label, unit):
    if key in ["LOLP", "LPSP"]:
        fmt = CommaScalarFormatter(useMathText=True)
        fmt.set_powerlimits((0, 0))
        cb.formatter = fmt
    else:
        cb.formatter = comma_formatter(2)

    cb.update_ticks()
    cb.ax.tick_params(labelsize=CB_TICK_SIZE)

    cbar_label = f"{label}, {unit}" if unit else label
    cb.set_label(cbar_label, fontsize=CB_LABEL_SIZE)


# =========================================================
# ЗАГРУЗКА
# =========================================================
wb = load_workbook(EXCEL_PATH, data_only=True)
ws = wb[SHEET_NAME]

rows = len(METRICS)

fig = plt.figure(figsize=(22, 4.6 * rows))

gs = fig.add_gridspec(
    rows + 1,
    4,
    width_ratios=[1.2, 2.2, 2.2, 2.2],
    hspace=0.25,
    wspace=0.35
)

# =========================================================
# ШАПКА
# =========================================================
headers = [
    "ПОКАЗАТЕЛЬ",
    SCHEME_1_NAME,
    SCHEME_2_NAME,
    "Δ = (A − B)/A ·100%"
]

for j in range(4):
    ax = fig.add_subplot(gs[0, j])
    ax.axis("off")
    ax.text(
        0.5, 0.6,
        headers[j],
        ha="center",
        va="center",
        fontsize=HEADER_SIZE,
        fontweight="bold"
    )

# =========================================================
# ГРАФИКИ
# =========================================================
for i, (key, label, unit) in enumerate(METRICS):
    row = i + 1

    (r1, c1), (r2, c2) = find_metric(ws, key)

    x1, y1, z1 = trim(*read_block(ws, r1, c1))
    x2, y2, z2 = trim(*read_block(ws, r2, c2))

    d = diff(z1, z2)

    vmin = np.nanmin([z1, z2])
    vmax = np.nanmax([z1, z2])

    ax0 = fig.add_subplot(gs[row, 0])
    ax0.axis("off")
    ax0.text(
        0.5, 0.6,
        label,
        ha="center",
        va="center",
        fontsize=METRIC_LABEL_SIZE,
        fontweight="bold"
    )

    # Секционированная — без colorbar
    ax1 = fig.add_subplot(gs[row, 1])
    ax1.imshow(
        z1,
        aspect="auto",
        cmap=MATTE_METRIC_CMAP,
        vmin=vmin,
        vmax=vmax
    )
    style_axis(ax1, x1, y1, show_x=(i == rows - 1), show_y=True)

    # Двойная — общий colorbar для обеих схем
    ax2 = fig.add_subplot(gs[row, 2])
    im2 = ax2.imshow(
        z2,
        aspect="auto",
        cmap=MATTE_METRIC_CMAP,
        vmin=vmin,
        vmax=vmax
    )
    style_axis(ax2, x2, y2, show_x=(i == rows - 1), show_y=False)

    cb2 = fig.colorbar(im2, ax=ax2, fraction=0.046, pad=0.04)
    setup_metric_colorbar(cb2, key, label, unit)

    # Дельта
    ax3 = fig.add_subplot(gs[row, 3])
    lim = np.nanmax(np.abs(d))
    if not np.isfinite(lim) or lim == 0:
        lim = 1.0

    im3 = ax3.imshow(
        d,
        aspect="auto",
        cmap=MATTE_DIFF_CMAP,
        vmin=-lim,
        vmax=lim
    )
    style_axis(ax3, x1, y1, show_x=(i == rows - 1), show_y=False)

    cb3 = fig.colorbar(im3, ax=ax3, fraction=0.046, pad=0.04)
    cb3.formatter = comma_formatter(2)
    cb3.update_ticks()
    cb3.ax.tick_params(labelsize=CB_TICK_SIZE)
    cb3.set_label("Δ, %", fontsize=CB_LABEL_SIZE)

# =========================================================
# СОХРАНЕНИЕ
# =========================================================
os.makedirs(OUTPUT_DIR, exist_ok=True)

out = os.path.join(OUTPUT_DIR, "SS_vs_D_table_compare.png")
fig.savefig(out, dpi=300, bbox_inches="tight")
plt.close(fig)

print("Готово:")
print(out)