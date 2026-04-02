import os
import re
from pathlib import Path

import numpy as np
import matplotlib.pyplot as plt
from matplotlib.colors import LinearSegmentedColormap
from matplotlib.ticker import ScalarFormatter
from mpl_toolkits.mplot3d import Axes3D  # noqa: F401
from openpyxl import load_workbook
from openpyxl.utils import range_boundaries


# ============================================================
# НАСТРОЙКИ
# ============================================================
FILE = r"D:\2.xlsx"
SHEET = "SWEEP_2"
OUTPUT_DIR = r"D:\results"

# Выбор критерия по номеру
TARGET_INDEX = 7

# Типовой размер блока
MATRIX_RANGE = "A2:J13"

# Подписи осей X/Y
# X_AXIS_LABEL = "Макс. ток разряда"
X_AXIS_LABEL = "Остаточный заряд"
Y_AXIS_LABEL = "Доля емкости СНЭ"

FIGSIZE = (11, 8)
DPI = 220

# Поворот: на 90° против часовой относительно старого вида
ELEV = 30
AZIM = 25

SURFACE_ALPHA = 0.98
EDGE_COLOR = "#6f6f6f"
EDGE_LINEWIDTH = 0.35
SHOW_COLORBAR = True

MATTE_DIFF_CMAP = LinearSegmentedColormap.from_list(
    "matte_diff",
    [
        "#4d7a99",
        "#8fa8b7",
        "#e9e9e9",
        "#ddb48c",
        "#c98d5b",
    ]
)

# ============================================================
# ПОДПИСИ КРИТЕРИЕВ
# ключ = как в Excel
# значение = как подписывать на графике / colorbar
# ============================================================
METRICS = [
    ("LCOE, руб/кВт∙ч", "LCOE, руб/кВт·ч"),
    ("Расход топлива, тыс.тонн", "Расход топлива, тыс. т"),
    ("Моточасы, тыс.мч", "Моточасы, тыс. мч"),
    ("ENS,кВт∙ч", "ENS, кВт·ч"),
    ("LOLH", "LOLH, ч"),
    ("ENS_evtN", "Кол-во событий ENS"),
    ("ENS_evtMaxH", "Макс. длит. ENS, ч"),
    ("LOLP", "LOLP"),
    ("LPSP", "LPSP"),
    ("ENS1_mean", "ENS1_mean, кВт·ч"),
    ("ENS2_mean", "ENS2_mean, кВт·ч"),
    ("FailDg", "Отказы ДГУ, кол-во"),
    ("FailBt", "Отказы АКБ, кол-во"),
    ("BtRepl", "Замены АКБ, кол-во"),
]

METRIC_LABELS = dict(METRICS)


# ============================================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ============================================================
def ensure_dir(path: str):
    Path(path).mkdir(parents=True, exist_ok=True)


def safe_filename(name: str):
    name = str(name).strip()
    name = re.sub(r'[\\/*?:"<>|]', "_", name)
    name = re.sub(r"\s+", " ", name).strip()
    name = name.replace("∙", "_")
    name = name.replace("·", "_")
    name = name.replace(",", "_")
    return name


def is_numeric_like(value):
    if value is None:
        return False
    if isinstance(value, (int, float)):
        return True
    s = str(value).strip().replace(" ", "").replace(",", ".")
    try:
        float(s)
        return True
    except Exception:
        return False


def to_float(value):
    if value is None or str(value).strip() == "":
        return np.nan
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip().replace(" ", "").replace(",", ".")
    return float(s)


def read_matrix_shape(matrix_range: str):
    min_col, min_row, max_col, max_row = range_boundaries(matrix_range)

    total_rows = max_row - min_row + 1
    total_cols = max_col - min_col + 1

    if total_rows < 3 or total_cols < 3:
        raise ValueError("MATRIX_RANGE слишком мал. Нужен минимум 3x3.")

    return {
        "min_col": min_col,
        "min_row": min_row,
        "max_col": max_col,
        "max_row": max_row,
        "total_rows": total_rows,
        "total_cols": total_cols,
    }


def detect_criteria(ws, matrix_range: str):
    shape = read_matrix_shape(matrix_range)

    min_col = shape["min_col"]
    criteria = []

    for r in range(1, ws.max_row + 1):
        title = ws.cell(r, min_col).value
        if title is None:
            continue

        title_str = str(title).strip()
        if not title_str:
            continue

        if is_numeric_like(title_str):
            continue

        matrix_start_row = r + 1
        matrix_end_row = matrix_start_row + shape["total_rows"] - 1

        if matrix_end_row > ws.max_row:
            continue

        ok = True

        # X-подписи
        for c in range(min_col + 1, shape["max_col"] + 1):
            if not is_numeric_like(ws.cell(matrix_start_row, c).value):
                ok = False
                break

        # Y-подписи
        if ok:
            for rr in range(matrix_start_row + 1, matrix_end_row + 1):
                if not is_numeric_like(ws.cell(rr, min_col).value):
                    ok = False
                    break

        # тело матрицы
        if ok:
            for rr in range(matrix_start_row + 1, matrix_end_row + 1):
                for c in range(min_col + 1, shape["max_col"] + 1):
                    if not is_numeric_like(ws.cell(rr, c).value):
                        ok = False
                        break
                if not ok:
                    break

        if ok:
            criteria.append(
                {
                    "excel_title": title_str,
                    "display_title": METRIC_LABELS.get(title_str, title_str),
                    "matrix_start_row": matrix_start_row,
                }
            )

    return criteria


def extract_matrix(ws, matrix_start_row: int, matrix_range: str):
    shape = read_matrix_shape(matrix_range)

    min_col = shape["min_col"]
    max_col = shape["max_col"]
    total_rows = shape["total_rows"]

    matrix_end_row = matrix_start_row + total_rows - 1

    x = [to_float(ws.cell(matrix_start_row, c).value)
         for c in range(min_col + 1, max_col + 1)]

    y = [to_float(ws.cell(r, min_col).value)
         for r in range(matrix_start_row + 1, matrix_end_row + 1)]

    z = []
    for r in range(matrix_start_row + 1, matrix_end_row + 1):
        row_vals = []
        for c in range(min_col + 1, max_col + 1):
            row_vals.append(to_float(ws.cell(r, c).value))
        z.append(row_vals)

    x = np.array(x, dtype=float)
    y = np.array(y, dtype=float)
    z = np.array(z, dtype=float)

    X, Y = np.meshgrid(x, y)
    return x, y, X, Y, z


def plot_surface(excel_title, display_title, x, y, X, Y, Z):
    ensure_dir(OUTPUT_DIR)

    fig = plt.figure(figsize=FIGSIZE, dpi=DPI)
    ax = fig.add_subplot(111, projection="3d")

    surf = ax.plot_surface(
        X, Y, Z,
        cmap=MATTE_DIFF_CMAP,
        linewidth=EDGE_LINEWIDTH,
        edgecolor=EDGE_COLOR,
        alpha=SURFACE_ALPHA,
        antialiased=True,
    )

    # Заголовок можно оставить как display_title
    ax.set_title(display_title)

    ax.set_xlabel(X_AXIS_LABEL)
    ax.set_ylabel(Y_AXIS_LABEL)

    # ВАЖНО: подпись оси Z не ставим
    ax.set_zlabel("")
    ax.zaxis.label.set_visible(False)

    ax.view_init(elev=ELEV, azim=AZIM)

    ax.set_xticks(x)
    ax.set_yticks(y)

    z_formatter = ScalarFormatter(useMathText=True)
    z_formatter.set_powerlimits((-3, 4))
    ax.zaxis.set_major_formatter(z_formatter)

    if SHOW_COLORBAR:
        cbar = fig.colorbar(surf, ax=ax, shrink=0.7)
        cbar.set_label(display_title)

    out_path = os.path.join(OUTPUT_DIR, safe_filename(display_title) + ".png")
    fig.savefig(out_path, bbox_inches="tight")
    plt.close(fig)

    print(f"Готово: {out_path}")


# ============================================================
# ОСНОВНОЙ КОД
# ============================================================
def main():
    wb = load_workbook(FILE, data_only=True)
    ws = wb[SHEET]

    criteria = detect_criteria(ws, MATRIX_RANGE)

    if not criteria:
        raise RuntimeError("Критерии не найдены")

    print("\nСписок критериев:")
    for i, c in enumerate(criteria, start=1):
        print(f"{i}. {c['excel_title']}  -->  {c['display_title']}")

    if TARGET_INDEX < 1 or TARGET_INDEX > len(criteria):
        raise ValueError(f"Неверный TARGET_INDEX: {TARGET_INDEX}")

    selected = criteria[TARGET_INDEX - 1]

    print(f"\nВыбран критерий:")
    print(f"Excel:   {selected['excel_title']}")
    print(f"Подпись: {selected['display_title']}")

    x, y, X, Y, Z = extract_matrix(
        ws,
        selected["matrix_start_row"],
        MATRIX_RANGE
    )

    plot_surface(
        excel_title=selected["excel_title"],
        display_title=selected["display_title"],
        x=x, y=y, X=X, Y=Y, Z=Z
    )


if __name__ == "__main__":
    main()