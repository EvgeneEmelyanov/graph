import math

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter, MaxNLocator, LogLocator


# ============================================================
# НАСТРОЙКИ
# ============================================================
FILES = {
    "SS100": r"D:\SS100.xlsx",
    "SS200": r"D:\SS200.xlsx",
    "SS300": r"D:\SS300.xlsx",
}

RAW_SHEET = "RAW"

DG_COUNTS = [6, 8, 10]

COLORS = {
    6: "#426f91",
    8: "#c95f46",
    10: "#6f8f72",
}

FIGSIZE = (6.5, 6.5)
DPI = 300

BASE_FONT = 11
FONT_SIZE = BASE_FONT + 5
LEGEND_SIZE = BASE_FONT + 4
TICK_SIZE = BASE_FONT + 4

PLOT_LCOE = True
PLOT_LOLP = True
PLOT_MOTO = False

# "intersection" - только общий диапазон, где есть все линии
# "union"        - расширенный диапазон, линии могут местами обрываться
GRID_MODE = "intersection"

# шаг сетки по суммарной мощности, кВт
TOTAL_POWER_STEP = 50


# ============================================================
# ФОРМАТ ЧИСЕЛ
# ============================================================
def comma_formatter(decimals=2):
    def _fmt(x, pos):
        s = f"{x:.{decimals}f}"
        return s.replace(".", ",")
    return FuncFormatter(_fmt)


def int_formatter():
    def _fmt(x, pos):
        return f"{int(round(x))}"
    return FuncFormatter(_fmt)


def sci_power10_formatter(decimals=0):
    superscript_map = str.maketrans("0123456789-", "⁰¹²³⁴⁵⁶⁷⁸⁹⁻")

    def _fmt(x, pos):
        if x <= 0:
            return ""

        s = f"{x:.{decimals}e}"
        mantissa, exp = s.split("e")
        exp = int(exp)

        if decimals == 0:
            mantissa_str = str(int(round(float(mantissa))))
        else:
            mantissa_str = mantissa.replace(".", ",")

        exp_str = str(exp).translate(superscript_map)
        return f"{mantissa_str}·10{exp_str}"

    return FuncFormatter(_fmt)


# ============================================================
# ЗАГРУЗКА
# ============================================================
def load_raw(path):
    df = pd.read_excel(path, sheet_name=RAW_SHEET, header=1)
    df = df.dropna(how="all").copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


# ============================================================
# ПОДГОТОВКА: СУММАРНАЯ МОЩНОСТЬ = N_ДГУ * P_ОДНОЙ_ДГУ
# ============================================================
def prepare_total_power_df(df, y_col):
    work = df[df["param1"].isin(DG_COUNTS)].copy()

    work["param1"] = pd.to_numeric(work["param1"], errors="coerce")
    work["param2"] = pd.to_numeric(work["param2"], errors="coerce")
    work[y_col] = pd.to_numeric(work[y_col], errors="coerce")

    work = work.dropna(subset=["param1", "param2", y_col]).copy()
    work["total_power"] = work["param1"] * work["param2"]

    return work


def collapse_duplicate_total_power(sub, y_col):
    grouped = (
        sub.groupby("total_power", as_index=False)[y_col]
        .mean()
        .sort_values("total_power")
        .reset_index(drop=True)
    )
    return grouped


# ============================================================
# ОБЩАЯ СЕТКА ДЛЯ ИНТЕРПОЛЯЦИИ
# ============================================================
def build_common_grid(prepared_df):
    mins = []
    maxs = []

    for dg in DG_COUNTS:
        sub = prepared_df[prepared_df["param1"] == dg].copy()
        if sub.empty:
            continue

        mins.append(sub["total_power"].min())
        maxs.append(sub["total_power"].max())

    if not mins or not maxs:
        raise ValueError("Не удалось построить диапазон по суммарной мощности.")

    if GRID_MODE == "intersection":
        grid_min = max(mins)
        grid_max = min(maxs)
    elif GRID_MODE == "union":
        grid_min = min(mins)
        grid_max = max(maxs)
    else:
        raise ValueError("GRID_MODE должен быть 'intersection' или 'union'.")

    if grid_min >= grid_max:
        raise ValueError(
            "Нет общего диапазона суммарной мощности для всех кривых. "
            "Попробуй GRID_MODE='union'."
        )

    start = int(math.ceil(grid_min / TOTAL_POWER_STEP) * TOTAL_POWER_STEP)
    end = int(math.floor(grid_max / TOTAL_POWER_STEP) * TOTAL_POWER_STEP)

    if start > end:
        start = int(grid_min)
        end = int(grid_max)

    return np.arange(start, end + TOTAL_POWER_STEP, TOTAL_POWER_STEP, dtype=float)


def interpolate_on_grid(prepared_df, y_col, grid):
    result = {}

    for dg in DG_COUNTS:
        sub = prepared_df[prepared_df["param1"] == dg].copy()
        sub = collapse_duplicate_total_power(sub[["total_power", y_col]], y_col=y_col)

        x = sub["total_power"].to_numpy(dtype=float)
        y = sub[y_col].to_numpy(dtype=float)

        if len(x) < 2:
            result[dg] = np.full_like(grid, np.nan, dtype=float)
            continue

        if GRID_MODE == "intersection":
            y_interp = np.interp(grid, x, y)
        else:
            y_interp = np.full_like(grid, np.nan, dtype=float)
            mask = (grid >= x.min()) & (grid <= x.max())
            y_interp[mask] = np.interp(grid[mask], x, y)

        result[dg] = y_interp

    return result


# ============================================================
# ГРАНИЦЫ ОСЕЙ Y ПО ВСЕМ ФАЙЛАМ ПОСЛЕ ИНТЕРПОЛЯЦИИ
# ============================================================
def collect_global_y_limits_interpolated(y_col, positive_only=False):
    all_values = []

    for _, path in FILES.items():
        df = load_raw(path)
        prepared = prepare_total_power_df(df, y_col)
        grid = build_common_grid(prepared)
        interp_data = interpolate_on_grid(prepared, y_col, grid)

        for dg in DG_COUNTS:
            vals = pd.Series(interp_data[dg]).dropna()

            if positive_only:
                vals = vals[vals > 0]

            all_values.extend(vals.tolist())

    if not all_values:
        raise ValueError(f"Нет данных для столбца: {y_col}")

    return min(all_values), max(all_values)


def nice_limits_regular(y_min, y_max):
    return math.floor(y_min), math.ceil(y_max)


def nice_limits_log(y_min, y_max):
    lower = 10 ** math.floor(math.log10(y_min))
    upper = 10 ** math.ceil(math.log10(y_max))
    return lower, upper


# ============================================================
# ПОСТРОЕНИЕ ОДНОГО ГРАФИКА
# ============================================================
def plot_case_interpolated(case_name, path, y_col, y_label, output_suffix, y_mode, y_limits):
    df = load_raw(path)
    prepared = prepare_total_power_df(df, y_col)
    grid = build_common_grid(prepared)
    interp_data = interpolate_on_grid(prepared, y_col, grid)

    plt.figure(figsize=FIGSIZE)

    for dg in DG_COUNTS:
        x = grid.copy()
        y = pd.Series(interp_data[dg]).copy()

        if y_mode == "log":
            mask = y > 0
            x = x[mask.to_numpy()]
            y = y[mask]

        plt.plot(
            x,
            y,
            marker="o",
            linewidth=2.8,
            markersize=8,
            color=COLORS[dg],
            label=f"{dg} ДГУ"
        )

    plt.xlabel("Суммарная установленная мощность ДГУ, кВт", fontsize=FONT_SIZE)
    plt.ylabel(y_label, fontsize=FONT_SIZE)

    plt.xticks(fontsize=TICK_SIZE)
    plt.yticks(fontsize=TICK_SIZE)

    ax = plt.gca()
    ax.xaxis.set_major_formatter(comma_formatter(0))

    if y_mode == "regular_int":
        ax.set_ylim(y_limits)
        ax.yaxis.set_major_locator(MaxNLocator(integer=True))
        ax.yaxis.set_major_formatter(int_formatter())

    elif y_mode == "log":
        ax.set_yscale("log")
        ax.set_ylim(y_limits)
        ax.yaxis.set_major_locator(LogLocator(base=10.0))
        ax.yaxis.set_major_formatter(sci_power10_formatter(decimals=0))

    plt.grid(True, alpha=0.3, which="both")
    plt.legend(fontsize=LEGEND_SIZE)

    plt.tight_layout()
    plt.savefig(f"D:\\{case_name}_{output_suffix}.png", dpi=DPI)
    plt.show()


# ============================================================
# ОСНОВНОЙ БЛОК
# ============================================================

# ---------- LCOE ----------
if PLOT_LCOE:
    lmin, lmax = collect_global_y_limits_interpolated("LCOE, руб/кВт∙ч")
    l_limits = nice_limits_regular(lmin, lmax)

    for name, path in FILES.items():
        plot_case_interpolated(
            case_name=name,
            path=path,
            y_col="LCOE, руб/кВт∙ч",
            y_label="LCOE",
            output_suffix="LCOE_vs_total_DG_power_interp",
            y_mode="regular_int",
            y_limits=l_limits
        )

# ---------- LOLP ----------
if PLOT_LOLP:
    lolp_min, lolp_max = collect_global_y_limits_interpolated("LOLP", positive_only=True)
    lolp_limits = nice_limits_log(lolp_min, lolp_max)

    for name, path in FILES.items():
        plot_case_interpolated(
            case_name=name,
            path=path,
            y_col="LOLP",
            y_label="LOLP",
            output_suffix="LOLP_vs_total_DG_power_interp",
            y_mode="log",
            y_limits=lolp_limits
        )

# ---------- Моточасы ----------
if PLOT_MOTO:
    mmin, mmax = collect_global_y_limits_interpolated("Моточасы")
    m_limits = nice_limits_regular(mmin, mmax)

    for name, path in FILES.items():
        plot_case_interpolated(
            case_name=name,
            path=path,
            y_col="Моточасы",
            y_label="Моточасы",
            output_suffix="Moto_vs_total_DG_power_interp",
            y_mode="regular_int",
            y_limits=m_limits
        )