import math

import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter, MaxNLocator, LogLocator


# ============================================================
# НАСТРОЙКИ
# ============================================================
FILES = {
    "SS200": r"D:\2.xlsx",
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
PLOT_MOTO = True

# База для перевода оси X в %
BASE_DG_POWER = 1346.0

# В Excel здесь уже лежит суммарная мощность ДГУ
X_COL = "param2"

# Фиксированная ось X
X_MIN = 100
X_MAX = 200
X_STEP = 10


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
# ПОДГОТОВКА ОСИ X
# X = готовая суммарная мощность ДГУ из Excel, переведенная в %
# ============================================================
def prepare_x_percent(df, y_col):
    work = df[df["param1"].isin(DG_COUNTS)].copy()

    work["param1"] = pd.to_numeric(work["param1"], errors="coerce")
    work[X_COL] = pd.to_numeric(work[X_COL], errors="coerce")
    work[y_col] = pd.to_numeric(work[y_col], errors="coerce")

    work = work.dropna(subset=["param1", X_COL, y_col]).copy()
    work["x_percent"] = work[X_COL] / BASE_DG_POWER * 100.0

    return work


# ============================================================
# ГРАНИЦЫ ОСИ Y
# ============================================================
def collect_global_y_limits(y_col, positive_only=False):
    all_values = []

    for _, path in FILES.items():
        df = load_raw(path)
        prepared = prepare_x_percent(df, y_col)

        for dg in DG_COUNTS:
            sub = prepared[prepared["param1"] == dg].copy()
            if sub.empty:
                continue

            vals = pd.to_numeric(sub[y_col], errors="coerce").dropna()

            if positive_only:
                vals = vals[vals > 0]

            all_values.extend(vals.tolist())

    if not all_values:
        raise ValueError(f"Нет данных для столбца: {y_col}")

    return min(all_values), max(all_values)


def nice_limits_regular(y_min, y_max):
    return math.floor(y_min), math.ceil(y_max)


def nice_limits_log_auto(y_min, y_max, pad_decades=0.08):
    """
    Более аккуратные границы для логарифмической оси.
    Не прыгает слишком широко, но оставляет небольшой отступ.
    """
    if y_min <= 0 or y_max <= 0:
        raise ValueError("Для логарифмической шкалы все значения должны быть > 0")

    log_min = math.log10(y_min)
    log_max = math.log10(y_max)

    if math.isclose(log_min, log_max):
        log_min -= 0.2
        log_max += 0.2
    else:
        log_min -= pad_decades
        log_max += pad_decades

    lower = 10 ** log_min
    upper = 10 ** log_max

    return lower, upper


# ============================================================
# ГРАФИК
# ============================================================
def plot_case(case_name, path, y_col, y_label, output_suffix, y_mode, y_limits):
    df = load_raw(path)
    prepared = prepare_x_percent(df, y_col)

    plt.figure(figsize=FIGSIZE)

    for dg in DG_COUNTS:
        sub = prepared[prepared["param1"] == dg].copy()
        sub = sub.sort_values("x_percent")

        x = pd.to_numeric(sub["x_percent"], errors="coerce")
        y = pd.to_numeric(sub[y_col], errors="coerce")

        if y_mode == "log":
            mask = y > 0
            x = x[mask]
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

    plt.xlabel("Суммарная мощность ДГУ, %", fontsize=FONT_SIZE)
    plt.ylabel(y_label, fontsize=FONT_SIZE)

    plt.xticks(range(X_MIN, X_MAX + 1, X_STEP), fontsize=TICK_SIZE)
    plt.yticks(fontsize=TICK_SIZE)

    ax = plt.gca()
    ax.set_xlim(X_MIN, X_MAX)
    ax.xaxis.set_major_formatter(comma_formatter(0))

    if y_mode == "regular_int":
        ax.set_ylim(y_limits)
        ax.yaxis.set_major_locator(MaxNLocator(integer=True))
        ax.yaxis.set_major_formatter(int_formatter())

    elif y_mode == "log":
        ax.set_yscale("log")
        ax.set_ylim(y_limits)

        # основные деления только по степеням 10
        ax.yaxis.set_major_locator(LogLocator(base=10.0))
        ax.yaxis.set_major_formatter(sci_power10_formatter(decimals=0))

        # можно отключить подписи у промежуточных делений
        ax.yaxis.set_minor_formatter(FuncFormatter(lambda x, pos: ""))

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
    lmin, lmax = collect_global_y_limits("LCOE, руб/кВт∙ч")
    l_limits = nice_limits_regular(lmin, lmax)

    for name, path in FILES.items():
        plot_case(
            case_name=name,
            path=path,
            y_col="LCOE, руб/кВт∙ч",
            y_label="LCOE, руб/кВт∙ч",
            output_suffix="LCOE_vs_total_DG_power_percent",
            y_mode="regular_int",
            y_limits=l_limits
        )

# ---------- LOLP ----------
if PLOT_LOLP:
    lolp_min, lolp_max = collect_global_y_limits("LOLP", positive_only=True)
    lolp_limits = nice_limits_log_auto(lolp_min, lolp_max, pad_decades=0.08)

    for name, path in FILES.items():
        plot_case(
            case_name=name,
            path=path,
            y_col="LOLP",
            y_label="LOLP",
            output_suffix="LOLP_vs_total_DG_power_percent",
            y_mode="log",
            y_limits=lolp_limits
        )

# ---------- Моточасы ----------
if PLOT_MOTO:
    mmin, mmax = collect_global_y_limits("Моточасы")
    m_limits = nice_limits_regular(mmin, mmax)

    for name, path in FILES.items():
        plot_case(
            case_name=name,
            path=path,
            y_col="Моточасы",
            y_label="Моточасы, тыс. мч.",
            output_suffix="Moto_vs_total_DG_power_percent",
            y_mode="regular_int",
            y_limits=m_limits
        )