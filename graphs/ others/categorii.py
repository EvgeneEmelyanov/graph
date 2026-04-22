import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.colors import LinearSegmentedColormap


# ============================================================
# НАСТРОЙКИ
# ============================================================
FILE = r"D:\SS.xlsx"
RAW_SHEET = "RAW"

METRICS = [
    {
        "value_col": "LCOE, руб/кВт∙ч",
        "cbar_label": "LCOE, руб/кВт·ч",
        "formatter": "comma",
        "decimals": 2,
    },
    {
        "value_col": "LOLP",
        "cbar_label": "LOLP",
        "formatter": "sci10",
        "decimals": 2,
    },
    {
        "value_col": "LPSP",
        "cbar_label": "LPSP",
        "formatter": "sci10",
        "decimals": 2,
    },
]

CMAP = LinearSegmentedColormap.from_list(
    "custom_blue_terracotta",
    ["#f2f2f2", "#4d7a99"]
)

FIGSIZE = (12.8, 9.2)
DPI = 160

BASE_FONT = 15
AXIS_LABEL_SIZE = BASE_FONT + 18
TICK_SIZE = BASE_FONT + 14
CBAR_SIZE = BASE_FONT + 8


# ============================================================
# ФОРМАТ ЧИСЕЛ
# ============================================================
SUPERSCRIPTS = str.maketrans("0123456789-", "⁰¹²³⁴⁵⁶⁷⁸⁹⁻")


def format_comma_value(x, decimals=2):
    if not np.isfinite(x):
        return ""
    return f"{x:.{decimals}f}".replace(".", ",")


def format_sci10_value(x, decimals=2):
    if not np.isfinite(x):
        return ""
    if x == 0:
        return "0"

    ax = abs(x)
    exp = int(np.floor(np.log10(ax)))
    mantissa = x / (10 ** exp)

    mantissa_str = f"{mantissa:.{decimals}f}".replace(".", ",")
    exp_str = str(exp).translate(SUPERSCRIPTS)

    return f"{mantissa_str}×10{exp_str}"


# ============================================================
# ЗАГРУЗКА RAW ОДИН РАЗ
# ============================================================
def load_raw(path, sheet_name="RAW"):
    df = pd.read_excel(path, sheet_name=sheet_name, header=1)
    df = df.dropna(how="all").copy()
    df.columns = [str(c).strip() for c in df.columns]

    required_base_cols = ["param1", "param2"]
    for col in required_base_cols:
        if col not in df.columns:
            raise KeyError(
                f"В файле {path} не найден столбец '{col}'. "
                f"Доступные столбцы: {df.columns.tolist()}"
            )

    df["param1"] = pd.to_numeric(df["param1"], errors="coerce")
    df["param2"] = pd.to_numeric(df["param2"], errors="coerce")
    df = df.dropna(subset=["param1", "param2"]).copy()

    return df


# ============================================================
# МАТРИЦА ДЛЯ ОДНОЙ МЕТРИКИ
# ============================================================
def build_metric_matrix(df, value_col):
    if value_col not in df.columns:
        raise KeyError(
            f"Не найден столбец '{value_col}'. "
            f"Доступные столбцы: {df.columns.tolist()}"
        )

    work = df[["param1", "param2", value_col]].copy()
    work[value_col] = pd.to_numeric(work[value_col], errors="coerce")
    work = work.dropna(subset=[value_col]).copy()

    pivot = work.pivot_table(
        index="param1",
        columns="param2",
        values=value_col,
        aggfunc="mean"
    )

    pivot = pivot.sort_index().sort_index(axis=1)

    y_vals = pivot.index.to_numpy(dtype=float)
    x_vals = pivot.columns.to_numpy(dtype=float)
    z = pivot.to_numpy(dtype=float)

    return x_vals, y_vals, z


# ============================================================
# ТРЕУГОЛЬНАЯ ДОПУСТИМАЯ ОБЛАСТЬ
# ============================================================
def build_triangle_mask(x_vals, y_vals, z):
    x_grid, y_grid = np.meshgrid(x_vals, y_vals)
    triangle_mask = (x_grid + y_grid) > 1.0 + 1e-12
    full_mask = np.isnan(z) | triangle_mask
    return np.ma.masked_where(full_mask, z)


# ============================================================
# ПОДПИСИ ОСЕЙ: 0,0 0,2 0,4 0,6 0,8 1,0
# ============================================================
def build_sparse_ticks(values, step=0.2, decimals=1):
    tick_positions = []
    tick_labels = []

    target_values = np.arange(0.0, 1.0 + 1e-9, step)

    for tv in target_values:
        idx = np.argmin(np.abs(values - tv))
        if np.isclose(values[idx], tv, atol=1e-9):
            tick_positions.append(idx)
            tick_labels.append(format_comma_value(values[idx], decimals))

    return tick_positions, tick_labels


# ============================================================
# TICKS И LABELS ДЛЯ COLORBAR
# ============================================================
def make_decimal_ticks_and_labels(vmin, vmax, n=5, decimals=2):
    if not np.isfinite(vmin) or not np.isfinite(vmax):
        return np.array([]), []

    if vmin == vmax:
        return np.array([vmin]), [format_comma_value(vmin, decimals)]

    ticks = np.linspace(vmin, vmax, n)
    labels = [format_comma_value(t, decimals) for t in ticks]

    while len(set(labels)) < len(labels) and n > 2:
        n -= 1
        ticks = np.linspace(vmin, vmax, n)
        labels = [format_comma_value(t, decimals) for t in ticks]

    return ticks, labels


def make_sci_ticks_and_labels(vmin, vmax, n=5, decimals=2):
    if not np.isfinite(vmin) or not np.isfinite(vmax):
        return np.array([]), []

    if vmin == vmax:
        return np.array([vmin]), [format_sci10_value(vmin, decimals)]

    ticks = np.linspace(vmin, vmax, n)
    labels = [format_sci10_value(t, decimals) for t in ticks]

    while len(set(labels)) < len(labels) and n > 2:
        n -= 1
        ticks = np.linspace(vmin, vmax, n)
        labels = [format_sci10_value(t, decimals) for t in ticks]

    return ticks, labels


# ============================================================
# ПОДГОТОВКА ВСЕХ МЕТРИК ЗАРАНЕЕ
# ============================================================
def prepare_all_metrics(df, metrics):
    prepared = []

    for metric in metrics:
        x_vals, y_vals, z = build_metric_matrix(df, metric["value_col"])
        masked = build_triangle_mask(x_vals, y_vals, z)

        vals = masked.compressed()
        if len(vals) == 0:
            raise ValueError(f"Для метрики '{metric['value_col']}' нет допустимых данных.")

        xtick_positions, xtick_labels = build_sparse_ticks(x_vals, step=0.2, decimals=1)
        ytick_positions, ytick_labels = build_sparse_ticks(y_vals, step=0.2, decimals=1)

        prepared.append({
            "metric": metric,
            "x_vals": x_vals,
            "y_vals": y_vals,
            "masked": masked,
            "vmin": vals.min(),
            "vmax": vals.max(),
            "xtick_positions": xtick_positions,
            "xtick_labels": xtick_labels,
            "ytick_positions": ytick_positions,
            "ytick_labels": ytick_labels,
        })

    return prepared


# ============================================================
# ОТРИСОВКА ОДНОГО ГРАФИКА
# ============================================================
def plot_single_metric(prepared_item):
    metric = prepared_item["metric"]
    x_vals = prepared_item["x_vals"]
    y_vals = prepared_item["y_vals"]
    masked = prepared_item["masked"]

    fig, ax = plt.subplots(figsize=FIGSIZE, dpi=DPI)

    im = ax.imshow(
        masked,
        origin="upper",
        aspect="equal",
        cmap=CMAP,
        vmin=prepared_item["vmin"],
        vmax=prepared_item["vmax"],
        interpolation="nearest"
    )

    ax.set_xlabel("Доля нагрузки 2 категории", fontsize=AXIS_LABEL_SIZE, labelpad=18)
    ax.xaxis.set_label_position("top")
    ax.xaxis.tick_top()

    ax.set_ylabel("Доля нагрузки 1 категории", fontsize=AXIS_LABEL_SIZE, labelpad=18)

    ax.set_xticks(prepared_item["xtick_positions"])
    ax.set_yticks(prepared_item["ytick_positions"])

    ax.set_xticklabels(prepared_item["xtick_labels"], fontsize=TICK_SIZE)
    ax.set_yticklabels(prepared_item["ytick_labels"], fontsize=TICK_SIZE)

    ax.set_xticks(np.arange(-0.5, len(x_vals), 1), minor=True)
    ax.set_yticks(np.arange(-0.5, len(y_vals), 1), minor=True)
    ax.grid(which="minor", color="white", linestyle="-", linewidth=0.8)
    ax.tick_params(which="minor", bottom=False, left=False)

    ax.tick_params(axis="x", pad=10)
    ax.tick_params(axis="y", pad=8)

    cbar = fig.colorbar(
        im,
        ax=ax,
        fraction=0.040,
        pad=0.030
    )
    cbar.set_label(metric["cbar_label"], fontsize=CBAR_SIZE)
    cbar.ax.tick_params(labelsize=CBAR_SIZE)

    if metric["formatter"] == "sci10":
        ticks, labels = make_sci_ticks_and_labels(
            prepared_item["vmin"],
            prepared_item["vmax"],
            n=5,
            decimals=metric["decimals"]
        )
    else:
        ticks, labels = make_decimal_ticks_and_labels(
            prepared_item["vmin"],
            prepared_item["vmax"],
            n=5,
            decimals=metric["decimals"]
        )

    cbar.set_ticks(ticks)
    cbar.set_ticklabels(labels)

    plt.subplots_adjust(left=0.12, right=0.80, top=0.87, bottom=0.10)
    plt.show()


# ============================================================
# ОСНОВНОЙ БЛОК
# ============================================================
df = load_raw(FILE, RAW_SHEET)
prepared_metrics = prepare_all_metrics(df, METRICS)

for item in prepared_metrics:
    plot_single_metric(item)