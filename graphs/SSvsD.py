import os
import numpy as np
import matplotlib.pyplot as plt
from openpyxl import load_workbook

# =========================
# НАСТРОЙКИ
# =========================
EXCEL_PATH = r"D:\comb_results.xlsx"  # при необходимости замени на свой путь
SHEET_NAME = "SWEEP_2"
OUTPUT_DIR = r"D:\results"

SCHEME_1_NAME = "Секционированная система"
SCHEME_2_NAME = "Двойная система"

PEAK_LOAD_KW = 1346.0   # пиковая нагрузка
WT_COUNT = 2            # число ВЭУ
SIM_YEARS = 20          # длительность моделирования для LOLH

# Метрики:
# 1) как искать в Excel
# 2) как подписывать на графике
METRICS = [
    ("ENS", "ENS"),
    ("LOLH", "LOLH, ч/год"),
    ("ENS_evtN", "Кол-во событий ENS"),
    ("ENS_evtMaxH", "Макс. длит. ENS, ч"),
]

# Возможные варианты имен в заголовках Excel
HEADER_ALIASES = {
    "ENS": [
        "ENS",
        "ENS,кВт∙ч",
        "ENS, кВт∙ч",
        "ENS,кВт·ч",
        "ENS, кВт·ч",
    ],
    "LOLH": [
        "LOLH",
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


# =========================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# =========================
def normalize_text(x):
    if x is None:
        return ""
    return str(x).strip().replace(" ", "").replace("·", "∙")


def to_float(v):
    if v is None:
        return np.nan
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(" ", "").replace(",", ".")
    if s == "":
        return np.nan
    return float(s)


def find_metric_start_row(ws, metric_key):
    aliases_norm = {normalize_text(a) for a in HEADER_ALIASES[metric_key]}

    for r in range(1, ws.max_row + 1):
        v1 = normalize_text(ws.cell(r, 1).value)
        v14 = normalize_text(ws.cell(r, 14).value)
        if v1 in aliases_norm or v14 in aliases_norm:
            return r

    raise ValueError(f"Не найден блок для метрики: {metric_key}")


def extract_two_scheme_blocks(ws, metric_key):
    """
    Ожидаемая структура блока:
      row      : название метрики
      row + 1  : строка X (емкость СНЭ)
      row + 2 ... row + 13 : строки Y (мощность одной ВЭУ) + значения

    Левая схема:
      Y-подписи: col=1
      X: cols 2..12
      data: rows row+2..row+13, cols 2..12

    Правая схема:
      Y-подписи: col=14
      X: cols 15..25
      data: rows row+2..row+13, cols 15..25
    """
    start_row = find_metric_start_row(ws, metric_key)

    x1_raw = [to_float(ws.cell(start_row + 1, c).value) for c in range(2, 13)]
    y1_raw = [to_float(ws.cell(r, 1).value) for r in range(start_row + 2, start_row + 14)]
    z1 = np.array(
        [[to_float(ws.cell(r, c).value) for c in range(2, 13)]
         for r in range(start_row + 2, start_row + 14)],
        dtype=float
    )

    x2_raw = [to_float(ws.cell(start_row + 1, c).value) for c in range(15, 26)]
    y2_raw = [to_float(ws.cell(r, 14).value) for r in range(start_row + 2, start_row + 14)]
    z2 = np.array(
        [[to_float(ws.cell(r, c).value) for c in range(15, 26)]
         for r in range(start_row + 2, start_row + 14)],
        dtype=float
    )

    return x1_raw, y1_raw, z1, x2_raw, y2_raw, z2


def transform_metric(metric_key, arr):
    arr = arr.astype(float).copy()

    # LOLH из "за 20 лет" -> "ч/год"
    if metric_key == "LOLH":
        arr = arr / SIM_YEARS

    return arr


def battery_capacity_labels(x_raw):
    # 0.0 -> 0, 0.5 -> 50, 1.0 -> 100
    return [f"{int(round(v * 100))}" for v in x_raw]


def wt_power_percent_labels(y_raw, wt_count=WT_COUNT, peak_load_kw=PEAK_LOAD_KW):
    # round((P_one_WT * WT_COUNT / peak_load_kw) * 100)
    return [f"{int(round(v * wt_count / peak_load_kw * 100))}" for v in y_raw]


def safe_rel_diff(a, b):
    """
    Δ% = (b - a) / a * 100
    Если a == 0 -> nan
    """
    out = np.full_like(a, np.nan, dtype=float)
    mask = np.abs(a) > 1e-12
    out[mask] = (b[mask] - a[mask]) / a[mask] * 100.0
    return out


def setup_heatmap_axis(ax, xlabels, ylabels, title, show_xlabel=False, show_ylabel=False):
    ax.set_title(title, fontsize=12, fontweight="bold")

    ax.set_xticks(np.arange(len(xlabels)))
    ax.set_xticklabels(xlabels, fontsize=9)

    ax.set_yticks(np.arange(len(ylabels)))
    ax.set_yticklabels(ylabels, fontsize=9)

    if show_xlabel:
        ax.set_xlabel("Емкость СНЭ, %", fontsize=10)
    if show_ylabel:
        ax.set_ylabel("Мощность ВЭУ, %", fontsize=10)

    ax.set_xticks(np.arange(-0.5, len(xlabels), 1), minor=True)
    ax.set_yticks(np.arange(-0.5, len(ylabels), 1), minor=True)
    ax.grid(which="minor", color="white", linestyle="-", linewidth=0.7, alpha=0.5)
    ax.tick_params(which="minor", bottom=False, left=False)


def save_individual_figure(metric_label, z1, z2, diff, xlabels, ylabels, out_path):
    fig, axes = plt.subplots(1, 3, figsize=(16, 4.8), constrained_layout=True)

    common_vmin = np.nanmin([np.nanmin(z1), np.nanmin(z2)])
    common_vmax = np.nanmax([np.nanmax(z1), np.nanmax(z2)])

    im1 = axes[0].imshow(
        z1, aspect="auto", origin="upper",
        vmin=common_vmin, vmax=common_vmax, cmap="viridis"
    )
    setup_heatmap_axis(
        axes[0], xlabels, ylabels,
        f"{metric_label} — {SCHEME_1_NAME}",
        show_xlabel=True, show_ylabel=True
    )

    im2 = axes[1].imshow(
        z2, aspect="auto", origin="upper",
        vmin=common_vmin, vmax=common_vmax, cmap="viridis"
    )
    setup_heatmap_axis(
        axes[1], xlabels, ylabels,
        f"{metric_label} — {SCHEME_2_NAME}",
        show_xlabel=True, show_ylabel=False
    )

    lim = np.nanmax(np.abs(diff))
    if not np.isfinite(lim) or lim == 0:
        lim = 1.0

    im3 = axes[2].imshow(
        diff, aspect="auto", origin="upper",
        vmin=-lim, vmax=lim, cmap="RdBu_r"
    )
    setup_heatmap_axis(
        axes[2], xlabels, ylabels,
        f"{metric_label} — Δ, %",
        show_xlabel=True, show_ylabel=False
    )

    cbar1 = fig.colorbar(im2, ax=axes[:2], fraction=0.025, pad=0.02)
    cbar1.ax.set_ylabel(metric_label, rotation=90)

    cbar2 = fig.colorbar(im3, ax=axes[2], fraction=0.046, pad=0.04)
    cbar2.ax.set_ylabel("Δ, %", rotation=90)

    fig.savefig(out_path, dpi=300, bbox_inches="tight")
    plt.close(fig)


# =========================
# ОСНОВНОЙ КОД
# =========================
def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    wb = load_workbook(EXCEL_PATH, data_only=True)
    ws = wb[SHEET_NAME]

    prepared = []

    for metric_key, metric_label in METRICS:
        x1_raw, y1_raw, z1, x2_raw, y2_raw, z2 = extract_two_scheme_blocks(ws, metric_key)

        # Преобразование метрик
        z1 = transform_metric(metric_key, z1)
        z2 = transform_metric(metric_key, z2)

        # Проверка совпадения сетки
        if len(x1_raw) != len(x2_raw) or len(y1_raw) != len(y2_raw):
            raise ValueError(f"Размер сеток у схем не совпадает для метрики {metric_key}")

        if not np.allclose(np.array(x1_raw), np.array(x2_raw), equal_nan=True):
            raise ValueError(f"Ось X у схем не совпадает для метрики {metric_key}")

        if not np.allclose(np.array(y1_raw), np.array(y2_raw), equal_nan=True):
            raise ValueError(f"Ось Y у схем не совпадает для метрики {metric_key}")

        xlabels = battery_capacity_labels(x1_raw)
        ylabels = wt_power_percent_labels(y1_raw)

        diff = safe_rel_diff(z1, z2)

        prepared.append({
            "metric_key": metric_key,
            "metric_label": metric_label,
            "xlabels": xlabels,
            "ylabels": ylabels,
            "z1": z1,
            "z2": z2,
            "diff": diff,
        })

        save_individual_figure(
            metric_label=metric_label,
            z1=z1,
            z2=z2,
            diff=diff,
            xlabels=xlabels,
            ylabels=ylabels,
            out_path=os.path.join(OUTPUT_DIR, f"{metric_key}_compare.png")
        )

    # Сводная фигура: 4 строки × 3 столбца
    fig, axes = plt.subplots(
        len(prepared), 3,
        figsize=(17, 4.2 * len(prepared)),
        constrained_layout=True
    )

    if len(prepared) == 1:
        axes = np.array([axes])

    for i, item in enumerate(prepared):
        z1 = item["z1"]
        z2 = item["z2"]
        diff = item["diff"]
        xlabels = item["xlabels"]
        ylabels = item["ylabels"]
        metric_label = item["metric_label"]

        common_vmin = np.nanmin([np.nanmin(z1), np.nanmin(z2)])
        common_vmax = np.nanmax([np.nanmax(z1), np.nanmax(z2)])

        # Левая панель
        ax = axes[i, 0]
        im_left = ax.imshow(
            z1, aspect="auto", origin="upper",
            vmin=common_vmin, vmax=common_vmax, cmap="viridis"
        )
        setup_heatmap_axis(
            ax, xlabels, ylabels,
            f"{metric_label} — {SCHEME_1_NAME}",
            show_xlabel=(i == len(prepared) - 1),
            show_ylabel=True
        )

        # Средняя панель
        ax = axes[i, 1]
        im_mid = ax.imshow(
            z2, aspect="auto", origin="upper",
            vmin=common_vmin, vmax=common_vmax, cmap="viridis"
        )
        setup_heatmap_axis(
            ax, xlabels, ylabels,
            f"{metric_label} — {SCHEME_2_NAME}",
            show_xlabel=(i == len(prepared) - 1),
            show_ylabel=False
        )

        # Правая панель: относительная разница
        ax = axes[i, 2]
        lim = np.nanmax(np.abs(diff))
        if not np.isfinite(lim) or lim == 0:
            lim = 1.0

        im_right = ax.imshow(
            diff, aspect="auto", origin="upper",
            vmin=-lim, vmax=lim, cmap="RdBu_r"
        )
        setup_heatmap_axis(
            ax, xlabels, ylabels,
            f"{metric_label} — Δ, %",
            show_xlabel=(i == len(prepared) - 1),
            show_ylabel=False
        )

        cbar_left = fig.colorbar(im_mid, ax=[axes[i, 0], axes[i, 1]], fraction=0.018, pad=0.015)
        cbar_left.ax.set_ylabel(metric_label, rotation=90)

        cbar_right = fig.colorbar(im_right, ax=axes[i, 2], fraction=0.046, pad=0.02)
        cbar_right.ax.set_ylabel("Δ, %", rotation=90)

    fig.suptitle(
        "Сравнение схем по критериям надежности\n"
        "Y: Мощность ВЭУ, %   |   X: Емкость СНЭ, %",
        fontsize=16,
        fontweight="bold"
    )

    summary_path = os.path.join(OUTPUT_DIR, "reliability_compare_summary.png")
    fig.savefig(summary_path, dpi=300, bbox_inches="tight")
    plt.close(fig)

    print("Готово.")
    print(f"Сводная фигура: {summary_path}")
    print(f"Отдельные фигуры сохранены в папку: {OUTPUT_DIR}")


if __name__ == "__main__":
    main()