import re
from typing import List, Tuple, Dict, Optional

import numpy as np
import matplotlib.pyplot as plt
from matplotlib.colors import LinearSegmentedColormap

# =====================================================
# НАСТРОЙКИ (МЕНЯТЬ ЗДЕСЬ)
# =====================================================
PLOT_VARIANT = 2
# 1 — отдельный график для каждой метрики (bars)
# 2 — тепловая карта ST (метрики по X, параметры по Y) + таблица статистик сверху + числа в ячейках

SORT_PARAMS_BY_TOTAL_ST = True
TITLE_PREFIX = "Sobol"

FONT_BASE = 11
FONT_SMALL = 10
MAX_PARAM_LABEL_LEN_FOR_WIDE = 14

# все значения <= 0 или |x| < ZERO_EPS считаются нулём и подписываются как "0"
ZERO_EPS = 5e-4

# Градиент: белый (0) -> оранжевый (max)
WHITE_ORANGE = LinearSegmentedColormap.from_list(
    "white_orange",
    [
        (1.0, 1.0, 1.0),   # white
        (1.0, 0.85, 0.6),  # light orange
        (1.0, 0.55, 0.0),  # orange
    ],
)
# =====================================================

EXPECTED_METRICS = ["LCOE", "ENS", "Fuel", "Moto"]

# Пример:
# LCOE: var=12481,5 std=111,720 range=[29,0663..818,510]
METRIC_STATS_RE = re.compile(
    r"^\s*(?P<metric>LCOE|ENS|Fuel|Moto)\s*:\s*"
    r"var=(?P<var>[^ ]+)\s+"
    r"std=(?P<std>[^ ]+)\s+"
    r"range=\[(?P<min>[^.]+)\.\.(?P<max>[^\]]+)\]\s*$"
)

# -------------------------
# Parsing helpers
# -------------------------

def parse_number(s: str) -> float:
    """Парсит число с ',' как десятичным разделителем. Поддерживает 7,95e+13."""
    s = s.strip().replace("\u00a0", "").replace(" ", "")
    s = s.replace(",", ".")
    return float(s)


def read_lines_from_console() -> List[str]:
    """Вставляешь данные, затем ОДНА пустая строка — конец ввода."""
    print("Вставьте данные, затем ОДНА пустая строка — конец ввода.\n")
    lines = []
    while True:
        try:
            line = input()
        except EOFError:
            break
        if line.strip() == "":
            break
        lines.append(line.rstrip("\n"))
    if not lines:
        raise ValueError("Не введено ни одной строки")
    return lines


def split_stats_and_table(lines: List[str]) -> Tuple[List[str], List[str]]:
    """Делит ввод на блок статистик (до 'param') и таблицу (с 'param')."""
    header_idx = None
    for i, line in enumerate(lines):
        if line.strip().lower().startswith("param"):
            header_idx = i
            break
    if header_idx is None:
        raise ValueError("Не найден табличный блок: отсутствует строка 'param ...'")
    return lines[:header_idx], lines[header_idx:]


def parse_metric_stats(lines_stats: List[str]) -> Dict[str, Dict[str, float]]:
    """stats[metric] = {'var','std','min','max'}"""
    stats: Dict[str, Dict[str, float]] = {}
    for line in lines_stats:
        m = METRIC_STATS_RE.match(line.strip())
        if not m:
            continue
        metric = m.group("metric")
        stats[metric] = {
            "var": parse_number(m.group("var")),
            "std": parse_number(m.group("std")),
            "min": parse_number(m.group("min")),
            "max": parse_number(m.group("max")),
        }
    return stats


def parse_format2_table(lines_table: List[str]) -> Dict[str, Dict[str, Tuple[float, float]]]:
    """
    Таблица:
      param S_LCOE ST_LCOE S_ENS ST_ENS ...
    Возвращает:
      data[param][metric] = (S, ST)
    ВАЖНО: если параметр встречается несколько раз — добавляем суффикс #2, #3, ...
    """
    header = lines_table[0].strip().split()
    if not header or header[0].lower() != "param":
        raise ValueError("Ожидается заголовок таблицы, начинающийся с: param ...")

    col_idx = {name: i for i, name in enumerate(header)}

    missing = []
    for metric in EXPECTED_METRICS:
        if f"S_{metric}" not in col_idx or f"ST_{metric}" not in col_idx:
            missing.append(metric)
    if missing:
        raise ValueError("Не найдены пары колонок S_*/ST_* для: " + ", ".join(missing))

    data: Dict[str, Dict[str, Tuple[float, float]]] = {}
    name_counts: Dict[str, int] = {}

    for line in lines_table[1:]:
        parts = line.strip().split()
        if not parts:
            continue

        base_name = parts[0]
        name_counts[base_name] = name_counts.get(base_name, 0) + 1
        param = base_name if name_counts[base_name] == 1 else f"{base_name}#{name_counts[base_name]}"

        if len(parts) < len(header):
            raise ValueError(f"Строка короче заголовка: {line}")

        data[param] = {}
        for metric in EXPECTED_METRICS:
            s = parse_number(parts[col_idx[f"S_{metric}"]])
            st = parse_number(parts[col_idx[f"ST_{metric}"]])
            data[param][metric] = (s, st)

    if not data:
        raise ValueError("Не найдено ни одной строки с параметрами")
    return data


# -------------------------
# Common helpers
# -------------------------

def sort_params_by_total_st(data: Dict[str, Dict[str, Tuple[float, float]]]) -> List[str]:
    """Сортировка по суммарному ST по всем метрикам (убывание)."""
    params = list(data.keys())
    scored = []
    for p in params:
        total = 0.0
        for m in EXPECTED_METRICS:
            total += max(data[p][m][1], 0.0)
        scored.append((p, total))
    scored.sort(key=lambda x: x[1], reverse=True)
    return [p for p, _ in scored]


def figure_size_for_labels(params: List[str], metrics: List[str]) -> Tuple[float, float]:
    """Подбор figsize так, чтобы подписи влезали."""
    max_len = max((len(p) for p in params), default=10)
    width = 9.5
    if max_len >= MAX_PARAM_LABEL_LEN_FOR_WIDE:
        width = 11.5
    width += max(0, (len(metrics) - 4) * 0.8)
    height = max(4.8, len(params) * 0.45 + 2.0)
    return width, height


def fmt_sig(x: float) -> str:
    """Компактный формат числа для подписи в таблице статистики."""
    return f"{x:.4g}"


def zeroize(v: float) -> float:
    """Отрицательные и почти нулевые значения -> 0."""
    return 0.0 if (v <= 0.0 or abs(v) < ZERO_EPS) else v


def fmt_cell(v: float) -> str:
    """
    Подписи в ячейках теплокарты:
    - если v <= 0 или почти 0 -> "0"
    - иначе -> 2 знака после точки
    ВАЖНО: чтобы нигде не появлялось "0.00".
    """
    v = zeroize(v)
    if v == 0.0:
        return "0"
    # дополнительно: если округление до 2 знаков даёт 0.00 -> печатаем "0"
    if round(v, 2) == 0.0:
        return "0"
    return f"{v:.2f}"


# =====================================================
# ВАРИАНТ 1: 4 отдельных графика (bars)
# =====================================================

def plot_single_bars(rows: List[Tuple[str, float, float]], title="", caption: str = ""):
    names = [r[0] for r in rows]
    S = np.maximum(np.array([r[1] for r in rows], dtype=float), 0.0)
    ST = np.maximum(np.array([r[2] for r in rows], dtype=float), 0.0)

    n = len(rows)
    x = np.arange(n)
    width = 0.38

    fig, ax = plt.subplots(figsize=(max(10, n * 0.9), 5), constrained_layout=True)
    ax.bar(x - width / 2, S, width, label="S")
    ax.bar(x + width / 2, ST, width, label="ST")

    ax.set_title(title, fontsize=FONT_BASE)
    ax.set_ylabel("Индекс Соболя (доля дисперсии)", fontsize=FONT_BASE)
    ax.set_xticks(x)
    ax.set_xticklabels([str(i + 1) for i in range(n)], fontsize=FONT_SMALL)
    ax.legend()
    ax.grid(axis="y", alpha=0.3)

    ymax = max(S.max(), ST.max()) if n > 0 else 1.0
    pad = max(ymax * 0.05, 0.05)
    ax.set_ylim(0, ymax + pad)

    if caption:
        ax.text(0.01, 0.99, caption, transform=ax.transAxes,
                ha="left", va="top", fontsize=FONT_SMALL)

    for i, name in enumerate(names):
        y = max(S[i], ST[i])
        ax.text(
            x[i] + 0.06,
            y + pad * 0.15,
            name,
            rotation=35,
            ha="left",
            va="bottom",
            fontsize=FONT_SMALL,
        )

    plt.show()


def make_metric_caption_ru(metric: str, stats: Optional[Dict[str, Dict[str, float]]]) -> str:
    if not stats or metric not in stats:
        return ""
    s = stats[metric]
    return (
        f"σ= {fmt_sig(s['std'])}   "
        f"Диапазон=[{fmt_sig(s['min'])}; {fmt_sig(s['max'])}]"
    )


def plot_variant_1(data: Dict[str, Dict[str, Tuple[float, float]]],
                   stats: Optional[Dict[str, Dict[str, float]]],
                   title_prefix: str,
                   sort_params: bool):
    params = sort_params_by_total_st(data) if sort_params else list(data.keys())
    for metric in EXPECTED_METRICS:
        rows = [(p, data[p][metric][0], data[p][metric][1]) for p in params]
        caption = make_metric_caption_ru(metric, stats)
        plot_single_bars(rows, title=f"{title_prefix}: {metric}", caption=caption)


# =====================================================
# ВАРИАНТ 2: теплокарта ST + таблица статистик + числа в ячейках + белый->оранжевый
# =====================================================

def build_stats_table_cells(stats: Optional[Dict[str, Dict[str, float]]],
                            metrics: List[str]) -> Tuple[List[str], List[List[str]]]:
    row_labels = ["σ", "Диапазон"]
    cell_text = [["" for _ in metrics] for _ in row_labels]

    if not stats:
        return row_labels, cell_text

    for j, m in enumerate(metrics):
        if m not in stats:
            continue
        s = stats[m]
        cell_text[0][j] = fmt_sig(s["std"])
        cell_text[1][j] = f"[{fmt_sig(s['min'])}; {fmt_sig(s['max'])}]"
    return row_labels, cell_text


def plot_variant_2_heatmap(data: Dict[str, Dict[str, Tuple[float, float]]],
                           stats: Optional[Dict[str, Dict[str, float]]],
                           title: str,
                           sort_params: bool):
    params = sort_params_by_total_st(data) if sort_params else list(data.keys())
    metrics = EXPECTED_METRICS

    # --- ST matrix ---
    M = np.zeros((len(params), len(metrics)), dtype=float)
    for i, p in enumerate(params):
        for j, m in enumerate(metrics):
            st = data[p][m][1]
            M[i, j] = zeroize(st)

    figsize = figure_size_for_labels(params, metrics)

    # --- Layout: table (top) + heatmap (bottom) ---
    fig = plt.figure(figsize=figsize, constrained_layout=True)
    gs = fig.add_gridspec(nrows=2, ncols=1, height_ratios=[0.65, 2.55])

    ax_table = fig.add_subplot(gs[0, 0])
    ax = fig.add_subplot(gs[1, 0])

    # --- Stats table ---
    ax_table.axis("off")
    row_labels, cell_text = build_stats_table_cells(stats, metrics)

    tbl = ax_table.table(
        cellText=cell_text,
        rowLabels=row_labels,
        colLabels=metrics,
        cellLoc="center",
        rowLoc="center",
        loc="center",
    )
    tbl.auto_set_font_size(False)
    tbl.set_fontsize(FONT_SMALL)
    tbl.scale(1.0, 1.35)

    ax_table.set_title(title, fontsize=FONT_BASE, pad=6)

    # --- Heatmap: white->orange, 0 must be white ---
    vmax = float(M.max()) if M.size else 1.0
    vmax = max(vmax, 1e-12)

    im = ax.imshow(
        M,
        aspect="auto",
        cmap=WHITE_ORANGE,
        vmin=0.0,
        vmax=vmax
    )

    ax.set_xlabel("Метрика", fontsize=FONT_BASE)
    ax.set_ylabel("Параметр", fontsize=FONT_BASE)

    ax.set_xticks(np.arange(len(metrics)))
    ax.set_xticklabels(metrics, fontsize=FONT_BASE)
    ax.set_yticks(np.arange(len(params)))
    ax.set_yticklabels(params, fontsize=FONT_BASE)

    ax.tick_params(axis="x", pad=6)
    ax.tick_params(axis="y", pad=6)

    # --- Numeric labels inside cells ---
    threshold = 0.60 * vmax  # для контраста

    for i in range(M.shape[0]):
        for j in range(M.shape[1]):
            val = M[i, j]
            text_color = "white" if val >= threshold else "black"
            ax.text(
                j, i,
                fmt_cell(val),
                ha="center",
                va="center",
                fontsize=FONT_SMALL,
                color=text_color
            )

    # --- Colorbar ---
    cbar = fig.colorbar(im, ax=ax, shrink=0.95)
    cbar.set_label("ST (доля дисперсии)", fontsize=FONT_BASE)

    plt.show()


# =====================================================
# ENTRYPOINT
# =====================================================

def main():
    lines = read_lines_from_console()
    lines_stats, lines_table = split_stats_and_table(lines)

    stats = parse_metric_stats(lines_stats)
    data = parse_format2_table(lines_table)

    title = TITLE_PREFIX.strip()

    if PLOT_VARIANT == 1:
        plot_variant_1(data, stats, title_prefix=title, sort_params=SORT_PARAMS_BY_TOTAL_ST)
    elif PLOT_VARIANT == 2:
        plot_variant_2_heatmap(data, stats, title=f"{title}: теплокарта ST", sort_params=SORT_PARAMS_BY_TOTAL_ST)
    else:
        raise ValueError(f"Неизвестный PLOT_VARIANT={PLOT_VARIANT}. Используй 1 или 2.")


if __name__ == "__main__":
    main()
