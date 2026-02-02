import re
from typing import List, Tuple, Dict, Optional

import numpy as np
import matplotlib.pyplot as plt
from matplotlib.colors import LinearSegmentedColormap

# =====================================================
# НАСТРОЙКИ (МЕНЯТЬ ЗДЕСЬ)
# =====================================================
# Варианты:
# 1) "triptych" — 1..3 теплокарты ST рядом (по числу введённых схем) + таблицы статистик сверху
# 2) "bars"     — для каждой метрики отдельный график: сгруппированные горизонтальные бары ST по введённым схемам
COMPARE_MODE = "triptych"   # "triptych" | "bars"

SCHEMES_ORDER = ["SN", "SS", "D", "H"]

# Подписи схем:
# - для двухстрочных: две строки рисуются на фиксированных высотах
# - для однострочных: строка рисуется по центру между двумя высотами (как ты хотел)
SCHEME_TITLES = {
    "SN": ["Одиночная", "несекционированная"],
    "SS": ["Одиночная", "секционированная"],
    "D":  ["Двойная"],
    "H": ["Хуевая"],
}

SORT_PARAMS_BY = "base_total_st"  # "base_total_st" | "max_total_st" | "none"

# ВАЖНО: верхнюю подпись (suptitle) убираем полностью — TITLE_PREFIX не используется в triptych
TITLE_PREFIX = "Sobol: сравнение схем"

FONT_BASE = 11
FONT_SMALL = 10
MAX_PARAM_LABEL_LEN_FOR_WIDE = 14
ZERO_EPS = 5e-4  # все <=0 или |x|<eps считаем 0

EXPECTED_METRICS = ["LCOE", "ENS", "Fuel", "Moto"]

# белый -> оранжевый для ST
WHITE_ORANGE = LinearSegmentedColormap.from_list(
    "white_orange",
    [
        (1.0, 1.0, 1.0),
        (1.0, 0.85, 0.6),
        (1.0, 0.55, 0.0),
    ],
)

# =====================================================

# Пример:
# LCOE: var=12481,5 std=111,720 range=[29,0663..818,510]
METRIC_STATS_RE = re.compile(
    r"^\s*(?P<metric>LCOE|ENS|Fuel|Moto)\s*:\s*"
    r"var=(?P<var>[^ ]+)\s+"
    r"std=(?P<std>[^ ]+)\s+"
    r"range=\[(?P<min>[^.]+)\.\.(?P<max>[^\]]+)\]\s*$"
)

SCHEME_HEADER_RE = re.compile(r"^\s*(SN|SS|D|H)\s*$", re.IGNORECASE)


# -------------------------
# Parsing helpers
# -------------------------
def parse_number(s: str) -> float:
    s = s.strip().replace("\u00a0", "").replace(" ", "")
    s = s.replace(",", ".")
    return float(s)


def read_lines_from_console() -> List[str]:
    """
    Вставляешь данные для 1..3 схем подряд (SN/SS/D),
    затем ОДНА пустая строка — конец ввода.
    ВАЖНО: строка с именем схемы должна быть отдельной строкой: SN или SS или D
    """
    print("Вставьте данные (SN/SS/D), затем ОДНА пустая строка — конец ввода.\n")
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
    header_idx = None
    for i, line in enumerate(lines):
        if line.strip().lower().startswith("param"):
            header_idx = i
            break
    if header_idx is None:
        raise ValueError("Не найден табличный блок: отсутствует строка 'param ...'")
    return lines[:header_idx], lines[header_idx:]


def parse_metric_stats(lines_stats: List[str]) -> Dict[str, Dict[str, float]]:
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
    Возвращает:
      data[param][metric] = (S, ST)
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
    for line in lines_table[1:]:
        parts = line.strip().split()
        if not parts:
            continue

        param = parts[0]
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


def parse_multi_scheme_input(
    lines: List[str],
) -> Dict[str, Tuple[Dict[str, Dict[str, float]], Dict[str, Dict[str, Tuple[float, float]]]]]:
    """
    Вход: блоки вида (можно 1-2-3 блока):
      SN
      ...
      SS
      ...
      D
      ...
    Возвращает: schemes[scheme] = (stats, data)
    """
    blocks: Dict[str, List[str]] = {}
    current_scheme: Optional[str] = None

    for line in lines:
        m = SCHEME_HEADER_RE.match(line)
        if m:
            current_scheme = m.group(1).upper()
            blocks[current_scheme] = []
            continue
        if current_scheme is None:
            continue
        blocks[current_scheme].append(line)

    if not blocks:
        raise ValueError("Не найдено ни одной схемы. Добавьте строку 'SN' или 'SS' или 'D' перед блоком данных.")

    parsed: Dict[str, Tuple[Dict[str, Dict[str, float]], Dict[str, Dict[str, Tuple[float, float]]]]] = {}
    for scheme in SCHEMES_ORDER:
        if scheme not in blocks:
            continue
        b = blocks[scheme]
        ls, lt = split_stats_and_table(b)
        stats = parse_metric_stats(ls)
        data = parse_format2_table(lt)
        parsed[scheme] = (stats, data)

    if not parsed:
        raise ValueError("Не удалось распарсить ни одного блока схемы.")
    return parsed


# -------------------------
# Formatting helpers for stats table (σ/min/max)
# -------------------------
def fmt_1dp_trim(x: float) -> str:
    """
    1 знак после запятой, но если ,0 — не показываем.
    Десятичный разделитель: запятая.
    """
    s = f"{x:.1f}"
    if s.endswith(".0"):
        s = s[:-2]
    return s.replace(".", ",")


def scale_stat(metric: str, x: float) -> float:
    """
    Масштабирование статистики:
      Fuel -> /1e6
      Moto -> /1e3
      остальное без изменений
    """
    if metric == "Fuel":
        return x / 1e6
    if metric == "Moto":
        return x / 1e3
    return x


def fmt_stat(metric: str, x: float) -> str:
    """
    Формат для σ/min/max:
      - 1 знак после запятой (и обрезать ,0)
      - Fuel делим на 1e6, Moto на 1e3
      - без экспоненты
    """
    return fmt_1dp_trim(scale_stat(metric, x))


# -------------------------
# Common helpers
# -------------------------
def zeroize(v: float) -> float:
    return 0.0 if (v <= 0.0 or abs(v) < ZERO_EPS) else v


def fmt_cell(v: float) -> str:
    """
    Подписи в ячейках теплокарты:
    - если v <= 0 или почти 0 -> "0"
    - иначе -> 2 знака после точки
    """
    v = zeroize(v)
    if v == 0.0:
        return "0"
    if round(v, 2) == 0.0:
        return "0"
    return f"{v:.2f}"


def build_stats_table_cells(
    stats: Optional[Dict[str, Dict[str, float]]],
    metrics: List[str],
) -> Tuple[List[str], List[List[str]]]:
    row_labels = ["σ", "min", "max"]
    cell_text = [["" for _ in metrics] for _ in row_labels]

    if not stats:
        return row_labels, cell_text

    for j, m in enumerate(metrics):
        if m not in stats:
            continue
        s = stats[m]
        cell_text[0][j] = fmt_stat(m, s["std"])
        cell_text[1][j] = fmt_stat(m, s["min"])
        cell_text[2][j] = fmt_stat(m, s["max"])

    return row_labels, cell_text


def figure_size_for_labels(params: List[str], ncols: int) -> Tuple[float, float]:
    # ОСТАВЛЕНО КАК В ТВОЁМ КОДЕ (масштабирование/ширина и т.д.)
    max_len = max((len(p) for p in params), default=10)
    width = 10.5
    if max_len >= MAX_PARAM_LABEL_LEN_FOR_WIDE:
        width = 12.0
    width += (ncols - 3) * 0.9
    height = max(5.2, len(params) * 0.45 + 2.4)
    return width, height


def union_params(parsed: Dict[str, Tuple[Dict, Dict]]) -> List[str]:
    # объединение параметров по введённым схемам
    s = set()
    for scheme in parsed.keys():
        _, data = parsed[scheme]
        s.update(data.keys())
    return sorted(s)


def get_st_matrix(
    data: Dict[str, Dict[str, Tuple[float, float]]],
    params: List[str],
    metrics: List[str],
) -> np.ndarray:
    M = np.zeros((len(params), len(metrics)), dtype=float)
    for i, p in enumerate(params):
        for j, m in enumerate(metrics):
            st = 0.0
            if p in data and m in data[p]:
                st = data[p][m][1]
            M[i, j] = zeroize(st)
    return M


def sort_params(
    parsed: Dict[str, Tuple[Dict, Dict]],
    params: List[str],
    mode: str,
    base_scheme: str,
) -> List[str]:
    if mode == "none":
        return params

    metrics = EXPECTED_METRICS

    # если базовая схема не введена — берём первую введённую в порядке SCHEMES_ORDER
    if base_scheme not in parsed:
        for s in SCHEMES_ORDER:
            if s in parsed:
                base_scheme = s
                break

    if mode == "base_total_st":
        _, base_data = parsed[base_scheme]
        scores = []
        for p in params:
            total = 0.0
            if p in base_data:
                for m in metrics:
                    total += max(base_data[p][m][1], 0.0)
            scores.append((p, total))
        scores.sort(key=lambda x: x[1], reverse=True)
        return [p for p, _ in scores]

    if mode == "max_total_st":
        scores = []
        for p in params:
            best = 0.0
            for scheme in parsed.keys():
                _, d = parsed[scheme]
                total = 0.0
                if p in d:
                    for m in metrics:
                        total += max(d[p][m][1], 0.0)
                best = max(best, total)
            scores.append((p, best))
        scores.sort(key=lambda x: x[1], reverse=True)
        return [p for p, _ in scores]

    raise ValueError(f"Unknown SORT_PARAMS_BY={mode}")


def draw_scheme_title(ax, lines: List[str], fontsize: int):
    """
    Заголовки над таблицами:
      - если 2 строки: фиксированные уровни y1 и y2
      - если 1 строка: по центру между y1 и y2
    Это даёт "Двойная" ровно между строками двухстрочных заголовков.
    """
    ax.set_title("")  # отключаем стандартный title

    y1 = 1.10
    y2 = 0.98
    y_mid = (y1 + y2) / 2.0

    if len(lines) >= 2:
        ax.text(0.5, y1, lines[0], transform=ax.transAxes,
                ha="center", va="bottom", fontsize=fontsize)
        ax.text(0.5, y2, lines[1], transform=ax.transAxes,
                ha="center", va="bottom", fontsize=fontsize)
    elif len(lines) == 1:
        ax.text(0.5, y_mid, lines[0], transform=ax.transAxes,
                ha="center", va="bottom", fontsize=fontsize)


# -------------------------
# Plot modes
# -------------------------
def plot_triptych(parsed: Dict[str, Tuple[Dict, Dict]]):
    metrics = EXPECTED_METRICS

    # 1..3 схемы в порядке SCHEMES_ORDER
    schemes = [s for s in SCHEMES_ORDER if s in parsed]
    ncols = len(schemes)

    params = union_params(parsed)
    # сортировка: базовая = SN если введена, иначе первая введенная
    params = sort_params(parsed, params, SORT_PARAMS_BY, base_scheme="SN")

    # общий vmax по введённым схемам
    vmax = 0.0
    Ms: Dict[str, np.ndarray] = {}
    for scheme in schemes:
        _, data = parsed[scheme]
        M = get_st_matrix(data, params, metrics)
        Ms[scheme] = M
        vmax = max(vmax, float(M.max()) if M.size else 0.0)
    vmax = max(vmax, 1e-12)

    figsize = figure_size_for_labels(params, ncols=ncols)
    fig = plt.figure(figsize=figsize, constrained_layout=True)

    gs = fig.add_gridspec(nrows=2, ncols=ncols, height_ratios=[0.72, 2.65])

    # TOP: таблица над каждой схемой; подписи σ/min/max только у первой
    for col, scheme in enumerate(schemes):
        stats, _ = parsed[scheme]
        ax_t = fig.add_subplot(gs[0, col])
        ax_t.axis("off")

        row_labels, cell_text = build_stats_table_cells(stats, metrics)
        use_row_labels = row_labels if col == 0 else None

        tbl = ax_t.table(
            cellText=cell_text,
            rowLabels=use_row_labels,
            colLabels=metrics,
            cellLoc="center",
            rowLoc="center",
            loc="center",
        )
        tbl.auto_set_font_size(False)
        tbl.set_fontsize(FONT_SMALL)
        tbl.scale(1.0, 1.25)

        # Рисуем заголовки как 1/2 строки на фиксированных уровнях
        title_lines = SCHEME_TITLES.get(scheme, [scheme])
        draw_scheme_title(ax_t, title_lines, fontsize=FONT_BASE)

    # BOTTOM: теплокарты
    heatmap_axes = []
    for col, scheme in enumerate(schemes):
        ax = fig.add_subplot(gs[1, col])
        heatmap_axes.append(ax)

        M = Ms[scheme]
        im = ax.imshow(M, aspect="auto", cmap=WHITE_ORANGE, vmin=0.0, vmax=vmax)

        ax.set_xticks(np.arange(len(metrics)))
        ax.set_xticklabels(metrics, fontsize=FONT_BASE)

        if col == 0:
            ax.set_yticks(np.arange(len(params)))
            ax.set_yticklabels(params, fontsize=FONT_BASE)
        else:
            ax.set_yticks(np.arange(len(params)))
            ax.set_yticklabels([])

        # подписи "Параметр" и "Метрика" убраны
        ax.tick_params(axis="x", pad=6)
        ax.tick_params(axis="y", pad=6)

        threshold = 0.60 * vmax
        for i in range(M.shape[0]):
            for j in range(M.shape[1]):
                val = M[i, j]
                color = "white" if val >= threshold else "black"
                ax.text(
                    j, i, fmt_cell(val),
                    ha="center", va="center",
                    fontsize=FONT_SMALL, color=color
                )

    # colorbar только по высоте теплокарт
    cbar = fig.colorbar(
        plt.cm.ScalarMappable(cmap=WHITE_ORANGE, norm=plt.Normalize(0.0, vmax)),
        ax=heatmap_axes,
        shrink=1.0,
        location="right",
        pad=0.04,
        fraction=0.045
    )
    cbar.set_label("ST (доля дисперсии)", fontsize=FONT_BASE, labelpad=18)

    plt.show()


def plot_bars(parsed: Dict[str, Tuple[Dict, Dict]]):
    metrics = EXPECTED_METRICS
    schemes = [s for s in SCHEMES_ORDER if s in parsed]

    params = union_params(parsed)
    params = sort_params(parsed, params, SORT_PARAMS_BY, base_scheme="SN")

    # сгруппированные горизонтальные бары: по числу введенных схем
    for metric in metrics:
        fig_h = max(5.0, len(params) * 0.45 + 1.8)
        fig_w = 11.5
        fig, ax = plt.subplots(figsize=(fig_w, fig_h), constrained_layout=True)

        y = np.arange(len(params))

        # ширина группы фиксирована как раньше для 3 схем; для 1-2 схем подстраиваем offsets
        bar_h = 0.22 if len(schemes) >= 3 else (0.26 if len(schemes) == 2 else 0.32)
        offsets = {}
        if len(schemes) == 1:
            offsets[schemes[0]] = 0.0
        elif len(schemes) == 2:
            offsets[schemes[0]] = -bar_h / 2
            offsets[schemes[1]] = +bar_h / 2
        else:
            offsets = {"SN": -bar_h, "SS": 0.0, "D": +bar_h}

        vmax = 0.0
        for scheme in schemes:
            _, d = parsed[scheme]
            vals = []
            for p in params:
                st = 0.0
                if p in d:
                    st = d[p][metric][1]
                vals.append(zeroize(st))
            vals = np.array(vals, dtype=float)
            vmax = max(vmax, float(vals.max()) if vals.size else 0.0)

            label = " ".join(SCHEME_TITLES.get(scheme, [scheme]))
            ax.barh(y + offsets.get(scheme, 0.0), vals, height=bar_h, label=label)

        ax.set_yticks(y)
        ax.set_yticklabels(params, fontsize=FONT_BASE)
        ax.set_xlabel("ST (доля дисперсии)", fontsize=FONT_BASE)
        ax.set_title(f"{TITLE_PREFIX}: {metric} — ST по схемам", fontsize=FONT_BASE)
        ax.grid(axis="x", alpha=0.25)
        ax.legend()

        ax.set_xlim(0, max(vmax * 1.10, 0.05))
        plt.show()


# =====================================================
# ENTRYPOINT
# =====================================================
def main():
    lines = read_lines_from_console()
    parsed = parse_multi_scheme_input(lines)

    if COMPARE_MODE == "triptych":
        plot_triptych(parsed)
    elif COMPARE_MODE == "bars":
        plot_bars(parsed)
    else:
        raise ValueError(f"Unknown COMPARE_MODE={COMPARE_MODE}")


if __name__ == "__main__":
    main()
