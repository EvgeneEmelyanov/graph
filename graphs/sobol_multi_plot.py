import re
from dataclasses import dataclass
from typing import Dict, List, Optional, Set, Tuple

import numpy as np
import matplotlib.pyplot as plt
from matplotlib.colors import LinearSegmentedColormap

# =====================================================
# НАСТРОЙКИ (МЕНЯТЬ ЗДЕСЬ)
# =====================================================
# Варианты:
# 1) "triptych" — 1..N теплокарт ST рядом (по числу введённых схем) + таблицы статистик сверху
# 2) "bars"     — для каждой метрики отдельный график: сгруппированные горизонтальные бары ST по введённым схемам
COMPARE_MODE = "triptych"   # "triptych" | "bars"

SCHEMES_ORDER = ["SN", "SS", "D", "H", "A", "B", "C", "4", "5"]

SCHEME_TITLES = {
    "SN": ["Одиночная", "несекционированная"],
    "SS": ["Одиночная", "секционированная"],
    "D":  ["Двойная"],
    "H":  ["хехе"],
    "A":  ["Базовые диапазоны"],
    "B":  ["ДТ +-10%"],
    "C":  ["МЧТ+-10%"],
    "4":  ["Вариант 4"],
    "5":  ["Вариант 5"],
}

# Строка с именем схемы должна быть отдельной строкой
SCHEME_HEADER_RE = re.compile(r"^\s*(SN|SS|D|H|A|B|C|4|5)\s*$", re.IGNORECASE)

# Русские подписи параметров (как было)
PARAM_LABELS_RU = {
    "FIRST_CAT": "Доля потребителей I категории",
    "SECOND_CAT": "Доля потребителей II категории",
    "WT_COUNT": "Количество ВЭУ",
    "WT_POWER": "Мощность одной ВЭУ (кВт)",
    "WT_COUNT_TOTAL": "Количество ВЭУ",
    "WT_POWER_KW": "Мощность одной ВЭУ (кВт)",
    "WT_FAILURE_RATE": "Интенсивность отказов ВЭУ",
    "WT_REPAIR_TIME": "Время восстановления ВЭУ",
    "DG_COUNT": "Количество ДГУ",
    "DG_POWER": "Мощность одного ДГУ",
    "DG_COUNT_TOTAL": "Количество ДГУ",
    "DG_POWER_KW": "Мощность одного ДГУ",
    "DG_FAILURE_RATE": "Интенсивность отказов ДГУ",
    "DG_REPAIR_TIME": "Время восстановления ДГУ",
    "BT_CAPACITY_PER_BUS": "Емкость СНЭ на одной шине",
    "BT_CAPACITY_KWH_PER_BUS": "Емкость СНЭ на одной шине",
    "BT_MAX_CHARGE_CURRENT": "Максимальный ток заряда СНЭ",
    "BT_MAX_DISCHARGE_CURRENT": "Максимальный ток разряда СНЭ",
    "BT_NON_RESERVE_DISCHARGE_LEVEL": "Допустимая глубина разряда СНЭ",
    "BT_FAILURE_RATE": "Интенсивность отказов СНЭ",
    "BT_REPAIR_TIME": "Время восстановления СНЭ",
    "BUS_FAILURE_RATE": "Интенсивность отказов шины",
    "BUS_REPAIR_TIME": "Время восстановления шины",
    "BRK_FAILURE_RATE": "Интенсивность отказов СВ/МШВ",
    "BRK_REPAIR_TIME": "Время восстановления СВ/МШВ",
    "SWITCHGEAR_ROOM_FAILURE_RATE": "Интенсивность отказов Р",
    "SWITCHGEAR_ROOM_REPAIR_TIME": "Время восстановления РУ",
    "BUS_CCF_BETA_SECTIONAL": "Коэф. ООП шин (секционир.) β",
    "BUS_CCF_BETA_DOUBLE": "Коэф. ООП шин (двойная) β",
    "DISCOUNT_RATE": "Ставка дисконтирования",
    "COST_RU_RUB": "Стоимость РУ",
    "COST_DG_RUB_PER_KW": "Стоимость ДГУ",
    "COST_DG_RUB_PER_KW_PER_KMH": "Эксплуатационные затраты ДГУ",
    "COST_FUEL_RUB_PER_KT": "Стоимость топлива",
    "COST_WT_RUB_PER_KW": "Стоимость ВЭУ",
    "COST_WT_RUB_PER_KW": "Эксплуатационные затраты ВЭУ",
    "COST_BT_RUB_PER_KWH": "Стоимость СНЭ",
    "COST_BT_RUB_PER_KWH": "Эксплуатационные затраты СНЭ",
    "DAMAGE_RUB_PER_KWH_CAT1": "Ущерб недоотпуска Cat1 (руб/кВт·ч)",
    "DAMAGE_RUB_PER_KWH_CAT2": "Ущерб недоотпуска Cat2 (руб/кВт·ч)",
    "DAMAGE_RUB_PER_KWH_CAT3": "Ущерб недоотпуска",
}

USE_RU_PARAM_LABELS = True

SORT_PARAMS_BY = "avg_lcoe_st"  # "avg_lcoe_st" | "avg_lcoe_s" | "none"

TITLE_PREFIX = "Sobol: сравнение схем"

FONT_BASE = 11
FONT_SMALL = 10
MAX_PARAM_LABEL_LEN_FOR_WIDE = 14
ZERO_EPS = 5e-4  # все <=0 или |x|<eps считаем 0

# ОБНОВЛЕНО: добавили LOLE (S_LOLE/ST_LOLE)
PREFERRED_METRIC_ORDER = ["LCOE", "LOLE", "ENS", "Fuel", "Moto"]

WHITE_ORANGE = LinearSegmentedColormap.from_list(
    "white_orange",
    [(1.0, 1.0, 1.0), (1.0, 0.85, 0.6), (1.0, 0.55, 0.0)],
)
# =====================================================

METRIC_STATS_RE = re.compile(
    r"^\s*(?P<metric>[A-Za-z0-9_]+)\s*:\s*"
    r"var=(?P<var>[^ ]+)\s+"
    r"std=(?P<std>[^ ]+)\s+"
    r"range=\[(?P<min>[^.]+)\.\.(?P<max>[^\]]+)\]\s*$"
)

SINGLE_METRIC_PARAM_RE = re.compile(
    r"^\s*(?P<param>\S+)\s+S\s*=\s*(?P<s>[-+0-9eE.,]+)\s+ST\s*=\s*(?P<st>[-+0-9eE.,]+)\s*$"
)


# -------------------------
# Data structures
# -------------------------
@dataclass(frozen=True)
class SchemeData:
    stats: Dict[str, Dict[str, float]]  # metric -> {var,std,min,max}
    data: Dict[str, Dict[str, Tuple[float, float]]]  # param -> metric -> (S, ST)


# -------------------------
# Parsing helpers
# -------------------------
def parse_number(s: str) -> float:
    s = s.strip().replace("\u00a0", "").replace(" ", "").replace(",", ".")
    return float(s)


def read_lines_from_console() -> List[str]:
    """
    Вставляешь данные для 1..N схем подряд (SN/SS/D/H/A/B/C/4/5),
    затем ОДНА пустая строка — конец ввода.
    ВАЖНО: строка с именем схемы должна быть отдельной строкой.
    """
    print("Вставьте данные (SN/SS/D/H/A/B/C/4/5...), затем ОДНА пустая строка — конец ввода.\n")
    lines: List[str] = []
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


def split_stats_and_body(lines: List[str]) -> Tuple[List[str], List[str]]:
    """
    Делит блок схемы на:
      - lines_stats: строки со статистиками "Metric: var=... std=... range=[..]"
      - lines_body:  дальше либо табличный формат (param S_... ST_...), либо строчный формат (PARAM S=.. ST=..)
    """
    idx: Optional[int] = None
    for i, line in enumerate(lines):
        if METRIC_STATS_RE.match(line.strip()):
            continue
        if line.strip() == "":
            continue
        idx = i
        break

    if idx is None:
        return lines, []
    return lines[:idx], lines[idx:]


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


def parse_format2_table(lines_body: List[str]) -> Dict[str, Dict[str, Tuple[float, float]]]:
    """
    Формат 2:
      param  S_LCOE ST_LCOE S_LOLE ST_LOLE S_ENS ST_ENS ...
    Возвращает:
      data[param][metric] = (S, ST)
    """
    header = lines_body[0].strip().split()
    if not header or header[0].lower() != "param":
        raise ValueError("Ожидается заголовок таблицы, начинающийся с: param ...")

    col_idx = {name: i for i, name in enumerate(header)}

    # Ищем пары S_x / ST_x динамически (теперь поддерживает LOLE и любые будущие метрики)
    metrics: List[str] = []
    for name in header[1:]:
        if not name.startswith("S_"):
            continue
        metric = name[2:]
        if f"ST_{metric}" in col_idx:
            metrics.append(metric)

    if not metrics:
        raise ValueError("Не найдены пары колонок S_*/ST_* в табличном формате")

    data: Dict[str, Dict[str, Tuple[float, float]]] = {}
    for line in lines_body[1:]:
        parts = line.strip().split()
        if not parts:
            continue
        if len(parts) < len(header):
            raise ValueError(f"Строка короче заголовка: {line}")

        param = parts[0]
        data[param] = {}
        for metric in metrics:
            s = parse_number(parts[col_idx[f"S_{metric}"]])
            st = parse_number(parts[col_idx[f"ST_{metric}"]])
            data[param][metric] = (s, st)

    if not data:
        raise ValueError("Не найдено ни одной строки с параметрами (табличный формат)")
    return data


def choose_single_metric_name(stats: Dict[str, Dict[str, float]]) -> str:
    # приоритеты обновлены: если есть LCOE — он, иначе LOLE, иначе первая по алфавиту
    if "LCOE" in stats:
        return "LCOE"
    if "LOLE" in stats:
        return "LOLE"
    if len(stats) == 1:
        return next(iter(stats.keys()))
    if len(stats) == 0:
        return "LCOE"
    return sorted(stats.keys())[0]


def parse_format1_single_metric(
    lines_body: List[str], metric_name: str
) -> Dict[str, Dict[str, Tuple[float, float]]]:
    """
    Формат 1:
      PARAM  S=0.123  ST=0.456
    """
    data: Dict[str, Dict[str, Tuple[float, float]]] = {}
    for line in lines_body:
        if line.strip() == "":
            continue
        m = SINGLE_METRIC_PARAM_RE.match(line)
        if not m:
            continue
        param = m.group("param")
        s = parse_number(m.group("s"))
        st = parse_number(m.group("st"))
        data[param] = {metric_name: (s, st)}

    if not data:
        raise ValueError("Не найдено строк PARAM S=.. ST=.. (формат одной метрики)")
    return data


def parse_body_auto(
    lines_body: List[str], stats: Dict[str, Dict[str, float]]
) -> Dict[str, Dict[str, Tuple[float, float]]]:
    first: Optional[str] = None
    for ln in lines_body:
        if ln.strip():
            first = ln.strip()
            break
    if first is None:
        raise ValueError("Пустой блок данных параметров")

    if first.lower().startswith("param"):
        return parse_format2_table(lines_body)

    metric_name = choose_single_metric_name(stats)
    return parse_format1_single_metric(lines_body, metric_name)


def parse_multi_scheme_input(lines: List[str]) -> Dict[str, SchemeData]:
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
        raise ValueError("Не найдено ни одной схемы. Добавьте строку 'SN'/'SS'/'D'/'H'/'A'/'B'/'C'/'4'/'5' перед блоком данных.")

    parsed: Dict[str, SchemeData] = {}
    for scheme in SCHEMES_ORDER:
        if scheme not in blocks:
            continue

        ls, lb = split_stats_and_body(blocks[scheme])
        stats = parse_metric_stats(ls)
        data = parse_body_auto(lb, stats)
        parsed[scheme] = SchemeData(stats=stats, data=data)

    if not parsed:
        raise ValueError("Не удалось распарсить ни одного блока схемы.")
    return parsed


# -------------------------
# Formatting helpers for stats table (σ/min/max)
# -------------------------
def fmt_1dp_trim(x: float) -> str:
    s = f"{x:.1f}"
    if s.endswith(".0"):
        s = s[:-2]
    return s.replace(".", ",")


def scale_stat(metric: str, x: float) -> float:
    # оставляем прежнюю логику единиц
    if metric == "Fuel":
        return x / 1e6
    if metric == "Moto":
        return x / 1e3
    return x


def fmt_stat(metric: str, x: float) -> str:
    return fmt_1dp_trim(scale_stat(metric, x))


# -------------------------
# Common helpers
# -------------------------
def zeroize(v: float) -> float:
    return 0.0 if (v <= 0.0 or abs(v) < ZERO_EPS) else v


def fmt_cell(v: float) -> str:
    v = zeroize(v)
    if v == 0.0:
        return "0"
    if round(v, 2) == 0.0:
        return "0"
    return f"{v:.2f}"


def param_label(p: str) -> str:
    if not USE_RU_PARAM_LABELS:
        return p
    return PARAM_LABELS_RU.get(p, p)


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


def figure_size_for_labels(y_labels: List[str], ncols: int) -> Tuple[float, float]:
    max_len = max((len(x) for x in y_labels), default=10)
    width = 10.5
    if max_len >= MAX_PARAM_LABEL_LEN_FOR_WIDE:
        width = 12.0
    width += (ncols - 3) * 0.9
    height = max(5.2, len(y_labels) * 0.45 + 2.4)
    return width, height


def union_params(parsed: Dict[str, SchemeData]) -> List[str]:
    s: Set[str] = set()
    for scheme in parsed:
        s.update(parsed[scheme].data.keys())
    return sorted(s)


def union_metrics(parsed: Dict[str, SchemeData]) -> List[str]:
    ms: Set[str] = set()
    for scheme in parsed:
        data = parsed[scheme].data
        for p in data:
            ms.update(data[p].keys())
    ordered = [m for m in PREFERRED_METRIC_ORDER if m in ms]
    rest = sorted([m for m in ms if m not in ordered])
    return ordered + rest


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
    parsed: Dict[str, SchemeData],
    params: List[str],
    mode: str,
) -> List[str]:
    if mode == "none":
        return params
    if mode not in ("avg_lcoe_st", "avg_lcoe_s"):
        raise ValueError(f"Unknown SORT_PARAMS_BY={mode}")

    schemes = list(parsed.keys())
    if not schemes:
        return params

    metrics_all = union_metrics(parsed)
    if not metrics_all:
        return params

    target_metric = "LCOE" if "LCOE" in metrics_all else metrics_all[0]
    use_total = (mode == "avg_lcoe_st")  # True -> ST, False -> S

    scored: List[Tuple[str, float]] = []
    for p in params:
        acc = 0.0
        cnt = 0
        for scheme in schemes:
            d = parsed[scheme].data
            if p not in d or target_metric not in d[p]:
                continue
            s_val, st_val = d[p][target_metric]
            v = st_val if use_total else s_val
            acc += zeroize(v)
            cnt += 1
        scored.append((p, (acc / cnt) if cnt else 0.0))

    scored.sort(key=lambda x: x[1], reverse=True)
    return [p for p, _ in scored]


def draw_scheme_title(ax, lines: List[str], fontsize: int):
    ax.set_title("")
    y1 = 1.10
    y2 = 0.98
    y_mid = (y1 + y2) / 2.0

    if len(lines) >= 2:
        ax.text(0.5, y1, lines[0], transform=ax.transAxes, ha="center", va="bottom", fontsize=fontsize)
        ax.text(0.5, y2, lines[1], transform=ax.transAxes, ha="center", va="bottom", fontsize=fontsize)
    elif len(lines) == 1:
        ax.text(0.5, y_mid, lines[0], transform=ax.transAxes, ha="center", va="bottom", fontsize=fontsize)


# -------------------------
# Plot modes
# -------------------------
def plot_triptych(parsed: Dict[str, SchemeData]):
    metrics = union_metrics(parsed)
    if not metrics:
        raise ValueError("Не найдено ни одной метрики в данных")

    schemes = [s for s in SCHEMES_ORDER if s in parsed]
    ncols = len(schemes)

    params = union_params(parsed)
    params = sort_params(parsed, params, SORT_PARAMS_BY)
    y_labels = [param_label(p) for p in params]

    vmax = 0.0
    Ms: Dict[str, np.ndarray] = {}
    for scheme in schemes:
        M = get_st_matrix(parsed[scheme].data, params, metrics)
        Ms[scheme] = M
        vmax = max(vmax, float(M.max()) if M.size else 0.0)
    vmax = max(vmax, 1e-12)

    figsize = figure_size_for_labels(y_labels, ncols=ncols)
    fig = plt.figure(figsize=figsize, constrained_layout=True)

    gs = fig.add_gridspec(nrows=2, ncols=ncols, height_ratios=[0.72, 2.65])

    # TOP: stats tables
    for col, scheme in enumerate(schemes):
        ax_t = fig.add_subplot(gs[0, col])
        ax_t.axis("off")

        row_labels, cell_text = build_stats_table_cells(parsed[scheme].stats, metrics)
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

        draw_scheme_title(ax_t, SCHEME_TITLES.get(scheme, [scheme]), fontsize=FONT_BASE)

    # BOTTOM: heatmaps
    heatmap_axes = []
    for col, scheme in enumerate(schemes):
        ax = fig.add_subplot(gs[1, col])
        heatmap_axes.append(ax)

        M = Ms[scheme]
        ax.imshow(M, aspect="auto", cmap=WHITE_ORANGE, vmin=0.0, vmax=vmax)

        ax.set_xticks(np.arange(len(metrics)))
        ax.set_xticklabels(metrics, fontsize=FONT_BASE)

        ax.set_yticks(np.arange(len(params)))
        if col == 0:
            ax.set_yticklabels(y_labels, fontsize=FONT_BASE)
        else:
            ax.set_yticklabels([])

        ax.tick_params(axis="x", pad=6)
        ax.tick_params(axis="y", pad=6)

        threshold = 0.60 * vmax
        for i in range(M.shape[0]):
            for j in range(M.shape[1]):
                val = M[i, j]
                color = "white" if val >= threshold else "black"
                ax.text(j, i, fmt_cell(val), ha="center", va="center", fontsize=FONT_SMALL, color=color)

    cbar = fig.colorbar(
        plt.cm.ScalarMappable(cmap=WHITE_ORANGE, norm=plt.Normalize(0.0, vmax)),
        ax=heatmap_axes,
        shrink=1.0,
        location="right",
        pad=0.04,
        fraction=0.045,
    )
    cbar.set_label("ST (доля дисперсии)", fontsize=FONT_BASE, labelpad=18)

    plt.show()


def plot_bars(parsed: Dict[str, SchemeData]):
    metrics = union_metrics(parsed)
    if not metrics:
        raise ValueError("Не найдено ни одной метрики в данных")

    schemes = [s for s in SCHEMES_ORDER if s in parsed]

    params = union_params(parsed)
    params = sort_params(parsed, params, SORT_PARAMS_BY)
    y_labels = [param_label(p) for p in params]

    for metric in metrics:
        fig_h = max(5.0, len(params) * 0.45 + 1.8)
        fig_w = 11.5
        fig, ax = plt.subplots(figsize=(fig_w, fig_h), constrained_layout=True)

        y = np.arange(len(params))

        bar_h = 0.22 if len(schemes) >= 3 else (0.26 if len(schemes) == 2 else 0.32)
        offsets: Dict[str, float] = {}
        if len(schemes) == 1:
            offsets[schemes[0]] = 0.0
        elif len(schemes) == 2:
            offsets[schemes[0]] = -bar_h / 2
            offsets[schemes[1]] = +bar_h / 2
        else:
            idx = {s: i for i, s in enumerate(schemes)}
            mid = (len(schemes) - 1) / 2.0
            for s in schemes:
                offsets[s] = (idx[s] - mid) * bar_h

        vmax = 0.0
        for scheme in schemes:
            d = parsed[scheme].data
            vals = []
            for p in params:
                st = 0.0
                if p in d and metric in d[p]:
                    st = d[p][metric][1]
                vals.append(zeroize(st))
            vals = np.array(vals, dtype=float)
            vmax = max(vmax, float(vals.max()) if vals.size else 0.0)

            label = " ".join(SCHEME_TITLES.get(scheme, [scheme]))
            ax.barh(y + offsets.get(scheme, 0.0), vals, height=bar_h, label=label)

        ax.set_yticks(y)
        ax.set_yticklabels(y_labels, fontsize=FONT_BASE)
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