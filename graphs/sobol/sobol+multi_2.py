import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import numpy as np
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter

# =====================================================
# НАСТРОЙКИ
# =====================================================

INPUT_FILES = [
    Path(r"C:\Users\Balt_\Desktop\256 100 параметры .txt"),
    Path(r"C:\Users\Balt_\Desktop\256 100 надежность.txt"),
    Path(r"C:\Users\Balt_\Desktop\экономика.txt"),
]

OUTPUT_DIR = Path(r"C:\Users\Balt_\Desktop")

# какие метрики строить отдельными графиками
TARGET_METRICS = ["LCOE", "LPSP", "LOLP"]

# если в файлах вдруг вместо LPSP/LOLP окажутся другие имена, можно задать замену
METRIC_ALIASES = {
    "LCOE": ["LCOE"],
    "LPSP": ["ENS"],
    "LOLP": ["LOLE"],
}

SCHEMES_ORDER = ["SS", "D", "SN", "H", "1", "2", "3", "4", "5"]

SCHEME_TITLES = {
    "SS": "Одиночная секционированная система шин",
    "D": "Двойная система шин",
    "SN": "Одиночная несекционированная система шин",
    "H": "хехе",
    "1": "1",
    "2": "2",
    "3": "3",
    "4": "4",
    "5": "5",
}

SCHEME_BAR_COLORS = {
    "SS": "#426f91",
    "D": "#c98d5b",
}

GROUPS = [
    (
        "Параметры оборудования",
        [
            "BT_CAPACITY_PER_BUS",
            "DG_POWER",
            "DG_COUNT",
            "WT_POWER",
            "BT_MAX_DISCHARGE_CURRENT",
            "BT_NON_RESERVE_DISCHARGE_LVL",
            "BT_MAX_CHARGE_CURRENT",
        ],
    ),
    (
        "Параметры надёжности",
        [
            "WT_FAILURE_RATE",
            "WT_REPAIR_TIME",
            "DG_FAILURE_RATE",
            "DG_REPAIR_TIME",
            "BT_FAILURE_RATE",
            "BT_REPAIR_TIME",
            "BUS_FAILURE_RATE",
            "BUS_REPAIR_TIME",
            "BRK_FAILURE_RATE",
            "BRK_REPAIR_TIME",
            "FIRST_CAT",
            "SECOND_CAT",
        ],
    ),
    (
        "Экономические параметры",
        [
            "COST_FUEL_RUB_PER_KT",
            "DISCOUNT_RATE",
            "COST_WT_RUB_PER_KW",
            "COST_DG_RUB_PER_KW_PER_KMH",
            "COST_DG_RUB_PER_KW",
            "COST_BT_RUB_PER_KWH",
            "COST_WT_RUB_PER_KW_PER_YEAR",
            "COST_BT_RUB_PER_KWH_PER_YEAR",
            "COST_RU_RUB",
            "DAMAGE_RUB_PER_KWH_CAT3",
        ],
    ),
]

SORT_WITHIN_GROUP = "mean"   # "mean" | "max" | "none"

FIGURE_DPI = 200
SAVE_DPI = 300

TITLE_FONT = 14
FONT_BASE = 12
FONT_SMALL = 10

BAR_GROUP_WIDTH = 0.78
GROUP_GAP = 1.6

GRID_COLOR = "#aeb7bd"
GRID_ALPHA = 0.28

ZERO_EPS = 5e-4

plt.rcParams["figure.dpi"] = FIGURE_DPI
plt.rcParams["savefig.dpi"] = SAVE_DPI
plt.rcParams["font.family"] = "DejaVu Sans"

# =====================================================
# ПОДПИСИ ПАРАМЕТРОВ
# =====================================================

PARAM_LABELS_RU = {
    "FIRST_CAT": "Доля потребителей I категории",
    "SECOND_CAT": "Доля потребителей II категории",
    "WT_COUNT": "Количество ВЭУ",
    "WT_POWER": "Мощность одной ВЭУ",
    "WT_COUNT_TOTAL": "Количество ВЭУ",
    "WT_POWER_KW": "Мощность одной ВЭУ",
    "WT_FAILURE_RATE": "Наработка на отказ ВЭУ",
    "WT_REPAIR_TIME": "Время восстановления ВЭУ",
    "DG_COUNT": "Количество ДГУ",
    "DG_POWER": "Мощность одного ДГУ",
    "DG_COUNT_TOTAL": "Количество ДГУ",
    "DG_POWER_KW": "Мощность одного ДГУ",
    "DG_FAILURE_RATE": "Наработка на отказ ДГУ",
    "DG_REPAIR_TIME": "Время восстановления ДГУ",
    "BT_CAPACITY_PER_BUS": "Емкость СНЭ",
    "BT_CAPACITY_KWH_PER_BUS": "Емкость СНЭ",
    "BT_MAX_CHARGE_CURRENT": "Максимальный ток заряда СНЭ",
    "BT_MAX_DISCHARGE_CURRENT": "Максимальный ток разряда СНЭ",
    "BT_NON_RESERVE_DISCHARGE_LVL": "Минимальный уровень заряда СНЭ",
    "BT_FAILURE_RATE": "Наработка на отказ СНЭ",
    "BT_REPAIR_TIME": "Время восстановления СНЭ",
    "BUS_FAILURE_RATE": "Наработка на отказ шины",
    "BUS_REPAIR_TIME": "Время восстановления шины",
    "BRK_FAILURE_RATE": "Наработка на отказ СВ/МШВ",
    "BRK_REPAIR_TIME": "Время восстановления СВ/МШВ",
    "SWITCHGEAR_ROOM_FAILURE_RATE": "Наработка на отказ РУ",
    "SWITCHGEAR_ROOM_REPAIR_TIME": "Время восстановления РУ",
    "DISCOUNT_RATE": "Ставка дисконтирования",
    "COST_RU_RUB": "Стоимость РУ",
    "COST_DG_RUB_PER_KW": "Стоимость ДГУ",
    "COST_DG_RUB_PER_KW_PER_KMH": "Эксплуатационные затраты ДГУ",
    "COST_FUEL_RUB_PER_KT": "Стоимость топлива",
    "COST_WT_RUB_PER_KW": "Стоимость ВЭУ",
    "COST_WT_RUB_PER_KW_PER_YEAR": "Эксплуатационные затраты ВЭУ",
    "COST_BT_RUB_PER_KWH": "Стоимость СНЭ",
    "COST_BT_RUB_PER_KWH_PER_YEAR": "Эксплуатационные затраты СНЭ",
    "DAMAGE_RUB_PER_KWH_CAT3": "Ущерб от недоотпуска",
}

# =====================================================
# REGEX
# =====================================================

SCHEME_HEADER_RE = re.compile(r"^\s*(SN|SS|D|H|1|2|3|4|5)\s*$", re.IGNORECASE)

METRIC_STATS_RE = re.compile(
    r"^\s*(?P<metric>[A-Za-z0-9_]+)\s*:\s*"
    r"var=(?P<var>[^ ]+)\s+"
    r"std=(?P<std>[^ ]+)\s+"
    r"range=\[(?P<min>[^.]+)\.\.(?P<max>[^\]]+)\]\s*$"
)

SINGLE_METRIC_PARAM_RE = re.compile(
    r"^\s*(?P<param>\S+)\s+S\s*=\s*(?P<s>[-+0-9eE.,]+)\s+ST\s*=\s*(?P<st>[-+0-9eE.,]+)\s*$"
)

# =====================================================
# DATA STRUCTURES
# =====================================================

@dataclass(frozen=True)
class SchemeData:
    stats: Dict[str, Dict[str, float]]
    data: Dict[str, Dict[str, Tuple[float, float]]]

# =====================================================
# PARSE
# =====================================================

def parse_number(s: str) -> float:
    s = s.strip().replace("\u00a0", "").replace(" ", "").replace(",", ".")
    return float(s)

def read_text_file(path: Path) -> List[str]:
    if not path.exists():
        raise FileNotFoundError(f"Файл не найден: {path}")
    text = path.read_text(encoding="utf-8-sig")
    return [line.rstrip("\n").rstrip("\r") for line in text.splitlines()]

def split_stats_and_body(lines: List[str]) -> Tuple[List[str], List[str]]:
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
    header = lines_body[0].strip().split()
    if not header or header[0].lower() != "param":
        raise ValueError("Ожидается заголовок таблицы, начинающийся с: param ...")

    col_idx = {name: i for i, name in enumerate(header)}

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
    if "LCOE" in stats:
        return "LCOE"
    if len(stats) == 1:
        return next(iter(stats.keys()))
    if len(stats) == 0:
        return "LCOE"
    return sorted(stats.keys())[0]

def parse_format1_single_metric(
    lines_body: List[str],
    metric_name: str
) -> Dict[str, Dict[str, Tuple[float, float]]]:
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
    lines_body: List[str],
    stats: Dict[str, Dict[str, float]]
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
        raise ValueError("Не найдено ни одной схемы.")

    parsed: Dict[str, SchemeData] = {}
    for scheme in SCHEMES_ORDER:
        if scheme not in blocks:
            continue

        ls, lb = split_stats_and_body(blocks[scheme])
        stats = parse_metric_stats(ls)
        data = parse_body_auto(lb, stats)
        parsed[scheme] = SchemeData(stats=stats, data=data)

    if not parsed:
        raise ValueError("Не удалось распарсить данные ни по одной схеме.")
    return parsed

def merge_scheme_dicts(parts: List[Dict[str, SchemeData]]) -> Dict[str, SchemeData]:
    merged: Dict[str, SchemeData] = {}

    for part in parts:
        for scheme, scheme_data in part.items():
            if scheme not in merged:
                merged[scheme] = SchemeData(
                    stats=dict(scheme_data.stats),
                    data={p: dict(v) for p, v in scheme_data.data.items()}
                )
                continue

            old = merged[scheme]
            new_stats = dict(old.stats)
            new_stats.update(scheme_data.stats)

            new_data = {p: dict(v) for p, v in old.data.items()}
            for p, metric_map in scheme_data.data.items():
                if p not in new_data:
                    new_data[p] = {}
                new_data[p].update(metric_map)

            merged[scheme] = SchemeData(stats=new_stats, data=new_data)

    return merged

# =====================================================
# HELPERS
# =====================================================

def ensure_output_dir() -> Path:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    return OUTPUT_DIR

def zeroize(v: float) -> float:
    return 0.0 if (v <= 0.0 or abs(v) < ZERO_EPS) else v

def comma_tick(x, pos):
    if abs(x) < 1e-12:
        x = 0.0
    s = f"{x:.2f}".replace(".", ",")
    while s.endswith("0"):
        s = s[:-1]
    if s.endswith(","):
        s = s[:-1]
    return s

def param_label(code: str) -> str:
    return PARAM_LABELS_RU.get(code, code)

def get_present_schemes(parsed: Dict[str, SchemeData]) -> List[str]:
    return [s for s in SCHEMES_ORDER if s in parsed and s in ("SS", "D")]

def resolve_metric_name_for_scheme(scheme_data: SchemeData, target_metric: str) -> Optional[str]:
    aliases = METRIC_ALIASES.get(target_metric, [target_metric])
    for candidate in aliases:
        for p in scheme_data.data.values():
            if candidate in p:
                return candidate
    for candidate in aliases:
        if candidate in scheme_data.stats:
            return candidate
    return None

def collect_rows_for_metric(parsed: Dict[str, SchemeData], target_metric: str):
    schemes = get_present_schemes(parsed)
    rows = []

    for group_name, params_in_group in GROUPS:
        group_rows = []
        for param_code in params_in_group:
            vals = {}
            exists = False
            for scheme in schemes:
                real_metric = resolve_metric_name_for_scheme(parsed[scheme], target_metric)
                if (
                    real_metric is not None
                    and param_code in parsed[scheme].data
                    and real_metric in parsed[scheme].data[param_code]
                ):
                    vals[scheme] = zeroize(parsed[scheme].data[param_code][real_metric][1])
                    exists = True
                else:
                    vals[scheme] = 0.0

            if exists:
                group_rows.append((group_name, param_code, vals))

        if SORT_WITHIN_GROUP == "mean":
            group_rows.sort(
                key=lambda rec: float(np.mean([rec[2].get(s, 0.0) for s in schemes])),
                reverse=True,
            )
        elif SORT_WITHIN_GROUP == "max":
            group_rows.sort(
                key=lambda rec: float(max([rec[2].get(s, 0.0) for s in schemes])),
                reverse=True,
            )
        elif SORT_WITHIN_GROUP == "none":
            pass
        else:
            raise ValueError(f"Неизвестный режим SORT_WITHIN_GROUP={SORT_WITHIN_GROUP}")

        rows.extend(group_rows)

    return rows

def save_figure(fig, metric_name: str) -> Path:
    ensure_output_dir()
    out_path = OUTPUT_DIR / f"sobol_{metric_name}.png"
    fig.savefig(out_path, dpi=SAVE_DPI, bbox_inches="tight")
    print(f"Сохранено: {out_path}")
    return out_path

# =====================================================
# PLOT
# =====================================================

def plot_metric(parsed: Dict[str, SchemeData], metric_name: str) -> None:
    schemes = get_present_schemes(parsed)
    if not schemes:
        raise ValueError("Нет схем SS/D для построения.")

    rows = collect_rows_for_metric(parsed, metric_name)
    if not rows:
        print(f"Пропуск {metric_name}: метрика не найдена в данных.")
        return

    x_positions = []
    x_labels = []
    values_by_scheme = {s: [] for s in schemes}
    group_meta = []

    x = 0.0
    first_group = True
    idx = 0

    while idx < len(rows):
        group_name = rows[idx][0]

        if not first_group:
            x += GROUP_GAP
        first_group = False

        start_x = x

        while idx < len(rows) and rows[idx][0] == group_name:
            _, param_code, vals = rows[idx]
            x_positions.append(x)
            x_labels.append(param_label(param_code))
            for s in schemes:
                values_by_scheme[s].append(vals.get(s, 0.0))
            x += 1.0
            idx += 1

        end_x = x - 1.0
        group_meta.append((group_name, start_x, end_x))

    max_val = 0.0
    for s in schemes:
        if values_by_scheme[s]:
            max_val = max(max_val, max(values_by_scheme[s]))
    max_val = max(max_val, 0.05)

    fig_w = max(14.0, len(x_labels) * 0.55)
    fig_h = 7.2
    fig, ax = plt.subplots(figsize=(fig_w, fig_h), dpi=FIGURE_DPI)

    if len(schemes) == 1:
        offsets = {schemes[0]: 0.0}
        bar_width = BAR_GROUP_WIDTH
    elif len(schemes) == 2:
        bar_width = BAR_GROUP_WIDTH / 2.0
        offsets = {
            schemes[0]: -bar_width / 2.0,
            schemes[1]: +bar_width / 2.0,
        }
    else:
        bar_width = BAR_GROUP_WIDTH / len(schemes)
        mid = (len(schemes) - 1) / 2.0
        offsets = {s: (i - mid) * bar_width for i, s in enumerate(schemes)}

    for s in schemes:
        ax.bar(
            np.array(x_positions) + offsets[s],
            values_by_scheme[s],
            width=bar_width,
            color=SCHEME_BAR_COLORS.get(s, "#8ea9b8"),
            edgecolor="none",
            label=SCHEME_TITLES.get(s, s),
            zorder=3,
        )

    ax.set_xticks(x_positions)
    ax.set_xticklabels(x_labels, rotation=58, ha="right", fontsize=FONT_SMALL)

    for i, (group_name, start_x, end_x) in enumerate(group_meta):
        x_center = (start_x + end_x) / 2.0
        ax.text(
            x_center,
            max_val * 1.14,
            group_name,
            ha="center",
            va="bottom",
            fontsize=FONT_BASE,
            fontweight="bold",
        )
        if i < len(group_meta) - 1:
            ax.axvline(end_x + GROUP_GAP / 2.0, color="#9a9a9a", lw=0.8, alpha=0.55, zorder=1)

    ax.set_ylabel("ST (доля дисперсии)", fontsize=FONT_BASE)
    ax.set_title(f"Чувствительность параметров по метрике {metric_name}", fontsize=TITLE_FONT)

    ax.set_ylim(0, max_val * 1.24)
    ax.grid(axis="y", alpha=GRID_ALPHA, color=GRID_COLOR, zorder=1)
    ax.yaxis.set_major_formatter(FuncFormatter(comma_tick))

    ax.legend(
        frameon=True,
        fontsize=FONT_SMALL,
        loc="upper left",
        bbox_to_anchor=(1.01, 1.0),
        borderaxespad=0.0,
    )

    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    plt.tight_layout()
    save_figure(fig, metric_name)
    plt.show()

# =====================================================
# MAIN
# =====================================================

def main():
    parsed_parts: List[Dict[str, SchemeData]] = []

    for path in INPUT_FILES:
        print(f"Чтение: {path}")
        lines = read_text_file(path)
        parsed_parts.append(parse_multi_scheme_input(lines))

    parsed = merge_scheme_dicts(parsed_parts)

    schemes = get_present_schemes(parsed)
    print("Найдены схемы:", ", ".join(schemes) if schemes else "нет")

    for metric in TARGET_METRICS:
        plot_metric(parsed, metric)

if __name__ == "__main__":
    main()