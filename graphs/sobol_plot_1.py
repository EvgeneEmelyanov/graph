import re
import argparse
from typing import List, Tuple, Dict, Optional

import numpy as np
import matplotlib.pyplot as plt

# ====== Format 1 (single metric per run):  PARAM S=... ST=... ======
LINE_RE = re.compile(
    r"^\s*(?P<name>[A-Z0-9_]+)\s+"
    r"S=(?P<S>[-+]?(\d+(\.\d*)?|\.\d+)([eE][-+]?\d+)?)\s+"
    r"ST=(?P<ST>[-+]?(\d+(\.\d*)?|\.\d+)([eE][-+]?\d+)?)\s*$"
)

# ====== Format 2 (table): param S_LCOE ST_LCOE S_ENS ST_ENS ... ======
EXPECTED_METRICS = ["LCOE", "ENS", "Fuel", "Moto"]

# ====== New: metric stats lines (with comma decimals)
# Example:
# LCOE: var=12481,5 std=111,720 range=[29,0663..818,510]
METRIC_STATS_RE = re.compile(
    r"^\s*(?P<metric>LCOE|ENS|Fuel|Moto)\s*:\s*"
    r"var=(?P<var>[^ ]+)\s+"
    r"std=(?P<std>[^ ]+)\s+"
    r"range=\[(?P<min>[^.]+)\.\.(?P<max>[^\]]+)\]\s*$"
)

def parse_number(s: str) -> float:
    """
    Парсит число, где десятичный разделитель может быть ','.
    Поддерживает научную нотацию вида 7,95e+13.
    """
    s = s.strip().replace("\u00a0", "").replace(" ", "")
    s = s.replace(",", ".")
    return float(s)

def read_lines_from_console() -> List[str]:
    """
    Читает строки из консоли.
    Вставляешь данные, потом ОДНА пустая строка — ввод заканчивается.
    """
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

def parse_format1(lines: List[str]) -> List[Tuple[str, float, float]]:
    rows = []
    for line in lines:
        m = LINE_RE.match(line.strip())
        if not m:
            raise ValueError(f"Не удалось распарсить строку (format1): {line}")
        rows.append((m.group("name"), float(m.group("S")), float(m.group("ST"))))
    return rows

def split_stats_and_table(lines: List[str]) -> Tuple[List[str], List[str]]:
    """
    Делит ввод на блок статистик метрик (сверху) и табличный блок, начиная с 'param ...'.
    """
    header_idx = None
    for i, line in enumerate(lines):
        if line.strip().lower().startswith("param"):
            header_idx = i
            break
    if header_idx is None:
        return lines, []
    return lines[:header_idx], lines[header_idx:]

def parse_metric_stats(lines_stats: List[str]) -> Dict[str, Dict[str, float]]:
    """
    Возвращает:
      stats[metric] = {"var":..., "std":..., "min":..., "max":...}
    """
    stats: Dict[str, Dict[str, float]] = {}
    for line in lines_stats:
        m = METRIC_STATS_RE.match(line.strip())
        if not m:
            # игнорируем любые не-статистические строки в верхнем блоке (на всякий случай)
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
    Парсит таблицу:
      param S_LCOE ST_LCOE S_ENS ST_ENS ...
    Возвращает:
      data[param][metric] = (S, ST)
    """
    if not lines_table:
        raise ValueError("Format2: не найден табличный блок (строка 'param ...')")

    header = lines_table[0].strip().split()
    if not header or header[0].lower() != "param":
        raise ValueError("Format2 ожидает заголовок, начинающийся с: param ...")

    col_idx = {name: i for i, name in enumerate(header)}
    missing = []
    for metric in EXPECTED_METRICS:
        if f"S_{metric}" not in col_idx or f"ST_{metric}" not in col_idx:
            missing.append(metric)
    if missing:
        raise ValueError(
            "В format2 не найдены пары колонок S_*/ST_* для метрик: " + ", ".join(missing)
        )

    data: Dict[str, Dict[str, Tuple[float, float]]] = {}
    for line in lines_table[1:]:
        parts = line.strip().split()
        if not parts:
            continue
        param = parts[0]
        if len(parts) < len(header):
            raise ValueError(f"{line}")

        data[param] = {}
        for metric in EXPECTED_METRICS:
            s = parse_number(parts[col_idx[f"S_{metric}"]])
            st = parse_number(parts[col_idx[f"ST_{metric}"]])
            data[param][metric] = (s, st)

    if not data:
        raise ValueError("Format2: не найдено ни одной строки с параметрами")
    return data

def make_metric_caption_ru(metric: str, stats: Optional[Dict[str, Dict[str, float]]]) -> str:
    """
    Краткая подпись на русском:
      Дисп.=..., СКО=..., Диапазон=[..;..]
    """
    if not stats or metric not in stats:
        return ""
    s = stats[metric]
    # Короткая русская расшифровка:
    # var -> Дисп. (дисперсия)
    # std -> СКО (среднеквадратическое отклонение)
    # range -> Диапазон
    return f"Var.= {s['var']:.4g}   std= {s['std']:.4g}   Диапазон=[{s['min']:.4g}; {s['max']:.4g}]"

def plot_sobol_single(rows: List[Tuple[str, float, float]], title="", caption: str = ""):
    names = [r[0] for r in rows]
    S = np.array([r[1] for r in rows]) * 1.0
    ST = np.array([r[2] for r in rows]) * 1.0

    # если число меньше 0 — приравниваем к 0
    S = np.maximum(S, 0.0)
    ST = np.maximum(ST, 0.0)

    n = len(rows)
    x = np.arange(n)
    width = 0.38

    fig, ax = plt.subplots(figsize=(max(10, n * 0.9), 5))
    ax.bar(x - width / 2, S, width, label="S")
    ax.bar(x + width / 2, ST, width, label="ST")

    ax.set_title(title)
    ax.set_ylabel("Доля дисперсии выходной метрики")
    ax.set_xticks(x)
    ax.set_xticklabels([str(i + 1) for i in range(n)])
    ax.legend()
    ax.grid(axis="y", alpha=0.3)

    ymax = max(S.max(), ST.max()) if n > 0 else 1.0
    pad = max(ymax * 0.05, 0.5)
    ax.set_ylim(0, ymax + pad)

    # верхняя подпись со статистикой метрики (по-русски)
    if caption:
        ax.text(
            0.01, 0.99,
            caption,
            transform=ax.transAxes,
            ha="left",
            va="top",
            fontsize=10
        )

    # подписи параметров
    for i, name in enumerate(names):
        y = max(S[i], ST[i])
        ax.text(
            x[i] + 0.06,
            y + pad * 0.15,
            name,
            rotation=35,
            ha="left",
            va="bottom",
            fontsize=9,
        )

    plt.tight_layout()
    plt.show()

def plot_sobol_multi(
    data: Dict[str, Dict[str, Tuple[float, float]]],
    title_prefix="",
    stats: Optional[Dict[str, Dict[str, float]]] = None
):
    # общий порядок параметров
    params = list(data.keys())

    for metric in EXPECTED_METRICS:
        rows = [(p, data[p][metric][0], data[p][metric][1]) for p in params]
        caption = make_metric_caption_ru(metric, stats)
        plot_sobol_single(rows, title=f"{title_prefix}: {metric}", caption=caption)

def autodetect_mode(lines: List[str]) -> str:
    """
    Теперь режим авто:
    - если есть строка, начинающаяся с 'param' => format2 (с возможными строками статистик сверху)
    - иначе: format1
    """
    for line in lines:
        if line.strip().lower().startswith("param"):
            return "format2"
    first = lines[0].strip()
    if LINE_RE.match(first):
        return "format1"
    # если не совпало, но табличного 'param' тоже нет — считаем что это format1 и упадём в parse_format1 с ошибкой
    return "format1"

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--mode",
        choices=["auto", "format1", "format2"],
        default="auto",
        help="auto: определить по наличию строки 'param ...'; format1: 'NAME S=.. ST=..'; format2: таблица + (опц.) строки статистик метрик сверху",
    )
    parser.add_argument("--title", default="", help="Заголовок/префикс заголовка графиков")
    args = parser.parse_args()

    lines = read_lines_from_console()
    mode = autodetect_mode(lines) if args.mode == "auto" else args.mode

    if mode == "format1":
        rows = parse_format1(lines)
        plot_sobol_single(rows, title=args.title)

    elif mode == "format2":
        lines_stats, lines_table = split_stats_and_table(lines)
        stats = parse_metric_stats(lines_stats)
        data = parse_format2_table(lines_table)
        plot_sobol_multi(data, title_prefix=args.title, stats=stats)

    else:
        raise RuntimeError("Неизвестный режим")

if __name__ == "__main__":
    main()
