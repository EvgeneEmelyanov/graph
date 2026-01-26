import re
import argparse
from typing import List, Tuple, Dict

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


def parse_format2(lines: List[str]) -> Dict[str, Dict[str, Tuple[float, float]]]:
    """
    Возвращает:
      data[param][metric] = (S, ST)
    где metric ∈ {"LCOE","ENS","Fuel","Moto"}
    """
    header = lines[0].strip().split()
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
    for line in lines[1:]:
        parts = line.strip().split()
        if not parts:
            continue
        param = parts[0]
        if len(parts) < len(header):
            raise ValueError(f"Строка короче заголовка (format2): {line}")

        data[param] = {}
        for metric in EXPECTED_METRICS:
            s = float(parts[col_idx[f"S_{metric}"]])
            st = float(parts[col_idx[f"ST_{metric}"]])
            data[param][metric] = (s, st)

    if not data:
        raise ValueError("Format2: не найдено ни одной строки с параметрами")
    return data


def plot_sobol_single(rows: List[Tuple[str, float, float]], title="Sobol indices"):
    names = [r[0] for r in rows]
    S = np.array([r[1] for r in rows]) * 100.0
    ST = np.array([r[2] for r in rows]) * 100.0

    n = len(rows)
    x = np.arange(n)
    width = 0.38

    fig, ax = plt.subplots(figsize=(max(10, n * 0.9), 5))
    ax.bar(x - width / 2, S, width, label="S")
    ax.bar(x + width / 2, ST, width, label="ST")

    ax.set_title(title)
    ax.set_ylabel("Contribution to variance, %")
    ax.set_xticks(x)
    ax.set_xticklabels([str(i + 1) for i in range(n)])
    ax.legend()
    ax.grid(axis="y", alpha=0.3)

    ymin = min(S.min(), ST.min())
    ymax = max(S.max(), ST.max())
    pad = max((ymax - ymin) * 0.05, 0.5)

    # более корректный ylim при отрицательных значениях
    ax.set_ylim(ymin - pad, ymax + pad)

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


def plot_sobol_multi(data: Dict[str, Dict[str, Tuple[float, float]]], title_prefix="Sobol indices"):
    # общий порядок параметров
    params = list(data.keys())

    for metric in EXPECTED_METRICS:
        rows = [(p, data[p][metric][0], data[p][metric][1]) for p in params]
        plot_sobol_single(rows, title=f"{title_prefix}: {metric}")


def autodetect_mode(lines: List[str]) -> str:
    first = lines[0].strip()
    if LINE_RE.match(first):
        return "format1"
    if first.lower().startswith("param " ) or first.lower() == "param":
        return "format2"
    # попробуем угадать по первой строке: если там 1-й токен "param" — format2
    if first.split() and first.split()[0].lower() == "param":
        return "format2"
    raise ValueError("Не удалось определить формат входных данных (ожидается format1 или format2).")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--mode",
        choices=["auto", "format1", "format2"],
        default="auto",
        help="auto: определить по первой строке; format1: 'NAME S=.. ST=..'; format2: таблица с колонками S_LCOE/ST_LCOE/...",
    )
    parser.add_argument("--title", default="Sobol indices", help="Заголовок/префикс заголовка графиков")
    args = parser.parse_args()

    lines = read_lines_from_console()
    mode = autodetect_mode(lines) if args.mode == "auto" else args.mode

    if mode == "format1":
        rows = parse_format1(lines)
        plot_sobol_single(rows, title=args.title)
    elif mode == "format2":
        data = parse_format2(lines)
        plot_sobol_multi(data, title_prefix=args.title)
    else:
        raise RuntimeError("Неизвестный режим")


if __name__ == "__main__":
    main()
