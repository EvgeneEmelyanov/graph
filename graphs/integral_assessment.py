# =========================
# CODE 1 (grid -> best (param1,param2) per scenario) + LOLE_h
# Вывод ТОЛЬКО в консоль, никаких Excel-файлов
# =========================
from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Dict, List, Tuple, Optional

import numpy as np
import pandas as pd


# =========================
# НАСТРОЙКИ ПОЛЬЗОВАТЕЛЯ
# =========================

EXCEL_PATH = r"D:\results.xlsx"
SHEET_NAME: Optional[str] = "SWEEP_2"  # None = первый лист

# Порядок параметров:
# param1 = подписи строк (колонка A под таблицей)
# param2 = подписи столбцов (строка заголовков)
# Если у тебя теперь наоборот — поставь SWAP_AXES=True
PARAM1_NAME = "param1"
PARAM2_NAME = "param2"
SWAP_AXES = False

# Метрики — точно как написано в колонке A в Excel
METRICS = [
    "LCOE, руб/кВт∙ч",

    "ENS,кВт∙ч",
    "LOLE_h",
    "ENS_evtN",
    "ENS_evtMaxH",
    # "ENS1_mean",
    # "ENS2_mean",

    "Расход топлива, тыс.тонн",
    "Моточасы, тыс.мч",
]

# Направления: 'min' (меньше лучше) или 'max' (больше лучше)
DIRECTIONS = {m: "min" for m in METRICS}

# Веса (если не указать — 1.0)
WEIGHTS = {m: 1.0 for m in METRICS}

# =========================
# КОНЕЦ НАСТРОЕК
# =========================


def to_num(x) -> float:
    """Понимает числа с точкой и с запятой. Если не число — NaN."""
    if x is None:
        return float("nan")
    if isinstance(x, (int, float, np.integer, np.floating)):
        return float(x)
    s = str(x).strip()
    if s == "" or s.lower() == "nan":
        return float("nan")
    s = s.replace("\xa0", "").replace(" ", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return float("nan")


def is_text_label(x) -> bool:
    if x is None:
        return False
    return bool(re.search(r"[A-Za-zА-Яа-я]", str(x)))


def find_blocks(header_vals: List) -> List[Tuple[int, int]]:
    """
    Находит блоки по строке заголовков (B..):
    [числа...][пусто][числа...][пусто]...
    Возвращает (start,end) индексы относительно B.. (без колонки A).
    """
    hdr = np.array([to_num(v) for v in header_vals], dtype=float)
    blocks: List[Tuple[int, int]] = []
    c = 0
    while c < len(hdr):
        if not np.isnan(hdr[c]):
            start = c
            while c < len(hdr) and not np.isnan(hdr[c]):
                c += 1
            end = c
            blocks.append((start, end))
        else:
            c += 1
    return blocks


@dataclass
class TableMeta:
    metric_name: str
    label_row: int
    header_row: int
    blocks: List[Tuple[int, int]]


def locate_metric_table(raw: pd.DataFrame, metric_name: str) -> TableMeta:
    """
    Ищет строку, где в колонке A написано metric_name.
    Строка заголовка (числа по столбцам) — следующая.
    """
    col0 = raw.iloc[:, 0].astype(str).str.strip()
    idx = raw.index[col0.eq(metric_name)].tolist()
    if not idx:
        idx = raw.index[col0.str.contains(re.escape(metric_name), na=False)].tolist()
    if not idx:
        raise ValueError(f"Метрика не найдена в колонке A: '{metric_name}'")

    r = idx[0]
    header_row = r + 1
    if header_row >= len(raw):
        raise ValueError(f"У метрики '{metric_name}' нет строки заголовков ниже.")

    header_vals = raw.iloc[header_row, 1:].tolist()
    blocks = find_blocks(header_vals)
    if not blocks:
        raise ValueError(
            f"Метрика '{metric_name}' найдена, но не удалось распознать блоки заголовков."
        )
    return TableMeta(metric_name, r, header_row, blocks)


def extract_block_matrix(
    raw: pd.DataFrame, meta: TableMeta, block: Tuple[int, int]
) -> Tuple[np.ndarray, np.ndarray, np.ndarray]:
    """
    Для одного блока (сценария) одной метрики вытаскивает:
    row_labels (колонка A ниже таблицы)
    col_labels (строка заголовка по столбцам)
    values (матрица [nrow, ncol])
    """
    start, end = block
    col_labels = np.array(
        [to_num(v) for v in raw.iloc[meta.header_row, 1 + start : 1 + end].tolist()],
        dtype=float,
    )

    r = meta.label_row + 2
    row_labels: List[float] = []
    data: List[np.ndarray] = []

    while r < len(raw):
        a = raw.iat[r, 0]
        if is_text_label(a):
            break

        row_lab = to_num(a)
        row_vals = raw.iloc[r, 1 + start : 1 + end].tolist()
        row_nums = np.array([to_num(v) for v in row_vals], dtype=float)

        if np.isnan(row_lab) and np.isnan(row_nums).all():
            break

        row_labels.append(row_lab)
        data.append(row_nums)
        r += 1

    values = np.vstack(data) if data else np.zeros((0, len(col_labels)), dtype=float)
    return np.array(row_labels, dtype=float), col_labels, values


def normalize_minmax(x: np.ndarray) -> np.ndarray:
    xmin = np.nanmin(x)
    xmax = np.nanmax(x)
    if not np.isfinite(xmin) or not np.isfinite(xmax) or xmax <= xmin:
        return np.zeros_like(x, dtype=float)
    return (x - xmin) / (xmax - xmin)


def score_grid(
    metric_grids: Dict[str, np.ndarray],
    directions: Dict[str, str],
    weights: Dict[str, float],
) -> np.ndarray:
    """
    Нормализация ВНУТРИ сценария по каждой метрике.
    Score ниже = лучше.
    """
    score = None
    for m, grid in metric_grids.items():
        gn = normalize_minmax(grid.astype(float))
        if directions.get(m, "min").lower() == "max":
            gn = 1.0 - gn
        w = float(weights.get(m, 1.0))
        score = (w * gn) if score is None else (score + w * gn)
    return score


def _fmt(v: float) -> str:
    if not np.isfinite(v):
        return "NaN"
    av = abs(v)
    if av != 0 and (av < 1e-3 or av >= 1e4):
        return f"{v:.6e}"
    return f"{v:.6f}".rstrip("0").rstrip(".")


def run():
    raw = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME, header=None)

    metas = [locate_metric_table(raw, m) for m in METRICS]

    blocks = metas[0].blocks
    for meta in metas[1:]:
        if meta.blocks != blocks:
            raise ValueError(
                "Структура блоков (сценариев) различается между метриками:\n"
                f"  '{metas[0].metric_name}' vs '{meta.metric_name}'"
            )

    best_rows = []

    for scen_idx, block in enumerate(blocks, start=1):
        row_labels, col_labels, _ = extract_block_matrix(raw, metas[0], block)

        metric_grids: Dict[str, np.ndarray] = {}
        for meta in metas:
            rlab, clab, vals = extract_block_matrix(raw, meta, block)
            if vals.shape != (len(row_labels), len(col_labels)):
                raise ValueError(
                    f"Сценарий {scen_idx}, метрика '{meta.metric_name}': размер {vals.shape}, "
                    f"ожидалось {(len(row_labels), len(col_labels))}"
                )
            metric_grids[meta.metric_name] = vals

        if SWAP_AXES:
            row_labels, col_labels = col_labels, row_labels
            metric_grids = {m: v.T for m, v in metric_grids.items()}

        S = score_grid(metric_grids, DIRECTIONS, WEIGHTS)
        k = int(np.nanargmin(S))
        r = k // S.shape[1]
        c = k % S.shape[1]

        rec = {
            "scenario": scen_idx,
            PARAM1_NAME: float(row_labels[r]),
            PARAM2_NAME: float(col_labels[c]),
            "Score": float(S[r, c]),
        }
        for m in METRICS:
            rec[m] = float(metric_grids[m][r, c])
        best_rows.append(rec)

    best = pd.DataFrame(best_rows)

    # IQR и частоты
    def q(col, p):
        x = best[col].dropna().astype(float)
        return float(x.quantile(p)) if len(x) else float("nan")

    p1_q25, p1_q50, p1_q75 = q(PARAM1_NAME, 0.25), q(PARAM1_NAME, 0.50), q(PARAM1_NAME, 0.75)
    p2_q25, p2_q50, p2_q75 = q(PARAM2_NAME, 0.25), q(PARAM2_NAME, 0.50), q(PARAM2_NAME, 0.75)

    freq_p1 = best[PARAM1_NAME].value_counts()
    freq_p2 = best[PARAM2_NAME].value_counts()

    # ---------- PRINT ----------
    cols = ["scenario", PARAM1_NAME, PARAM2_NAME, "Score"] + METRICS

    print("\n=== Best point per scenario (min Score) ===")
    print(" | ".join(cols))
    for _, row in best[cols].iterrows():
        parts = [
            f"{int(row['scenario'])}",
            _fmt(row[PARAM1_NAME]),
            _fmt(row[PARAM2_NAME]),
            _fmt(row["Score"]),
        ] + [_fmt(row[m]) for m in METRICS]
        print(" | ".join(parts))

    print("\n=== Robust IQR of selected params ===")
    print(f"{PARAM1_NAME}: q25={_fmt(p1_q25)}  median={_fmt(p1_q50)}  q75={_fmt(p1_q75)}")
    print(f"{PARAM2_NAME}: q25={_fmt(p2_q25)}  median={_fmt(p2_q50)}  q75={_fmt(p2_q75)}")

    print("\n=== Frequency (selected params) ===")
    print(f"[{PARAM1_NAME}]")
    for v, cnt in freq_p1.items():
        print(f"  { _fmt(float(v)) }: {int(cnt)}")
    print(f"[{PARAM2_NAME}]")
    for v, cnt in freq_p2.items():
        print(f"  { _fmt(float(v)) }: {int(cnt)}")


if __name__ == "__main__":
    run()