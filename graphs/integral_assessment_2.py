# =========================
# CODE 2 (compare scenarios/groups) rewritten WITHOUT tab-splitting + LOLE_h
# =========================
import re
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

import numpy as np
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet


# =========================
# НАСТРОЙКИ ПОЛЬЗОВАТЕЛЯ
# =========================

EXCEL_PATH = r"D:\123.xlsx"
SHEET_NAME: Optional[str] = "SWEEP_2"  # None = первый лист

METRICS = [
    # "LCOE, руб/кВт∙ч",

    "ENS,кВт∙ч",
    "LOLE_h",
    "ENS1_mean",
    "ENS2_mean",
    "ENS_evtN",
    "ENS_evtMaxH",

    # "Расход топлива, тыс.тонн",
    # "Моточасы, тыс.мч",
]

DIRECTIONS: Dict[str, str] = {m: "min" for m in METRICS}
WEIGHTS: Dict[str, float] = {m: 1.0 for m in METRICS}

# Как оценивать каждую группу по каждой метрике:
#  - "best"   : берём лучший достижимый (min или max) в таблице группы
#  - "mean"   : среднее по таблице
#  - "median" : медиана по таблице
AGGREGATOR = "best"  # "best" | "mean" | "median"

# Как объединять метрики в общий итог:
#  - "sum"  : взвешенная сумма нормализованных значений (0=лучше)
#  - "rank" : взвешенная сумма рангов (1=лучше)
COMBINE_MODE = "sum"  # "sum" | "rank"

# Названия групп (сценариев) слева-направо (если None — G1..Gk)
GROUP_NAMES: Optional[List[str]] = None

# Скан справа, чтобы найти конец заголовка
MAX_COL_SCAN = 500

# =========================
# Data model
# =========================

@dataclass
class Table2D:
    x_labels: np.ndarray  # float labels
    y_labels: np.ndarray  # float labels
    values: np.ndarray    # [ny, nx]


# =========================
# Helpers
# =========================

def _to_float_ru_or_nan(v: object) -> float:
    if v is None:
        return np.nan
    if isinstance(v, (int, float, np.integer, np.floating)):
        return float(v)
    s = str(v).replace("\xa0", "").replace(" ", "").strip()
    if s == "":
        return np.nan
    s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return np.nan

def _is_text(v: object) -> bool:
    if v is None:
        return False
    return bool(re.search(r"[A-Za-zА-Яа-я]", str(v)))

def _find_metric_row(ws: Worksheet, metric: str) -> int:
    for r in range(1, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if v is None:
            continue
        if str(v).strip() == metric:
            return r
    raise ValueError(f"Metric '{metric}' not found in column A.")

def _detect_table_height(ws: Worksheet, header_row: int) -> int:
    """
    Таблица данных идёт с header_row (строка с X) и далее строки с Y.
    Заканчиваем на первой "пустой" строке или на строке с текстовой меткой в колонке A.
    """
    r = header_row + 1
    last = header_row
    while r <= ws.max_row:
        a = ws.cell(row=r, column=1).value
        if _is_text(a):
            break
        # строка пустая по первым 20 колонкам
        any_non_empty = False
        for c in range(1, 21):
            if ws.cell(row=r, column=c).value not in (None, ""):
                any_non_empty = True
                break
        if not any_non_empty:
            break
        last = r
        r += 1
    return last

def _find_blocks_in_header(ws: Worksheet, header_row: int) -> List[Tuple[int, int]]:
    """
    Возвращает список блоков (start_col, end_col_exclusive) по строке header_row,
    где X-метки числовые, а блоки разделены пустыми колонками.
    """
    # Определяем правую границу заголовка
    last_non_empty = 1
    for c in range(1, min(ws.max_column, MAX_COL_SCAN) + 1):
        v = ws.cell(row=header_row, column=c).value
        if v not in (None, ""):
            last_non_empty = c

    # Сканируем B..last_non_empty (A обычно пустая в заголовке)
    blocks: List[Tuple[int, int]] = []
    c = 2
    while c <= last_non_empty:
        v = _to_float_ru_or_nan(ws.cell(row=header_row, column=c).value)
        if np.isfinite(v):
            start = c
            c += 1
            while c <= last_non_empty:
                vv = _to_float_ru_or_nan(ws.cell(row=header_row, column=c).value)
                if not np.isfinite(vv):
                    break
                c += 1
            end = c
            blocks.append((start, end))
        else:
            c += 1
    if not blocks:
        raise ValueError("Could not detect any numeric header blocks (groups) in the header row.")
    return blocks

def load_metric_tables_from_excel(xlsx_path: str, sheet_name: Optional[str], metric: str) -> List[Table2D]:
    wb = load_workbook(xlsx_path, data_only=True)
    try:
        ws = wb[sheet_name] if sheet_name else wb[wb.sheetnames[0]]
        mr = _find_metric_row(ws, metric)
        header_row = mr + 1
        if header_row > ws.max_row:
            raise ValueError(f"Metric '{metric}' has no header row below it.")

        max_row = _detect_table_height(ws, header_row)
        blocks = _find_blocks_in_header(ws, header_row)

        tables: List[Table2D] = []
        for (c0, c1) in blocks:
            x = np.array([_to_float_ru_or_nan(ws.cell(row=header_row, column=c).value) for c in range(c0, c1)], dtype=float)

            y_list: List[float] = []
            vals: List[List[float]] = []
            for r in range(header_row + 1, max_row + 1):
                yv = _to_float_ru_or_nan(ws.cell(row=r, column=1).value)
                row_vals = [_to_float_ru_or_nan(ws.cell(row=r, column=c).value) for c in range(c0, c1)]
                # стоп на полностью пустой строке
                if (not np.isfinite(yv)) and all(not np.isfinite(v) for v in row_vals):
                    break
                y_list.append(yv)
                vals.append(row_vals)

            arr = np.array(vals, dtype=float)
            tables.append(Table2D(x_labels=x, y_labels=np.array(y_list, dtype=float), values=arr))

        return tables
    finally:
        wb.close()


# =========================
# Scoring
# =========================

def _finite(vals: np.ndarray) -> np.ndarray:
    return vals[np.isfinite(vals)]

def _aggregate_value(table: Table2D, direction: str, aggregator: str) -> float:
    v = _finite(table.values)
    if v.size == 0:
        return np.nan
    if aggregator == "best":
        return float(np.min(v) if direction == "min" else np.max(v))
    if aggregator == "mean":
        return float(np.mean(v))
    if aggregator == "median":
        return float(np.median(v))
    raise ValueError(f"Unknown AGGREGATOR={aggregator}")

def _normalize(values: np.ndarray, direction: str) -> np.ndarray:
    """
    0..1 (0=лучше).
    min: (x - best)/(worst-best)
    max: (best - x)/(best-worst)
    """
    v = values.astype(float)
    if np.all(~np.isfinite(v)):
        return np.full_like(v, np.nan, dtype=float)

    finite = v[np.isfinite(v)]
    vmin = float(np.min(finite))
    vmax = float(np.max(finite))
    if abs(vmax - vmin) < 1e-15:
        out = np.zeros_like(v, dtype=float)
        out[~np.isfinite(v)] = np.nan
        return out

    if direction == "min":
        out = (v - vmin) / (vmax - vmin)
    elif direction == "max":
        out = (vmax - v) / (vmax - vmin)
    else:
        raise ValueError(f"Direction must be 'min' or 'max', got: {direction}")

    out[~np.isfinite(v)] = np.nan
    return out

def compare_groups_from_excel(
    xlsx_path: str,
    sheet_name: Optional[str],
    metrics: List[str],
    directions: Dict[str, str],
    weights: Dict[str, float],
    aggregator: str,
    combine_mode: str,
    group_names: Optional[List[str]] = None,
) -> None:
    metric_tables: Dict[str, List[Table2D]] = {}
    k_expected: Optional[int] = None

    for m in metrics:
        tabs = load_metric_tables_from_excel(xlsx_path, sheet_name, m)
        if k_expected is None:
            k_expected = len(tabs)
        elif len(tabs) != k_expected:
            raise ValueError(f"Metric '{m}' has {len(tabs)} groups, expected {k_expected}.")
        metric_tables[m] = tabs

    assert k_expected is not None
    k = k_expected

    if group_names is None or len(group_names) != k:
        group_names = [f"G{i+1}" for i in range(k)]

    # raw aggregated per metric per group
    raw: Dict[str, np.ndarray] = {}
    for m in metrics:
        dirn = directions.get(m, "min")
        raw[m] = np.array([_aggregate_value(t, dirn, aggregator) for t in metric_tables[m]], dtype=float)

    if combine_mode == "sum":
        total = np.zeros(k, dtype=float)
        any_used = np.zeros(k, dtype=bool)

        for m in metrics:
            w = float(weights.get(m, 1.0))
            dirn = directions.get(m, "min")
            norm = _normalize(raw[m], dirn)  # 0..1 (0=лучше)
            mask = np.isfinite(norm)
            total[mask] += w * norm[mask]
            any_used[mask] = True

        total[~any_used] = np.nan
        order = np.argsort(total)

        print(f"\n=== Итог (COMBINE_MODE=sum), AGGREGATOR={aggregator} ===")
        for idx in order:
            g = group_names[idx]
            sc = total[idx]
            print(f"{g}: score={'NaN' if not np.isfinite(sc) else f'{sc:.6f}'}")

    elif combine_mode == "rank":
        rank_sum = np.zeros(k, dtype=float)

        for m in metrics:
            w = float(weights.get(m, 1.0))
            dirn = directions.get(m, "min")
            v = raw[m].copy()

            nan_mask = ~np.isfinite(v)
            if np.all(nan_mask):
                continue

            if dirn == "min":
                sort_idx = np.argsort(np.where(nan_mask, np.inf, v))
            elif dirn == "max":
                sort_idx = np.argsort(np.where(nan_mask, -np.inf, -v))
            else:
                raise ValueError(f"Direction must be 'min' or 'max', got: {dirn}")

            ranks = np.empty(k, dtype=float)
            ranks[sort_idx] = np.arange(1, k + 1, dtype=float)
            rank_sum += w * ranks

        order = np.argsort(rank_sum)
        print(f"\n=== Итог (COMBINE_MODE=rank), AGGREGATOR={aggregator} ===")
        for idx in order:
            print(f"{group_names[idx]}: rank_score={rank_sum[idx]:.6f}")

    else:
        raise ValueError(f"Unknown COMBINE_MODE={combine_mode}")

    print("\n=== Детализация по метрикам ===")
    for m in metrics:
        dirn = directions.get(m, "min")
        w = float(weights.get(m, 1.0))
        vals = raw[m]
        print(f"\n[{m}] dir={dirn}, weight={w}, aggregator={aggregator}")
        for i in range(k):
            print(f"  {group_names[i]}: {vals[i]}")


if __name__ == "__main__":
    compare_groups_from_excel(
        xlsx_path=EXCEL_PATH,
        sheet_name=SHEET_NAME,
        metrics=METRICS,
        directions=DIRECTIONS,
        weights=WEIGHTS,
        aggregator=AGGREGATOR,
        combine_mode=COMBINE_MODE,
        group_names=GROUP_NAMES,
    )