from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Dict, List, Tuple, Optional, Iterable

import numpy as np
import pandas as pd

EXCEL_PATH = r"D:\results.xlsx"
SHEET_NAME: Optional[str] = "SWEEP_2"  # None = первый лист

PARAM1_NAME = "param1"
PARAM2_NAME = "param2"
SWAP_AXES = False

# --- 1) РЕЖИМ СКОРИНГА ---
# "metrics" -> считаем score по метрикам (как было)
# "super"   -> считаем score по супер-критериям
SCORE_MODE = "super"  # "metrics" | "super"

# --- 2) МЕТРИКИ, которые будем читать из Excel ---
# (можно читать больше, чем используется в score — они просто будут выводиться)
METRICS = [
    # экономика
    "LCOE, руб/кВт∙ч",

    # надежность
    "ENS,кВт∙ч",
    "LOLE_h",
    "ENS_evtN",
    "ENS_evtMaxH",

    # операционные
    "Расход топлива, тыс.тонн",
    "Моточасы, тыс.мч",
    "FailDg"
]

# Направления: 'min' (меньше лучше) или 'max' (больше лучше)
DIRECTIONS: Dict[str, str] = {m: "min" for m in METRICS}

# Веса метрик (актуально для SCORE_MODE="metrics" и для агрегирования супер-критерия, если method="wsum")
METRIC_WEIGHTS: Dict[str, float] = {m: 1.0 for m in METRICS}

# --- 3) СУПЕР-КРИТЕРИИ ---
# Выбираешь, какие метрики входят в какой критерий
SUPER_GROUPS: Dict[str, List[str]] = {
    "economy": ["LCOE, руб/кВт∙ч"],
    "reliability": ["ENS,кВт∙ч", "LOLE_h", "ENS_evtN", "ENS_evtMaxH"],
    "operations": ["Расход топлива, тыс.тонн", "Моточасы, тыс.мч"],
    # "operations": ["Моточасы, тыс.мч"],
}

# Как агрегировать метрики внутри каждого супер-критерия (по нормированным 0..1 значениям):
#   "wsum"  : взвешенная сумма (по SUPER_INNER_WEIGHTS или METRIC_WEIGHTS)
#   "mean"  : среднее
#   "max"   : худший показатель (worst-case) -> полезно для надежности
#   "pnorm" : p-норма (p задаётся в SUPER_P)
SUPER_AGG_METHOD: Dict[str, str] = {
    "economy": "wsum",
    "reliability": "wsum",   # <-- обычно так и надо, чтобы плохая надёжность не "компенсировалась"
    "operations": "mean",
}

# p для p-нормы (если где-то SUPER_AGG_METHOD="pnorm")
SUPER_P: Dict[str, float] = {"reliability": 4.0}

# Веса СУПЕР-критериев в итоговом score (актуально для SCORE_MODE="super")
SUPER_WEIGHTS: Dict[str, float] = {"economy": 1, "reliability": 1.0, "operations": 0.2}

# (опционально) Пороговые ограничения: если нарушено -> score=+inf
# Пример:
# CONSTRAINTS = {"LOLE_h": ("<=", 5.0), "ENS,кВт∙ч": ("<=", 1000.0)}
CONSTRAINTS: Dict[str, Tuple[str, float]] = {}

# =========================
# КОНЕЦ НАСТРОЕК
# =========================


def to_num(x) -> float:
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
        raise ValueError(f"Метрика '{metric_name}' найдена, но не удалось распознать блоки заголовков.")
    return TableMeta(metric_name, r, header_row, blocks)


def extract_block_matrix(raw: pd.DataFrame, meta: TableMeta, block: Tuple[int, int]) -> Tuple[np.ndarray, np.ndarray, np.ndarray]:
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


def _fmt(v: float) -> str:
    if not np.isfinite(v):
        return "NaN"
    av = abs(v)
    if av != 0 and (av < 1e-3 or av >= 1e4):
        return f"{v:.6e}"
    return f"{v:.6f}".rstrip("0").rstrip(".")


def _check_constraints(cell_vals: Dict[str, float]) -> bool:
    # True if OK, False if violates any constraint
    for m, (op, thr) in CONSTRAINTS.items():
        v = cell_vals.get(m, np.nan)
        if not np.isfinite(v):
            return False
        if op == "<=" and not (v <= thr):
            return False
        if op == "<" and not (v < thr):
            return False
        if op == ">=" and not (v >= thr):
            return False
        if op == ">" and not (v > thr):
            return False
        if op == "==" and not (v == thr):
            return False
        if op == "!=" and not (v != thr):
            return False
        if op not in ("<=", "<", ">=", ">", "==", "!="):
            raise ValueError(f"Unknown constraint operator: {op} (metric {m})")
    return True


def _super_aggregate(group_name: str, norm_grids: Dict[str, np.ndarray]) -> np.ndarray:
    metrics = SUPER_GROUPS[group_name]
    method = SUPER_AGG_METHOD.get(group_name, "mean").lower()

    grids = [norm_grids[m] for m in metrics]

    if method == "mean":
        return np.nanmean(np.stack(grids, axis=0), axis=0)

    if method == "max":
        return np.nanmax(np.stack(grids, axis=0), axis=0)

    if method == "pnorm":
        p = float(SUPER_P.get(group_name, 2.0))
        a = np.stack(grids, axis=0)
        return (np.nanmean(np.power(a, p), axis=0)) ** (1.0 / p)

    if method == "wsum":
        w = np.array([float(METRIC_WEIGHTS.get(m, 1.0)) for m in metrics], dtype=float)
        if np.all(w == 0):
            w = np.ones_like(w)
        w = w / np.sum(w)
        a = np.stack(grids, axis=0)
        return np.nansum(a * w[:, None, None], axis=0)

    raise ValueError(f"Unknown SUPER_AGG_METHOD for {group_name}: {method}")


def score_grid_metrics(metric_grids: Dict[str, np.ndarray]) -> np.ndarray:
    # нормирование внутри сценария по метрике
    norm_grids: Dict[str, np.ndarray] = {}
    for m, grid in metric_grids.items():
        gn = normalize_minmax(grid.astype(float))
        if DIRECTIONS.get(m, "min").lower() == "max":
            gn = 1.0 - gn
        norm_grids[m] = gn

    # сумма по метрикам
    score = None
    for m in METRICS:
        w = float(METRIC_WEIGHTS.get(m, 1.0))
        g = norm_grids[m]
        score = (w * g) if score is None else (score + w * g)
    return score


def score_grid_super(metric_grids: Dict[str, np.ndarray]) -> Tuple[np.ndarray, Dict[str, np.ndarray]]:
    # нормирование внутри сценария по каждой метрике
    norm_grids: Dict[str, np.ndarray] = {}
    for m, grid in metric_grids.items():
        gn = normalize_minmax(grid.astype(float))
        if DIRECTIONS.get(m, "min").lower() == "max":
            gn = 1.0 - gn
        norm_grids[m] = gn

    # супер-критерии
    super_grids: Dict[str, np.ndarray] = {}
    for gname in SUPER_GROUPS.keys():
        # пропускаем супер-группу, если там пусто
        if not SUPER_GROUPS[gname]:
            continue
        # проверка, что метрики существуют
        for m in SUPER_GROUPS[gname]:
            if m not in norm_grids:
                raise ValueError(f"Metric '{m}' is used in SUPER_GROUPS['{gname}'] but not present in METRICS.")
        super_grids[gname] = _super_aggregate(gname, norm_grids)

    # итоговый score по супер-критериям
    score = None
    for gname, grid in super_grids.items():
        w = float(SUPER_WEIGHTS.get(gname, 1.0))
        score = (w * grid) if score is None else (score + w * grid)

    if score is None:
        raise ValueError("No super-criteria produced. Check SUPER_GROUPS.")
    return score, super_grids


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
            _, _, vals = extract_block_matrix(raw, meta, block)
            if vals.shape != (len(row_labels), len(col_labels)):
                raise ValueError(
                    f"Сценарий {scen_idx}, метрика '{meta.metric_name}': размер {vals.shape}, "
                    f"ожидалось {(len(row_labels), len(col_labels))}"
                )
            metric_grids[meta.metric_name] = vals

        if SWAP_AXES:
            row_labels, col_labels = col_labels, row_labels
            metric_grids = {m: v.T for m, v in metric_grids.items()}

        # --- score ---
        if SCORE_MODE == "metrics":
            S = score_grid_metrics(metric_grids)
            super_grids = {}
        elif SCORE_MODE == "super":
            S, super_grids = score_grid_super(metric_grids)
        else:
            raise ValueError("SCORE_MODE must be 'metrics' or 'super'")

        # --- constraints ---
        # если есть ограничения, запрещаем ячейки, которые не проходят
        if CONSTRAINTS:
            mask_bad = np.zeros_like(S, dtype=bool)
            for rr in range(S.shape[0]):
                for cc in range(S.shape[1]):
                    cell_vals = {m: float(metric_grids[m][rr, cc]) for m in METRICS}
                    if not _check_constraints(cell_vals):
                        mask_bad[rr, cc] = True
            S = S.copy()
            S[mask_bad] = np.inf

        k = int(np.nanargmin(S))
        r = k // S.shape[1]
        c = k % S.shape[1]

        rec = {
            "scenario": scen_idx,
            PARAM1_NAME: float(row_labels[r]),
            PARAM2_NAME: float(col_labels[c]),
            "Score": float(S[r, c]),
            "ScoreMode": SCORE_MODE,
        }

        # добавить значения супер-критериев для выбранной точки (если режим super)
        for gname, grid in super_grids.items():
            rec[f"SC_{gname}"] = float(grid[r, c])

        # добавить исходные значения метрик в выбранной точке
        for m in METRICS:
            rec[m] = float(metric_grids[m][r, c])

        best_rows.append(rec)

    best = pd.DataFrame(best_rows)

    # ---------- PRINT ----------
    base_cols = ["scenario", PARAM1_NAME, PARAM2_NAME, "Score", "ScoreMode"]
    sc_cols = [c for c in best.columns if c.startswith("SC_")]
    metric_cols = METRICS

    cols = base_cols + sc_cols + metric_cols

    print("\n=== Best point per scenario (min Score) ===")
    print(" | ".join(cols))
    for _, row in best[cols].iterrows():
        parts = [
            f"{int(row['scenario'])}",
            _fmt(row[PARAM1_NAME]),
            _fmt(row[PARAM2_NAME]),
            _fmt(row["Score"]),
            str(row["ScoreMode"]),
        ]
        # супер-критерии
        for c in sc_cols:
            parts.append(_fmt(row[c]))
        # метрики
        for m in METRICS:
            parts.append(_fmt(row[m]))
        print(" | ".join(parts))

    # IQR и частоты по параметрам (всё равно полезно)
    def q(col, p):
        x = best[col].dropna().astype(float)
        return float(x.quantile(p)) if len(x) else float("nan")

    p1_q25, p1_q50, p1_q75 = q(PARAM1_NAME, 0.25), q(PARAM1_NAME, 0.50), q(PARAM1_NAME, 0.75)
    p2_q25, p2_q50, p2_q75 = q(PARAM2_NAME, 0.25), q(PARAM2_NAME, 0.50), q(PARAM2_NAME, 0.75)

    freq_p1 = best[PARAM1_NAME].value_counts()
    freq_p2 = best[PARAM2_NAME].value_counts()

    print("\n=== Robust IQR of selected params ===")
    print(f"{PARAM1_NAME}: q25={_fmt(p1_q25)}  median={_fmt(p1_q50)}  q75={_fmt(p1_q75)}")
    print(f"{PARAM2_NAME}: q25={_fmt(p2_q25)}  median={_fmt(p2_q50)}  q75={_fmt(p2_q75)}")

    print("\n=== Frequency (selected params) ===")
    print(f"[{PARAM1_NAME}]")
    for v, cnt in freq_p1.items():
        print(f"  {_fmt(float(v))}: {int(cnt)}")
    print(f"[{PARAM2_NAME}]")
    for v, cnt in freq_p2.items():
        print(f"  {_fmt(float(v))}: {int(cnt)}")


if __name__ == "__main__":
    run()