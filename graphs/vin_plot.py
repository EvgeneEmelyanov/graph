import os
import re
from dataclasses import dataclass
from typing import List, Optional, Tuple

import numpy as np
import matplotlib.pyplot as plt
from matplotlib.patches import Patch
from matplotlib.colors import to_rgb, LinearSegmentedColormap

from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries, get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

# =========================
# DEFAULTS (EDIT THESE ONCE)
# =========================

DEFAULT_EXCEL_PATH = r"D:\comb.xlsx"
DEFAULT_SHEET = "SWEEP_2"

DEFAULT_BLOCK_RANGE = "A2:BY13"
DEFAULT_GAP_ROWS = 3

DEFAULT_SCENARIO_LABELS: Optional[List[str]] = ["SS без ВР", "D без ВР", "SS ВР для II", "D ВР для II", "SS ВР", "D ВР"]  # legend labels
CELL_LABELS_SHORT: List[str] = ["SS-0", "D-0", "SS-II", "D-II", "SS-A", "D-A"]  # short names (independent of legend)

USE_DEFAULT_LABELS = True

# =========================
# WINNER MAP SETTINGS
# =========================

# Title template: parameter name will be substituted from Excel header cell above the block
TITLE_TEMPLATE = "Выигрышный сценарий (min {param})"

# "Уверенность" = Δ = (2-й лучший − лучший) по оптимизируемому параметру
SHOW_CONFIDENCE_COLORBAR = False
CONF_COLORBAR_LABEL_TEMPLATE = "Δ{param} = (2-й − 1-й), (чем больше, тем увереннее выбор)"

# Optional: near ties shown as gray
USE_NEAR_TIE_GRAY = True
NEAR_TIE_THRESHOLD = 0.01  # in parameter units (e.g., руб/кВт·ч)

# Layout
RIGHT_MARGIN = 0.90
FIGSIZE = (14, 6)

# Grid
SHOW_CELL_GRID = True
CELL_GRID_LW = 0.35
CELL_GRID_ALPHA = 0.35

# Ticks: show ALL labels on both axes, X rotated 90 deg
X_LABEL_ROTATION = 90
X_LABEL_FONTSIZE = 8
Y_LABEL_FONTSIZE = 8

# Space for rotated X labels
BOTTOM_MARGIN = 0.32

# =========================
# Cell text overlay
# =========================

DRAW_CELL_LABELS = True
CELL_LABEL_FONTSIZE = 7
CELL_LABEL_ALPHA = 0.95
CELL_LABEL_EVERY_N = 1  # 1=every cell, 2=every 2nd cell, ...

# =========================
# Data model
# =========================

@dataclass
class Table2D:
    x_labels: List[str]
    y_labels: List[str]
    values: np.ndarray


# =========================
# Parsing helpers (tabs)
# =========================

def _to_float_ru(s: str) -> float:
    s = str(s).strip().replace("\xa0", "").replace(" ", "").replace(",", ".")
    return float(s)

def _split_blocks_by_tabs(line: str) -> List[str]:
    return [b for b in re.split(r"\t{2,}", line.rstrip("\n")) if b.strip()]

def _split_cells_by_tabs(block: str) -> List[str]:
    return [c for c in re.split(r"\t+", block.strip()) if c.strip()]

def parse_tables_from_paste(paste_text: str) -> List[Table2D]:
    lines = [ln for ln in paste_text.splitlines() if ln.strip()]
    if not lines:
        raise ValueError("Empty input.")

    header_blocks = _split_blocks_by_tabs(lines[0])
    if not header_blocks:
        raise ValueError("Header line could not be split into table blocks. Keep tabs in the input.")

    x_labels_per_table: List[List[str]] = []
    for hb in header_blocks:
        cells = _split_cells_by_tabs(hb)
        x_labels_per_table.append([c.strip() for c in cells])

    n_tables = len(x_labels_per_table)
    n_cols = len(x_labels_per_table[0])
    if any(len(x) != n_cols for x in x_labels_per_table):
        raise ValueError("Not all tables have the same number of X columns.")

    y_labels: List[str] = []
    values_per_table: List[List[List[float]]] = [[] for _ in range(n_tables)]

    for ln in lines[1:]:
        blocks = _split_blocks_by_tabs(ln)
        if len(blocks) != n_tables:
            raise ValueError(
                f"Line has {len(blocks)} blocks, expected {n_tables}. "
                f"Make sure table gaps are preserved as multiple TABs."
            )

        row_y: Optional[str] = None
        for ti, blk in enumerate(blocks):
            cells = _split_cells_by_tabs(blk)
            if len(cells) != (1 + n_cols):
                raise ValueError(
                    f"Table {ti+1}: row has {len(cells)} cells, expected {1+n_cols} (y + {n_cols} values)."
                )

            yv = cells[0].strip()
            vals = [_to_float_ru(c) for c in cells[1:]]

            if row_y is None:
                row_y = yv
            elif row_y != yv:
                raise ValueError("Y labels differ across tables on the same row; input misaligned.")

            values_per_table[ti].append(vals)

        y_labels.append(str(row_y))

    tables: List[Table2D] = []
    for ti in range(n_tables):
        arr = np.array(values_per_table[ti], dtype=float)
        tables.append(Table2D(x_labels=x_labels_per_table[ti], y_labels=y_labels, values=arr))

    return tables


# =========================
# Excel reading
# =========================

def _block_range_for_graph_index(base_range: str, graph_index: int, gap_rows: int) -> Tuple[str, int, int, int, int]:
    if graph_index < 1:
        raise ValueError("graph_index must be >= 1")

    min_col, min_row, max_col, max_row = range_boundaries(base_range)
    height = (max_row - min_row + 1)

    shift = (graph_index - 1) * (height + gap_rows)
    new_min_row = min_row + shift
    new_max_row = max_row + shift

    a1 = f"{get_column_letter(min_col)}{new_min_row}:{get_column_letter(max_col)}{new_max_row}"
    return a1, min_col, new_min_row, max_col, new_max_row

def _cells_to_tabbed_text(ws: Worksheet, min_col: int, min_row: int, max_col: int, max_row: int) -> str:
    lines = []
    for r in range(min_row, max_row + 1):
        row_vals = []
        for c in range(min_col, max_col + 1):
            v = ws.cell(row=r, column=c).value
            row_vals.append("" if v is None else str(v))
        lines.append("\t".join(row_vals))
    return "\n".join(lines)

def parse_tables_from_excel(xlsx_path: str, sheet_name: str, a1_range: str) -> Tuple[List[Table2D], str]:
    """
    Returns (tables, param_name), where param_name is read from the cell directly above the block:
    (left column of block, row = min_row-1).
    Example:
      base range A2:AY13 -> param in A1
      2nd block -> range A17:AY28 -> param in A16
    """
    wb = load_workbook(xlsx_path, data_only=True)
    try:
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found in workbook.")
        ws = wb[sheet_name]

        min_col, min_row, max_col, max_row = range_boundaries(a1_range)

        title_row = min_row - 1
        title_col = min_col
        param_name = ""
        if title_row >= 1:
            v = ws.cell(row=title_row, column=title_col).value
            param_name = "" if v is None else str(v).strip()

        block_text = _cells_to_tabbed_text(ws, min_col, min_row, max_col, max_row)
        tables = parse_tables_from_paste(block_text)
        return tables, param_name
    finally:
        wb.close()


# =========================
# Winner-map computation
# =========================

def compute_winner_and_margin(tables: List[Table2D]) -> Tuple[np.ndarray, np.ndarray]:
    if not tables:
        raise ValueError("No tables provided.")

    ny = len(tables[0].y_labels)
    nx = len(tables[0].x_labels)
    for t in tables:
        if t.values.shape != (ny, nx):
            raise ValueError("All tables must have identical shape and grids.")

    stack = np.stack([t.values for t in tables], axis=0)  # [k, ny, nx]
    winner_idx = np.argmin(stack, axis=0)                 # [ny, nx]

    part = np.partition(stack, kth=1, axis=0)
    best = part[0, :, :]
    second = part[1, :, :]
    margin = second - best

    return winner_idx, margin


# =========================
# Plotting (NO LINES): hue = winner, saturation = confidence
# =========================

def _base_rgb_colors(k: int) -> List[np.ndarray]:
    base = plt.rcParams["axes.prop_cycle"].by_key().get("color", ["C0", "C1", "C2", "C3"])
    return [np.array(to_rgb(base[i % len(base)]), dtype=float) for i in range(k)]

def _confidence_cmap():
    return LinearSegmentedColormap.from_list(
        "conf_cmap",
        [(0.98, 0.98, 0.98), (0.35, 0.35, 0.35)]
    )

def _ideal_text_color(rgb: np.ndarray) -> str:
    r, g, b = float(rgb[0]), float(rgb[1]), float(rgb[2])
    lum = 0.2126 * r + 0.7152 * g + 0.0722 * b
    return "black" if lum > 0.6 else "white"

def plot_winner_map_confidence(
    tables: List[Table2D],
    scenario_labels: List[str],
    title_main: str,
    param_name: str,
):
    ref = tables[0]
    k = len(tables)

    winner_idx, margin = compute_winner_and_margin(tables)

    # Confidence normalization
    m = margin.copy()
    m[~np.isfinite(m)] = 0.0
    mmax = float(np.max(m)) if float(np.max(m)) > 0 else 1.0
    conf = np.clip(m / mmax, 0.0, 1.0)  # 0..1

    # RGB: white -> scenario color by conf
    ny, nx = winner_idx.shape
    img = np.ones((ny, nx, 3), dtype=float)
    colors = _base_rgb_colors(k)

    for i in range(k):
        mask = (winner_idx == i)
        if not np.any(mask):
            continue
        a = conf[mask][:, None]
        img[mask] = (1.0 - a) * 1.0 + a * colors[i]

    if USE_NEAR_TIE_GRAY:
        tie_mask = margin < float(NEAR_TIE_THRESHOLD)
        img[tie_mask] = np.array([0.92, 0.92, 0.92], dtype=float)

    fig, ax = plt.subplots(figsize=FIGSIZE)
    ax.imshow(img, origin="lower", aspect="auto", interpolation="nearest")

    # Grid aligned to cells
    if SHOW_CELL_GRID:
        ax.set_xticks(np.arange(-0.5, nx, 1), minor=True)
        ax.set_yticks(np.arange(-0.5, ny, 1), minor=True)
        ax.grid(which="minor", color="black", linewidth=CELL_GRID_LW, alpha=CELL_GRID_ALPHA)

    # Labels and title (use parameter name from Excel)
    ax.set_xlabel("Емкость АКБ")
    ax.set_ylabel("Доля потребителей I категории")
    ax.set_title(title_main)

    # Show ALL tick labels on both axes
    ax.set_xticks(np.arange(len(ref.x_labels)))
    ax.set_yticks(np.arange(len(ref.y_labels)))
    ax.set_xticklabels(ref.x_labels, rotation=X_LABEL_ROTATION, ha="center", va="top", fontsize=X_LABEL_FONTSIZE)
    ax.set_yticklabels(ref.y_labels, fontsize=Y_LABEL_FONTSIZE)

    ax.tick_params(axis="x", which="both", bottom=True, top=True, labelbottom=True, length=4)
    ax.tick_params(axis="y", which="both", left=True, right=True, labelleft=True, length=4)

    # Draw short scenario label inside each cell (optional)
    if DRAW_CELL_LABELS:
        if len(CELL_LABELS_SHORT) != k:
            raise ValueError(f"CELL_LABELS_SHORT has {len(CELL_LABELS_SHORT)} items, but there are {k} scenarios.")

        for y in range(0, ny, CELL_LABEL_EVERY_N):
            for x in range(0, nx, CELL_LABEL_EVERY_N):
                idx = int(winner_idx[y, x])
                txt = str(CELL_LABELS_SHORT[idx])

                cell_rgb = img[y, x, :]
                color_txt = _ideal_text_color(cell_rgb)

                if USE_NEAR_TIE_GRAY and margin[y, x] < float(NEAR_TIE_THRESHOLD):
                    color_txt = "black"

                ax.text(
                    x, y, txt,
                    ha="center", va="center",
                    fontsize=CELL_LABEL_FONTSIZE,
                    color=color_txt,
                    alpha=CELL_LABEL_ALPHA,
                    clip_on=True,
                )

    # Confidence colorbar (optional) — label includes param name
    if SHOW_CONFIDENCE_COLORBAR:
        sm = plt.cm.ScalarMappable(cmap=_confidence_cmap(), norm=plt.Normalize(vmin=0, vmax=mmax))
        cbar = fig.colorbar(sm, ax=ax, fraction=0.046, pad=0.02)
        cbar.set_label(CONF_COLORBAR_LABEL_TEMPLATE.format(param=param_name if param_name else "показатель"))

    # Legend — title includes parameter name
    legend_title = f"Сценарий с min {param_name}" if param_name else "Сценарий с min показателем"
    handles = [Patch(facecolor=colors[i], edgecolor="black", label=str(scenario_labels[i])) for i in range(k)]
    if USE_NEAR_TIE_GRAY:
        handles.append(Patch(facecolor=(0.92, 0.92, 0.92), edgecolor="black",
                             label=f"Почти равно (Δ<{NEAR_TIE_THRESHOLD:g})"))

    ax.legend(
        handles=handles,
        title=legend_title,
        loc="upper left",
        bbox_to_anchor=(1.22, 1.0) if SHOW_CONFIDENCE_COLORBAR else (1.02, 1.0),
        borderaxespad=0.0,
    )

    plt.tight_layout(rect=(0, 0, RIGHT_MARGIN, 1))
    plt.subplots_adjust(bottom=BOTTOM_MARGIN)

    return fig


# =========================
# Main
# =========================

if __name__ == "__main__":
    xlsx_path = DEFAULT_EXCEL_PATH
    sheet_name = DEFAULT_SHEET

    graph_idx_s = input("Graph index (1..N) (default 1): ").strip()
    graph_idx = int(graph_idx_s) if graph_idx_s else 1
    if graph_idx < 1:
        raise ValueError("Graph index must be >= 1")

    block_range, _, start_row, _, _ = _block_range_for_graph_index(
        DEFAULT_BLOCK_RANGE, graph_idx, DEFAULT_GAP_ROWS
    )

    print(f"Using block range: {block_range}")

    tables, param_name = parse_tables_from_excel(xlsx_path, sheet_name, block_range)

    labels = DEFAULT_SCENARIO_LABELS if USE_DEFAULT_LABELS else None
    if labels is None:
        labels = [f"S{i+1}" for i in range(len(tables))]
    if len(labels) != len(tables):
        raise ValueError(
            f"DEFAULT_SCENARIO_LABELS has {len(labels)} items, but block contains {len(tables)} tables."
        )

    # Title uses parameter name from Excel
    pname = param_name if param_name else "показатель"
    title = TITLE_TEMPLATE.format(param=pname)

    fig = plot_winner_map_confidence(
        tables=tables,
        scenario_labels=labels,
        title_main=title,
        param_name=pname,
    )

    # optional save
    out_dir = r"D:\10_results\plots"
    os.makedirs(out_dir, exist_ok=True)
    fig_path = os.path.join(out_dir, f"winner_map_{graph_idx}.png")
    fig.savefig(fig_path, dpi=300)

    plt.show()
