import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

# -----------------------------
# INPUT DATA (оставляем как есть)
# -----------------------------

min_by_current = pd.DataFrame(
    {
        ("50%", "SS"): [33.40, 33.40, 33.40, 33.40, 33.40],
        ("50%", "D") : [33.13, 33.13, 33.13, 33.13, 33.13],
        ("100%","SS"): [32.59, 32.59, 32.59, 32.59, 32.59],
        ("100%","D") : [32.38, 32.38, 32.38, 32.38, 32.38],
        ("200%","SS"): [32.18, 31.95, 31.74, 31.60, 31.60],
        ("200%","D") : [32.00, 31.67, 31.46, 31.38, 31.37],
        ("400%","SS"): [34.17, 32.13, 31.88, 31.77, 31.77],
        ("400%","D") : [34.02, 31.94, 31.71, 31.60, 31.60],
    }
)
min_by_current.columns = pd.MultiIndex.from_tuples(min_by_current.columns, names=["WT", "Scheme"])
min_by_current.index = [f"I{i+1}" for i in range(len(min_by_current))]

# -----------------------------
# "НАШЕ" ОФОРМЛЕНИЕ: двойные подписи X + разделители групп + чёрные линии
# -----------------------------

def _infer_double_labels_from_multiindex(df: pd.DataFrame, x_order, schemes):
    top_labels, bottom_labels = [], []
    for wt in x_order:
        for sch in schemes:
            top_labels.append(str(wt))
            bottom_labels.append(str(sch))
    return top_labels, bottom_labels

def plot_points_with_median_double_x(
    ax,
    df: pd.DataFrame,
    title: str,
    y_label: str = "LCOE, руб/кВт·ч",
    x_order=("50%","100%","200%","400%"),
    schemes=("SS","D"),
    jitter=0.06,
    show_median=True,
    show_minmax=True,
    # grouping:
    group_mode="fixed",   # "fixed" | "none"
    group_size=2,
    group_separators=True,
    # captions
    top_axis_name="Мощность ВЭУ",
    bottom_axis_name="Схема",
):
    # ---------- DATA ----------
    groups = []
    for wt in x_order:
        for sch in schemes:
            groups.append(df[(wt, sch)].dropna().values.astype(float))

    n = len(groups)
    x = np.arange(n, dtype=float)

    # ---------- POINTS ----------
    for i, vals in enumerate(groups):
        if len(vals) == 0:
            continue

        offs = np.linspace(-jitter, jitter, num=len(vals)) if len(vals) > 1 else np.array([0.0])
        ax.scatter(np.full_like(vals, x[i]) + offs, vals)

        if show_minmax and len(vals) > 1:
            ax.vlines(x[i], float(np.min(vals)), float(np.max(vals)), color="black")

        if show_median:
            med = float(np.median(vals))
            ax.hlines(med, x[i] - 0.18, x[i] + 0.18, color="black")

    # ---------- X LABELS ----------
    top_labels, bottom_labels = _infer_double_labels_from_multiindex(df, x_order, schemes)

    ax.set_xticks(x)
    ax.set_xticklabels(bottom_labels)

    ax_bottom2 = ax.secondary_xaxis("bottom")
    ax_bottom2.spines["bottom"].set_position(("outward", 22))

    centers = []
    uniq_top = []
    i = 0
    while i < len(top_labels):
        label = top_labels[i]
        j = i
        while j < len(top_labels) and top_labels[j] == label:
            j += 1
        centers.append((i + j - 1) / 2)
        uniq_top.append(label)
        i = j

    ax_bottom2.set_xticks(centers)
    ax_bottom2.set_xticklabels(uniq_top)

    # ---------- FORCE DRAW (CRITICAL) ----------
    ax.figure.canvas.draw()

    # ---------- CAPTIONS EXACTLY ON LABEL LINES ----------
    renderer = ax.figure.canvas.get_renderer()

    # bottom row (SS / D)
    lbls_bottom = ax.get_xticklabels()
    if lbls_bottom:
        bb = lbls_bottom[0].get_window_extent(renderer=renderer)
        y_bottom = ax.transData.inverted().transform((0, bb.y0))[1]

        ax.text(
            x[0] - 0.8, y_bottom,
            bottom_axis_name,
            ha="right", va="center"
        )

    # top row (50 / 100 / ...)
    lbls_top = ax_bottom2.get_xticklabels()
    if lbls_top:
        bb = lbls_top[0].get_window_extent(renderer=renderer)
        y_top = ax.transData.inverted().transform((0, bb.y0))[1]

        ax.text(
            x[0] - 0.8, y_top,
            top_axis_name,
            ha="right", va="center"
        )

    # ---------- GROUP SEPARATORS ----------
    if group_separators and group_mode != "none":
        lw = float(ax.spines["left"].get_linewidth() or 1.0)
        for k in range(group_size, n, group_size):
            ax.axvline(k - 0.5, color="black", linewidth=lw)

    # ---------- FINAL ----------
    ax.set_title(title)
    ax.set_ylabel(y_label)
    ax.grid(True, axis="y")

# -----------------------------
# ОДИН ГРАФИК
# -----------------------------

fig, ax = plt.subplots(figsize=(10, 4))
plot_points_with_median_double_x(
    ax,
    min_by_current,
    title="",
    y_label="LCOE, руб/кВт·ч",
    group_mode="fixed",
    group_size=2,          # по умолчанию 2
    group_separators=True,
    top_axis_name="",
    bottom_axis_name="",
)
plt.tight_layout(rect=(0, 0.18, 1, 1))  # место под 2 ряда подписей
plt.show()
