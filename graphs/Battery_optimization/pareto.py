import csv
from pathlib import Path

import matplotlib.pyplot as plt
from matplotlib.lines import Line2D
from matplotlib.ticker import FuncFormatter, MultipleLocator


# ============================================================
# НАСТРОЙКИ
# ============================================================
OUTPUT_PNG = r"D:\adaptive_tune_pareto_multi.png"
FIG_DPI = 300

SHOW_ONLY_FEASIBLE_FOR_PRIMARY = False
SHOW_LABELS_FOR_PARETO = True
SHOW_BASELINE_LABELS = True
MULTI_FRONT_MODE = True

# Если True, Pareto-фронт строится только по feasible-точкам
# Если False, Pareto-фронт строится по всем точкам
PARETO_ONLY_ON_FEASIBLE = True

CSV_FILE = r"C:\Users\Balt_\Desktop\Диссер\01_Результаты моделирования\04_Оптимизация весов _2С\25-300\125_adaptive_tune.csv"

# Время моделирования для перевода LOLH -> LOLP
SIM_HOURS = 175320.0

# Фиксированный порядок для оси Y:
# подписи будут 1,0 ... 10,0 и т.д., а общий множитель будет указан в названии оси
Y_EXPONENT = -5
Y_TICK_STEP = 1e-5

# Мягкие, приглушённые, но разные цвета
SOFT_COLORS = [
    "#426f91",  # мягкий синий
    "#c95f46",  # мягкий терракотовый
    "#6f8f72",  # мягкий зелёный
    "#8a6f9e",  # мягкий фиолетовый
    "#b2875a",  # мягкий охристый
    "#5f8f96",  # мягкий бирюзовый
    "#b06c7a",  # мягкий розово-коричневый
    "#7d7f8c",  # мягкий серо-синий
    "#9a845f",  # мягкий песочный
    "#5d7a66",  # мягкий оливковый
]

FRONTS = [
    {
        "label": "ВЭУ 100%",
        "csv": r"C:\Users\Balt_\Desktop\Диссер\01_Результаты моделирования\04_Оптимизация весов _2С\25-300\100_adaptive_tune.csv",
        "color": SOFT_COLORS[1],
    },
    {
        "label": "ВЭУ 150%",
        "csv": r"C:\Users\Balt_\Desktop\Диссер\01_Результаты моделирования\04_Оптимизация весов _2С\25-300\150_adaptive_tune.csv",
        "color": SOFT_COLORS[2],
    },
    {
        "label": "ВЭУ 200%",
        "csv": r"C:\Users\Balt_\Desktop\Диссер\01_Результаты моделирования\04_Оптимизация весов _2С\25-300\200_adaptive_tune.csv",
        "color": SOFT_COLORS[3],
    },
    {
        "label": "ВЭУ 250%",
        "csv": r"C:\Users\Balt_\Desktop\Диссер\01_Результаты моделирования\04_Оптимизация весов _2С\25-300\250_adaptive_tune.csv",
        "color": SOFT_COLORS[4],
    },
    {
        "label": "ВЭУ 300%",
        "csv": r"C:\Users\Balt_\Desktop\Диссер\01_Результаты моделирования\04_Оптимизация весов _2С\25-300\300_adaptive_tune.csv",
        "color": SOFT_COLORS[5],
    },
]


# ============================================================
# ФОРМАТ ЧИСЕЛ
# ============================================================
def comma_formatter(x, pos):
    return f"{x:.1f}".replace(".", ",")


def to_superscript(n: int) -> str:
    sup_map = {
        "-": "⁻",
        "0": "⁰", "1": "¹", "2": "²", "3": "³",
        "4": "⁴", "5": "⁵", "6": "⁶",
        "7": "⁷", "8": "⁸", "9": "⁹",
    }
    return "".join(sup_map[c] for c in str(n))


def sci_1_decimal(value: float) -> str:
    """
    Формат x,x·10^-y
    """
    if value == 0:
        return "0,0·10⁰"

    exp = 0
    v = abs(value)

    while v >= 10.0:
        v /= 10.0
        exp += 1
    while v < 1.0:
        v *= 10.0
        exp -= 1

    mantissa = round(v, 1)

    if mantissa >= 10.0:
        mantissa = 1.0
        exp += 1

    sign = "-" if value < 0 else ""
    exp_str = to_superscript(exp)
    return f"{sign}{mantissa:.1f}·10{exp_str}".replace(".", ",")


def y_formatter_fixed_exp(x, pos):
    """
    Фиксированный порядок для оси Y:
    например, при Y_EXPONENT = -5 будут подписи 1,0 ... 10,0,
    а общий множитель будет указан в подписи оси.
    """
    if abs(x) < 1e-20:
        return "0"

    scaled = x / (10 ** Y_EXPONENT)
    return f"{scaled:.1f}".replace(".", ",")


# ============================================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ============================================================
def parse_float(s: str) -> float:
    return float(s.replace(",", "."))


def parse_bool(s: str) -> bool:
    return s.strip().lower() == "true"


def lolh_to_lolp(lolh_hours: float) -> float:
    return lolh_hours / SIM_HOURS


def load_csv(path: str):
    rows = []
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f, delimiter=";")
        for row in reader:
            lcoe = parse_float(row["LCOE"])
            lolh = parse_float(row["LOLH"])

            item = {
                "LCOE": lcoe,
                "LOLH": lolh,
                "LOLP_FROM_LOLH": lolh_to_lolp(lolh),
                "isPareto": parse_bool(row["isPareto"]),
                "isFeasible": parse_bool(row["isFeasibleForPrimaryObjective"]),
            }

            if "lcoeNorm" in row and row["lcoeNorm"].strip():
                item["lcoeNorm"] = parse_float(row["lcoeNorm"])
            if "loleNorm" in row and row["loleNorm"].strip():
                item["loleNorm"] = parse_float(row["loleNorm"])

            rows.append(item)
    return rows


def dominates(a, b):
    no_worse_lcoe = a["LCOE"] <= b["LCOE"]
    no_worse_lolp = a["LOLP_FROM_LOLH"] <= b["LOLP_FROM_LOLH"]
    strictly_better = (a["LCOE"] < b["LCOE"]) or (a["LOLP_FROM_LOLH"] < b["LOLP_FROM_LOLH"])
    return no_worse_lcoe and no_worse_lolp and strictly_better


def compute_pareto(rows):
    out = []
    for i, candidate in enumerate(rows):
        dominated = False
        for j, other in enumerate(rows):
            if i == j:
                continue
            if dominates(other, candidate):
                dominated = True
                break
        if not dominated:
            out.append(candidate)

    out.sort(key=lambda r: (r["LCOE"], r["LOLP_FROM_LOLH"]))
    return out


def infer_baseline_from_rows(rows):
    for r in rows:
        lcoe_norm = r.get("lcoeNorm")
        lole_norm = r.get("loleNorm")

        if lcoe_norm and lole_norm and lcoe_norm > 0.0 and lole_norm > 0.0:
            baseline_lcoe = r["LCOE"] / lcoe_norm
            baseline_lolh = r["LOLH"] / lole_norm
            return baseline_lcoe, baseline_lolh

    return None, None


def prepare_fronts():
    if MULTI_FRONT_MODE:
        raw_fronts = FRONTS
    else:
        raw_fronts = [
            {
                "label": Path(CSV_FILE).stem,
                "csv": CSV_FILE,
                "color": SOFT_COLORS[0],
            }
        ]

    prepared = []
    for front in raw_fronts:
        all_rows = load_csv(front["csv"])

        plot_rows = all_rows
        if SHOW_ONLY_FEASIBLE_FOR_PRIMARY:
            plot_rows = [r for r in all_rows if r["isFeasible"]]

        pareto_source_rows = plot_rows
        if PARETO_ONLY_ON_FEASIBLE:
            pareto_source_rows = [r for r in plot_rows if r["isFeasible"]]

        pareto = compute_pareto(pareto_source_rows)

        baseline_lcoe = front.get("baseline_lcoe")
        baseline_lolh = front.get("baseline_lolh")

        if baseline_lcoe is None or baseline_lolh is None:
            inferred_lcoe, inferred_lolh = infer_baseline_from_rows(all_rows)
            if baseline_lcoe is None:
                baseline_lcoe = inferred_lcoe
            if baseline_lolh is None:
                baseline_lolh = inferred_lolh

        baseline_lolp = lolh_to_lolp(baseline_lolh) if baseline_lolh is not None else None

        prepared.append({
            "label": front["label"],
            "csv": front["csv"],
            "color": front["color"],
            "rows": plot_rows,
            "pareto": pareto,
            "baseline_lcoe": baseline_lcoe,
            "baseline_lolh": baseline_lolh,
            "baseline_lolp": baseline_lolp,
        })

    return prepared


# ============================================================
# ОСНОВНОЙ КОД
# ============================================================
fronts = prepare_fronts()

if not fronts:
    raise ValueError("Нет наборов данных для построения.")

if all(len(f["rows"]) == 0 for f in fronts):
    raise ValueError("Во всех CSV нет данных для построения.")

fig, ax = plt.subplots(figsize=(11, 8))

# Формат осей
ax.xaxis.set_major_formatter(FuncFormatter(comma_formatter))
ax.yaxis.set_major_formatter(FuncFormatter(y_formatter_fixed_exp))
ax.yaxis.set_major_locator(MultipleLocator(Y_TICK_STEP))

single_mode = not MULTI_FRONT_MODE
show_pareto_labels_now = SHOW_LABELS_FOR_PARETO and single_mode
show_baseline_labels_now = SHOW_BASELINE_LABELS and single_mode

for front in fronts:
    rows = front["rows"]
    pareto = front["pareto"]
    color = front["color"]

    if not rows:
        continue

    x_all = [r["LCOE"] for r in rows]
    y_all = [r["LOLP_FROM_LOLH"] for r in rows]

    x_p = [r["LCOE"] for r in pareto]
    y_p = [r["LOLP_FROM_LOLH"] for r in pareto]

    # Все решения
    ax.scatter(
        x_all,
        y_all,
        s=42,
        alpha=0.22,
        color=color
    )

    # Pareto-фронт
    if pareto:
        ax.plot(
            x_p,
            y_p,
            marker="o",
            markersize=6,
            linewidth=2.5,
            color=color
        )

    # Подписи Pareto-точек только в single mode
    if show_pareto_labels_now:
        for i, r in enumerate(pareto, start=1):
            txt = f"#{i}\n{r['LCOE']:.2f} / {sci_1_decimal(r['LOLP_FROM_LOLH'])}".replace(".", ",")
            ax.annotate(
                txt,
                (r["LCOE"], r["LOLP_FROM_LOLH"]),
                textcoords="offset points",
                xytext=(6, 6),
                fontsize=8,
                color=color
            )

    baseline_lcoe = front["baseline_lcoe"]
    baseline_lolp = front["baseline_lolp"]

    if baseline_lcoe is not None and baseline_lolp is not None:
        ax.scatter(
            [baseline_lcoe],
            [baseline_lolp],
            s=130,
            marker="X",
            color=color,
            edgecolors="black",
            linewidths=0.8
        )

        # Подпись baseline только в single mode
        if show_baseline_labels_now:
            base_txt = f"baseline\n{baseline_lcoe:.2f} / {sci_1_decimal(baseline_lolp)}".replace(".", ",")
            ax.annotate(
                base_txt,
                (baseline_lcoe, baseline_lolp),
                textcoords="offset points",
                xytext=(8, -10),
                fontsize=8,
                color=color
            )

ax.set_xlabel("LCOE, руб/кВт·ч", fontsize=16)
ax.set_ylabel(f"LOLP · 10{to_superscript(Y_EXPONENT)}", fontsize=16)
ax.tick_params(axis="both", labelsize=16)
ax.grid(True, alpha=0.3)

# ============================================================
# ДВЕ ЛЕГЕНДЫ
# ============================================================
style_handles = [
    Line2D([0], [0], color="black", linewidth=2.5, marker="o", markersize=6, label="Парето-фронт"),
    Line2D([0], [0], color="black", linestyle="None", marker="o", markersize=6, alpha=0.35, label="Все решения"),
    Line2D([0], [0], color="black", linestyle="None", marker="X", markersize=10, label="Базовый вариант"),
]

color_handles = []
for front in fronts:
    color_handles.append(
        Line2D([0], [0], color=front["color"], linewidth=2.5, label=front["label"])
    )

legend1 = ax.legend(
    handles=style_handles,
    loc="upper left",
    bbox_to_anchor=(0.69, 0.99),
    fontsize=14,
    frameon=True
)
ax.add_artist(legend1)

legend2 = ax.legend(
    handles=color_handles,
    loc="upper right",
    bbox_to_anchor=(0.89, 0.85),
    fontsize=14,
    frameon=True
)

fig.tight_layout()

output_path = Path(OUTPUT_PNG)
output_path.parent.mkdir(parents=True, exist_ok=True)
fig.savefig(output_path, dpi=FIG_DPI, bbox_inches="tight")
plt.show()

print(f"Готово: {output_path}")