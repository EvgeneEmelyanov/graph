import csv
from pathlib import Path

import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter


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
PARETO_ONLY_ON_FEASIBLE = False

CSV_FILE = r"D:\1adaptive_tune.csv"

# Время моделирования для перевода LOLH -> LOLP
SIM_HOURS = 175320.0

FRONTS = [
    {
        "label": "run_1",
        "csv": r"D:\1adaptive_tune.csv",
        "color": "#426f91",
        # "baseline_lcoe": 28.501221,
        # "baseline_lolh": 9.666667,
    },
    {
        "label": "run_2",
        "csv": r"D:\2adaptive_tune.csv",
        "color": "#c95f46",
    },
]


# ============================================================
# ФОРМАТ ЧИСЕЛ
# ОСИ: 1 ЗНАК ПОСЛЕ ЗАПЯТОЙ
# ============================================================
def comma_formatter(x, pos):
    return f"{x:.1f}".replace(".", ",")


# ============================================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ============================================================
def parse_float(s: str) -> float:
    return float(s.replace(",", "."))


def parse_bool(s: str) -> bool:
    return s.strip().lower() == "true"


def lolh_to_lolp(lolh_hours: float) -> float:
    return lolh_hours / SIM_HOURS


def sci_1_decimal(value: float) -> str:
    """
    Формат x,x·10^-y
    """
    if value == 0:
        return "0,0·10^0"

    exp = 0
    v = abs(value)

    while v >= 10.0:
        v /= 10.0
        exp += 1
    while v < 1.0:
        v *= 10.0
        exp -= 1

    mantissa = round(v, 1)

    # если после округления получилось 10,0
    if mantissa >= 10.0:
        mantissa = 1.0
        exp += 1

    sign = "-" if value < 0 else ""
    return f"{sign}{mantissa:.1f}·10^{exp}".replace(".", ",")


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
                "color": "#426f91",
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

plt.figure(figsize=(11, 8))
ax = plt.gca()

ax.xaxis.set_major_formatter(FuncFormatter(comma_formatter))
def y_formatter(x, pos):
    if x == 0:
        return "0"

    # если число "обычное"
    if abs(x) >= 0.1:
        return f"{x:.1f}".replace(".", ",")

    # научная запись
    exp = 0
    v = abs(x)

    while v < 1.0:
        v *= 10.0
        exp -= 1

    mantissa = round(v, 1)

    if mantissa >= 10.0:
        mantissa = 1.0
        exp += 1

    sign = "-" if x < 0 else ""
    return f"{sign}{mantissa:.1f}·10^{exp}".replace(".", ",")


ax.yaxis.set_major_formatter(FuncFormatter(y_formatter))

single_mode = not MULTI_FRONT_MODE
show_pareto_labels_now = SHOW_LABELS_FOR_PARETO and single_mode
show_baseline_labels_now = SHOW_BASELINE_LABELS and single_mode

for front in fronts:
    rows = front["rows"]
    pareto = front["pareto"]
    color = front["color"]
    label = front["label"]

    if not rows:
        continue

    x_all = [r["LCOE"] for r in rows]
    y_all = [r["LOLP_FROM_LOLH"] for r in rows]

    x_p = [r["LCOE"] for r in pareto]
    y_p = [r["LOLP_FROM_LOLH"] for r in pareto]

    all_label = "Все решения" if single_mode else f"{label}: все решения"
    pareto_label = "Pareto" if single_mode else f"{label}: Pareto"
    baseline_label = "Базовый вариант" if single_mode else f"{label}: baseline"

    plt.scatter(
        x_all,
        y_all,
        s=42,
        alpha=0.22,
        color=color,
        label=all_label
    )

    if pareto:
        plt.plot(
            x_p,
            y_p,
            marker="o",
            markersize=6,
            linewidth=2.5,
            color=color,
            label=pareto_label
        )

    # Подписи Pareto-точек только в single mode
    if show_pareto_labels_now:
        for i, r in enumerate(pareto, start=1):
            txt = f"#{i}\n{r['LCOE']:.2f} / {sci_1_decimal(r['LOLP_FROM_LOLH'])}".replace(".", ",")
            plt.annotate(
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
        plt.scatter(
            [baseline_lcoe],
            [baseline_lolp],
            s=130,
            marker="X",
            color=color,
            edgecolors="black",
            linewidths=0.8,
            label=baseline_label
        )

        # Подпись baseline только в single mode
        if show_baseline_labels_now:
            base_txt = f"baseline\n{baseline_lcoe:.2f} / {sci_1_decimal(baseline_lolp)}".replace(".", ",")
            plt.annotate(
                base_txt,
                (baseline_lcoe, baseline_lolp),
                textcoords="offset points",
                xytext=(8, -10),
                fontsize=8,
                color=color
            )

plt.xlabel("LCOE, руб/кВт·ч")
plt.ylabel("LOLP")
plt.grid(True, alpha=0.3)
plt.legend()
plt.tight_layout()

output_path = Path(OUTPUT_PNG)
output_path.parent.mkdir(parents=True, exist_ok=True)
plt.savefig(output_path, dpi=FIG_DPI, bbox_inches="tight")
plt.show()

print(f"Готово: {output_path}")