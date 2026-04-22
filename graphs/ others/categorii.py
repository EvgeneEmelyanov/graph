import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
from mpl_toolkits.mplot3d import Axes3D  # noqa: F401


# ============================================================
# НАСТРОЙКИ
# ============================================================
FILES = {
    "ВЭУ 2": r"D:\res2.xlsx",
}

RAW_SHEET = "RAW"
OUTPUT = r"D:\3d_LCOE_LOLP_LPSP.png"

FIGSIZE = (11, 8)
DPI = 300

BASE_FONT = 11
FONT_SIZE = BASE_FONT + 4
TITLE_SIZE = BASE_FONT + 6
TICK_SIZE = BASE_FONT + 3
LEGEND_SIZE = BASE_FONT + 2

COLORS = [
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


# ============================================================
# ФОРМАТ ЧИСЕЛ
# ============================================================
def comma_formatter(decimals=2):
    def _fmt(x, pos):
        return f"{x:.{decimals}f}".replace(".", ",")
    return FuncFormatter(_fmt)


def sci_formatter_comma(decimals=1):
    def _fmt(x, pos):
        s = f"{x:.{decimals}e}"
        mantissa, exp = s.split("e")
        mantissa = mantissa.replace(".", ",")
        exp = int(exp)
        return f"{mantissa}e{exp}"
    return FuncFormatter(_fmt)


# ============================================================
# ЗАГРУЗКА
# ============================================================
def load_raw_points(path, sheet_name="RAW"):
    df = pd.read_excel(path, sheet_name=sheet_name, header=1)
    df = df.dropna(how="all").copy()
    df.columns = [str(c).strip() for c in df.columns]

    required_cols = ["param1", "param2", "LCOE, руб/кВт∙ч", "LOLP", "LPSP"]
    for col in required_cols:
        if col not in df.columns:
            raise KeyError(
                f"В файле {path} не найден столбец '{col}'. "
                f"Доступные столбцы: {df.columns.tolist()}"
            )

    for col in required_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df = df.dropna(subset=required_cols).copy()

    return df


# ============================================================
# ПОСТРОЕНИЕ 3D-ГРАФИКА
# ============================================================
def plot_3d_for_file(title, path):
    df = load_raw_points(path, RAW_SHEET)

    fig = plt.figure(figsize=FIGSIZE, dpi=DPI)
    ax = fig.add_subplot(111, projection="3d")

    unique_param1 = sorted(df["param1"].unique())

    color_map = {
        p1: COLORS[i % len(COLORS)]
        for i, p1 in enumerate(unique_param1)
    }

    for p1 in unique_param1:
        sub = df[df["param1"] == p1].copy()

        ax.scatter(
            sub["LCOE, руб/кВт∙ч"],
            sub["LOLP"],
            sub["LPSP"],
            s=55,
            color=color_map[p1],
            label=f"param1 = {str(p1).replace('.', ',')}"
        )

        # Подписи точек
        for _, row in sub.iterrows():
            x = row["LCOE, руб/кВт∙ч"]
            y = row["LOLP"]
            z = row["LPSP"]
            p2 = row["param2"]

            ax.text(
                x, y, z,
                f"p2={str(round(p2, 3)).replace('.', ',')}",
                fontsize=9
            )

    ax.set_title(title, fontsize=TITLE_SIZE, pad=18)

    ax.set_xlabel("LCOE, руб/кВт·ч", fontsize=FONT_SIZE, labelpad=12)
    ax.set_ylabel("LOLP", fontsize=FONT_SIZE, labelpad=12)
    ax.set_zlabel("LPSP", fontsize=FONT_SIZE, labelpad=12)

    ax.xaxis.set_major_formatter(comma_formatter(decimals=2))
    ax.yaxis.set_major_formatter(sci_formatter_comma(decimals=1))
    ax.zaxis.set_major_formatter(sci_formatter_comma(decimals=1))

    ax.tick_params(axis="both", labelsize=TICK_SIZE)
    ax.tick_params(axis="z", labelsize=TICK_SIZE)

    ax.legend(fontsize=LEGEND_SIZE, loc="best")

    plt.tight_layout()
    plt.savefig(OUTPUT, dpi=DPI, bbox_inches="tight")
    plt.show()


# ============================================================
# ОСНОВНОЙ БЛОК
# ============================================================
for title, path in FILES.items():
    plot_3d_for_file(title, path)