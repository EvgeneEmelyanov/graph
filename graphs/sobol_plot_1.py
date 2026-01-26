import re
import numpy as np
import matplotlib.pyplot as plt

LINE_RE = re.compile(
    r"^\s*(?P<name>[A-Z0-9_]+)\s+S=(?P<S>[-+]?(\d+(\.\d*)?|\.\d+)([eE][-+]?\d+)?)\s+ST=(?P<ST>[-+]?(\d+(\.\d*)?|\.\d+)([eE][-+]?\d+)?)\s*$"
)

def read_from_console():
    """
    Читает строки из консоли.
    Вставляешь данные, потом ОДНА пустая строка — ввод заканчивается.
    """
    print("Вставьте строки вида:")
    print("PARAM_NAME   S=...   ST=...")
    print("Пустая строка — конец ввода.\n")

    rows = []
    while True:
        try:
            line = input()
        except EOFError:
            break

        line = line.strip()
        if line == "":
            break

        m = LINE_RE.match(line)
        if not m:
            raise ValueError(f"Не удалось распарсить строку: {line}")

        rows.append((
            m.group("name"),
            float(m.group("S")),
            float(m.group("ST"))
        ))

    if not rows:
        raise ValueError("Не введено ни одной строки")

    return rows

def plot_sobol(rows, title="Sobol indices for LCOE"):
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

    # Ось X — номера 1..N
    ax.set_xticks(x)
    ax.set_xticklabels([str(i + 1) for i in range(n)])

    ax.legend()
    ax.grid(axis="y", alpha=0.3)

    ymax = max(S.max(), ST.max())
    ax.set_ylim(0, ymax * 1.2 if ymax > 0 else 1.0)

    # Подписи параметров над группами
    pad = ymax * 0.02 if ymax > 0 else 0.5
    for i, name in enumerate(names):
        ax.text(
            x[i] + 0.06,
            max(S[i], ST[i]) + pad,
            name,
            rotation=35,
            ha="left",
            va="bottom",
            fontsize=9
        )

    plt.tight_layout()
    plt.show()

if __name__ == "__main__":
    rows = read_from_console()
    plot_sobol(rows)
