import matplotlib
matplotlib.use("TkAgg")

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.widgets import Slider


FILE_PATH = r"D:\adapt_trace.xlsx"
SHEET_NAME = "TRACE"
LIMIT_ROWS = 8760
INITIAL_WINDOW = 720
MIN_WINDOW = 24
MAX_WINDOW = 8760

MAIN_LINE_WIDTH = 1.2
BATTERY_LINE_WIDTH = MAIN_LINE_WIDTH * 2.0


def load_data_from_xlsx(file_path, sheet_name, limit_rows=None):
    df = pd.read_excel(
        file_path,
        sheet_name=sheet_name,
        usecols="E,G,L",
        skiprows=1,
        engine="openpyxl"
    )

    df.columns = ["B1_L", "B1_W", "B1_B"]

    df = df.dropna(how="all").copy()

    for col in ["B1_L", "B1_W", "B1_B"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df = df.dropna(subset=["B1_L", "B1_W"]).copy()
    df["B1_B"] = df["B1_B"].fillna(0.0)

    if limit_rows is not None:
        df = df.head(limit_rows).copy()

    df["power_diff"] = df["B1_L"] - df["B1_W"]
    df["interval"] = np.arange(1, len(df) + 1)

    return df


def masked_line(y, mask):
    result = np.array(y, dtype=float, copy=True)
    result[~mask] = np.nan
    return result


def plot_with_two_sliders(df):
    x = df["interval"].to_numpy()
    y = df["power_diff"].to_numpy()
    battery = df["B1_B"].to_numpy()

    n = len(df)
    if n == 0:
        raise ValueError("Нет данных для построения графика")

    max_window = min(MAX_WINDOW, n)
    initial_window = min(INITIAL_WINDOW, n)
    min_window = min(MIN_WINDOW, n)

    plt.rcParams["toolbar"] = "toolbar2"

    fig, ax = plt.subplots(figsize=(16, 8))
    plt.subplots_adjust(bottom=0.22)

    line_main, = ax.plot([], [], linewidth=MAIN_LINE_WIDTH, label="B1_L - B1_W")
    line_discharge, = ax.plot([], [], linewidth=BATTERY_LINE_WIDTH, color="orange", label="АКБ разряд")
    line_charge, = ax.plot([], [], linewidth=BATTERY_LINE_WIDTH, color="green", label="АКБ заряд")

    ax.set_title("Разница мощности B1_L - B1_W с режимами АКБ")
    ax.set_xlabel("Интервал")
    ax.set_ylabel("Мощность")
    ax.grid(True, alpha=0.3)
    ax.legend(loc="upper right")

    ax_scale = plt.axes([0.12, 0.10, 0.76, 0.03])
    ax_shift = plt.axes([0.12, 0.05, 0.76, 0.03])

    slider_scale = Slider(
        ax=ax_scale,
        label="Масштаб окна",
        valmin=min_window,
        valmax=max_window,
        valinit=initial_window,
        valstep=1
    )

    slider_shift = Slider(
        ax=ax_shift,
        label="Сдвиг",
        valmin=0,
        valmax=max(0, n - min_window),
        valinit=0,
        valstep=1
    )

    def update(_=None):
        window = int(slider_scale.val)
        start = int(slider_shift.val)

        max_start = max(0, n - window)
        start = min(start, max_start)
        end = start + window

        xs = x[start:end]
        ys = y[start:end]
        bs = battery[start:end]

        if len(xs) == 0:
            return

        discharge_local = bs > 0
        charge_local = bs < 0

        ys_discharge = masked_line(ys, discharge_local)
        ys_charge = masked_line(ys, charge_local)

        line_main.set_data(xs, ys)
        line_discharge.set_data(xs, ys_discharge)
        line_charge.set_data(xs, ys_charge)

        if len(xs) > 1:
            ax.set_xlim(xs[0], xs[-1])
        else:
            ax.set_xlim(xs[0], xs[0] + 1)

        y_min = np.nanmin(ys)
        y_max = np.nanmax(ys)

        if np.isfinite(y_min) and np.isfinite(y_max):
            pad = max((y_max - y_min) * 0.08, 1e-6)
            ax.set_ylim(y_min - pad, y_max + pad)

        fig.canvas.draw_idle()

    def on_scale_change(val):
        window = int(val)
        max_start = max(0, n - window)

        if slider_shift.val > max_start:
            slider_shift.set_val(max_start)
        else:
            update()

    def on_shift_change(val):
        update()

    slider_scale.on_changed(on_scale_change)
    slider_shift.on_changed(on_shift_change)

    update()
    plt.show()


def main():
    df = load_data_from_xlsx(FILE_PATH, SHEET_NAME, LIMIT_ROWS)
    print(f"Загружено строк: {len(df)}")
    plot_with_two_sliders(df)


if __name__ == "__main__":
    main()