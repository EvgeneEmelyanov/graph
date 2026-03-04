import re
import numpy as np
import matplotlib.pyplot as plt

# Формат строки (9 колонок):
# 1..3 : нагрузка шина1 (cat3, cat2, cat1)
# 4..6 : нагрузка шина2 (cat3, cat2, cat1)
# 7..8 : генерация на шине (шина1, шина2)
# 9    : OPEN / CLOSED / FAIL / FAIL2

def read_matrix_from_console(expected_numeric_cols=8):
    print("Вставьте таблицу (таб/пробел). В конце каждой строки: OPEN / CLOSED / FAIL / FAIL2.")
    print("Запятая в дробях допускается. Пустая строка завершает ввод.\n")

    lines = []
    while True:
        s = input()
        if not s.strip():
            break
        lines.append(s)

    if not lines:
        raise ValueError("Ввод пустой.")

    rows, states = [], []
    for i, line in enumerate(lines, start=1):
        tokens = re.split(r"[ \t]+", line.strip())
        tokens = [t for t in tokens if t.strip()]

        st = "OPEN"
        if tokens and tokens[-1].upper() in ("OPEN", "CLOSED", "FAIL", "FAIL2"):
            st = tokens[-1].upper()
            tokens = tokens[:-1]
        states.append(st)

        vals = []
        for t in tokens:
            t_num = t.replace(",", ".")
            try:
                vals.append(float(t_num))
            except ValueError:
                vals.append(0.0)

        if len(vals) != expected_numeric_cols:
            raise ValueError(
                f"Строка {i}: получено {len(vals)} чисел, ожидается {expected_numeric_cols} + (OPEN/CLOSED/FAIL/FAIL2).\n"
                f"Строка: {line}"
            )
        rows.append(vals)

    return np.array(rows, dtype=float), states


def plot_scenario(data, states):
    H = data.shape[0]
    hours = np.arange(1, H + 1)

    # Нагрузки (ввод: cat3, cat2, cat1)
    L1_c3, L1_c2, L1_c1 = data[:, 0], data[:, 1], data[:, 2]
    L2_c3, L2_c2, L2_c1 = data[:, 3], data[:, 4], data[:, 5]

    # Генерация на шине (нужна только для расчёта дефицита)
    G1, G2 = data[:, 6], data[:, 7]

    # Цвета
    bus1_ok = "tab:blue"
    bus2_ok = "tab:orange"
    agg_ok  = "tab:purple"   # CLOSED (общая шина)
    deficit_color = "lightgray"

    # Границы сегментов
    edge = "black"
    lw = 1.0

    # Геометрия
    w = 0.35  # ширина одного столбца
    x_shift = w / 2

    plt.figure(figsize=(14, 6))

    def split_by_served(cat1, cat2, cat3, served):
        s1 = min(cat1, served); served -= s1
        s2 = min(cat2, max(0.0, served)); served -= s2
        s3 = min(cat3, max(0.0, served))
        d1 = cat1 - s1
        d2 = cat2 - s2
        d3 = cat3 - s3
        return s1, s2, s3, d1, d2, d3

    def label_I_II_III(x, c1, c2, c3):
        if c1 > 0: plt.text(x, c1 / 2, "I", ha="center", va="center", fontsize=9)
        if c2 > 0: plt.text(x, c1 + c2 / 2, "II", ha="center", va="center", fontsize=9)
        if c3 > 0: plt.text(x, c1 + c2 + c3 / 2, "III", ha="center", va="center", fontsize=9)

    # Для легенды
    added = {"load1": False, "load2": False, "loadA": False, "def": False}

    states_u = [s.upper() for s in states]

    # Счётчики
    fail_seen = 0    # для FAIL: 1-й раз -> I, со 2-го -> I+II
    fail2_seen = 0   # для FAIL2: 1-й раз -> I, со 2-го -> I+II+III

    for idx, h in enumerate(hours):
        st = states_u[idx]

        c1_1, c2_1, c3_1 = float(L1_c1[idx]), float(L1_c2[idx]), float(L1_c3[idx])
        c1_2, c2_2, c3_2 = float(L2_c1[idx]), float(L2_c2[idx]), float(L2_c3[idx])
        g1, g2 = float(G1[idx]), float(G2[idx])

        split_I = split_II = split_III = False

        if st == "FAIL":
            fail_seen += 1
            split_I = True
            split_II = (fail_seen >= 2)
            split_III = False
        elif st == "FAIL2":
            fail2_seen += 1
            split_I = True
            split_II = (fail2_seen >= 2)
            split_III = (fail2_seen >= 2)

        if st == "CLOSED":
            # Один столбец по центру часа, шириной как обычный одиночный столбец
            x = h

            c1 = c1_1 + c1_2
            c2 = c2_1 + c2_2
            c3 = c3_1 + c3_2
            g = g1 + g2

            Ltot = c1 + c2 + c3
            served = min(Ltot, g)

            s1, s2, s3, d1, d2, d3 = split_by_served(c1, c2, c3, served)

            plt.bar(x, s1, width=w, color=agg_ok, edgecolor=edge, linewidth=lw,
                    label=("Нагрузка (CLOSED)" if not added["loadA"] else None))
            added["loadA"] = True
            plt.bar(x, s2, width=w, bottom=s1, color=agg_ok, edgecolor=edge, linewidth=lw)
            plt.bar(x, s3, width=w, bottom=s1 + s2, color=agg_ok, edgecolor=edge, linewidth=lw)

            plt.bar(x, d1, width=w, bottom=s1, color=deficit_color, edgecolor=edge, linewidth=lw,
                    label=("Дефицит" if not added["def"] else None))
            added["def"] = True
            plt.bar(x, d2, width=w, bottom=c1 + s2, color=deficit_color, edgecolor=edge, linewidth=lw)
            plt.bar(x, d3, width=w, bottom=c1 + c2 + s3, color=deficit_color, edgecolor=edge, linewidth=lw)

            label_I_II_III(x, c1, c2, c3)

        else:
            # Два столбца в час: левый/правый
            x1 = h - x_shift
            x2 = h + x_shift

            Ltot1 = c1_1 + c2_1 + c3_1
            Ltot2 = c1_2 + c2_2 + c3_2
            served1 = min(Ltot1, g1)
            served2 = min(Ltot2, g2)

            s11, s12, s13, d11, d12, d13 = split_by_served(c1_1, c2_1, c3_1, served1)
            s21, s22, s23, d21, d22, d23 = split_by_served(c1_2, c2_2, c3_2, served2)

            def draw_hsplit(x, height, bottom, color_low, color_high, label=None):
                plt.bar(x, height * 0.5, width=w, bottom=bottom, color=color_low,
                        edgecolor=edge, linewidth=lw, label=label)
                plt.bar(x, height * 0.5, width=w, bottom=bottom + height * 0.5, color=color_high,
                        edgecolor=edge, linewidth=lw)

            # I (served)
            if split_I:
                draw_hsplit(
                    x1, s11, 0.0, bus1_ok, bus2_ok,
                    label=("Нагрузка (шина 1)" if not added["load1"] else None)
                )
                added["load1"] = True
                draw_hsplit(
                    x2, s21, 0.0, bus2_ok, bus1_ok,
                    label=("Нагрузка (шина 2)" if not added["load2"] else None)
                )
                added["load2"] = True
            else:
                plt.bar(
                    x1, s11, width=w, color=bus1_ok, edgecolor=edge, linewidth=lw,
                    label=("Нагрузка (шина 1)" if not added["load1"] else None)
                )
                added["load1"] = True
                plt.bar(
                    x2, s21, width=w, color=bus2_ok, edgecolor=edge, linewidth=lw,
                    label=("Нагрузка (шина 2)" if not added["load2"] else None)
                )
                added["load2"] = True

            # II (served)
            if split_II:
                draw_hsplit(x1, s12, s11, bus1_ok, bus2_ok)
                draw_hsplit(x2, s22, s21, bus2_ok, bus1_ok)
            else:
                plt.bar(x1, s12, width=w, bottom=s11, color=bus1_ok, edgecolor=edge, linewidth=lw)
                plt.bar(x2, s22, width=w, bottom=s21, color=bus2_ok, edgecolor=edge, linewidth=lw)

            # III (served)
            if split_III:
                draw_hsplit(x1, s13, s11 + s12, bus1_ok, bus2_ok)
                draw_hsplit(x2, s23, s21 + s22, bus2_ok, bus1_ok)
            else:
                plt.bar(x1, s13, width=w, bottom=s11 + s12, color=bus1_ok, edgecolor=edge, linewidth=lw)
                plt.bar(x2, s23, width=w, bottom=s21 + s22, color=bus2_ok, edgecolor=edge, linewidth=lw)

            # deficit (серый)
            plt.bar(
                x1, d11, width=w, bottom=s11, color=deficit_color, edgecolor=edge, linewidth=lw,
                label=("Дефицит" if not added["def"] else None)
            )
            added["def"] = True
            plt.bar(x2, d21, width=w, bottom=s21, color=deficit_color, edgecolor=edge, linewidth=lw)

            plt.bar(x1, d12, width=w, bottom=c1_1 + s12, color=deficit_color, edgecolor=edge, linewidth=lw)
            plt.bar(x2, d22, width=w, bottom=c1_2 + s22, color=deficit_color, edgecolor=edge, linewidth=lw)

            plt.bar(x1, d13, width=w, bottom=c1_1 + c2_1 + s13, color=deficit_color, edgecolor=edge, linewidth=lw)
            plt.bar(x2, d23, width=w, bottom=c1_2 + c2_2 + s23, color=deficit_color, edgecolor=edge, linewidth=lw)

            label_I_II_III(x1, c1_1, c2_1, c3_1)
            label_I_II_III(x2, c1_2, c2_2, c3_2)

    plt.xticks(hours)
    plt.xlim(0.4, H + 0.6)

    L1_total = L1_c1 + L1_c2 + L1_c3
    L2_total = L2_c1 + L2_c2 + L2_c3
    ymax = max(
        np.nanmax(L1_total),
        np.nanmax(L2_total),
        np.nanmax(L1_total + L2_total)
    )
    plt.ylim(0, ymax * 1.25)

    plt.xlabel("Время, ч")
    plt.ylabel("")  # убрали подпись оси мощности
    plt.title("Нагрузка по категориям и дефицит: OPEN/FAIL/FAIL2 (2 шины) / CLOSED (объединение шин)")
    plt.grid(True, axis="y", alpha=0.35)
    plt.yticks([])  # оставить деления без подписей

    handles, labels = plt.gca().get_legend_handles_labels()
    uniq = {}
    for hh, ll in zip(handles, labels):
        if ll and ll not in uniq:
            uniq[ll] = hh
    plt.legend(list(uniq.values()), list(uniq.keys()), ncols=3, frameon=False, loc="upper left")

    plt.tight_layout()
    plt.show()


if __name__ == "__main__":
    data, states = read_matrix_from_console(expected_numeric_cols=8)
    plot_scenario(data, states)
