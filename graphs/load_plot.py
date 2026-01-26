import matplotlib.pyplot as plt
import numpy as np

print("Введите 3 столбца по 24 значения (часовая нагрузка).")
print("Формат ввода: в каждой строке 3 числа через пробел/таб/; (десятичные можно через запятую).")
print("Пример: 12,5  10  8,2")
print("Нужно ровно 24 строки. Пустая строка — завершить ввод (если уже введено 24 строки).\n")

rows = []
while True:
    s = input().strip()
    if s == "":
        break

    # поддержка разделителей: пробелы, табы, ';'
    s = s.replace(";", " ").replace("\t", " ")
    parts = [p for p in s.split(" ") if p != ""]

    if len(parts) != 3:
        print("⚠️ В строке должно быть ровно 3 числа (3 типа нагрузки).")
        continue

    try:
        a = float(parts[0].replace(",", "."))
        b = float(parts[1].replace(",", "."))
        c = float(parts[2].replace(",", "."))
        rows.append((a, b, c))
        if len(rows) == 24:
            break
    except ValueError:
        print("⚠️ Введите корректные числа (десятичные можно через запятую).")

if len(rows) != 24:
    print(f"⚠️ Введено {len(rows)} строк. Нужно ровно 24 строки (по одному часу).")
    raise SystemExit

data = np.array(rows, dtype=float)  # shape (24, 3)
h = np.arange(1, 25)  # часы 1..24

y1 = data[:, 0]
y2 = data[:, 1]
y3 = data[:, 2]

plt.figure(figsize=(11, 5))
plt.plot(h, y1, linewidth=1.2, label="Промышленная (ГОК)")
plt.plot(h, y2, linewidth=1.2, label="Коммунально-бытовая (поселение)")
plt.plot(h, y3, linewidth=1.2, label="Сельхоз (орошение)")

plt.xticks(np.arange(1, 25, 1))
plt.xlim(1, 24)

plt.xlabel("Время (часы)", fontsize=11)
plt.ylabel("Часовая нагрузка, о.е", fontsize=11)
plt.grid(True, linestyle="--", alpha=0.35)
plt.legend()

# запас по Y
all_y = np.concatenate([y1, y2, y3])
ymin, ymax = float(np.min(all_y)), float(np.max(all_y))
pad = (ymax - ymin) * 0.05 if ymax > ymin else 1.0
plt.ylim(ymin - pad, ymax + pad)

plt.tight_layout()
plt.show()
