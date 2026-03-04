import re
from typing import Dict, Tuple, Optional

import pandas as pd


# -----------------------------
# USER SETTINGS
# -----------------------------
EXCEL_PATH = r"D:\1.xlsx"          # <-- поменяйте на ваш путь (у вас файл на диске D)
SHEET_NAME = "RAW"                # в вашем файле данные лежат в RAW

DELTA_RU_RUB = 1_000_000          # двойная шина дороже на 1 млн руб
K2 = 5.0                          # штраф 2 кат = 5 * 3 кат
K1 = 10.0                         # штраф 1 кат = 10 * 3 кат
K3 = 1.0                          # штраф 3 кат = 1 * 3 кат

# Какие столбцы должны быть в таблице (как в вашем RAW)
COL_PARAM = "param1"
COL_ENS_TOTAL = "ENS,кВт∙ч"
COL_ENS1 = "ENS1_mean"
COL_ENS2 = "ENS2_mean"

BUS_SINGLE_KEYWORD = "SINGLE_SECTIONAL_BUS"
BUS_DOUBLE_KEYWORD = "DOUBLE_BUS"

OUT_XLSX = "thresholds_out.xlsx"


# -----------------------------
# Helpers
# -----------------------------
def parse_float_ru(x) -> float:
    """Парсит числа вида 5537,53 или 1,596E-04"""
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return float("nan")
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).strip().replace(" ", "").replace("\u2212", "-")
    if s == "":
        return float("nan")
    s = s.replace(",", ".")
    return float(s)


def read_raw_sheet_as_matrix(path: str, sheet: str) -> pd.DataFrame:
    # header=None чтобы получить все строки "как есть"
    return pd.read_excel(path, sheet_name=sheet, header=None, dtype=object)


def extract_blocks(mat: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    """
    В RAW у вас блоки выглядят так:
      row i: "bus=...."
      row i+1: заголовки (param1, DG_kW, ..., FailBrk)
      rows i+2..: данные
      затем пустая строка или следующий bus=
    Возвращает словарь: bus_name -> DataFrame с нормальными колонками.
    """
    blocks: Dict[str, pd.DataFrame] = {}

    # Найти строки, где в первом столбце есть bus=...
    bus_rows = []
    for r in range(len(mat)):
        v = mat.iat[r, 0]
        if isinstance(v, str) and v.strip().startswith("bus="):
            bus_rows.append(r)

    if not bus_rows:
        raise ValueError("В листе RAW не найдено строк вида 'bus=...' в первом столбце.")

    # Добавим "конец" как длину матрицы
    bus_rows.append(len(mat))

    for idx in range(len(bus_rows) - 1):
        bus_r = bus_rows[idx]
        end_r = bus_rows[idx + 1]

        bus_line = str(mat.iat[bus_r, 0]).strip()
        # bus=XXX; ...
        m = re.match(r"bus\s*=\s*([^;]+)", bus_line)
        bus_name = m.group(1).strip() if m else bus_line

        header_r = bus_r + 1
        if header_r >= len(mat):
            continue

        # Заголовки лежат в строке header_r по столбцам 0..N пока не пусто
        header = []
        for c in range(mat.shape[1]):
            hv = mat.iat[header_r, c]
            if hv is None or (isinstance(hv, float) and pd.isna(hv)):
                break
            header.append(str(hv).strip())

        if not header or header[0] != "param1":
            # иногда между bus= и header может быть пустая строка — попробуем найти "param1" ниже
            found = False
            for rr in range(bus_r + 1, min(bus_r + 6, end_r)):
                v0 = mat.iat[rr, 0]
                if isinstance(v0, str) and v0.strip() == "param1":
                    header_r = rr
                    header = []
                    for c in range(mat.shape[1]):
                        hv = mat.iat[header_r, c]
                        if hv is None or (isinstance(hv, float) and pd.isna(hv)):
                            break
                        header.append(str(hv).strip())
                    found = True
                    break
            if not found:
                raise ValueError(f"Не нашёл строку заголовков (param1...) для блока {bus_name}")

        # Данные начинаются со следующей строки
        data_start = header_r + 1
        rows = []
        for r in range(data_start, end_r):
            v0 = mat.iat[r, 0]
            # конец блока: пустая строка
            if v0 is None or (isinstance(v0, float) and pd.isna(v0)):
                break
            # или случайно встретили следующий bus=
            if isinstance(v0, str) and v0.strip().startswith("bus="):
                break

            row = []
            for c in range(len(header)):
                row.append(mat.iat[r, c])
            rows.append(row)

        df = pd.DataFrame(rows, columns=header)

        # Преобразуем числовые колонки
        for col in df.columns:
            # param1 тоже числовой
            try:
                df[col] = df[col].map(parse_float_ru)
            except Exception:
                pass

        blocks[bus_name] = df

    return blocks


def compute_thresholds(single_df: pd.DataFrame, double_df: pd.DataFrame) -> pd.DataFrame:
    # Проверки
    for col in [COL_PARAM, COL_ENS_TOTAL, COL_ENS1, COL_ENS2]:
        if col not in single_df.columns:
            raise ValueError(f"В SINGLE нет колонки '{col}'")
        if col not in double_df.columns:
            raise ValueError(f"В DOUBLE нет колонки '{col}'")

    s = single_df[[COL_PARAM, COL_ENS_TOTAL, COL_ENS1, COL_ENS2]].copy()
    d = double_df[[COL_PARAM, COL_ENS_TOTAL, COL_ENS1, COL_ENS2]].copy()

    # ENS3 = total - ENS1 - ENS2
    s["ENS3_mean"] = s[COL_ENS_TOTAL] - s[COL_ENS1] - s[COL_ENS2]
    d["ENS3_mean"] = d[COL_ENS_TOTAL] - d[COL_ENS1] - d[COL_ENS2]

    m = pd.merge(
        s.rename(columns={
            COL_ENS_TOTAL: "ENS_s", COL_ENS1: "ENS1_s", COL_ENS2: "ENS2_s", "ENS3_mean": "ENS3_s"
        }),
        d.rename(columns={
            COL_ENS_TOTAL: "ENS_d", COL_ENS1: "ENS1_d", COL_ENS2: "ENS2_d", "ENS3_mean": "ENS3_d"
        }),
        on=COL_PARAM, how="inner"
    )

    # Δ = double - single
    m["dENS1"] = m["ENS1_d"] - m["ENS1_s"]
    m["dENS2"] = m["ENS2_d"] - m["ENS2_s"]
    m["dENS3"] = m["ENS3_d"] - m["ENS3_s"]
    m["dENS_total"] = m["ENS_d"] - m["ENS_s"]

    # Взвешенная дельта на 1 руб/кВт·ч базового штрафа 3 кат
    m["dW_per_c3"] = K1 * m["dENS1"] + K2 * m["dENS2"] + K3 * m["dENS3"]

    # Порог: DELTA_RU + c3 * dW = 0  =>  c3* = DELTA_RU / (-dW) если dW < 0
    def c3_star(dw: float) -> Optional[float]:
        if pd.isna(dw):
            return None
        if dw < 0:
            return DELTA_RU_RUB / (-dw)
        return None  # если dw >= 0, двойная не окупается при DELTA_RU>0

    m["c3_star_rub_per_kwh"] = m["dW_per_c3"].apply(c3_star)
    m["c2_star_rub_per_kwh"] = m["c3_star_rub_per_kwh"] * K2
    m["c1_star_rub_per_kwh"] = m["c3_star_rub_per_kwh"] * K1

    out = m[[COL_PARAM,
             "ENS_s", "ENS_d", "dENS_total",
             "dENS1", "dENS2", "dENS3",
             "dW_per_c3",
             "c3_star_rub_per_kwh", "c2_star_rub_per_kwh", "c1_star_rub_per_kwh"]].copy()

    return out.sort_values(COL_PARAM).reset_index(drop=True)


# -----------------------------
# Main
# -----------------------------
if __name__ == "__main__":
    mat = read_raw_sheet_as_matrix(EXCEL_PATH, SHEET_NAME)
    blocks = extract_blocks(mat)

    # найти нужные блоки по ключевым словам
    single_name = next((k for k in blocks.keys() if BUS_SINGLE_KEYWORD in k), None)
    double_name = next((k for k in blocks.keys() if BUS_DOUBLE_KEYWORD in k), None)
    if single_name is None or double_name is None:
        raise ValueError(f"Не нашёл блоки SINGLE/DOUBLE. Найдено: {list(blocks.keys())}")

    df_single = blocks[single_name]
    df_double = blocks[double_name]

    res = compute_thresholds(df_single, df_double)

    print(res.to_string(index=False))

    # сохранить в Excel
    with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as w:
        res.to_excel(w, sheet_name="thresholds", index=False)
    print(f"\nSaved: {OUT_XLSX}")