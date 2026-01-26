import os
from pathlib import Path
from copy import copy

from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet


SRC_SHEET = "SWEEP_2"
DST_SHEET = "SWEEP_2"

DEFAULT_SRC_DIR = r"D:\10_results"
DEFAULT_OUT_PATH = r"D:\10_results\combined.xlsx"


# Порядок вставки (имена файлов БЕЗ .xlsx)
ORDER = ["2.1.1", "2.2.1", "2.1.2", "2.2.2", "2.1.3", "2.2.3", "2.1.4", "2.2.4"]

# Сдвиг между блоками (1 колонка пустая)
GAP_COLS = 1


def used_range(ws: Worksheet):
    """
    Определяет используемый диапазон по max_row/max_column.
    Если нужно жёстко ограничить область (например A1:F200) — можно заменить этой логикой.
    """
    return 1, 1, ws.max_row, ws.max_column  # min_row, min_col, max_row, max_col


def copy_cell(src, dst):
    dst.value = src.value

    # стиль/форматы можно оставить
    if src.has_style:
        dst._style = copy(src._style)
    dst.number_format = src.number_format
    dst.font = copy(src.font)
    dst.fill = copy(src.fill)
    dst.border = copy(src.border)
    dst.alignment = copy(src.alignment)
    dst.protection = copy(src.protection)
    dst.comment = copy(src.comment) if src.comment else None


def copy_dimensions(src_ws: Worksheet, dst_ws: Worksheet, col_offset: int):
    # column widths
    for col_letter, dim in src_ws.column_dimensions.items():
        if not dim.width:
            continue
        src_col_idx = openpyxl_col_to_int(col_letter)
        dst_col_idx = src_col_idx + col_offset
        dst_letter = get_column_letter(dst_col_idx)
        dst_ws.column_dimensions[dst_letter].width = dim.width

    # row heights
    for row_idx, dim in src_ws.row_dimensions.items():
        if dim.height:
            dst_ws.row_dimensions[row_idx].height = dim.height


def openpyxl_col_to_int(col_letter: str) -> int:
    # openpyxl has column_index_from_string but keep local to avoid extra imports
    col_letter = col_letter.upper()
    n = 0
    for ch in col_letter:
        n = n * 26 + (ord(ch) - ord('A') + 1)
    return n


def copy_merged_cells(src_ws: Worksheet, dst_ws: Worksheet, col_offset: int):
    for mr in src_ws.merged_cells.ranges:
        # mr: e.g. "A1:C1"
        min_col, min_row, max_col, max_row = mr.bounds
        dst_ws.merge_cells(
            start_row=min_row,
            start_column=min_col + col_offset,
            end_row=max_row,
            end_column=max_col + col_offset,
        )


def copy_sheet_block(src_ws: Worksheet, dst_ws: Worksheet, col_offset: int):
    min_row, min_col, max_row, max_col = used_range(src_ws)

    # ячейки
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            src_cell = src_ws.cell(row=r, column=c)
            dst_cell = dst_ws.cell(row=r, column=c + col_offset)
            copy_cell(src_cell, dst_cell)

    # объединения
    copy_merged_cells(src_ws, dst_ws, col_offset)

    # размеры (ширины колонок/высоты строк)
    copy_dimensions(src_ws, dst_ws, col_offset)


def main(src_dir: str, out_path: str):
    src_dir = Path(src_dir)
    out_path = Path(out_path)

    # создаём книгу-результат
    out_wb = Workbook()
    # удалить дефолтный лист
    if out_wb.active.title == "Sheet":
        out_wb.remove(out_wb.active)
    dst_ws = out_wb.create_sheet(DST_SHEET)

    col_offset = 0
    block_width = None

    for name in ORDER:
        src_file = src_dir / f"{name}.xlsx"
        if not src_file.exists():
            raise FileNotFoundError(f"Не найден файл: {src_file}")

        wb = load_workbook(src_file, data_only=True)
        if SRC_SHEET not in wb.sheetnames:
            raise ValueError(f"В файле {src_file.name} нет листа {SRC_SHEET}")
        src_ws = wb[SRC_SHEET]

        # определяем ширину блока по первой книге
        min_row, min_col, max_row, max_col = used_range(src_ws)
        current_width = (max_col - min_col + 1)

        if block_width is None:
            block_width = current_width
        else:
            if current_width != block_width:
                raise ValueError(
                    f"Размер блока отличается: {src_file.name} имеет ширину {current_width}, ожидалось {block_width}"
                )

        # копируем блок
        copy_sheet_block(src_ws, dst_ws, col_offset)

        # сдвигаем offset на ширину блока + 1 пустая колонка
        col_offset += block_width + GAP_COLS

        wb.close()

    out_wb.save(out_path)
    print(f"OK: сохранено {out_path}")


if __name__ == "__main__":
    src_dir = DEFAULT_SRC_DIR
    out_path = DEFAULT_OUT_PATH
    main(src_dir, out_path)

