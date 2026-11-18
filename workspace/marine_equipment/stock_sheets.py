#!/usr/bin/env python3
"""Остатки склада по локациям, файлы менеджера за сентябрь 2026.

Четыре файла, пять листов, у каждого своя раскладка:

  Склад_1А_1Б   — лист «Склад 1А, 1Б», колонки A локация, B наименование,
                  C единица, D количество, E цена, F сумма;
  Фойе_и_цех    — листы «Фойе» и «цех», БЕЗ колонки единицы: количество
                  стоит в C, цена в D, сумма в E;
  Гараж_2       — лист «Гараж 2», раскладка как у склада 1А;
  Контейнер_2   — лист «Контейнер 2», раскладка как у фойе.

Поэтому колонка количества ищется по шапке, а не берется по номеру:
в одном файле это D, в другом C, и перепутать их — значит принять
цену за количество.

Старый формат .xls читается через xlrd, новый через openpyxl.
"""

import pathlib
import re

ROOT = pathlib.Path(__file__).resolve().parents[2]
SRC = ROOT / "data" / "marine_equipment" / "stock_09.2026"

FILES = [
    ("Склад_1А_1Б_сент2026.xlsx", "Склад 1А, 1Б"),
    ("Фойе_и_цех_сент2026.xls", "Фойе и цех"),
    ("Гараж_2_сент2026.xlsx", "Гараж 2"),
    ("Контейнер_2_сент2026.xls", "Контейнер 2"),
]

# строки, которые не являются товаром
SKIP = re.compile(r"^\s*(итого|всего|сумма|наименование)\b", re.I)


def text(v):
    return re.sub(r"\s+", " ", str(v)).strip() if v is not None else ""


def num(v):
    if isinstance(v, (int, float)):
        return float(v)
    s = text(v).replace(",", ".").replace(" ", "")
    return float(s) if re.match(r"^-?\d+(\.\d+)?$", s) else None


def read_xlsx(path):
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True)
    for ws in wb.worksheets:
        yield ws.title, [[c for c in row]
                         for row in ws.iter_rows(values_only=True)]


def read_xls(path):
    import xlrd
    wb = xlrd.open_workbook(path)
    for sh in wb.sheets():
        yield sh.name, [[sh.cell_value(r, c) for c in range(sh.ncols)]
                        for r in range(sh.nrows)]


def layout(rows):
    """Где количество, цена и сумма. Ищем по шапке, а не по номеру.

    Признак шапки — ячейка со словом «цена»: справа от нее сумма, слева
    количество. Если шапки нет, берем раскладку по умолчанию.
    """
    for row in rows[:6]:
        for i, c in enumerate(row):
            if re.search(r"цена", text(c), re.I):
                return dict(qty=i - 1, price=i, total=i + 1)
    return dict(qty=3, price=4, total=5)


# Подзаголовки внутри листа: строка с названием, но без количества.
# В «Фойе» ими размечены группы менеджера — электрика, насосы,
# клапана, фильтры, клинкет, корпусная часть. Это ее собственное
# деление товара, и терять его нельзя: по нему она и считает склад.
HEADING = re.compile(r"^(судовое оборудование|фойе|цех|гараж|контейнер)",
                     re.I)


def read():
    """Строки остатков по всем локациям, с подгруппами менеджера."""
    out = []
    for fname, label in FILES:
        path = SRC / fname
        reader = read_xls if path.suffix == ".xls" else read_xlsx
        for sheet, rows in reader(path):
            col = layout(rows)
            group = ""
            for i, row in enumerate(rows, start=1):
                def cell(k):
                    j = col[k]
                    return row[j] if 0 <= j < len(row) else None

                name = text(row[1] if len(row) > 1 else "")
                if not name or SKIP.match(name):
                    continue
                qty = num(cell("qty"))
                if not qty or qty <= 0:
                    # название без количества — это подзаголовок группы,
                    # кроме шапки самого листа
                    if not HEADING.match(name):
                        group = name
                    continue
                out.append({
                    "Подгруппа": group,
                    "Файл": label,
                    "Лист": sheet.strip(),
                    "Строка": i,
                    "Локация": text(row[0]) or sheet.strip(),
                    "Наименование": name,
                    "Штук": int(qty),
                    "Цена": num(cell("price")) or 0,
                    "Сумма": num(cell("total")) or 0,
                })
    return out


if __name__ == "__main__":
    rows = read()
    from collections import defaultdict
    by = defaultdict(lambda: [0, 0, 0.0])
    for r in rows:
        a = by[f'{r["Файл"]} / {r["Лист"]}']
        a[0] += 1
        a[1] += r["Штук"]
        a[2] += r["Сумма"]
    print(f"строк всего: {len(rows)}, штук {sum(r['Штук'] for r in rows):,}"
          .replace(",", " "))
    print()
    for k, (n, q, s) in by.items():
        print(f"  {k:<28}{n:>4} поз, {q:>6} шт, {s:>14,.0f} руб"
              .replace(",", " "))
    print(f"\nитого по указанным ценам: "
          f"{sum(r['Сумма'] for r in rows):,.0f} руб".replace(",", " "))
