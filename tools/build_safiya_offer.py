"""Персональное предложение SAFIYA по брендам и товарам.

Лист на каждый бренд плюс своды. Где та же позиция есть в публичном
прайсе, показываем, на сколько персональная цена ниже. Публичная цена
идет с НДС 22%, персональная в долларах без НДС, поэтому считаем две
скидки: как есть и к публичной цене, очищенной от НДС.

Запуск:
    python3 tools/build_safiya_offer.py
"""

import csv
import os
import re
import sys

import pandas as pd
from openpyxl import Workbook
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from rates import USD_RUB

OFFER = "data/price_lists/safiya/safiya_2026-08-12.xlsx"
PUBLIC = "data/price_lists/safiya/official/safiya_2026-08-04_official.csv"
TARGET = "outputs/SAFIYA_персональное_предложение.xlsx"
VAT = 1.22

HEAD = PatternFill("solid", fgColor="1F3864")
TOTAL = PatternFill("solid", fgColor="E2EFDA")
WHITE = Font(color="FFFFFF", bold=True)


def read_offer():
    table = pd.read_excel(OFFER, skiprows=4, header=None)
    table.columns = ["Бренд", "Категория", "Наименование", "Штрихкод",
                     "Объем", "Доступно, шт", "Цена, $"]
    table = table.dropna(subset=["Штрихкод", "Цена, $"])
    table["Штрихкод"] = table["Штрихкод"].astype(str).str.replace(r"\D", "", regex=True)
    table["Цена, руб"] = (table["Цена, $"] * USD_RUB).round()
    table["Доступно, шт"] = pd.to_numeric(table["Доступно, шт"], errors="coerce").fillna(0)
    return table


def read_public():
    if not os.path.exists(PUBLIC):
        return {}
    with open(PUBLIC, encoding="utf-8") as handle:
        return {row["Штрихкод"]: row for row in csv.DictReader(handle, delimiter=";")}


def with_public(table, public):
    """Добавляем публичную цену и скидку по совпавшему штрихкоду."""
    rows = []
    for code, price in zip(table["Штрихкод"], table["Цена, руб"]):
        card = public.get(code)
        if not card:
            rows.append(("", "", ""))
            continue
        base = int(card["Цена A, руб"])
        # Скидка положительным числом: сколько процентов мы не платим.
        rows.append((base, round(100 - price / base * 100, 1),
                     round(100 - price / (base / VAT) * 100, 1)))
    table["Публичная цена A, руб"] = [row[0] for row in rows]
    table["Дешевле публичной, %"] = [row[1] for row in rows]
    table["То же без НДС, %"] = [row[2] for row in rows]
    return table


def sheet_name(brand):
    """Имя листа: Excel не берет длиннее 31 знака и служебные символы."""
    return re.sub(r"[\\/*?:\[\]]", "-", brand)[:31]


def write(ws, columns, rows, widths, total=None):
    ws.append(columns)
    for cell in ws[1]:
        cell.fill, cell.font = HEAD, WHITE
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    for row in rows:
        ws.append(row)
    if total:
        ws.append(total)
        for cell in ws[ws.max_row]:
            cell.fill, cell.font = TOTAL, Font(bold=True)
    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width
    ws.freeze_panes = "A2"
    return ws


def money(ws, letters, fmt="# ##0"):
    for letter in letters:
        for cell in ws[letter][1:]:
            cell.number_format = fmt


def scale(ws, letter, rows):
    """Чем ниже персональная цена, тем зеленее ячейка."""
    if rows < 2:
        return
    ws.conditional_formatting.add(
        f"{letter}2:{letter}{rows}",
        ColorScaleRule(start_type="num", start_value=0, start_color="F8696B",
                       mid_type="num", mid_value=20, mid_color="FFEB84",
                       end_type="num", end_value=45, end_color="63BE7B"))


def main():
    table = with_public(read_offer(), read_public())
    book = Workbook()
    book.remove(book.active)

    brands = []
    for brand, part in table.groupby("Бренд", sort=True):
        matched = part[part["Дешевле публичной, %"] != ""]
        brands.append({
            "Бренд": brand,
            "Позиций": len(part),
            "Доступно, шт": int(part["Доступно, шт"].sum()),
            "Средняя цена, руб": round(part["Цена, руб"].mean()),
            "Мин цена, руб": int(part["Цена, руб"].min()),
            "Макс цена, руб": int(part["Цена, руб"].max()),
            "Сверено с публичным, поз.": len(matched),
            "Дешевле публичной, %": (round(matched["Дешевле публичной, %"].median(), 1)
                                     if len(matched) else ""),
        })

    итого = write(
        book.create_sheet("ИТОГО"),
        list(brands[0]),
        [list(row.values()) for row in brands],
        [22, 9, 13, 17, 14, 15, 14, 14],
        ["ВСЕГО", len(table), int(table["Доступно, шт"].sum()),
         round(table["Цена, руб"].mean()), int(table["Цена, руб"].min()),
         int(table["Цена, руб"].max()),
         int((table["Дешевле публичной, %"] != "").sum()),
         round(table.loc[table["Дешевле публичной, %"] != "",
                         "Дешевле публичной, %"].median(), 1)],
    )
    money(итого, "CDEF")
    scale(итого, "H", итого.max_row - 1)

    сверка = table[table["Дешевле публичной, %"] != ""].sort_values(
        "Дешевле публичной, %", ascending=False)
    columns = ["Бренд", "Наименование", "Объем", "Штрихкод", "Цена, $", "Цена, руб",
               "Публичная цена A, руб", "Дешевле публичной, %", "То же без НДС, %"]
    лист = write(
        book.create_sheet("СКИДКА К ПУБЛИЧНОМУ"),
        columns,
        сверка[columns].values.tolist(),
        [16, 58, 12, 16, 10, 12, 15, 14, 14],
    )
    money(лист, "FG")
    money(лист, "E", "0.00")
    scale(лист, "H", лист.max_row)

    for brand, part in table.groupby("Бренд", sort=True):
        columns = ["Категория", "Наименование", "Объем", "Штрихкод", "Доступно, шт",
                   "Цена, $", "Цена, руб", "Публичная цена A, руб", "Дешевле публичной, %"]
        part = part.sort_values("Наименование")
        ws = write(
            book.create_sheet(sheet_name(brand)),
            columns,
            part[columns].values.tolist(),
            [16, 58, 12, 16, 12, 10, 12, 15, 14],
            ["ИТОГО", f"позиций: {len(part)}", "", "", int(part["Доступно, шт"].sum()),
             "", round(part["Цена, руб"].mean()), "", ""],
        )
        money(ws, "EGH")
        money(ws, "F", "0.00")
        scale(ws, "I", ws.max_row - 1)

    os.makedirs("outputs", exist_ok=True)
    book.save(TARGET)
    matched = table[table["Дешевле публичной, %"] != ""]
    print(f"Брендов: {len(brands)}   позиций: {len(table)}")
    print(f"Сверено с публичным прайсом: {len(matched)} позиций, "
          f"персональная цена ниже на {matched['Дешевле публичной, %'].median():.1f}% "
          f"(к цене без НДС — на {matched['То же без НДС, %'].median():.1f}%)")
    print(f"Сохранено: {TARGET}")


if __name__ == "__main__":
    main()
