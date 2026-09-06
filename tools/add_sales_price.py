"""Проставить наши продажные цены к списку остатков.

В остатках есть себестоимость, но нет цены, по которой мы продаем. Цену
берем из нашего оптового прайса и считаем, на какую сумму лежит склад в
продажных ценах.

Общего кода у файлов нет, сводим по названию: в прайсе и в остатках оно
пишется одинаково, поэтому сначала сравниваем строку целиком, а остаток
догоняем разбором по словам.

Запуск:
    python3 tools/add_sales_price.py <остатки.xlsx> [--out файл.xlsx]
"""

import argparse
import os
import re
import sys

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import name_match

PRICES = "data/pricing/pricelist_2026-08-24.xlsx"
TARGET = "outputs/Остатки с продажными ценами.xlsx"

HEAD = PatternFill("solid", fgColor="1F3864")
MISS = PatternFill("solid", fgColor="FCE4D6")
TOTAL = PatternFill("solid", fgColor="E2EFDA")
WHITE = Font(color="FFFFFF", bold=True)

COLUMNS = ["Наименование", "Остаток, шт", "Себестоимость, руб", "Срок годности",
           "Продажная цена, руб", "Продажа х остаток, руб"]


def normal(name):
    return re.sub(r"[^0-9a-zа-яё]+", "", str(name).lower())


def read_prices():
    table = pd.read_excel(PRICES, skiprows=4)
    table.columns = ["№", "Наименование", "Штрихкод", "В наличии, шт", "Цена, ₽"]
    table = table.dropna(subset=["Цена, ₽"])
    return table.reset_index(drop=True)


def read_stock(path):
    table = pd.read_excel(path, header=0)
    table = table.iloc[:, :4]
    table.columns = COLUMNS[:4]
    table = table[table["Наименование"].notna()]
    table = table[table["Остаток, шт"].notna()]
    table["Остаток, шт"] = pd.to_numeric(
        table["Остаток, шт"].astype(str).str.replace(r"\s", "", regex=True),
        errors="coerce")
    return table[table["Остаток, шт"].notna()].reset_index(drop=True)


def attach(stock, prices):
    """Цена из прайса: сначала точное совпадение строки, потом по словам."""
    by_text = {}
    for index, name in prices["Наименование"].items():
        by_text.setdefault(normal(name), index)
    found = pd.Series([by_text.get(normal(name)) for name in stock["Наименование"]],
                      index=stock.index, dtype="object")

    left = stock.loc[found.isna(), "Наименование"]
    free = prices.drop(index=[i for i in found.dropna()])
    for index, other in name_match.match(left, free["Наименование"]).items():
        found[index] = other

    stock["Продажная цена, руб"] = [
        prices.loc[index, "Цена, ₽"] if pd.notna(index) else None for index in found]
    stock["Продажа х остаток, руб"] = (
        stock["Продажная цена, руб"] * stock["Остаток, шт"]).round()
    return stock


def main(source, target):
    stock = attach(read_stock(source), read_prices())
    book = Workbook()
    ws = book.active
    ws.title = "ОСТАТКИ"
    ws.append(COLUMNS)
    for cell in ws[1]:
        cell.fill, cell.font = HEAD, WHITE
        cell.alignment = Alignment(wrap_text=True, vertical="center")

    for row in stock[COLUMNS].where(stock.notna(), None).values.tolist():
        ws.append(row)
        if row[4] is None:
            # Цены в прайсе нет — строка не попадает в сумму, это видно.
            for cell in ws[ws.max_row]:
                cell.fill = MISS

    ws.append(["ИТОГО", int(stock["Остаток, шт"].sum()), "", "", "",
               int(stock["Продажа х остаток, руб"].sum(skipna=True))])
    for cell in ws[ws.max_row]:
        cell.fill, cell.font = TOTAL, Font(bold=True)

    for index, width in enumerate([96, 12, 16, 13, 16, 18], start=1):
        ws.column_dimensions[get_column_letter(index)].width = width
    for letter in "BEF":
        for cell in ws[letter][1:]:
            cell.number_format = "# ##0"
    for cell in ws["C"][1:]:
        cell.number_format = "# ##0.00"
    ws.freeze_panes = "A2"

    os.makedirs(os.path.dirname(target) or ".", exist_ok=True)
    book.save(target)
    priced = stock["Продажная цена, руб"].notna().sum()
    print(f"Позиций: {len(stock)}   с продажной ценой: {priced}   "
          f"без цены: {len(stock) - priced}")
    print(f"Склад в продажных ценах: "
          f"{int(stock['Продажа х остаток, руб'].sum(skipna=True)):,} руб".replace(",", " "))
    print(f"Сохранено: {target}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Продажные цены к остаткам")
    parser.add_argument("source")
    parser.add_argument("--out", default=TARGET)
    args = parser.parse_args()
    main(args.source, args.out)
