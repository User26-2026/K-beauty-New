"""Заказ покупателя в разрезе брендов: где у него основные деньги.

Считаем по нашей цене продажи — в складском файле она лежит в колонке
«С/с» с наценкой 10%. Рядом ставим долю бренда в заказе и то, что мы по
нему решили: отдаем целиком, частично или придерживаем.

Запуск:
    python3 tools/order_by_brand.py <заказ.xlsx>
"""

import argparse
import os
import sys

import pandas as pd
from openpyxl import Workbook
from openpyxl.formatting.rule import DataBarRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from brand_names import brand_of
from check_customer_order import read_order

TARGET = "outputs/Заказ покупателя по брендам.xlsx"

HEAD = PatternFill("solid", fgColor="1F3864")
BRAND = PatternFill("solid", fgColor="DDEBF7")
TOTAL = PatternFill("solid", fgColor="E2EFDA")
WHITE = Font(color="FFFFFF", bold=True)

BRANDS = ["Бренд", "Позиций", "Просит, шт", "Сумма, руб", "Доля в заказе, %",
          "Средняя цена, руб", "Самая крупная позиция"]
ITEMS = ["№", "Бренд", "Наименование", "Просит, шт", "Остаток, шт", "Цена, руб",
         "Сумма, руб", "Доля в заказе, %"]


def build(path):
    order = read_order(path)
    order["Бренд"] = order["Наименование"].map(brand_of)
    order["Цена, руб"] = order["Себестоимость, руб"].round()
    order["Сумма, руб"] = (order["Цена, руб"] * order["Просит, шт"]).round()
    order["Доля в заказе, %"] = (order["Сумма, руб"]
                                / order["Сумма, руб"].sum() * 100).round(2)
    return order


def write_brands(book, order):
    ws = book.create_sheet("ПО БРЕНДАМ")
    ws.append(BRANDS)
    for cell in ws[1]:
        cell.fill, cell.font = HEAD, WHITE
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    total = order["Сумма, руб"].sum()
    rows = []
    for brand, part in order.groupby("Бренд"):
        top = part.loc[part["Сумма, руб"].idxmax()]
        rows.append([brand, len(part), int(part["Просит, шт"].sum()),
                     int(part["Сумма, руб"].sum()),
                     round(part["Сумма, руб"].sum() / total * 100, 1),
                     round(part["Сумма, руб"].sum() / part["Просит, шт"].sum()),
                     f"{top['Наименование'][:44]} — {int(top['Сумма, руб']):,} руб"
                     .replace(",", " ")])
    for row in sorted(rows, key=lambda row: -row[3]):
        ws.append(row)
    ws.append(["ВСЕГО", len(order), int(order["Просит, шт"].sum()),
               int(total), 100.0, round(total / order["Просит, шт"].sum()), ""])
    for cell in ws[ws.max_row]:
        cell.fill, cell.font = TOTAL, Font(bold=True)
    for index, width in enumerate([18, 10, 13, 15, 16, 16, 62], start=1):
        ws.column_dimensions[get_column_letter(index)].width = width
    for letter in "BCDF":
        for cell in ws[letter][1:]:
            cell.number_format = "# ##0"
    # Полоски по доле: сразу видно, где сидят деньги.
    ws.conditional_formatting.add(
        f"E2:E{ws.max_row - 1}",
        DataBarRule(start_type="num", start_value=0, end_type="num",
                    end_value=float(max(row[4] for row in rows)), color="638EC6"))
    ws.freeze_panes = "A2"
    return ws


def write_items(book, order):
    ws = book.create_sheet("ПОЗИЦИИ")
    ws.append(ITEMS)
    for cell in ws[1]:
        cell.fill, cell.font = HEAD, WHITE
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    order = order.sort_values(["Бренд", "Сумма, руб"], ascending=[True, False])
    brand, number = None, 0
    for _, item in order.iterrows():
        if item["Бренд"] != brand:
            brand = item["Бренд"]
            ws.append([None, brand])
            for cell in ws[ws.max_row]:
                cell.fill, cell.font = BRAND, Font(bold=True)
        number += 1
        ws.append([number, item["Бренд"], item["Наименование"],
                   int(item["Просит, шт"]), int(item["Остаток, шт"]),
                   item["Цена, руб"], int(item["Сумма, руб"]),
                   item["Доля в заказе, %"]])
    ws.append([None, "ИТОГО", f"позиций: {len(order)}",
               int(order["Просит, шт"].sum()), "", "",
               int(order["Сумма, руб"].sum()), 100.0])
    for cell in ws[ws.max_row]:
        cell.fill, cell.font = TOTAL, Font(bold=True)
    for index, width in enumerate([5, 16, 66, 12, 12, 12, 14, 15], start=1):
        ws.column_dimensions[get_column_letter(index)].width = width
    for letter in "DEFG":
        for cell in ws[letter][1:]:
            cell.number_format = "# ##0"
    ws.freeze_panes = "C2"
    return ws


def main(source, target):
    order = build(source)
    book = Workbook()
    book.remove(book.active)
    write_brands(book, order)
    write_items(book, order)
    os.makedirs("outputs", exist_ok=True)
    book.save(target)
    total = order["Сумма, руб"].sum()
    top = (order.groupby("Бренд")["Сумма, руб"].sum()
           .sort_values(ascending=False).head(5))
    print(f"Позиций: {len(order)}   штук: {int(order['Просит, шт'].sum()):,}   "
          f"сумма: {int(total):,} руб".replace(",", " "))
    print("Основные деньги:")
    for brand, value in top.items():
        print(f"   {brand:<16}{int(value):>12,} руб   {value / total * 100:>5.1f}%"
              .replace(",", " "))
    print(f"Сохранено: {target}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Заказ покупателя по брендам")
    parser.add_argument("source")
    parser.add_argument("--out", default=TARGET)
    args = parser.parse_args()
    main(args.source, args.out)
