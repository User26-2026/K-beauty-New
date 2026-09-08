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
from add_sales_price import sales_by_stock
from brand_names import brand_of
from check_customer_order import read_order
from wholesale_vs_wb import MARGIN, SKIP

TARGET = "outputs/Заказ покупателя по брендам.xlsx"

HEAD = PatternFill("solid", fgColor="1F3864")
BRAND = PatternFill("solid", fgColor="DDEBF7")
TOTAL = PatternFill("solid", fgColor="E2EFDA")
WHITE = Font(color="FFFFFF", bold=True)

BRANDS = ["Бренд", "Позиций", "Просит, шт", "Сумма, руб", "Доля в заказе, %",
          "Прибыль оптом, руб", "Прибыль на WB, руб", "Разница, руб",
          "Во сколько раз WB выгоднее", "Сверено с WB, поз.",
          "Самая крупная позиция"]
ITEMS = ["№", "Бренд", "Наименование", "Просит, шт", "Остаток, шт", "Цена, руб",
         "Сумма, руб", "Доля в заказе, %", "Прибыль оптом, руб",
         "Прибыль на WB, руб", "Разница, руб"]


def read_margin(path):
    """Фактическая прибыль WB по карточкам за месяц."""
    table = pd.read_excel(path, sheet_name=0, header=None).iloc[SKIP:]
    table = table[list(MARGIN)].rename(columns=MARGIN)
    table = table[table["Товар"].notna()]
    for column in ("Продано, шт", "Маржа, руб"):
        table[column] = pd.to_numeric(table[column], errors="coerce").fillna(0)
    return table.rename(columns={"Маржа, руб": "Выручка, руб"}).reset_index(drop=True)


def build(path, sales_path):
    order = read_order(path)
    order["Бренд"] = order["Наименование"].map(brand_of)
    # В колонке «С/с» лежит цена с наценкой 10%, себестоимость получаем делением.
    order["Цена, руб"] = order["Себестоимость, руб"].round()
    order["Сумма, руб"] = (order["Цена, руб"] * order["Просит, шт"]).round()
    order["Доля в заказе, %"] = (order["Сумма, руб"]
                                / order["Сумма, руб"].sum() * 100).round(2)
    order["Прибыль оптом, руб"] = (
        (order["Цена, руб"] - order["Себестоимость, руб"] / 1.1)
        * order["Просит, шт"]).round()

    found = sales_by_stock(order, read_margin(sales_path))
    order["Прибыль на WB, руб"] = [
        round(margin / sold * want) if sold else None
        for margin, sold, want in zip(found["Выручка, руб"], found["Продано, шт"],
                                      order["Просит, шт"])]
    order["Разница, руб"] = [
        round(wb - opt) if pd.notna(wb) else None
        for wb, opt in zip(order["Прибыль на WB, руб"], order["Прибыль оптом, руб"])]
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
        known = part[part["Прибыль на WB, руб"].notna()]
        # Прибыль оптом берем по тем же позициям, что сверились с WB,
        # иначе сравниваются разные наборы товара.
        opt = known["Прибыль оптом, руб"].sum()
        wb = known["Прибыль на WB, руб"].sum()
        rows.append([brand, len(part), int(part["Просит, шт"].sum()),
                     int(part["Сумма, руб"].sum()),
                     round(part["Сумма, руб"].sum() / total * 100, 1),
                     int(opt), int(wb), int(wb - opt),
                     round(wb / opt, 1) if opt > 0 else "", len(known),
                     f"{top['Наименование'][:44]} — {int(top['Сумма, руб']):,} руб"
                     .replace(",", " ")])
    for row in sorted(rows, key=lambda row: -row[3]):
        ws.append(row)
    known = order[order["Прибыль на WB, руб"].notna()]
    ws.append(["ВСЕГО", len(order), int(order["Просит, шт"].sum()), int(total), 100.0,
               int(known["Прибыль оптом, руб"].sum()),
               int(known["Прибыль на WB, руб"].sum()),
               int(known["Разница, руб"].sum()),
               round(known["Прибыль на WB, руб"].sum()
                     / known["Прибыль оптом, руб"].sum(), 1), len(known), ""])
    for cell in ws[ws.max_row]:
        cell.fill, cell.font = TOTAL, Font(bold=True)
    for index, width in enumerate([18, 10, 13, 15, 16, 16, 16, 15, 16, 15, 58],
                                  start=1):
        ws.column_dimensions[get_column_letter(index)].width = width
    for letter in "BCDFGHJ":
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
                   item["Доля в заказе, %"], int(item["Прибыль оптом, руб"]),
                   int(item["Прибыль на WB, руб"])
                   if pd.notna(item["Прибыль на WB, руб"]) else "",
                   int(item["Разница, руб"])
                   if pd.notna(item["Разница, руб"]) else ""])
    known = order[order["Прибыль на WB, руб"].notna()]
    ws.append([None, "ИТОГО", f"позиций: {len(order)}",
               int(order["Просит, шт"].sum()), "", "",
               int(order["Сумма, руб"].sum()), 100.0,
               int(known["Прибыль оптом, руб"].sum()),
               int(known["Прибыль на WB, руб"].sum()),
               int(known["Разница, руб"].sum())])
    for cell in ws[ws.max_row]:
        cell.fill, cell.font = TOTAL, Font(bold=True)
    for index, width in enumerate([5, 16, 60, 12, 12, 12, 14, 15, 16, 16, 14],
                                  start=1):
        ws.column_dimensions[get_column_letter(index)].width = width
    for letter in "DEFGIJK":
        for cell in ws[letter][1:]:
            cell.number_format = "# ##0"
    ws.freeze_panes = "C2"
    return ws


def main(source, sales_path, target):
    order = build(source, sales_path)
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
    parser.add_argument("--sales", default="data/sales/wb_sales_2026-08.xls")
    parser.add_argument("--out", default=TARGET)
    args = parser.parse_args()
    main(args.source, args.sales, args.out)
