"""Заказ покупателя: отдать оптом или доторговать на WB.

По каждой позиции заказа считаем две прибыли на одно и то же
количество: оптом по нашему прайсу и на WB по фактической марже из
отчета за месяц. Рядом — что останется на складе после отгрузки и на
сколько этого хватит с учетом сезона.

Запуск:
    python3 tools/check_customer_order.py <заказ.xlsx> --sales <отчет WB.xls>
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
from add_sales_price import attach, read_prices, sales_by_stock
from stock_forecast import run_out
from wholesale_vs_wb import read_margin

TARGET = "outputs/Заказ покупателя — опт или WB.xlsx"

HEAD = PatternFill("solid", fgColor="1F3864")
TOTAL = PatternFill("solid", fgColor="E2EFDA")
STOP = PatternFill("solid", fgColor="FCE4D6")
WHITE = Font(color="FFFFFF", bold=True)

SHOWN = ["Наименование", "Остаток, шт", "Просит, шт", "Останется, шт",
         "Себестоимость, руб", "Цена опт, руб", "Выручка опт, руб",
         "Прибыль опт, руб", "Продажи в месяц, шт", "Маржа WB на штуку, руб",
         "Прибыль WB на это кол-во, руб", "Разница, руб",
         "Остатка хватит на, мес", "Вывод"]
WIDTHS = [84, 12, 12, 13, 15, 12, 14, 14, 14, 15, 16, 14, 14, 26]


# Колонку ищем по заголовку: складской файл и файл с заказом идут с
# разной раскладкой, в заказе есть лишняя колонка с остатком после
# отгрузки, и по номерам колонки не совпадают.
HEADERS = {"Наименование": r"наименование", "Остаток, шт": r"^остаток",
           # С/с в шапке набрано латиницей, хотя выглядит как кириллица.
           "Себестоимость, руб": r"^[cс]\s*/?\s*[cс]$|себестоим",
           "Просит, шт": r"^заказ"}


def find_columns(table):
    found = {}
    for field, pattern in HEADERS.items():
        for column in table.columns:
            title = str(table[column].name)
            if re.search(pattern, title, re.IGNORECASE) and field not in found:
                found[field] = column
    return found


def read_order(path, everything=False):
    """Заказ покупателя: остаток, себестоимость и запрошенное количество.

    С everything берем все строки склада, а не только те, что он просит.
    """
    table = pd.read_excel(path, header=0)
    columns = find_columns(table)
    missing = set(HEADERS) - set(columns) - {"Просит, шт"}
    if missing:
        raise SystemExit(f"в файле не найдены колонки: {', '.join(sorted(missing))}")
    table = table[[columns[field] for field in columns]]
    table.columns = list(columns)
    if "Просит, шт" not in table:
        table["Просит, шт"] = 0
    table = table[table["Наименование"].notna()]
    for column in ("Остаток, шт", "Себестоимость, руб", "Просит, шт"):
        # Количество приходит текстом с пробелом-разделителем тысяч.
        table[column] = pd.to_numeric(
            table[column].astype(str).str.replace(r"\s|\u00a0", "", regex=True),
            errors="coerce")
    if everything:
        table = table[table["Остаток, шт"].notna()]
        table["Просит, шт"] = table["Просит, шт"].fillna(0)
        return table.reset_index(drop=True)
    return table[table["Просит, шт"].notna() & (table["Просит, шт"] > 0)].reset_index(drop=True)


def build(order, report):
    order = attach(order, read_prices())
    found = sales_by_stock(order, report.rename(columns={"Маржа, руб": "Выручка, руб"}))
    order["Продажи в месяц, шт"] = found["Продано, шт"]
    order["Маржа WB на штуку, руб"] = [
        round(margin / sold, 1) if sold else None
        for margin, sold in zip(found["Выручка, руб"], found["Продано, шт"])
    ]
    return order


def rows_of(order, average):
    rows = []
    for _, item in order.iterrows():
        take = item["Просит, шт"]
        rest = item["Остаток, шт"] - take
        cost, price = item["Себестоимость, руб"], item["Продажная цена, руб"]
        opt_sum = round(price * take) if pd.notna(price) else None
        opt = round((price - cost) * take) if pd.notna(price) else None
        unit = item["Маржа WB на штуку, руб"]
        wb = round((unit if pd.notna(unit) else average) * take)
        monthly = item["Продажи в месяц, шт"]
        # На сколько хватит того, что останется после отгрузки.
        if monthly and rest > 0:
            _, months, _ = run_out(rest, monthly)
            months = round(months, 1)
        else:
            months = 0 if monthly else ""
        diff = wb - opt if opt is not None else None
        verdict = ("нет оптовой цены" if opt is None else
                   "не отдавать: сами продадим дороже" if diff > 0 and monthly else
                   "отдать: не продается" if not monthly else "отдать")
        rows.append([item["Наименование"], int(item["Остаток, шт"]), int(take),
                     int(rest), cost, price, opt_sum, opt,
                     int(monthly) if monthly else 0,
                     unit if pd.notna(unit) else round(average, 1),
                     wb, diff, months, verdict])
    return pd.DataFrame(rows, columns=SHOWN)


def write(book, title, rows, total=None):
    ws = book.create_sheet(title)
    ws.append(SHOWN)
    for cell in ws[1]:
        cell.fill, cell.font = HEAD, WHITE
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    for row in rows:
        ws.append(row)
        if str(row[-1]).startswith("не отдавать"):
            for cell in ws[ws.max_row]:
                cell.fill = STOP
    if total:
        ws.append(total)
        for cell in ws[ws.max_row]:
            cell.fill, cell.font = TOTAL, Font(bold=True)
    for index, width in enumerate(WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width
    for letter in "BCDGHIKL":
        for cell in ws[letter][1:]:
            cell.number_format = "# ##0"
    for letter in "EFJ":
        for cell in ws[letter][1:]:
            cell.number_format = "# ##0.00"
    ws.freeze_panes = "B2"
    return ws


def main(source, sales_path, target):
    report = read_margin(sales_path)
    sold = report[report["Продано, шт"] > 0]
    average = sold["Маржа, руб"].sum() / sold["Продано, шт"].sum()

    table = rows_of(build(read_order(source), report), average)
    priced = table[table["Прибыль опт, руб"].notna()]
    держим = table[table["Вывод"].str.startswith("не отдавать")]
    отдаем = table[~table["Вывод"].str.startswith("не отдавать")]

    book = Workbook()
    book.remove(book.active)
    свод = book.create_sheet("ИТОГО")
    for row in [
        ["Показатель", "Значение"],
        ["Позиций в заказе", len(table)],
        ["Штук просит", int(table["Просит, шт"].sum())],
        ["Позиций с оптовой ценой", len(priced)],
        ["Штук с оптовой ценой", int(priced["Просит, шт"].sum())],
        ["", ""],
        # Дальше три строки об одном и том же товаре — только о позициях,
        # по которым есть оптовая цена, иначе сравнение будет нечестным.
        ["Выручка оптом, руб", int(priced["Выручка опт, руб"].sum())],
        ["Себестоимость этого товара, руб",
         round((priced["Себестоимость, руб"] * priced["Просит, шт"]).sum())],
        ["Прибыль оптом, руб", int(priced["Прибыль опт, руб"].sum())],
        ["Прибыль на WB с того же товара, руб",
         int(priced["Прибыль WB на это кол-во, руб"].sum())],
        ["Разница в пользу WB, руб", int(priced["Разница, руб"].sum())],
        ["", ""],
        ["Позиций отдать оптом", len(отдаем)],
        ["Позиций придержать", len(держим)],
        ["Прибыль оптом по тем, что отдаем, руб",
         int(отдаем["Прибыль опт, руб"].sum(skipna=True))],
    ]:
        свод.append(row)
    for cell in свод[1]:
        cell.fill, cell.font = HEAD, WHITE
    for index, width in enumerate([40, 20], start=1):
        свод.column_dimensions[get_column_letter(index)].width = width
    for row in свод.iter_rows(min_row=2):
        row[1].number_format = "# ##0"

    order = table.sort_values("Разница, руб", ascending=False,
                              key=lambda col: col.fillna(-10 ** 12))
    write(book, "ЗАКАЗ", order.values.tolist(),
          ["ИТОГО", int(table["Остаток, шт"].sum()), int(table["Просит, шт"].sum()),
           int(table["Останется, шт"].sum()), "", "",
           int(priced["Выручка опт, руб"].sum()), int(priced["Прибыль опт, руб"].sum()),
           int(table["Продажи в месяц, шт"].sum()), "",
           int(table["Прибыль WB на это кол-во, руб"].sum()),
           int(priced["Разница, руб"].sum()), "", ""])
    write(book, "ПРИДЕРЖАТЬ", держим.values.tolist())
    write(book, "МОЖНО ОТДАТЬ", отдаем.values.tolist())

    os.makedirs("outputs", exist_ok=True)
    book.save(target)
    print(f"Позиций в заказе: {len(table)}   штук: {int(table['Просит, шт'].sum()):,}"
          .replace(",", " "))
    print(f"Прибыль оптом:   {int(priced['Прибыль опт, руб'].sum()):,} руб".replace(",", " "))
    print(f"Прибыль на WB:   {int(priced['Прибыль WB на это кол-во, руб'].sum()):,} руб "
          f"(по тем же {len(priced)} позициям)".replace(",", " "))
    print(f"Придержать: {len(держим)} позиций   отдать: {len(отдаем)}")
    print(f"Сохранено: {target}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Заказ покупателя: опт или WB")
    parser.add_argument("source")
    parser.add_argument("--sales", required=True)
    parser.add_argument("--out", default=TARGET)
    args = parser.parse_args()
    main(args.source, args.sales, args.out)
