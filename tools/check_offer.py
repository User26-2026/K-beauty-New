"""Предложение поставщика против наших цен и прайсов конкурентов.

Поставщик прислал инвойс с ценами по нашей заявке. Смотрим три вещи:
во сколько позиция обойдется на складе в Москве, дешевле ли это нашей
текущей себестоимости и не дают ли ту же позицию дешевле другие
корейские поставщики.

Сверку с чужими прайсами ведем по штрихкоду: он есть в инвойсе, и это
единственный надежный ключ между поставщиками.

Запуск:
    python3 tools/check_offer.py <инвойс.xls> --stock <остатки.xlsx>
"""

import argparse
import os
import sys

import pandas as pd
from openpyxl import Workbook
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import name_match
from brand_names import brand_of, same
from check_customer_order import read_order
from rates import KRW_RUB

TARGET = "outputs/Предложение поставщика.xlsx"
PRICES = "outputs/prices_normalized.xlsx"
FIRST_ROW = 26          # с какой строки инвойса идут товары
IMPORT_COST = 1.10      # наши расходы на ввоз из Кореи

HEAD = PatternFill("solid", fgColor="1F3864")
BRAND = PatternFill("solid", fgColor="DDEBF7")
TOTAL = PatternFill("solid", fgColor="E2EFDA")
GOOD = PatternFill("solid", fgColor="E2EFDA")
BAD = PatternFill("solid", fgColor="FCE4D6")
WHITE = Font(color="FFFFFF", bold=True)

SHOWN = ["№", "Бренд", "Товар", "Штрихкод", "Кол-во, шт", "Цена, KRW",
         "Цена, руб", "С ввозом, руб", "Наша себестоимость, руб",
         "Разница, руб", "Разница, %", "Дешевле в Корее, KRW", "У кого", "Вывод"]
WIDTHS = [5, 15, 52, 15, 11, 11, 10, 13, 17, 12, 11, 15, 15, 26]


def read_offer(path):
    table = pd.read_excel(path, sheet_name=0, header=None).iloc[FIRST_ROW:]
    table = table[[1, 2, 5, 6, 12]]
    table.columns = ["Бренд", "Товар", "Кол-во, шт", "Цена, KRW", "Штрихкод"]
    table["Бренд"] = table["Бренд"].ffill()
    table = table[table["Товар"].notna()]
    table = table[table["Бренд"].astype(str).str.strip() != "Total"]
    for column in ("Кол-во, шт", "Цена, KRW"):
        table[column] = pd.to_numeric(table[column], errors="coerce")
    table = table[table["Цена, KRW"].notna() & (table["Цена, KRW"] > 0)]
    table["Штрихкод"] = (table["Штрихкод"].astype(str)
                         .str.replace(r"\D", "", regex=True))
    table["Бренд"] = table["Бренд"].map(same)
    return table.reset_index(drop=True)


def our_costs(stock_path):
    """Себестоимость по названию: в колонке С/с лежит цена с наценкой 10%."""
    stock = read_order(stock_path, everything=True)
    return stock.assign(**{"Себестоимость, руб": stock["Себестоимость, руб"] / 1.1})


def others(barcodes):
    """Самая дешевая корейская цена по штрихкоду, кроме этого поставщика."""
    table = pd.read_excel(PRICES, dtype={"Штрихкод": str})
    table = table[(table["Страна"] == "KR") & table["Штрихкод"].notna()].copy()
    table = table[table["Закупка, KRW"].notna() & (table["Закупка, KRW"] > 0)]
    table["Штрихкод"] = table["Штрихкод"].str.replace(r"\D", "", regex=True)
    table = table[table["Штрихкод"].isin(barcodes)]
    best = table.sort_values("Закупка, KRW").drop_duplicates("Штрихкод")
    return best.set_index("Штрихкод")[["Поставщик", "Закупка, KRW"]]


def build(offer, stock, cheaper):
    pairs = name_match.to_one(offer["Товар"], stock["Наименование"])
    rows = []
    for index, item in offer.iterrows():
        price = item["Цена, KRW"] * KRW_RUB
        landed = price * IMPORT_COST
        ours = (stock.at[pairs[index], "Себестоимость, руб"]
                if index in pairs else None)
        gap = round(landed - ours) if ours else ""
        percent = round(landed / ours * 100 - 100, 1) if ours else ""
        other = cheaper.loc[item["Штрихкод"]] if item["Штрихкод"] in cheaper.index else None
        rival = (round(other["Закупка, KRW"])
                 if other is not None and other["Закупка, KRW"] < item["Цена, KRW"]
                 else "")
        verdict = ("новая позиция" if not ours else
                   "дешевле нашей" if gap < 0 else "дороже нашей")
        if rival:
            verdict += ", есть дешевле в Корее"
        rows.append([0, item["Бренд"], item["Товар"], item["Штрихкод"],
                     int(item["Кол-во, шт"]), int(item["Цена, KRW"]), round(price),
                     round(landed), round(ours) if ours else "", gap, percent,
                     rival, other["Поставщик"] if rival else "", verdict])
    return pd.DataFrame(rows, columns=SHOWN)


def write(book, title, rows, total=None):
    ws = book.create_sheet(title)
    ws.append(SHOWN)
    for cell in ws[1]:
        cell.fill, cell.font = HEAD, WHITE
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    brand, number = None, 0
    for row in rows:
        if row[1] != brand:
            brand = row[1]
            ws.append([None, brand])
            for cell in ws[ws.max_row]:
                cell.fill, cell.font = BRAND, Font(bold=True)
        number += 1
        ws.append([number] + list(row[1:]))
        fill = (GOOD if str(row[13]).startswith("дешевле")
                else BAD if str(row[13]).startswith("дороже") else None)
        if fill:
            for cell in ws[ws.max_row]:
                cell.fill = fill
    if total:
        ws.append(total)
        for cell in ws[ws.max_row]:
            cell.fill, cell.font = TOTAL, Font(bold=True)
    for index, width in enumerate(WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width
    for letter in "EFGHIJL":
        for cell in ws[letter][1:]:
            cell.number_format = "# ##0"
    ws.freeze_panes = "C2"
    return ws


def main(source, stock_path, target):
    offer = read_offer(source)
    stock = our_costs(stock_path)
    table = build(offer, stock, others(set(offer["Штрихкод"])))
    table = table.sort_values(["Бренд", "Товар"])

    book = Workbook()
    book.remove(book.active)
    сумма = (table["С ввозом, руб"] * table["Кол-во, шт"]).sum()
    write(book, "ПРЕДЛОЖЕНИЕ", table.values.tolist(),
          ["", "ИТОГО", f"позиций: {len(table)}", "",
           int(table["Кол-во, шт"].sum()), "", "", round(сумма), "", "", "", "",
           "", ""])

    свод = book.create_sheet("ПО БРЕНДАМ")
    свод.append(["Бренд", "Позиций", "Кол-во, шт", "Сумма с ввозом, руб",
                 "Дешевле нашей", "Дороже нашей", "Новых позиций",
                 "Есть дешевле в Корее"])
    for cell in свод[1]:
        cell.fill, cell.font = HEAD, WHITE
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    for brand, part in table.groupby("Бренд", sort=True):
        свод.append([brand, len(part), int(part["Кол-во, шт"].sum()),
                     round((part["С ввозом, руб"] * part["Кол-во, шт"]).sum()),
                     int(part["Вывод"].str.startswith("дешевле").sum()),
                     int(part["Вывод"].str.startswith("дороже").sum()),
                     int(part["Вывод"].str.startswith("новая").sum()),
                     int((part["У кого"] != "").sum())])
    свод.append(["ВСЕГО", len(table), int(table["Кол-во, шт"].sum()), round(сумма),
                 int(table["Вывод"].str.startswith("дешевле").sum()),
                 int(table["Вывод"].str.startswith("дороже").sum()),
                 int(table["Вывод"].str.startswith("новая").sum()),
                 int((table["У кого"] != "").sum())])
    for cell in свод[свод.max_row]:
        cell.fill, cell.font = TOTAL, Font(bold=True)
    for index, width in enumerate([18, 10, 13, 19, 14, 14, 14, 18], start=1):
        свод.column_dimensions[get_column_letter(index)].width = width
    for letter in "BCDEFGH":
        for cell in свод[letter][1:]:
            cell.number_format = "# ##0"
    свод.freeze_panes = "A2"

    дороже = table[table["Вывод"].str.startswith("дороже")]
    if len(дороже):
        write(book, "ДОРОЖЕ НАШЕЙ ЦЕНЫ", дороже.values.tolist())
    рынок = table[table["У кого"] != ""]
    if len(рынок):
        write(book, "ЕСТЬ ДЕШЕВЛЕ В КОРЕЕ", рынок.values.tolist())

    book.move_sheet("ПО БРЕНДАМ", -(len(book.sheetnames) - 1))
    os.makedirs("outputs", exist_ok=True)
    book.save(target)
    print(f"Позиций: {len(table)}   штук: {int(table['Кол-во, шт'].sum()):,}"
          .replace(",", " "))
    print(f"Сумма с ввозом: {round(сумма):,} руб".replace(",", " "))
    print(f"Дешевле нашей: {int(table['Вывод'].str.startswith('дешевле').sum())}   "
          f"дороже: {int(table['Вывод'].str.startswith('дороже').sum())}   "
          f"новых: {int(table['Вывод'].str.startswith('новая').sum())}")
    print(f"Есть дешевле у других корейцев: {int((table['У кого'] != '').sum())}")
    print(f"Сохранено: {target}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Предложение поставщика")
    parser.add_argument("source")
    parser.add_argument("--stock", required=True)
    parser.add_argument("--out", default=TARGET)
    args = parser.parse_args()
    main(args.source, args.stock, args.out)
