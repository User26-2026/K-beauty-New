"""Остатки по брендам и позициям: сколько чего лежит и на какие деньги.

Количество берем из плана распределения — там свежий склад и артикул WB.
Себестоимость — из файла остатков, а где не нашлось, из отчета WB.

Скорость продаж сюда не выносим: после пожара на складах Wildberries
цифры продаж не показывают спрос. Ходовой товар отделяем от лежачего в
tools/illiquid_check.py, там для этого смотрят воронку по двум периодам.

Запуск:
    python3 tools/stock_by_brand.py
"""

import argparse
import os
import re

import openpyxl
import pandas as pd
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import name_match

PLAN = "data/shipments/2026-09_plan_raspredeleniya.xlsx"
STOCK = "data/stock_costs/Остатки_31.08.2026.xlsx"
SALES = "workspace/wb_audit/sales_profit_with_cogs_2026-04-01_2026-05-24.csv"
OUT = "outputs/stock_by_brand.xlsx"

HEADER_FILL = PatternFill("solid", fgColor="DDEBF7")
DEAD_FILL = PatternFill("solid", fgColor="FFC7CE")
HOT_FILL = PatternFill("solid", fgColor="C6EFCE")

# Один бренд пишут по-разному в трех документах — сводим к одному виду.
BRAND_FIX = {
    "FARM STAY": "FARMSTAY", "ROUND LAB": "ROUND LAB", "ETUDE HOUSE": "ETUDE",
    "DR.JART+": "DR. JART", "SOME BY MI": "SOME BY MI", "JMSOLUTION": "JMSOLUTION",
    "MANYO FACTORY": "MANYO", "I'M SORRY FOR MY SKIN": "IM SORRY FOR MY SKIN",
}
KNOWN = ["CELIMAX", "FARMSTAY", "ROUND LAB", "PETITFEE", "JIGOTT", "MANYO", "ETUDE",
         "ENOUGH", "JMSOLUTION", "ELIZAVECCA", "DEOPROCE", "LEBELAGE", "COSRX",
         "SKIN1004", "MEDICUBE", "DERMA FACTORY", "HEIMISH", "EKEL", "DR. JART",
         "SOME BY MI", "THE ORDINARY", "MASIL", "LANEIGE", "VT"]


def clean_brand(value):
    text = re.sub(r"\s+", " ", str(value or "").strip().upper())
    return BRAND_FIX.get(text, text)


def brand_from_name(name):
    text = str(name).upper()
    for brand in sorted(KNOWN, key=len, reverse=True):
        if text.startswith(brand) or f" {brand}" in text[:40]:
            return brand
    return text.split()[0] if text.split() else "БЕЗ БРЕНДА"


def read_plan(path):
    plan = pd.read_excel(path, header=0)
    plan = plan[plan["Артикул WB"].notna()].copy()
    plan["Артикул WB"] = pd.to_numeric(plan["Артикул WB"], errors="coerce").astype("Int64")
    for column in ("Склад, шт", "Резерв ВБ, шт", "Резерв Озон, шт",
                   "Предложить Егору, шт", "Егор уже брал, шт"):
        plan[column] = pd.to_numeric(plan[column], errors="coerce").fillna(0)
    return plan.reset_index(drop=True)


def read_stock(path):
    book = openpyxl.load_workbook(path, data_only=True)
    sheet = book[book.sheetnames[0]]
    rows = []
    for line in range(3, sheet.max_row + 1):
        name = sheet.cell(line, 1).value
        if not name:
            continue
        rows.append({"Товар в остатках": str(name).strip(),
                     "Себестоимость из остатков": sheet.cell(line, 4).value,
                     "Срок годности": sheet.cell(line, 5).value})
    book.close()
    return pd.DataFrame(rows)


BAD_CHARS = re.compile(r"[\\/*?:\[\]]")


def sheet_name(brand, used):
    """Имя вкладки: Excel не берет длиннее 31 символа и часть знаков."""
    name = BAD_CHARS.sub(" ", str(brand)).strip()[:31] or "БЕЗ БРЕНДА"
    if name in used:
        name = f"{name[:28]}_{len(used)}"
    used.add(name)
    return name


def brand_sheet(rows, money_total, units_total):
    """Лист одного бренда: его товары, деньги и доля в общем складе."""
    columns = ["Наименование", "Артикул WB", "Склад, шт", "Себестоимость, руб",
               "Остаток, руб", "Доля в бренде, %", "Доля в складе, %",
               "Приход в пути, шт", "Срок годности"]
    table = rows.copy()
    brand_money = table["Остаток, руб"].sum()
    table["Доля в бренде, %"] = (table["Остаток, руб"] / brand_money * 100).round(1) \
        if brand_money else 0.0
    table["Доля в складе, %"] = (table["Остаток, руб"] / money_total * 100).round(1)
    table = table[columns].sort_values("Остаток, руб", ascending=False)

    total = {column: None for column in columns}
    total.update({
        "Наименование": "ИТОГО ПО БРЕНДУ",
        "Склад, шт": table["Склад, шт"].sum(),
        "Остаток, руб": brand_money,
        "Доля в бренде, %": 100.0,
        "Доля в складе, %": round(brand_money / money_total * 100, 1),
        "Приход в пути, шт": table["Приход в пути, шт"].sum(),
        "Срок годности": f"{table['Склад, шт'].sum() / units_total * 100:.1f}% всех штук склада",
    })
    return pd.concat([table, pd.DataFrame([total])], ignore_index=True)


def main():
    plan = read_plan(PLAN)
    stock = read_stock(STOCK)
    # Из отчета WB берем только бренд и себестоимость: скорость продаж в
    # этот файл не выносим — после пожара на складах она не показательна.
    report = pd.read_csv(SALES, usecols=["Артикул WB", "Бренд", "cost"])
    report["Артикул WB"] = pd.to_numeric(report["Артикул WB"], errors="coerce").astype("Int64")
    report = report.groupby("Артикул WB", as_index=False).agg(
        **{"Себестоимость из отчета": ("cost", "max"), "Бренд из отчета": ("Бренд", "first")})

    table = plan.merge(report, on="Артикул WB", how="left")

    pairs = name_match.match(table["Наименование"], stock["Товар в остатках"])
    table["_pair"] = table.index.map(pairs.get)
    table = table.merge(stock, left_on="_pair", right_index=True, how="left").drop(columns="_pair")

    table["Себестоимость, руб"] = pd.to_numeric(
        table["Себестоимость из остатков"], errors="coerce").fillna(
        pd.to_numeric(table["Себестоимость из отчета"], errors="coerce"))
    table["Бренд"] = table["Бренд из отчета"].map(clean_brand).replace("", pd.NA)
    table["Бренд"] = table["Бренд"].fillna(table["Наименование"].map(brand_from_name))
    table.loc[~table["Бренд"].isin(KNOWN), "Бренд"] = \
        table.loc[~table["Бренд"].isin(KNOWN), "Наименование"].map(brand_from_name)

    table["Остаток, руб"] = (table["Склад, шт"] * table["Себестоимость, руб"]).round(0)
    table["Приход в пути, шт"] = (pd.to_numeric(table["Машина"], errors="coerce").fillna(0)
                                  + pd.to_numeric(table["Контейнер"], errors="coerce").fillna(0))

    by_brand = (table.groupby("Бренд")
                .agg(**{"Позиций": ("Наименование", "size"),
                        "Штук на складе": ("Склад, шт", "sum"),
                        "Остаток, руб": ("Остаток, руб", "sum"),
                        "Приход в пути, шт": ("Приход в пути, шт", "sum")})
                .reset_index())
    by_brand["Доля остатка, %"] = (by_brand["Остаток, руб"] /
                                   by_brand["Остаток, руб"].sum() * 100).round(1)
    by_brand["Доля по штукам, %"] = (by_brand["Штук на складе"] /
                                     by_brand["Штук на складе"].sum() * 100).round(1)
    by_brand = by_brand[["Бренд", "Позиций", "Штук на складе", "Доля по штукам, %",
                         "Остаток, руб", "Доля остатка, %", "Приход в пути, шт"]]
    by_brand = by_brand.sort_values("Остаток, руб", ascending=False)

    columns = ["Бренд", "Наименование", "Артикул WB", "Склад, шт", "Себестоимость, руб",
               "Остаток, руб", "Приход в пути, шт", "Резерв ВБ, шт",
               "Резерв Озон, шт", "Егор уже брал, шт", "Срок годности"]
    items = table[columns].sort_values("Остаток, руб", ascending=False)

    total = pd.DataFrame([
        {"Показатель": "Позиций на складе", "Значение": len(items)},
        {"Показатель": "Штук", "Значение": int(items["Склад, шт"].sum())},
        {"Показатель": "Остаток по себестоимости, руб", "Значение": round(items["Остаток, руб"].sum())},
        {"Показатель": "Брендов", "Значение": int(items["Бренд"].nunique())},
        {"Показатель": "Позиций без себестоимости",
         "Значение": int(items["Себестоимость, руб"].isna().sum())},
        {"Показатель": "Приход в пути, шт", "Значение": int(items["Приход в пути, шт"].sum())},
    ])

    money_total = float(items["Остаток, руб"].sum())
    units_total = float(items["Склад, шт"].sum())

    with pd.ExcelWriter(OUT) as writer:
        total.to_excel(writer, sheet_name="ИТОГО", index=False)
        by_brand.to_excel(writer, sheet_name="ПО БРЕНДАМ", index=False)
        # Дальше вкладка на каждый бренд, по убыванию денег в остатке.
        used = {"ИТОГО", "ПО БРЕНДАМ"}
        for brand in by_brand["Бренд"]:
            rows = items[items["Бренд"] == brand]
            brand_sheet(rows, money_total, units_total).to_excel(
                writer, sheet_name=sheet_name(brand, used), index=False)
        items.to_excel(writer, sheet_name="ПО ПОЗИЦИЯМ", index=False)

        book = writer.book
        for name in book.sheetnames:
            sheet = book[name]
            sheet.freeze_panes = "B2"
            titles = [str(c.value or "") for c in sheet[1]]
            for index, title in enumerate(titles, start=1):
                cell = sheet.cell(row=1, column=index)
                cell.font = Font(bold=True)
                cell.fill = HEADER_FILL
                cell.alignment = Alignment(wrap_text=True, vertical="center")
                sheet.column_dimensions[get_column_letter(index)].width = (
                    58 if title in ("Наименование", "Показатель") else
                    26 if title == "Бренд" else 13)
                if "%" in title:
                    for row in sheet.iter_rows(min_row=2, min_col=index, max_col=index):
                        row[0].number_format = "0.0"
                if "руб" in title or "шт" in title:
                    for row in sheet.iter_rows(min_row=2, min_col=index, max_col=index):
                        row[0].number_format = "# ##0"
            # Строку ИТОГО в листе бренда выделяем жирным.
            last = sheet.cell(row=sheet.max_row, column=1).value
            if isinstance(last, str) and last.startswith("ИТОГО"):
                for cell in sheet[sheet.max_row]:
                    cell.font = Font(bold=True)
            if "Доля остатка, %" in titles:
                letter = get_column_letter(titles.index("Доля остатка, %") + 1)
                sheet.conditional_formatting.add(
                    f"{letter}2:{letter}{sheet.max_row}",
                    ColorScaleRule(start_type="num", start_value=0, start_color="FFFFFF",
                                   mid_type="num", mid_value=10, mid_color="FFEB84",
                                   end_type="num", end_value=30, end_color="F8696B"))

    print(total.to_string(index=False))
    print("\nПо брендам:")
    print(by_brand.to_string(index=False))
    print(f"\nФайл: {OUT}")


if __name__ == "__main__":
    argparse.ArgumentParser(description="Остатки по брендам и ходовость").parse_args()
    main()
