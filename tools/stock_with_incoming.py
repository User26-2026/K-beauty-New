"""Склад вместе с товаром в пути: что продавать самим, а что отдавать.

Решение по опту нельзя принимать по одному складу: контейнер и машина
привозят те же самые позиции, и по ним запас после прихода становится
многолетним. Поэтому складываем остаток с товаром в пути, делим на
продажи с учетом сезона и смотрим, сколько месяцев торговли выходит.

Все, что сверх горизонта (--keep месяцев), — излишек: его и надо
отдавать оптом, а не то, что кончается в ноябре.

Запуск:
    python3 tools/stock_with_incoming.py --sales <отчет WB.xls> \
        --container <инвойс.xlsx> --keep 12
"""

import argparse
import os
import sys

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import name_match
from add_sales_price import normal
from rates import KRW_RUB
from sales_vs_stock import read_report
from stock_forecast import run_out

TARGET = "outputs/Склад с товаром в пути.xlsx"
# Машина, инвойс VEC-06 от 23.06.2026: цена за штуку в вонах.
TRUCK = [
    ("CELIMAX DUAL BARRIER CREAMY TONER 150ml", 15000, 10060),
    ("ENOUGH Premium Ultra X10 cover up Collagen foundation #13", 1000, 3880),
    ("ENOUGH Premium Rich Gold Double Wear Radiance Foundation #13", 200, 4006),
    ("ENOUGH Premium Rich Gold Double Wear Radiance Foundation #21", 400, 4006),
    ("ENOUGH Gold snail moisture foundation #13", 700, 3190),
    ("ENOUGH Gold snail moisture foundation #21", 400, 3094),
]
CONTAINER_COST = 1.30    # контейнер: доставка, пошлина и приемка к цене EXW
TRUCK_COST = 1.10        # машина: 1 млн руб на всю партию, это около 10%

HEAD = PatternFill("solid", fgColor="1F3864")
TOTAL = PatternFill("solid", fgColor="E2EFDA")
EXTRA = PatternFill("solid", fgColor="FCE4D6")
WHITE = Font(color="FFFFFF", bold=True)

SHOWN = ["Товар", "Остаток, шт", "Контейнер, шт", "Машина, шт", "Всего будет, шт",
         "Продажи в месяц, шт", "Хватит на, мес", "Излишек, шт",
         "Излишек в закупке, руб", "Вывод"]
WIDTHS = [62, 12, 14, 12, 15, 15, 13, 13, 18, 30]


def read_container(path):
    table = pd.read_excel(path, sheet_name=0, header=None).iloc[3:]
    table = table[[2, 4, 5]]
    table.columns = ["Товар", "Кол-во", "Цена KRW"]
    table = table[table["Товар"].notna()]
    for column in ("Кол-во", "Цена KRW"):
        table[column] = pd.to_numeric(table[column], errors="coerce")
    return table.dropna(subset=["Кол-во", "Цена KRW"]).reset_index(drop=True)


def truck():
    return pd.DataFrame(TRUCK, columns=["Товар", "Кол-во", "Цена KRW"])


def join(base, incoming, column):
    """Приход к карточкам кабинета по названию: общего кода у них нет."""
    by_text = {}
    for index, name in base["Товар"].items():
        by_text.setdefault(normal(name), index)
    found = pd.Series([by_text.get(normal(name)) for name in incoming["Товар"]],
                      index=incoming.index, dtype="object")
    left = incoming.loc[found.isna(), "Товар"]
    free = base.drop(index=[i for i in found.dropna()])
    for index, other in name_match.match(left, free["Товар"], loose=True).items():
        found[index] = other

    base[column] = 0.0
    base[f"{column} цена"] = 0.0
    lost = []
    for index, item in incoming.iterrows():
        target = found[index]
        if pd.isna(target):
            lost.append(item)
            continue
        base.at[target, column] += item["Кол-во"]
        base.at[target, f"{column} цена"] = item["Цена KRW"]
    return base, pd.DataFrame(lost)


def build(report, container, keep, markup):
    base = report.rename(columns={"Продано за месяц, шт": "Продажи в месяц, шт"})
    base = base[["Товар", "Остаток, шт", "Продажи в месяц, шт"]].copy()
    base, lost_box = join(base, container, "Контейнер, шт")
    base, lost_truck = join(base, truck(), "Машина, шт")

    rows = []
    for _, item in base.iterrows():
        total = (item["Остаток, шт"] + item["Контейнер, шт"] + item["Машина, шт"])
        monthly = item["Продажи в месяц, шт"]
        if not total:
            continue
        if monthly:
            _, months, _ = run_out(total, monthly)
            months = round(months, 1)
            # Сколько надо оставить, чтобы хватило на горизонт планирования.
            need, year, month = 0.0, 2026, 9
            from stock_forecast import SEASON
            for _ in range(keep):
                need += monthly * SEASON.get(month, 1.0)
                year, month = (year + 1, 1) if month == 12 else (year, month + 1)
            spare = max(0, round(total - need))
        else:
            months, spare = "", int(total)
        # Излишек оцениваем по цене прихода, а где ее нет — по контейнеру.
        price = item["Контейнер, шт цена"] or item["Машина, шт цена"]
        cost = (item["Контейнер, шт цена"] * CONTAINER_COST
                or item["Машина, шт цена"] * TRUCK_COST) * KRW_RUB
        verdict = ("лежит без продаж" if not monthly else
                   "запас в норме" if not spare else
                   f"излишек: хватит на {months} мес")
        rows.append([item["Товар"], int(item["Остаток, шт"]),
                     int(item["Контейнер, шт"]), int(item["Машина, шт"]),
                     int(total), int(monthly), months, spare,
                     round(spare * cost) if price else "", verdict])
    return pd.DataFrame(rows, columns=SHOWN), pd.concat([lost_box, lost_truck])


def write(book, title, rows, total=None):
    ws = book.create_sheet(title)
    ws.append(SHOWN)
    for cell in ws[1]:
        cell.fill, cell.font = HEAD, WHITE
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    for row in rows:
        ws.append(row)
        if str(row[-1]).startswith("излишек"):
            for cell in ws[ws.max_row]:
                cell.fill = EXTRA
    if total:
        ws.append(total)
        for cell in ws[ws.max_row]:
            cell.fill, cell.font = TOTAL, Font(bold=True)
    for index, width in enumerate(WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width
    for letter in "BCDEFHI":
        for cell in ws[letter][1:]:
            cell.number_format = "# ##0"
    ws.freeze_panes = "B2"
    return ws


def main(sales_path, container_path, keep, target):
    report = read_report(sales_path)
    table, lost = build(report, read_container(container_path), keep, CONTAINER_COST)
    live = table[(table["Всего будет, шт"] > 0)]
    излишек = live[live["Излишек, шт"] > 0].sort_values("Излишек, шт", ascending=False)

    book = Workbook()
    book.remove(book.active)
    write(book, "СКЛАД И ПРИХОД",
          live.sort_values("Хватит на, мес",
                           key=lambda col: col.replace("", 10 ** 6),
                           ascending=False).values.tolist(),
          ["ИТОГО", int(live["Остаток, шт"].sum()), int(live["Контейнер, шт"].sum()),
           int(live["Машина, шт"].sum()), int(live["Всего будет, шт"].sum()),
           int(live["Продажи в месяц, шт"].sum()), "", int(live["Излишек, шт"].sum()),
           int(pd.to_numeric(live["Излишек в закупке, руб"], errors="coerce").sum()), ""])
    write(book, "ОТДАТЬ ОПТОМ", излишек.values.tolist())

    os.makedirs("outputs", exist_ok=True)
    book.save(target)
    print(f"Позиций: {len(live)}   на складе {int(live['Остаток, шт'].sum()):,} шт   "
          f"в пути {int(live['Контейнер, шт'].sum() + live['Машина, шт'].sum()):,} шт"
          .replace(",", " "))
    print(f"Излишек сверх {keep} месяцев: {int(излишек['Излишек, шт'].sum()):,} шт "
          f"на {int(pd.to_numeric(излишек['Излишек в закупке, руб'], errors='coerce').sum()):,} руб "
          f"в закупке".replace(",", " "))
    if len(lost):
        print(f"Не легло на карточки: {len(lost)} позиций прихода")
        for _, item in lost.head(8).iterrows():
            print(f"   {item['Товар'][:60]}  {int(item['Кол-во'])} шт")
    print(f"Сохранено: {target}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Склад вместе с товаром в пути")
    parser.add_argument("--sales", default="data/sales/wb_sales_2026-08.xls")
    parser.add_argument("--container", required=True)
    parser.add_argument("--keep", type=int, default=12,
                        help="на сколько месяцев торговли оставляем товар")
    parser.add_argument("--out", default=TARGET)
    args = parser.parse_args()
    main(args.sales, args.container, args.keep, args.out)
