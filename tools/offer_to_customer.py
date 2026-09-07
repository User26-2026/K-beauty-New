"""Предложение покупателю по уже отправленным ценам.

Файл уходит покупателю, поэтому в нем нет ни себестоимости, ни нашей
прибыли, ни остатков — только позиции, количества и цены прайса.

Что можем отдать: складываем остаток с товаром в пути, вычитаем то, что
нужно нам самим на горизонт планирования с учетом сезона, и остаток
предлагаем. Отгрузить можем только то, что уже лежит на складе, поэтому
предложение ограничено остатком, а не будущим приходом.

Второй лист — по его заявке: его позиции и его количества, а рядом
сколько за тот же товар мы получаем на маркетплейсе после комиссии и
логистики. Разница между двумя суммами и есть то, что мы теряем на
сделке.

Запуск:
    python3 tools/offer_to_customer.py <заказ.xlsx> --sales <отчет WB.xls> \
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
from add_sales_price import attach, normal, read_prices, sales_by_stock
from check_customer_order import read_order
from stock_forecast import SEASON, START, run_out
from stock_with_incoming import read_container, truck
from wholesale_vs_wb import MARGIN, SKIP

TARGET = "outputs/Предложение покупателю.xlsx"

HEAD = PatternFill("solid", fgColor="1F3864")
TOTAL = PatternFill("solid", fgColor="E2EFDA")
ASKED = PatternFill("solid", fgColor="FFF2CC")
WHITE = Font(color="FFFFFF", bold=True)

OFFER = ["№", "Наименование", "Количество, шт", "Цена, руб", "Сумма, руб"]
OFFER_WIDTHS = [5, 88, 15, 12, 15]
ORDER = ["№", "Наименование", "Количество, шт", "Цена для вас, руб",
         "Сумма для вас, руб", "Выручка на маркетплейсе за штуку, руб",
         "Она же на это количество, руб", "Разница по выручке, руб"]
# Выручка маркетплейса — это уже за вычетом его комиссии и логистики, но
# до хранения, рекламы и налога. Пишем это в файле, чтобы разницу не
# читали как разницу в прибыли.
NOTE = ("Выручка на маркетплейсе указана после комиссии площадки и логистики, "
        "но до хранения, рекламы и налога.")
ORDER_WIDTHS = [5, 76, 15, 14, 16, 18, 18, 15]
# Внутренний лист: он остается у нас, поэтому здесь видно и остаток, и
# сколько останется после отгрузки, и на сколько этого хватит.
INSIDE = ["№", "Наименование", "Остаток, шт", "В пути, шт", "Продажи в месяц, шт",
          "Можем отдать, шт", "Останется у нас, шт", "Нам хватит на, мес",
          "Цена, руб", "Сумма, руб", "Прибыль, руб"]
INSIDE_WIDTHS = [5, 76, 12, 12, 15, 14, 15, 14, 11, 14, 14]


def read_net(path):
    """Сколько нам приходит с маркетплейса после комиссии и логистики."""
    columns = {1: "Товар", 15: "Продано, шт", 29: "К нам, руб"}
    table = pd.read_excel(path, sheet_name=0, header=None).iloc[SKIP:]
    table = table[list(columns)].rename(columns=columns)
    table = table[table["Товар"].notna()]
    for column in ("Продано, шт", "К нам, руб"):
        table[column] = pd.to_numeric(table[column], errors="coerce").fillna(0)
    return table.reset_index(drop=True)


def incoming(base):
    """Товар в пути к позициям остатков: общего кода нет, сводим по названию."""
    coming = pd.concat([read_container(PATHS["container"]), truck()], ignore_index=True)
    # Один товар приезжает и контейнером, и машиной — обе строки должны
    # лечь на одну позицию, поэтому сводим many-to-one.
    pairs = name_match.to_one(coming["Товар"], base["Наименование"])
    totals = pd.Series(0.0, index=base.index)
    for index, item in coming.iterrows():
        if index in pairs:
            totals[pairs[index]] += item["Кол-во"]
    return totals


def need_for(monthly, keep):
    """Сколько штук нужно нам самим на horizon месяцев с учетом сезона."""
    total, month = 0.0, START[1]
    for _ in range(keep):
        total += monthly * SEASON.get(month, 1.0)
        month = 1 if month == 12 else month + 1
    return total


def build(order, report, keep):
    base = attach(order, read_prices())
    found = sales_by_stock(base, report.rename(columns={"К нам, руб": "Выручка, руб"}))
    base["Продажи в месяц, шт"] = found["Продано, шт"]
    # Выручка к перечислению на штуку — цена сравнения, без себестоимости.
    base["Выручка на штуку, руб"] = [
        round(money / sold) if sold else None
        for money, sold in zip(found["Выручка, руб"], found["Продано, шт"])
    ]
    base["В пути, шт"] = incoming(base)

    offer, ask, inside = [], [], []
    for _, item in base.iterrows():
        rest, coming = item["Остаток, шт"], item["В пути, шт"]
        monthly, price = item["Продажи в месяц, шт"], item["Продажная цена, руб"]
        if pd.isna(price):
            continue
        give = int(max(0, min((rest + coming) - need_for(monthly, keep), rest)))
        if give:
            offer.append([len(offer) + 1, item["Наименование"], give, price,
                          round(price * give)])
            left = rest - give
            if monthly:
                _, months, _ = run_out(left + coming, monthly)
                months = round(months, 1)
            else:
                months = ""
            inside.append([len(inside) + 1, item["Наименование"], int(rest),
                           int(coming), int(monthly), give, int(left), months,
                           price, round(price * give),
                           round((price - item["Себестоимость, руб"]) * give)])
        want = int(item["Просит, шт"])
        net = item["Выручка на штуку, руб"]
        if want and pd.notna(net):
            ask.append([len(ask) + 1, item["Наименование"], want, price,
                        round(price * want), net, round(net * want),
                        round((net - price) * want)])
    return (pd.DataFrame(offer, columns=OFFER),
            pd.DataFrame(ask, columns=ORDER),
            pd.DataFrame(inside, columns=INSIDE))


def write(book, title, columns, rows, widths, total, money, price, note=None):
    ws = book.create_sheet(title)
    ws.append(columns)
    for cell in ws[1]:
        cell.fill, cell.font = HEAD, WHITE
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    for row in rows:
        ws.append(row)
    ws.append(total)
    for cell in ws[ws.max_row]:
        cell.fill, cell.font = TOTAL, Font(bold=True)
    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width
    for letter in money:
        for cell in ws[letter][1:]:
            cell.number_format = "# ##0"
    for letter in price:
        for cell in ws[letter][1:]:
            cell.number_format = "# ##0.00"
    if note:
        ws.append([])
        ws.append(["", note])
    ws.freeze_panes = "B2"
    return ws


PATHS = {}


def main(source, sales_path, container_path, keep, target, list_only=False,
         everything=False):
    PATHS["container"] = container_path
    offer, ask, inside = build(read_order(source, everything), read_net(sales_path), keep)
    offer = offer.sort_values("Сумма, руб", ascending=False)
    ask = ask.sort_values("Разница по выручке, руб", ascending=False)
    offer["№"] = range(1, len(offer) + 1)
    ask["№"] = range(1, len(ask) + 1)

    book = Workbook()
    book.remove(book.active)
    write(book, "ЧТО МОЖЕМ ОТГРУЗИТЬ", OFFER, offer.values.tolist(), OFFER_WIDTHS,
          ["", "ИТОГО", int(offer["Количество, шт"].sum()), "",
           int(offer["Сумма, руб"].sum())], "CE", "D")
    if everything:
        inside = inside.sort_values("Сумма, руб", ascending=False)
        inside["№"] = range(1, len(inside) + 1)
        write(book, "ЧТО ОСТАНЕТСЯ У НАС", INSIDE, inside.values.tolist(),
              INSIDE_WIDTHS,
              ["", "ИТОГО", int(inside["Остаток, шт"].sum()),
               int(inside["В пути, шт"].sum()), int(inside["Продажи в месяц, шт"].sum()),
               int(inside["Можем отдать, шт"].sum()),
               int(inside["Останется у нас, шт"].sum()), "", "",
               int(inside["Сумма, руб"].sum()), int(inside["Прибыль, руб"].sum())],
              "CDEFGJK", "I")
    if not list_only:
        write(book, "ВАША ЗАЯВКА", ORDER, ask.values.tolist(), ORDER_WIDTHS,
               ["", "ИТОГО", int(ask["Количество, шт"].sum()), "",
                int(ask["Сумма для вас, руб"].sum()), "",
                int(ask["Она же на это количество, руб"].sum()),
                int(ask["Разница по выручке, руб"].sum())], "CEGH", "DF", NOTE)

    os.makedirs("outputs", exist_ok=True)
    book.save(target)
    print(f"Можем отгрузить: {len(offer)} позиций, "
          f"{int(offer['Количество, шт'].sum()):,} шт на "
          f"{int(offer['Сумма, руб'].sum()):,} руб".replace(",", " "))
    print(f"Сохранено: {target}")
    if list_only:
        return
    print(f"Его заявка: {len(ask)} позиций, {int(ask['Количество, шт'].sum()):,} шт")
    print(f"  по нашим ценам:   {int(ask['Сумма для вас, руб'].sum()):,} руб"
          .replace(",", " "))
    print(f"  на маркетплейсе:  {int(ask['Она же на это количество, руб'].sum()):,} руб"
          .replace(",", " "))
    print(f"  разница:          "
          f"{int(ask['Разница по выручке, руб'].sum()):,} руб".replace(",", " "))


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Предложение покупателю из излишка")
    parser.add_argument("source")
    parser.add_argument("--sales", default="data/sales/wb_sales_2026-08.xls")
    parser.add_argument("--container", required=True)
    parser.add_argument("--keep", type=int, default=12)
    parser.add_argument("--out", default=TARGET)
    parser.add_argument("--list-only", action="store_true",
                        help="только список товаров, без сравнения с маркетплейсом")
    parser.add_argument("--all", action="store_true", dest="everything",
                        help="весь склад, а не только позиции из заявки покупателя")
    args = parser.parse_args()
    main(args.source, args.sales, args.container, args.keep, args.out,
         args.list_only, args.everything)
