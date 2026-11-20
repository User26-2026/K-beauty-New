"""На сколько хватит остатков с учетом сезонного роста.

Базу берем из отчета WB за август. Дальше идем по месяцам и списываем со
склада продажи с коэффициентом сезона: ноябрь и декабрь торгуют больше
обычного, потом спрос возвращается к обычному.

Позиции без продаж в отчете не прогнозируем: у них нет базы, и деление
на ноль дало бы бесконечный запас там, где товар просто не продается.

Запуск:
    python3 tools/stock_forecast.py <остатки.xlsx> --sales <отчет WB.xls>
"""

import argparse
import os
import sys

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from add_sales_price import read_sales, read_stock, sales_by_stock

TARGET = "outputs/Прогноз остатков.xlsx"
START = (2026, 9)
# Сезон: ноябрь и март в полтора раза, декабрь и февраль вдвое,
# остальные месяцы обычные.
SEASON = {11: 1.5, 12: 2.0, 2: 2.0, 3: 1.5}
MONTHS = ["январь", "февраль", "март", "апрель", "май", "июнь", "июль",
          "август", "сентябрь", "октябрь", "ноябрь", "декабрь"]
# Сколько месяцев показываем колонками — до апреля включительно.
SHOWN_MONTHS = 8
HORIZON = 36

HEAD = PatternFill("solid", fgColor="1F3864")
TOTAL = PatternFill("solid", fgColor="E2EFDA")
SOON = PatternFill("solid", fgColor="F8696B")
LATER = PatternFill("solid", fgColor="FFEB84")
NONE_ = PatternFill("solid", fgColor="D9D9D9")
WHITE = Font(color="FFFFFF", bold=True)

def calendar(count):
    """Названия месяцев прогноза с коэффициентом сезона в заголовке."""
    year, month, names = *START, []
    for _ in range(count):
        rate = SEASON.get(month, 1.0)
        title = MONTHS[month - 1].capitalize()
        names.append(f"{title} х{rate:g}".replace(".", ",") if rate != 1 else title)
        year, month = (year + 1, 1) if month == 12 else (year, month + 1)
    return names, f"Остаток на 01.{month:02d}, шт"


PLAN, REST_COLUMN = calendar(SHOWN_MONTHS)
SHOWN = (["Наименование", "Остаток, шт", "Продажи в месяц, шт"] + PLAN
         + [REST_COLUMN, "Хватит до", "Запас, месяцев"])
MONTH_FIRST = 3                     # с какой колонки идут месяцы
MONTH_LAST = MONTH_FIRST + SHOWN_MONTHS


def pace(stock, sales):
    """Продажи за месяц к каждой позиции остатков."""
    return sales_by_stock(stock, sales)["Продано, шт"]


def run_out(rest, monthly):
    """Идем по месяцам, пока склад не кончится.

    Возвращаем расход первых четырех месяцев, остаток на новый год,
    месяц окончания и запас в месяцах с учетом сезона.
    """
    year, month = START
    spent, left, months = [], rest, 0
    for step in range(HORIZON):
        need = monthly * SEASON.get(month, 1.0)
        take = min(left, need)
        if step < SHOWN_MONTHS:
            spent.append(round(take))
        if need and left < need:
            # Внутри месяца товар кончается не в первый день.
            months += left / need
            left = 0
            return spent, months, f"{MONTHS[month - 1]} {year}"
        left -= take
        months += 1
        year, month = (year + 1, 1) if month == 12 else (year, month + 1)
    return spent, HORIZON, f"больше {HORIZON // 12} лет"


def build(stock, sales):
    stock = stock.copy()
    stock["Продажи в месяц, шт"] = pace(stock, sales)
    rows = []
    for _, item in stock.iterrows():
        monthly = item["Продажи в месяц, шт"]
        if not monthly:
            rows.append([item["Наименование"], int(item["Остаток, шт"]), 0,
                         *[""] * SHOWN_MONTHS, int(item["Остаток, шт"]),
                         "нет продаж в августе", ""])
            continue
        spent, months, until = run_out(item["Остаток, шт"], monthly)
        spent += [0] * (SHOWN_MONTHS - len(spent))
        rows.append([item["Наименование"], int(item["Остаток, шт"]), int(monthly),
                     *spent, int(max(item["Остаток, шт"] - sum(spent), 0)),
                     until, round(months, 1)])
    return pd.DataFrame(rows, columns=SHOWN)


def write(book, title, rows, total=None):
    ws = book.create_sheet(title)
    ws.append(SHOWN)
    for cell in ws[1]:
        cell.fill, cell.font = HEAD, WHITE
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    for row in rows:
        ws.append(row)
        months = row[-1]
        fill = (NONE_ if months == "" else
                SOON if months <= 4 else LATER if months <= 8 else None)
        if fill:
            for cell in ws[ws.max_row]:
                cell.fill = fill
    if total:
        ws.append(total)
        for cell in ws[ws.max_row]:
            cell.fill, cell.font = TOTAL, Font(bold=True)
    widths = [92, 12, 14] + [11] * SHOWN_MONTHS + [16, 22, 12]
    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width
    for index in range(2, MONTH_LAST + 2):
        for cell in ws[get_column_letter(index)][1:]:
            cell.number_format = "# ##0"
    ws.freeze_panes = "A2"
    return ws


def main(source, sales_path, target):
    table = build(read_stock(source), read_sales(sales_path))
    book = Workbook()
    book.remove(book.active)

    order = table.sort_values(
        "Запас, месяцев", key=lambda col: col.replace("", 10 ** 6))
    write(book, "ПРОГНОЗ", order.values.tolist(),
          ["ИТОГО", int(table["Остаток, шт"].sum()),
           int(table["Продажи в месяц, шт"].sum()),
           *[int(pd.to_numeric(table[column], errors="coerce").sum())
             for column in PLAN],
           int(table[REST_COLUMN].sum()), "", ""])

    сгорит = order[[value != "" and value <= SHOWN_MONTHS
                    for value in order["Запас, месяцев"]]]
    write(book, "КОНЧИТСЯ ДО МАЯ", сгорит.values.tolist())

    os.makedirs("outputs", exist_ok=True)
    book.save(target)
    known = table[table["Запас, месяцев"] != ""]
    print(f"Позиций: {len(table)}   с продажами в августе: {len(known)}   "
          f"без продаж: {len(table) - len(known)}")
    print(f"Кончится до мая: {len(сгорит)} позиций")
    print(f"Медиана запаса: {known['Запас, месяцев'].median():.1f} мес")
    print(f"Сохранено: {target}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Прогноз остатков по месяцам")
    parser.add_argument("source")
    parser.add_argument("--sales", required=True)
    parser.add_argument("--out", default=TARGET)
    args = parser.parse_args()
    main(args.source, args.sales, args.out)
