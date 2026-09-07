"""Продажи против остатков по всем товарам кабинета.

Отчет WB несет и продажи, и остаток по каждой карточке, поэтому берем
его целиком, а не только те позиции, что попали в складской файл. Дальше
списываем остаток по месяцам с сезонным коэффициентом и смотрим, на
сколько его хватит.

Пустые карточки — без остатка и без продаж — в расчет не идут, это
архив кабинета.

Запуск:
    python3 tools/sales_vs_stock.py --sales data/sales/wb_sales_2026-08.xls
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
from stock_forecast import (HEAD, LATER, MONTHS, NONE_, PLAN, REST_COLUMN, SEASON,
                            SHOWN_MONTHS, SOON, START, TOTAL, WHITE, run_out)

TARGET = "outputs/Продажи и остатки.xlsx"
# Колонки отчета WB: артикул, название, остатки FBO и FBS, продажи.
REPORT = {0: "Артикул ВБ", 1: "Товар", 7: "FBO", 8: "FBS",
          15: "Продано за месяц, шт", 18: "Выручка за месяц, руб"}
SKIP = 5
# В кабинете остались карточки с прошлого товара — новогодние фигурки и
# китайская электроника. К косметике они отношения не имеют и в анализ
# ассортимента не идут.
NOT_OURS = re.compile(r"^dm7dm|^СГ-|гирлянд|рация|триммер|скалер|перкуссионн",
                      re.IGNORECASE)

SHOWN = (["Артикул ВБ", "Товар", "Остаток, шт", "Продано за месяц, шт",
          "Выручка за месяц, руб"] + PLAN + [REST_COLUMN, "Хватит до", "Запас, месяцев"])
WIDTHS = [13, 62, 12, 15, 16] + [11] * SHOWN_MONTHS + [16, 22, 12]
LAST_NUMBER = 5 + SHOWN_MONTHS + 1     # по какую колонку идут числа


def read_report(path):
    table = pd.read_excel(path, sheet_name=0, header=None).iloc[SKIP:]
    table = table[list(REPORT)].rename(columns=REPORT)
    table = table[table["Товар"].notna()]
    for column in ("FBO", "FBS", "Продано за месяц, шт", "Выручка за месяц, руб"):
        table[column] = pd.to_numeric(table[column], errors="coerce").fillna(0)
    table["Остаток, шт"] = table["FBO"] + table["FBS"]
    table = table[~table["Товар"].astype(str).str.contains(NOT_OURS)]
    return table.reset_index(drop=True)


def forecast(table):
    rows = []
    for _, item in table.iterrows():
        rest, monthly = item["Остаток, шт"], item["Продано за месяц, шт"]
        head = [item["Артикул ВБ"], item["Товар"], int(rest), int(monthly),
                round(item["Выручка за месяц, руб"])]
        if not monthly:
            rows.append(head + [""] * SHOWN_MONTHS + [int(rest), "нет продаж", ""])
            continue
        if not rest:
            rows.append(head + [0] * SHOWN_MONTHS + [0, "закончился", 0])
            continue
        spent, months, until = run_out(rest, monthly)
        spent += [0] * (SHOWN_MONTHS - len(spent))
        rows.append(head + spent + [int(max(rest - sum(spent), 0)), until,
                                    round(months, 1)])
    return pd.DataFrame(rows, columns=SHOWN)


def demand(table):
    """Спрос по месяцам против того, что реально сможем отгрузить.

    Спрос — это продажи августа с сезонным коэффициентом, без оглядки на
    склад. Отгрузка ограничена остатком: когда позиция кончилась, дальше
    она не продается, и разрыв между двумя строками — упущенные продажи.
    """
    left = dict(zip(table.index, table["Остаток, шт"]))
    base = dict(zip(table.index, table["Продано за месяц, шт"]))
    year, month, rows = *START, []
    for _ in range(SHOWN_MONTHS):
        rate = SEASON.get(month, 1.0)
        need = sum(base[key] * rate for key in left)
        sell = sum(min(left[key], base[key] * rate) for key in left)
        for key in left:
            left[key] = max(0.0, left[key] - base[key] * rate)
        rows.append([f"{MONTHS[month - 1].capitalize()} {year}",
                     f"х{rate:g}".replace(".", ","), round(need), round(sell),
                     round(need - sell)])
        year, month = (year + 1, 1) if month == 12 else (year, month + 1)
    rows.append(["ИТОГО", "", sum(row[2] for row in rows),
                 sum(row[3] for row in rows), sum(row[4] for row in rows)])
    return rows


def write_demand(book, rows):
    ws = book.create_sheet("СПРОС ПРОТИВ ОТГРУЗКИ")
    ws.append(["Месяц", "Сезон", "Спрос, шт", "Отгрузим, шт", "Не хватит, шт"])
    for cell in ws[1]:
        cell.fill, cell.font = HEAD, WHITE
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    for row in rows:
        ws.append(row)
        if row[4] > 0:
            for cell in ws[ws.max_row][2:]:
                cell.fill = SOON if row[4] > row[3] else LATER
    for cell in ws[ws.max_row]:
        cell.fill, cell.font = TOTAL, Font(bold=True)
    for index, width in enumerate([18, 9, 13, 14, 15], start=1):
        ws.column_dimensions[get_column_letter(index)].width = width
    for letter in "CDE":
        for cell in ws[letter][1:]:
            cell.number_format = "# ##0"
    ws.freeze_panes = "A2"
    return ws


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
    for index, width in enumerate(WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width
    for index in range(3, LAST_NUMBER + 2):
        for cell in ws[get_column_letter(index)][1:]:
            cell.number_format = "# ##0"
    ws.freeze_panes = "C2"
    return ws


def main(sales_path, target):
    report = read_report(sales_path)
    live = report[(report["Остаток, шт"] > 0) | (report["Продано за месяц, шт"] > 0)]
    table = forecast(live)

    book = Workbook()
    book.remove(book.active)
    order = table.sort_values("Запас, месяцев",
                              key=lambda col: col.replace("", 10 ** 6))
    write(book, "ВСЕ ТОВАРЫ", order.values.tolist(),
          ["", "ИТОГО", int(table["Остаток, шт"].sum()),
           int(table["Продано за месяц, шт"].sum()),
           int(table["Выручка за месяц, руб"].sum()),
           *[int(pd.to_numeric(table[column], errors="coerce").sum()) for column in PLAN],
           int(table[REST_COLUMN].sum()), "", ""])

    скоро = order[[value != "" and value <= SHOWN_MONTHS
                   for value in order["Запас, месяцев"]]]
    write_demand(book, demand(live))
    write(book, "КОНЧИТСЯ ДО МАЯ", скоро.values.tolist())
    write(book, "НЕТ ПРОДАЖ",
          order[order["Хватит до"] == "нет продаж"].values.tolist())
    # Продажи были, а товара нет — это уже упущенная выручка.
    write(book, "ЗАКОНЧИЛСЯ",
          order[order["Хватит до"] == "закончился"].values.tolist())

    os.makedirs("outputs", exist_ok=True)
    book.save(target)
    known = table[table["Запас, месяцев"] != ""]
    print(f"Карточек в отчете: {len(report)}   в работе: {len(table)}   "
          f"пустых: {len(report) - len(table)}")
    print(f"Остаток: {int(table['Остаток, шт'].sum()):,} шт   "
          f"продано за месяц: {int(table['Продано за месяц, шт'].sum()):,} шт"
          .replace(",", " "))
    print(f"Кончится до мая: {len(скоро)}   без продаж: "
          f"{(table['Хватит до'] == 'нет продаж').sum()}   "
          f"закончилось: {(table['Хватит до'] == 'закончился').sum()}")
    if len(known):
        print(f"Медиана запаса: {known['Запас, месяцев'].median():.1f} мес")
    plan = demand(live)[-1]
    print(f"Спрос до апреля: {plan[2]:,} шт   отгрузим {plan[3]:,} шт   "
          f"упустим {plan[4]:,} шт".replace(",", " "))
    print(f"Сохранено: {target}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Продажи против остатков")
    parser.add_argument("--sales", default="data/sales/wb_sales_2026-08.xls")
    parser.add_argument("--out", default=TARGET)
    args = parser.parse_args()
    main(args.sales, args.out)
