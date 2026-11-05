"""Продать склад оптом или торговать на WB — что даст больше прибыли.

Опт считается просто: цена из нашего прайса минус себестоимость, деньги
сразу. Для WB берем фактическую экономику из отчета за месяц — что
осталось после комиссии, логистики, хранения, рекламы и налога, — и
умножаем на остаток. Разница между сценариями и есть цена ожидания.

Сроки тоже важны: WB платит больше, но растянуто. Поэтому рядом ставим,
за сколько месяцев склад разойдется при текущих продажах, и считаем
прибыль в месяц.

Запуск:
    python3 tools/wholesale_vs_wb.py <остатки.xlsx> --sales <отчет WB.xls>
"""

import argparse
import os
import sys

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from add_sales_price import attach, read_prices, read_sales, read_stock, sales_by_stock
from stock_forecast import run_out

TARGET = "outputs/Опт против WB.xlsx"
# Колонки отчета WB: название, продажи, маржа после всех расходов.
MARGIN = {1: "Товар", 15: "Продано, шт", 123: "Маржа, руб"}
SKIP = 5

HEAD = PatternFill("solid", fgColor="1F3864")
TOTAL = PatternFill("solid", fgColor="E2EFDA")
OPT = PatternFill("solid", fgColor="FCE4D6")
WHITE = Font(color="FFFFFF", bold=True)

SHOWN = ["Наименование", "Остаток, шт", "Себестоимость, руб", "Цена опт, руб",
         "Прибыль опт, руб", "Продажи в месяц, шт", "Маржа WB на штуку, руб",
         "Прибыль WB, руб", "Распродам за, мес", "Прибыль WB в месяц, руб",
         "Разница WB минус опт, руб", "Что делать"]


def read_margin(path):
    """Фактическая маржа WB по карточкам за месяц."""
    table = pd.read_excel(path, sheet_name=0, header=None).iloc[SKIP:]
    table = table[list(MARGIN)].rename(columns=MARGIN)
    table = table[table["Товар"].notna()]
    for column in ("Продано, шт", "Маржа, руб"):
        table[column] = pd.to_numeric(table[column], errors="coerce").fillna(0)
    return table.reset_index(drop=True)


def build(stock, report):
    """Сводим остатки с продажами и маржой WB по каждой позиции."""
    sales = report.rename(columns={"Маржа, руб": "Выручка, руб"})
    found = sales_by_stock(stock, sales)
    stock = stock.copy()
    stock["Продажи в месяц, шт"] = found["Продано, шт"]
    stock["Маржа WB всего, руб"] = found["Выручка, руб"]
    # Маржа на штуку — из фактических продаж месяца, а не из прайса.
    stock["Маржа WB на штуку, руб"] = [
        round(margin / sold, 1) if sold else None
        for margin, sold in zip(stock["Маржа WB всего, руб"], stock["Продажи в месяц, шт"])
    ]
    return stock


def rows_of(stock, average):
    rows = []
    for _, item in stock.iterrows():
        rest = item["Остаток, шт"]
        cost = item["Себестоимость, руб"]
        price = item["Продажная цена, руб"]
        opt = round((price - cost) * rest) if pd.notna(price) else None
        unit = item["Маржа WB на штуку, руб"]
        # Где карточка не продавалась, берем среднюю маржу по кабинету:
        # без нее позиция выпала бы из сравнения совсем.
        known = unit if pd.notna(unit) else average
        wb = round(known * rest)
        monthly = item["Продажи в месяц, шт"]
        if monthly:
            _, months, _ = run_out(rest, monthly)
            months = round(months, 1)
            per_month = round(wb / months) if months else wb
        else:
            months, per_month = "", ""
        diff = wb - opt if opt is not None else None
        verdict = ("нет оптовой цены" if opt is None else
                   "оптом" if diff <= 0 else
                   "оптом: не продается" if not monthly else
                   "оптом: слишком долго" if months > 12 else "на WB")
        rows.append([item["Наименование"], int(rest), cost, price, opt,
                     int(monthly) if monthly else 0,
                     unit if pd.notna(unit) else round(average, 1),
                     wb, months, per_month, diff, verdict])
    return pd.DataFrame(rows, columns=SHOWN)


def write(book, title, rows, widths, total=None):
    ws = book.create_sheet(title)
    ws.append(SHOWN)
    for cell in ws[1]:
        cell.fill, cell.font = HEAD, WHITE
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    for row in rows:
        ws.append(row)
        if str(row[-1]).startswith("оптом"):
            for cell in ws[ws.max_row]:
                cell.fill = OPT
    if total:
        ws.append(total)
        for cell in ws[ws.max_row]:
            cell.fill, cell.font = TOTAL, Font(bold=True)
    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width
    for letter in "BEFHJK":
        for cell in ws[letter][1:]:
            cell.number_format = "# ##0"
    for letter in "CDG":
        for cell in ws[letter][1:]:
            cell.number_format = "# ##0.00"
    ws.freeze_panes = "B2"
    return ws


def main(source, sales_path, target):
    report = read_margin(sales_path)
    sold = report[report["Продано, шт"] > 0]
    average = sold["Маржа, руб"].sum() / sold["Продано, шт"].sum()

    stock = build(attach(read_stock(source), read_prices()), report)
    table = rows_of(stock, average)
    priced = table[table["Прибыль опт, руб"].notna()]

    book = Workbook()
    book.remove(book.active)
    widths = [86, 12, 15, 13, 15, 14, 15, 15, 13, 16, 17, 20]

    # Смешанный сценарий: ходовое оставляем на WB, лежачее сдаем оптом.
    держим = table[table["Что делать"] == "на WB"]
    сдаем = table[table["Что делать"].str.startswith("оптом")]
    mixed = держим["Прибыль WB, руб"].sum() + сдаем["Прибыль опт, руб"].sum(skipna=True)

    свод = book.create_sheet("ИТОГО")
    склад = (stock["Себестоимость, руб"] * stock["Остаток, шт"]).sum()
    for row in [
        ["Показатель", "Опт целиком", "Только WB", "Смешанно"],
        ["Позиций", len(priced), len(table), f"{len(держим)} на WB + {len(сдаем)} оптом"],
        ["Штук", int(priced["Остаток, шт"].sum()), int(table["Остаток, шт"].sum()),
         int(table["Остаток, шт"].sum())],
        ["Себестоимость склада, руб", round(склад), round(склад), round(склад)],
        ["Прибыль, руб", int(priced["Прибыль опт, руб"].sum()),
         int(table["Прибыль WB, руб"].sum()), int(mixed)],
        ["Прибыль к себестоимости, %",
         round(priced["Прибыль опт, руб"].sum()
               / (priced["Себестоимость, руб"] * priced["Остаток, шт"]).sum() * 100, 1),
         round(table["Прибыль WB, руб"].sum() / склад * 100, 1),
         round(mixed / склад * 100, 1)],
        ["Когда получим деньги", "сразу", "по мере продаж",
         "часть сразу, часть по мере продаж"],
        ["Маржа WB на штуку в среднем, руб", "", round(average, 1), ""],
    ]:
        свод.append(row)
    for cell in свод[1]:
        cell.fill, cell.font = HEAD, WHITE
    for index, width in enumerate([34, 18, 18, 34], start=1):
        свод.column_dimensions[get_column_letter(index)].width = width
    for row in свод.iter_rows(min_row=2):
        for cell in row[1:]:
            cell.number_format = "# ##0"

    order = table.sort_values("Разница WB минус опт, руб",
                              key=lambda col: col.fillna(-10 ** 12))
    write(book, "ПО ПОЗИЦИЯМ", order.values.tolist(), widths,
          ["ИТОГО", int(table["Остаток, шт"].sum()), "", "",
           int(priced["Прибыль опт, руб"].sum()),
           int(table["Продажи в месяц, шт"].sum()), "",
           int(table["Прибыль WB, руб"].sum()), "", "",
           int(priced["Разница WB минус опт, руб"].sum()), ""])
    оптом = order[order["Что делать"].str.startswith("оптом")]
    write(book, "ПРОДАТЬ ОПТОМ", оптом.values.tolist(), widths)

    os.makedirs("outputs", exist_ok=True)
    book.save(target)
    print(f"Средняя маржа WB: {average:.1f} руб на штуку")
    print(f"Опт целиком:  прибыль {int(priced['Прибыль опт, руб'].sum()):,} руб "
          f"по {len(priced)} позициям".replace(",", " "))
    print(f"Только WB:    прибыль {int(table['Прибыль WB, руб'].sum()):,} руб"
          .replace(",", " "))
    print(f"Смешанно:     прибыль {int(mixed):,} руб "
          f"(WB по {len(table) - len(оптом)} позициям, опт по {len(оптом)})"
          .replace(",", " "))
    print(f"Сохранено: {target}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Опт против торговли на WB")
    parser.add_argument("source")
    parser.add_argument("--sales", required=True)
    parser.add_argument("--out", default=TARGET)
    args = parser.parse_args()
    main(args.source, args.sales, args.out)
