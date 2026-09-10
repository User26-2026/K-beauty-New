#!/usr/bin/env python3
"""Полный каталог склада по товарным группам с ценами и суммами.

Цифры взяты из учетных файлов и на 10.09.2026 не подтверждены
инвентаризацией: сверка была 16.06.2025, часть товара уже продана.
Это оценка сверху, а не факт.

Результат: outputs/marine_equipment/Каталог_склада.xlsx
"""

import pathlib
import sys
from collections import defaultdict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
from build_marine_registry import SOURCES, parse  # noqa: E402

OUT = pathlib.Path(__file__).resolve().parents[2] / "outputs" / "marine_equipment"

FONT = "Arial"
HEAD_FILL = PatternFill("solid", fgColor="1F3864")
HEAD_FONT = Font(name=FONT, size=11, bold=True, color="FFFFFF")
TITLE = Font(name=FONT, size=14, bold=True, color="1F3864")
BASE = Font(name=FONT, size=10)
NOTE = Font(name=FONT, size=9, italic=True, color="666666")
GROUP_FONT = Font(name=FONT, size=11, bold=True, color="1F3864")
GROUP_FILL = PatternFill("solid", fgColor="D9E2F3")
SUB_FILL = PatternFill("solid", fgColor="EFF3FA")
NOPRICE_FILL = PatternFill("solid", fgColor="FFF2CC")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

ITEM_COLS = ["№", "Товарная группа", "Подгруппа", "Наименование", "Чертеж",
             "Наличие", "Цена за ед., руб", "Сумма, руб", "Состояние",
             "Комплектность", "Локация"]

# порядок листов по группам
GROUP_SHEETS = [
    ("Арматура", "Арматура"),
    ("Насосное и механическое", "Насосы и механика"),
    ("Электрика", "Электрика"),
    ("Корпусная часть", "Корпусная часть"),
    ("Трубопроводная обвязка", "Обвязка"),
    ("Прочее", "Прочее"),
]

DISCLAIMER = ("Цены и количества взяты из учетных файлов. Последняя сверка "
              "остатков — 16.06.2025, часть товара с тех пор продана, "
              "но числится в наличии. До инвентаризации это оценка сверху.")


def head(ws, row, ncols, height=28):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEAD_FILL
        cell.font = HEAD_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[row].height = height


def widths(ws, values):
    for i, w in enumerate(values, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def write_items(ws, rows, title, subtitle):
    ws["A1"] = title
    ws["A1"].font = TITLE
    ws["A2"] = subtitle
    ws["A2"].font = NOTE

    start = 4
    for i, h in enumerate(ITEM_COLS, start=1):
        ws.cell(row=start, column=i, value=h)
    head(ws, start, len(ITEM_COLS))

    for n, r in enumerate(rows, start=1):
        row = start + n
        values = [n, r["Товарная группа"], r["Подгруппа"], r["Наименование"],
                  r["Чертеж"], r["Наличие"], r["Цена за ед., руб"] or None,
                  r["Сумма"] or None, r["Состояние"], r["Комплектность"],
                  r["Локация"]]
        for i, v in enumerate(values, start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=(i == 4))
            if i in (7, 8):
                cell.number_format = "#,##0"
        if not r["Цена за ед., руб"]:
            ws.cell(row=row, column=7).fill = NOPRICE_FILL

    last = start + len(rows)
    total = last + 1
    ws.cell(row=total, column=4, value="ИТОГО").font = Font(name=FONT, bold=True)
    for col in (6, 8):
        letter = get_column_letter(col)
        cell = ws.cell(row=total, column=col,
                       value=f"=SUM({letter}{start + 1}:{letter}{last})")
        cell.font = Font(name=FONT, bold=True)
        cell.border = BORDER
        cell.number_format = "#,##0"
    ws.cell(row=total + 1, column=4, value="Позиций без цены").font = BASE
    ws.cell(row=total + 1, column=6,
            value=f"=COUNTBLANK(G{start + 1}:G{last})").font = Font(
                name=FONT, bold=True, color="C00000")

    ws.auto_filter.ref = f"A{start}:K{last}"
    ws.freeze_panes = f"D{start + 1}"
    widths(ws, [5, 24, 30, 56, 20, 9, 15, 15, 12, 14, 16])


def sheet_summary(wb, rows):
    ws = wb.create_sheet("Сводка по группам", 0)
    ws["A1"] = "Что на складе и на какую сумму"
    ws["A1"].font = TITLE
    ws["A2"] = DISCLAIMER
    ws["A2"].font = NOTE

    total_sum = sum(r["Сумма"] for r in rows)
    total_units = sum(r["Наличие"] for r in rows)

    ws["A4"] = "Всего позиций"
    ws["B4"] = len(rows)
    ws["A5"] = "Всего единиц хранения"
    ws["B5"] = total_units
    ws["A6"] = "Сумма по проставленным ценам"
    ws["B6"] = total_sum
    ws["B6"].number_format = "#,##0"
    ws["A7"] = "Позиций без цены"
    ws["B7"] = sum(1 for r in rows if not r["Цена за ед., руб"])
    ws["C7"] = "их стоимость в сумму не входит"
    for r in range(4, 8):
        ws.cell(row=r, column=1).font = Font(name=FONT, size=10, bold=True)
        ws.cell(row=r, column=2).font = Font(name=FONT, size=10, bold=True)
        ws.cell(row=r, column=3).font = NOTE

    headers = ["Группа и подгруппа", "Позиций", "Единиц", "Сумма, руб",
               "Доля", "Без цены"]
    start = 9
    for i, h in enumerate(headers, start=1):
        ws.cell(row=start, column=i, value=h)
    head(ws, start, len(headers))

    by_group = defaultdict(lambda: [0, 0, 0, 0])
    by_sub = defaultdict(lambda: [0, 0, 0, 0])
    for r in rows:
        for key, store in ((r["Товарная группа"], by_group),
                           ((r["Товарная группа"], r["Подгруппа"] or "прочее"),
                            by_sub)):
            a = store[key]
            a[0] += 1
            a[1] += r["Наличие"]
            a[2] += r["Сумма"]
            a[3] += 0 if r["Цена за ед., руб"] else 1

    row = start
    for group, (p, u, s, np_) in sorted(by_group.items(),
                                        key=lambda kv: -kv[1][2]):
        row += 1
        for i, v in enumerate([group, p, u, s or None,
                               s / total_sum if total_sum else 0, np_ or None],
                              start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = GROUP_FONT if i == 1 else Font(name=FONT, size=10,
                                                       bold=True)
            cell.fill = GROUP_FILL
            cell.border = BORDER
            if i == 4:
                cell.number_format = "#,##0"
            if i == 5:
                cell.number_format = "0.0%"
        for (g, sub), (p2, u2, s2, np2) in sorted(
                by_sub.items(), key=lambda kv: -kv[1][2]):
            if g != group:
                continue
            row += 1
            for i, v in enumerate([f"    {sub}", p2, u2, s2 or None,
                                   s2 / total_sum if total_sum else 0,
                                   np2 or None], start=1):
                cell = ws.cell(row=row, column=i, value=v)
                cell.font = BASE
                cell.fill = SUB_FILL
                cell.border = BORDER
                if i == 4:
                    cell.number_format = "#,##0"
                if i == 5:
                    cell.number_format = "0.0%"

    row += 1
    ws.cell(row=row, column=1, value="ИТОГО").font = Font(name=FONT, bold=True)
    for col, val in ((2, len(rows)), (3, total_units), (4, total_sum)):
        cell = ws.cell(row=row, column=col, value=val)
        cell.font = Font(name=FONT, bold=True)
        cell.border = BORDER
        if col == 4:
            cell.number_format = "#,##0"

    widths(ws, [46, 11, 11, 18, 10, 12])
    return ws


def sheet_locations(wb, rows):
    ws = wb.create_sheet("По локациям")
    ws["A1"] = "Где что лежит"
    ws["A1"].font = TITLE
    ws["A2"] = DISCLAIMER
    ws["A2"].font = NOTE

    headers = ["Локация", "Позиций", "Единиц", "Сумма, руб", "Без цены"]
    start = 4
    for i, h in enumerate(headers, start=1):
        ws.cell(row=start, column=i, value=h)
    head(ws, start, len(headers))

    agg = defaultdict(lambda: [0, 0, 0, 0])
    for r in rows:
        a = agg[r["Локация"]]
        a[0] += 1
        a[1] += r["Наличие"]
        a[2] += r["Сумма"]
        a[3] += 0 if r["Цена за ед., руб"] else 1

    row = start
    for loc, (p, u, s, np_) in sorted(agg.items(), key=lambda kv: -kv[1][2]):
        row += 1
        for i, v in enumerate([loc, p, u, s or None, np_ or None], start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            if i == 4:
                cell.number_format = "#,##0"

    row += 1
    ws.cell(row=row, column=1, value="ИТОГО").font = Font(name=FONT, bold=True)
    for col in (2, 3, 4):
        letter = get_column_letter(col)
        cell = ws.cell(row=row, column=col,
                       value=f"=SUM({letter}{start + 1}:{letter}{row - 1})")
        cell.font = Font(name=FONT, bold=True)
        cell.border = BORDER
        if col == 4:
            cell.number_format = "#,##0"

    widths(ws, [24, 11, 11, 18, 12])
    return ws


def main():
    rows = []
    for spec in SOURCES:
        rows.extend(parse(spec))
    for r in rows:
        r["Сумма"] = (r["Наличие"] or 0) * (r["Цена за ед., руб"] or 0)
    live = [r for r in rows if (r["Наличие"] or 0) > 0]
    live.sort(key=lambda r: (r["Товарная группа"], r["Подгруппа"] or "",
                             -r["Сумма"]))

    wb = Workbook()
    wb.remove(wb.active)
    sheet_summary(wb, live)
    sheet_locations(wb, live)
    write_items(wb.create_sheet("Все товары"), live,
                "Полный перечень склада",
                DISCLAIMER)
    for group, sheet_name in GROUP_SHEETS:
        part = [r for r in live if r["Товарная группа"] == group]
        if part:
            write_items(wb.create_sheet(sheet_name), part,
                        f"{group}: {len(part)} позиций",
                        DISCLAIMER)

    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / "Каталог_склада.xlsx"
    wb.save(path)

    total = sum(r["Сумма"] for r in live)
    print(f"готово: {path}")
    print(f"позиций: {len(live)}, единиц: {sum(r['Наличие'] for r in live):.0f}, "
          f"сумма: {total:,.0f} руб".replace(",", " "))
    print(f"листов: {len(wb.sheetnames)} — {', '.join(wb.sheetnames)}")


if __name__ == "__main__":
    main()
