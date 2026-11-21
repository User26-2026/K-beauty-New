#!/usr/bin/env python3
"""Каталог склада в рабочей группировке владельца.

Группы заданы так, как о товаре думает продавец, а не по формальной
номенклатуре: покупатель спрашивает «запчасти к двигателю» или
«брашпиль», а не «насосное и механическое оборудование».

Цифры взяты из учетных файлов и на 10.09.2026 не подтверждены
инвентаризацией: последняя сверка 16.06.2025, часть товара продана.
Это оценка сверху, а не факт.

Результат: outputs/marine_equipment/Каталог_склада.xlsx
"""

import pathlib
import re
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
GROUP_FILL = PatternFill("solid", fgColor="D9E2F3")
NOPRICE_FILL = PatternFill("solid", fgColor="FFF2CC")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

DISCLAIMER = ("Цены и количества из учетных файлов. Последняя сверка остатков "
              "16.06.2025, часть товара с тех пор продана, но числится "
              "в наличии. До инвентаризации это оценка сверху.")

# порядок групп и короткое имя листа
GROUPS = [
    ("Арматура", "Арматура"),
    ("Запчасти к двигателям", "Запчасти двигателей"),
    ("Насосы", "Насосы"),
    ("Двери и иллюминаторы", "Двери и иллюминаторы"),
    ("Сепараторы и компрессоры", "Сепараторы компрессоры"),
    ("Камбузное оборудование", "Камбузное"),
    ("Брашпили, лебедки, трал, подруливающие", "Палубные механизмы"),
    ("Головки, фланцы, гайки и стволы пожарные, стаканы", "Обвязка"),
    ("Электрика", "Электрика"),
    ("Прочее", "Прочее"),
]

# модели двигателей и насосов из подзаголовков раздела ЗИП
ENGINE_MODELS = re.compile(
    r"пилстик|vd26|nvd26|nvd48|8нвд|3d6|4ч ?8,5|4ч ?10,5|zd ?72|3д12|"
    r"6чн25|18/22|этф-3|д1\.м", re.I)
PUMP_MODELS = re.compile(r"нцв|нцкг|цвс|1эцну|эвн|эсн", re.I)

DECK = re.compile(
    r"брашпил|лебедк|трал|подрулив|ваероуклад|кран-балк|шпил|"
    r"контроллер|кранец", re.I)
GALLEY = re.compile(r"пкэ|плита камбузн|посуда к плите|вилка к пкэ", re.I)
DOORS = re.compile(r"иллюминатор|дверь|двери|крышка вгн|крышка легкая|"
                   r"крышка световая|крышка иллюминатора|крышка с коменсом", re.I)
SEPARATORS = re.compile(r"сепаратор|барабан сц|кулачки сц|ремкомплект сц|"
                        r"ключ на сц|компрессор|экп ?70", re.I)
FITTINGS = re.compile(
    r"головка воздушно|грибк|фланец|фланцы|гайка|гайки|ствол пожарн|"
    r"стакан переборочн|компенсатор|фонарь смотров|колонка указательн", re.I)
PUMPS = re.compile(r"^насос|^нцв|^нцвс|поплавок|крылатка", re.I)

# титан выносим отдельным списком: он дороже остальных цветных металлов
# и как изделие, и как лом
TITANIUM = re.compile(r"титан", re.I)


def group_of(r):
    """Группа в терминах, которыми пользуемся при продаже."""
    name = r["Наименование"]
    model = r["Оборудование"] or ""
    sub = r["Подгруппа"] or ""

    # палубные механизмы и камбуз опознаем первыми: их мало и они узнаваемы
    if DECK.search(name):
        return "Брашпили, лебедки, трал, подруливающие"
    if GALLEY.search(name):
        return "Камбузное оборудование"
    if SEPARATORS.search(name) or SEPARATORS.search(model):
        return "Сепараторы и компрессоры"
    if DOORS.search(name):
        return "Двери и иллюминаторы"

    # запчасть под подзаголовком модели относится к своему агрегату
    if ENGINE_MODELS.search(model) or ENGINE_MODELS.search(name):
        return "Запчасти к двигателям"
    if PUMP_MODELS.search(model):
        return "Насосы"
    if sub == "Двигатели и ЗИП":
        return "Запчасти к двигателям"
    if sub == "Насосы" or PUMPS.search(name):
        return "Насосы"

    if FITTINGS.search(name):
        return "Головки, фланцы, гайки и стволы пожарные, стаканы"
    if r["Товарная группа"] == "Арматура" or "фильтр" in name.lower():
        return "Арматура"
    if r["Товарная группа"] == "Электрика":
        return "Электрика"
    return "Прочее"


ITEM_COLS = ["№", "Группа", "Наименование", "Чертеж", "Наличие",
             "Цена за ед., руб", "Сумма, руб", "Состояние", "Комплектность",
             "К чему подходит", "Локация"]


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


def write_items(ws, rows, title):
    ws["A1"] = title
    ws["A1"].font = TITLE
    ws["A2"] = DISCLAIMER
    ws["A2"].font = NOTE

    start = 4
    for i, h in enumerate(ITEM_COLS, start=1):
        ws.cell(row=start, column=i, value=h)
    head(ws, start, len(ITEM_COLS))

    for n, r in enumerate(rows, start=1):
        row = start + n
        values = [n, r["Группа продажи"], r["Наименование"], r["Чертеж"],
                  r["Наличие"], r["Цена за ед., руб"] or None,
                  r["Сумма"] or None, r["Состояние"], r["Комплектность"],
                  r["Оборудование"], r["Локация"]]
        for i, v in enumerate(values, start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=(i == 3))
            if i in (6, 7):
                cell.number_format = "#,##0"
        if not r["Цена за ед., руб"]:
            ws.cell(row=row, column=6).fill = NOPRICE_FILL

    last = start + len(rows)
    total = last + 1
    ws.cell(row=total, column=3, value="ИТОГО").font = Font(name=FONT, bold=True)
    for col in (5, 7):
        letter = get_column_letter(col)
        cell = ws.cell(row=total, column=col,
                       value=f"=SUM({letter}{start + 1}:{letter}{last})")
        cell.font = Font(name=FONT, bold=True)
        cell.border = BORDER
        cell.number_format = "#,##0"
    ws.cell(row=total + 1, column=3, value="Позиций без цены").font = BASE
    ws.cell(row=total + 1, column=5,
            value=f"=COUNTBLANK(F{start + 1}:F{last})").font = Font(
                name=FONT, bold=True, color="C00000")

    ws.auto_filter.ref = f"A{start}:K{last}"
    ws.freeze_panes = f"C{start + 1}"
    widths(ws, [5, 34, 58, 20, 9, 15, 15, 12, 14, 22, 16])


def sheet_summary(wb, rows):
    ws = wb.create_sheet("Сводка", 0)
    ws["A1"] = "Что на складе и на какую сумму"
    ws["A1"].font = TITLE
    ws["A2"] = DISCLAIMER
    ws["A2"].font = NOTE

    total_sum = sum(r["Сумма"] for r in rows)
    facts = [
        ("Всего позиций", len(rows), ""),
        ("Всего единиц хранения", sum(r["Наличие"] for r in rows), ""),
        ("Сумма по проставленным ценам", total_sum, "оценка до инвентаризации"),
        ("Позиций без цены", sum(1 for r in rows if not r["Цена за ед., руб"]),
         "их стоимость в сумму не входит"),
    ]
    for i, (label, value, note) in enumerate(facts):
        row = 4 + i
        ws.cell(row=row, column=1, value=label).font = Font(name=FONT, size=10,
                                                            bold=True)
        cell = ws.cell(row=row, column=2, value=value)
        cell.font = Font(name=FONT, size=10, bold=True)
        if label.startswith("Сумма"):
            cell.number_format = "#,##0"
        ws.cell(row=row, column=3, value=note).font = NOTE

    headers = ["Группа", "Позиций", "Единиц", "Сумма, руб", "Доля",
               "Позиций без цены"]
    start = 9
    for i, h in enumerate(headers, start=1):
        ws.cell(row=start, column=i, value=h)
    head(ws, start, len(headers))

    agg = defaultdict(lambda: [0, 0, 0, 0])
    for r in rows:
        a = agg[r["Группа продажи"]]
        a[0] += 1
        a[1] += r["Наличие"]
        a[2] += r["Сумма"]
        a[3] += 0 if r["Цена за ед., руб"] else 1

    row = start
    for group, _ in GROUPS:
        if group not in agg:
            continue
        p, u, s, np_ = agg[group]
        row += 1
        for i, v in enumerate([group, p, u, s or None,
                               s / total_sum if total_sum else 0, np_ or None],
                              start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = Font(name=FONT, size=10, bold=True)
            cell.fill = GROUP_FILL
            cell.border = BORDER
            cell.alignment = Alignment(vertical="center", wrap_text=(i == 1))
            if i == 4:
                cell.number_format = "#,##0"
            if i == 5:
                cell.number_format = "0.0%"

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

    widths(ws, [44, 11, 11, 18, 10, 18])
    return ws


def main():
    rows = []
    for spec in SOURCES:
        rows.extend(parse(spec))
    for r in rows:
        r["Сумма"] = (r["Наличие"] or 0) * (r["Цена за ед., руб"] or 0)
    live = [r for r in rows if (r["Наличие"] or 0) > 0]
    for r in live:
        r["Группа продажи"] = group_of(r)

    order = {g: i for i, (g, _) in enumerate(GROUPS)}
    live.sort(key=lambda r: (order.get(r["Группа продажи"], 99), -r["Сумма"]))

    wb = Workbook()
    wb.remove(wb.active)
    sheet_summary(wb, live)
    write_items(wb.create_sheet("Все товары"), live, "Полный перечень склада")

    titanium = [r for r in live
                if TITANIUM.search(r["Наименование"] + " "
                                   + (r["Группа в файле"] or ""))]
    if titanium:
        ws = wb.create_sheet("Титановые клапаны")
        write_items(ws, titanium,
                    f"Титановые клапаны: {len(titanium)} позиции, "
                    f"{sum(r['Наличие'] for r in titanium):.0f} штук")
        row = ws.max_row + 3
        ws.cell(row=row, column=3,
                value="Цены нет ни по одной позиции — оценить отдельно.").font = NOTE
        ws.cell(row=row + 1, column=3,
                value="Титан дороже остальных цветных металлов и как изделие, "
                      "и как лом. В общий металл не отправлять.").font = NOTE
    for group, sheet_name in GROUPS:
        part = [r for r in live if r["Группа продажи"] == group]
        if part:
            write_items(wb.create_sheet(sheet_name), part,
                        f"{group}: {len(part)} позиций, "
                        f"{sum(r['Наличие'] for r in part):.0f} единиц")

    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / "Каталог_склада.xlsx"
    wb.save(path)

    print(f"готово: {path}")
    agg = defaultdict(lambda: [0, 0, 0, 0])
    for r in live:
        a = agg[r["Группа продажи"]]
        a[0] += 1
        a[1] += r["Наличие"]
        a[2] += r["Сумма"]
        a[3] += 0 if r["Цена за ед., руб"] else 1
    for group, _ in GROUPS:
        if group in agg:
            p, u, s, np_ = agg[group]
            print(f"  {group:<44} поз {p:>4} ед {u:>6.0f} {s:>12,.0f} "
                  f"без цены {np_}".replace(",", " "))
    print(f"  {'ИТОГО':<44} поз {len(live):>4} "
          f"ед {sum(r['Наличие'] for r in live):>6.0f} "
          f"{sum(r['Сумма'] for r in live):>12,.0f}".replace(",", " "))


if __name__ == "__main__":
    main()
