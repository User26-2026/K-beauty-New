#!/usr/bin/env python3
"""Прайс конкурента с Камчатки, разложенный по группам и ценам.

Исходник сверстан в две колонки и читается только глазами: 1 017 строк,
36 разделов, цена в последнем столбце. Здесь он разложен так, чтобы по
нему можно было работать: по нашим группам, внутри группы по разделу
конкурента и по убыванию диаметра.

Зачем это нужно. Это единственный внешний прайс, где цена стоит против
конкретного чертежа. По нему мы оцениваем свой склад и проверяем, не
продаем ли дешевле рынка. Цены Титова из файла покупателей — это он же:
из 75 сравнимых позиций 35 совпадают до рубля, так что это один
источник, а не два.

Чего в нем нет: веса. Поэтому цену килограмма по этому файлу не
посчитать, она берется из нашего учета.

Четыре вкладки:
  Сводка           — по группам: позиций, разброс цен, медиана
  Прайс по группам — все 1 017 строк
  Цены по Ду       — медианная цена в пересечении группы и диаметра
  Вне наших групп  — разделы, которых у нас в номенклатуре нет

Источник: data/marine_equipment/competitors/Прайс_Камчатка_03.10.26.xlsx
Результат: outputs/marine_equipment/Прайс_Камчатка_по_группам.xlsx
"""

import pathlib
import statistics
import sys
from collections import defaultdict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
from build_competitor_compare import read_competitor  # noqa: E402
from build_spb_list_xlsx import GROUPS  # noqa: E402

OUT = pathlib.Path(__file__).resolve().parents[2] / "outputs" / "marine_equipment"

FONT = "Arial"
HEAD_FILL = PatternFill("solid", fgColor="1F3864")
HEAD_FONT = Font(name=FONT, size=11, bold=True, color="FFFFFF")
TITLE = Font(name=FONT, size=14, bold=True, color="1F3864")
SUB = Font(name=FONT, size=11, bold=True, color="1F3864")
SECT = Font(name=FONT, size=10, bold=True, color="333333")
BOLD = Font(name=FONT, size=10, bold=True)
BASE = Font(name=FONT, size=10)
NOTE = Font(name=FONT, size=9, italic=True, color="666666")
GROUP_FILL = PatternFill("solid", fgColor="D9E2F3")
SECT_FILL = PatternFill("solid", fgColor="F2F2F2")
NEW_FILL = PatternFill("solid", fgColor="E2EFDA")
STEEL_FILL = PatternFill("solid", fgColor="FFF2CC")
NOPRICE = PatternFill("solid", fgColor="FCE4D6")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

OTHER = "Вне наших групп"


def head(ws, row, ncols, height=32):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEAD_FILL
        cell.font = HEAD_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=True,
                                   horizontal="center")
        cell.border = BORDER
    ws.row_dimensions[row].height = height


def widths(ws, values):
    for i, w in enumerate(values, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def put(ws, row, values, wrap=(), center=(), money=(), fill=None, font=None):
    for i, v in enumerate(values, start=1):
        cell = ws.cell(row=row, column=i, value=v)
        cell.font = font or BASE
        cell.border = BORDER
        cell.alignment = Alignment(vertical="top", wrap_text=(i in wrap),
                                   horizontal="center" if i in center
                                   else None)
        if i in money:
            cell.number_format = "#,##0"
        if fill:
            cell.fill = fill
    return row + 1


def plural(n, one, few, many):
    n = abs(int(n))
    if n % 10 == 1 and n % 100 != 11:
        return one
    if n % 10 in (2, 3, 4) and n % 100 not in (12, 13, 14):
        return few
    return many


def group_of(r):
    return r["Группа"] or OTHER


def group_order(groups):
    """Наши группы в принятом порядке, чужие разделы в конец."""
    order = {g: i for i, g in enumerate(GROUPS)}
    return sorted(groups, key=lambda g: (order.get(g, 98) if g != OTHER
                                         else 99, g))


def kind(r):
    """Состояние и материал одной строкой: это и объясняет цену."""
    bits = []
    if r["Новая"]:
        bits.append("новый")
    if r["Материал"]:
        bits.append(r["Материал"])
    low = r["Ду текст"].lower()
    if "(р)" in low:
        bits.append("после ремонта")
    if "б/с" in low:
        bits.append("бессальниковый")
    return ", ".join(bits)


def stats(prices):
    if not prices:
        return None, None, None
    return min(prices), statistics.median(prices), max(prices)


COLS_SUM = [
    ("Группа", 28), ("Позиций", 9), ("С ценой", 9), ("Штук в наличии", 12),
    ("Мин. цена, руб", 13), ("Медиана, руб", 13), ("Макс. цена, руб", 13),
    ("Разброс, раз", 11), ("Диаметры", 30), ("Разделов конкурента", 36),
]


def sheet_summary(wb, rows):
    ws = wb.active
    ws.title = "Сводка"
    by = defaultdict(list)
    for r in rows:
        by[group_of(r)].append(r)
    priced = [r for r in rows if r["Цена"]]
    ws["A1"] = "Прайс конкурента с Камчатки, 03.10.2026"
    ws["A1"].font = TITLE
    ws["A2"] = (f"{len(rows):,} "
                f"{plural(len(rows), 'позиция', 'позиции', 'позиций')}, "
                f"из них с ценой {len(priced):,}. Медианная цена "
                f"{statistics.median([r['Цена'] for r in priced]):,.0f} руб."
                .replace(",", " "))
    ws["A2"].font = NOTE
    ws["A3"] = ("Веса в прайсе нет, поэтому цену килограмма по нему не "
                "посчитать. Цены новых позиций и позиций после ремонта "
                "стоят рядом со старыми, усреднять их нельзя.")
    ws["A3"].font = NOTE

    start = 5
    for i, (h, _) in enumerate(COLS_SUM, start=1):
        ws.cell(row=start, column=i, value=h)
    head(ws, start, len(COLS_SUM))
    row = start + 1
    for g in group_order(by):
        items = by[g]
        pr = [r["Цена"] for r in items if r["Цена"]]
        lo, med, hi = stats(pr)
        dus = sorted({r["Ду"] for r in items if r["Ду"]})
        secs = sorted({r["Раздел"] for r in items if r["Раздел"]})
        qty = sum(r["Наличие"] for r in items
                  if isinstance(r["Наличие"], (int, float)))
        row = put(ws, row, [
            g, len(items), len(pr), qty or "",
            lo or "", round(med) if med else "", hi or "",
            round(hi / lo, 1) if lo else "",
            ", ".join(f"Ду{d}" for d in dus[:10])
            + (" и др." if len(dus) > 10 else ""),
            "; ".join(s.strip() for s in secs[:3])
            + (f" и еще {len(secs) - 3}" if len(secs) > 3 else ""),
        ], wrap=(1, 9, 10), center=(2, 3, 4, 8), money=(5, 6, 7),
            font=BOLD if g == OTHER else BASE)
    total = [r["Цена"] for r in priced]
    lo, med, hi = stats(total)
    row = put(ws, row, ["ВСЕГО", len(rows), len(priced), "",
                        lo, round(med), hi, round(hi / lo, 1), "", ""],
              center=(2, 3, 8), money=(5, 6, 7), font=BOLD,
              fill=GROUP_FILL)
    row += 1
    ws.cell(row=row, column=1, value="Как этим пользоваться").font = SUB
    row += 1
    for t in [
        "Разброс в разы внутри группы — это не торг, а разные позиции: "
        "новое против б/у, бронза против стали, разные диаметры. "
        "Сравнивать можно только позицию с позицией.",
        "Конкурент помечает стальные позиции «(ст.)», и они втрое "
        "дешевле бронзовых того же диаметра: задвижка Ду200 бронза "
        "150 000, сталь 40 000.",
        "«New» — новое, «(р)» — после ремонта, «б/с» — "
        "бессальниковый. Все три пометки подняты в колонку «Что это».",
        "Это прайс продавца, а не цена, по которой покупают у нас. "
        "Для нашей цены продажи остается договор с ПКФ «АРМА».",
    ]:
        cell = ws.cell(row=row, column=1, value="— " + t)
        cell.font = BASE
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[row].height = 28
        row += 1
    ws.freeze_panes = f"A{start + 1}"
    widths(ws, [w for _, w in COLS_SUM])
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1


COLS_BOOK = [
    ("№", 5), ("Раздел у конкурента", 40), ("Ду", 7), ("Ру", 7),
    ("Чертеж", 22), ("Что это", 28), ("Наличие, шт", 10),
    ("Цена, руб", 12), ("Как записано в прайсе", 16),
]


def sheet_book(wb, rows, name, which):
    ws = wb.create_sheet(name)
    data = [r for r in rows if which(r)]
    by = defaultdict(list)
    for r in data:
        by[group_of(r)].append(r)
    ws["A1"] = ("Прайс по группам" if name == "Прайс по группам"
                else "Разделы, которых нет в нашей номенклатуре")
    ws["A1"].font = TITLE
    pr = [r["Цена"] for r in data if r["Цена"]]
    ws["A2"] = (f"{len(data):,} "
                f"{plural(len(data), 'позиция', 'позиции', 'позиций')}, "
                f"с ценой {len(pr):,}, медиана "
                f"{statistics.median(pr):,.0f} руб. Внутри группы — по "
                f"разделу конкурента и по убыванию диаметра."
                .replace(",", " "))
    ws["A2"].font = NOTE

    start = 4
    for i, (h, _) in enumerate(COLS_BOOK, start=1):
        ws.cell(row=start, column=i, value=h)
    head(ws, start, len(COLS_BOOK))
    row = start + 1
    n = 0
    for g in group_order(by):
        items = by[g]
        gp = [r["Цена"] for r in items if r["Цена"]]
        lo, med, hi = stats(gp)
        row = put(ws, row, [
            f"{g} — {len(items)} "
            f"{plural(len(items), 'позиция', 'позиции', 'позиций')}"
            + (f", от {lo:,.0f} до {hi:,.0f} руб, медиана {med:,.0f}"
               .replace(",", " ") if gp else ", цен нет")
        ], fill=GROUP_FILL, font=SUB)
        bysec = defaultdict(list)
        for r in items:
            bysec[r["Раздел"] or "(без раздела)"].append(r)
        for sec in sorted(bysec, key=lambda s: -len(bysec[s])):
            sp = [r["Цена"] for r in bysec[sec] if r["Цена"]]
            row = put(ws, row, [
                "", sec.strip(), "", "", "", "", len(bysec[sec]),
                round(statistics.median(sp)) if sp else "", "медиана",
            ], wrap=(2,), center=(7,), money=(8,), fill=SECT_FILL,
                font=SECT)
            for r in sorted(bysec[sec],
                            key=lambda x: (-(x["Ду"] or 0),
                                           -(x["Цена"] or 0))):
                n += 1
                fill = (NOPRICE if not r["Цена"]
                        else STEEL_FILL if r["Материал"] == "сталь"
                        else NEW_FILL if r["Новая"] else None)
                row = put(ws, row, [
                    n, "", r["Ду"] or "", r["Ру"] or "", r["Чертеж"],
                    kind(r),
                    r["Наличие"] if isinstance(r["Наличие"], (int, float))
                    else "",
                    r["Цена"] or "", r["Ду текст"],
                ], wrap=(6,), center=(1, 3, 4, 7), money=(8,), fill=fill)
    ws.freeze_panes = f"A{start + 1}"
    ws.auto_filter.ref = (f"A{start}:"
                          f"{get_column_letter(len(COLS_BOOK))}{row - 1}")
    widths(ws, [w for _, w in COLS_BOOK])
    ws.print_title_rows = f"{start}:{start}"
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    return len(data)


def sheet_matrix(wb, rows):
    """Группа против диаметра: медианная цена и сколько позиций."""
    ws = wb.create_sheet("Цены по Ду", 2)
    dus = sorted({r["Ду"] for r in rows if r["Ду"]}, reverse=True)
    groups = group_order({group_of(r) for r in rows})
    ws["A1"] = "Медианная цена конкурента: группа против диаметра"
    ws["A1"].font = TITLE
    ws["A2"] = ("В ячейке медианная цена в рублях, под ней число позиций. "
                "Это ориентир, а не цена сделки: внутри одного диаметра "
                "стоят и новые, и б/у, и бронза, и сталь.")
    ws["A2"].font = NOTE

    start = 4
    ws.cell(row=start, column=1, value="Группа")
    for i, d in enumerate(dus, start=2):
        ws.cell(row=start, column=i, value=f"Ду{d}")
    head(ws, start, len(dus) + 1, height=22)
    row = start + 1
    cells = defaultdict(list)
    for r in rows:
        if r["Ду"] and r["Цена"]:
            cells[(group_of(r), r["Ду"])].append(r["Цена"])
    for g in groups:
        ws.cell(row=row, column=1, value=g).font = BOLD
        ws.cell(row=row, column=1).border = BORDER
        ws.cell(row=row, column=1).alignment = Alignment(wrap_text=True,
                                                         vertical="center")
        for i, d in enumerate(dus, start=2):
            pr = cells.get((g, d))
            cell = ws.cell(row=row, column=i,
                           value=round(statistics.median(pr)) if pr else None)
            cell.font = BASE
            cell.border = BORDER
            cell.number_format = "#,##0"
            cell.alignment = Alignment(horizontal="center")
            # жирным там, где цена посчитана по пяти и более позициям:
            # такой медиане можно верить, одиночной цене — нет
            if pr and len(pr) >= 5:
                cell.font = BOLD
        row += 1
    ws.freeze_panes = f"B{start + 1}"
    widths(ws, [28] + [11] * len(dus))
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1


def main():
    rows = read_competitor()
    wb = Workbook()
    sheet_summary(wb, rows)
    sheet_book(wb, rows, "Прайс по группам", lambda r: True)
    sheet_matrix(wb, rows)
    other = sheet_book(wb, rows, "Вне наших групп",
                       lambda r: not r["Группа"])

    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / "Прайс_Камчатка_по_группам.xlsx"
    wb.save(path)
    print(f"готово: {path}\n")

    priced = [r for r in rows if r["Цена"]]
    print(f"{len(rows)} позиций, с ценой {len(priced)}, "
          f"вне наших групп {other}")
    print(f"\n{'группа':<30}{'поз':>5}{'мин':>9}{'медиана':>10}{'макс':>10}")
    by = defaultdict(list)
    for r in rows:
        by[group_of(r)].append(r)
    for g in group_order(by):
        pr = [r["Цена"] for r in by[g] if r["Цена"]]
        lo, med, hi = stats(pr)
        print(f"{g:<30}{len(by[g]):>5}{lo or 0:>9,.0f}{med or 0:>10,.0f}"
              f"{hi or 0:>10,.0f}".replace(",", " "))


if __name__ == "__main__":
    main()
