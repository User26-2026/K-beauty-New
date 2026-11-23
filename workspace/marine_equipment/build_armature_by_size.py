#!/usr/bin/env python3
"""Вся арматура от большего к меньшему: чем грузить контейнер сначала.

Погрузка идет от крупного к мелкому, поэтому и список нужен в том же
порядке. Сверху крупные клинкеты, которые ложатся на пол первым слоем,
снизу мелочь, которой добивают ящики.

Разбивка по размерным блокам отвечает на вопрос «куда это в
контейнере»: Ду300-350 и Ду200-250 идут на пол, Ду100-150 на пол или
в низ ящиков, все что меньше — в ящики.

Нарастающий итог по весу показывает главное: крупная арматура дает
большую часть тонн, и именно она решает, влезет груз или нет.

Перечень тот же, что в планах отгрузки: список менеджера плюс склад с
вычетом задвоений и уже отгруженного.

Результат: outputs/marine_equipment/Арматура_по_размеру.xlsx
"""

import pathlib
import sys
from collections import defaultdict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
from build_container_spb import asrow, plural  # noqa: E402
from build_logistics_compare import rows_all  # noqa: E402
from build_shipment_spb import TIERS, item_kind, material, pack_type, tier  # noqa: E402
from build_spb_list_xlsx import group_of  # noqa: E402
from build_weight_xlsx import du as du_of  # noqa: E402
import gate_geometry as gg  # noqa: E402

OUT = pathlib.Path(__file__).resolve().parents[2] / "outputs" / "marine_equipment"

FONT = "Arial"
HEAD_FILL = PatternFill("solid", fgColor="1F3864")
HEAD_FONT = Font(name=FONT, size=11, bold=True, color="FFFFFF")
TITLE = Font(name=FONT, size=14, bold=True, color="1F3864")
SUB = Font(name=FONT, size=11, bold=True, color="1F3864")
BOLD = Font(name=FONT, size=10, bold=True)
BASE = Font(name=FONT, size=10)
NOTE = Font(name=FONT, size=9, italic=True, color="666666")
BLOCK_FILL = PatternFill("solid", fgColor="D9E2F3")
BIG_FILL = PatternFill("solid", fgColor="FCE4D6")
MID_FILL = PatternFill("solid", fgColor="FFF2CC")
SMALL_FILL = PatternFill("solid", fgColor="EAF3EA")
NODRAW_FILL = PatternFill("solid", fgColor="FFFF00")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# размерный блок: нижняя граница Ду, название, куда в контейнере, цвет
BLOCKS = [
    (300, "Ду300 и больше — самое крупное",
     "На пол первым слоем, у дальней стены. Поперек встает 2 штуки",
     BIG_FILL),
    (200, "Ду200-250 — крупное",
     "На пол, продолжение первого слоя. Ду200 встает по 3 поперек",
     BIG_FILL),
    (100, "Ду100-150 — среднее",
     "На пол рядом с клинкетами или в низ ящиков, самое тяжелое вниз",
     MID_FILL),
    (50, "Ду50-80 — мелкое",
     "В ящики на поддонах, нижний ярус",
     SMALL_FILL),
    (1, "Ду40 и меньше — мелочь",
     "В ящики, верхний ярус. Ей добиваем остаток места",
     SMALL_FILL),
    (0, "Без диаметра — уточнить при отборе",
     "Смотрят глазами на складе",
     SMALL_FILL),
]

COLS = [
    ("№", 5),
    ("Ду", 7),
    ("Тип", 24),
    ("Наименование", 48),
    ("Чертеж", 19),
    ("Материал", 16),
    ("Штук", 8),
    ("Вес ед., кг", 11),
    ("Вес всего, кг", 13),
    ("Нарастающим, кг", 15),
    ("Доля веса", 10),
    ("Куда в контейнере", 26),
    ("Как паковать", 30),
    ("Локация", 22),
    ("Цена", 14),
]


def head(ws, row, ncols, height=36):
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


def block_of(d):
    for lo, name, where, fill in BLOCKS:
        if d >= lo and lo > 0:
            return lo, name, where, fill
    return 0, BLOCKS[-1][1], BLOCKS[-1][2], BLOCKS[-1][3]


def load():
    """Вся арматура в наличии, с диаметром и весом."""
    out = []
    for r in rows_all():
        d = int(r.get("Ду") or du_of(r["Наименование"]) or 0)
        out.append(dict(r, **{
            "Ду": d,
            "Тип": item_kind(asrow(r)),
            "Группа": group_of(r["Наименование"]),
            "Материал": material(asrow(r)),
            "Ярус": TIERS[tier(asrow(r))][0],
            "Паковать": pack_type(asrow(r)),
        }))
    # от большего к меньшему: сначала диаметр, потом вес штуки
    out.sort(key=lambda r: (-r["Ду"], -r["вес_ед"], r["Наименование"]))
    return out


def sheet_list(wb, items):
    ws = wb.create_sheet("От большего к меньшему", 0)
    total = sum(r["вес"] for r in items)
    ws["A1"] = "Вся арматура от большего к меньшему"
    ws["A1"].font = TITLE
    ws["A2"] = (f"{len(items)} "
                f"{plural(len(items), 'позиция', 'позиции', 'позиций')}, "
                f"{sum(r['Штук'] for r in items):.0f} шт, "
                f"{total / 1000:.2f} т. Порядок тот же, что при погрузке: "
                "сверху то, что ложится на пол первым, снизу мелочь в "
                "ящики.")
    ws["A2"].font = NOTE

    start = 4
    for i, (h, _) in enumerate(COLS, start=1):
        ws.cell(row=start, column=i, value=h)
    head(ws, start, len(COLS))

    by_block = defaultdict(list)
    for r in items:
        lo, *_ = block_of(r["Ду"])
        by_block[lo].append(r)

    row = start
    n = 0
    run = 0.0
    for lo, name, where, fill in BLOCKS:
        rs = by_block.get(lo)
        if not rs:
            continue
        kg = sum(r["вес"] for r in rs)
        qty = sum(r["Штук"] for r in rs)
        row += 1
        kg_text = f"{kg:,.0f}".replace(",", " ")
        ws.cell(row=row, column=1,
                value=f"{name} — {len(rs)} "
                      f"{plural(len(rs), 'позиция', 'позиции', 'позиций')}, "
                      f"{qty:.0f} шт, {kg_text} кг "
                      f"({kg / total * 100:.0f}% веса). {where}")
        for c in range(1, len(COLS) + 1):
            cell = ws.cell(row=row, column=c)
            cell.fill = BLOCK_FILL
            cell.font = SUB
            cell.border = BORDER
        ws.row_dimensions[row].height = 20

        for r in rs:
            n += 1
            run += r["вес"]
            row += 1
            values = [n, r["Ду"] or "", r["Тип"], r["Наименование"],
                      r["Чертеж"], r["Материал"], r["Штук"],
                      round(r["вес_ед"], 1), round(r["вес"]), round(run),
                      run / total, r["Ярус"], r["Паковать"], r["Откуда"],
                      r["Цена"]]
            for i, v in enumerate(values, start=1):
                cell = ws.cell(row=row, column=i, value=v)
                cell.font = BASE
                cell.border = BORDER
                cell.alignment = Alignment(vertical="top",
                                           wrap_text=(i in (4, 13)))
                if i in (7, 9, 10):
                    cell.number_format = "#,##0"
                if i == 8:
                    cell.number_format = "#,##0.0"
                if i == 11:
                    cell.number_format = "0%"
            ws.cell(row=row, column=2).fill = fill
            if not r["Чертеж"]:
                ws.cell(row=row, column=5).fill = NODRAW_FILL

    last = row
    row += 1
    ws.cell(row=row, column=4, value="ИТОГО").font = BOLD
    for col in (7, 9):
        letter = get_column_letter(col)
        cell = ws.cell(row=row, column=col,
                       value=f"=SUM({letter}{start + 1}:{letter}{last})")
        cell.font = BOLD
        cell.border = BORDER
        cell.number_format = "#,##0"

    ws.auto_filter.ref = f"A{start}:O{last}"
    ws.freeze_panes = f"D{start + 1}"
    widths(ws, [w for _, w in COLS])
    return ws


def sheet_by_du(wb, items):
    ws = wb.create_sheet("Сводка по диаметрам")
    total = sum(r["вес"] for r in items)
    ws["A1"] = "Сколько чего по диаметрам"
    ws["A1"].font = TITLE
    ws["A2"] = ("Главное здесь — колонка «нарастающим». По ней видно, "
                "сколько тонн набирается на крупных диаметрах и когда "
                "можно переставать считать мелочь.")
    ws["A2"].font = NOTE

    row = 4
    cols = ["Ду", "Позиций", "Штук", "Вес, кг", "Доля веса",
            "Нарастающим, кг", "Нарастающим, %", "Вес штуки, кг",
            "В ряд поперек контейнера", "Глубина ряда, м"]
    for i, h in enumerate(cols, start=1):
        ws.cell(row=row, column=i, value=h)
    head(ws, row, len(cols), height=40)

    agg = defaultdict(lambda: [0, 0.0, 0.0])
    for r in items:
        a = agg[r["Ду"]]
        a[0] += 1
        a[1] += r["Штук"]
        a[2] += r["вес"]
    run = 0.0
    for d in sorted(agg, reverse=True):
        row += 1
        pos, qty, kg = agg[d]
        run += kg
        per = kg / qty if qty else 0
        values = [f"Ду{d}" if d else "без Ду", pos, qty, round(kg),
                  kg / total, round(run), run / total, round(per, 1),
                  gg.across(d) if d in gg.DIM else "",
                  round(gg.depth(d), 2) if d in gg.DIM else ""]
        for i, v in enumerate(values, start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            if i in (3, 4, 6):
                cell.number_format = "#,##0"
            if i in (5, 7):
                cell.number_format = "0%"
            if i == 8:
                cell.number_format = "#,##0.0"
            if i == 10:
                cell.number_format = "#,##0.00"
        if d >= 200:
            ws.cell(row=row, column=1).fill = BIG_FILL
        elif d >= 100:
            ws.cell(row=row, column=1).fill = MID_FILL

    row += 2
    ws.cell(row=row, column=1, value="Как этим пользоваться").font = SUB
    row += 1
    big = sum(r["вес"] for r in items if r["Ду"] >= 200)
    mid = sum(r["вес"] for r in items if 100 <= r["Ду"] < 200)
    small = sum(r["вес"] for r in items if 0 < r["Ду"] < 100)
    for text in [
        f"Ду200 и выше — {big / 1000:.2f} т, это {big / total * 100:.0f}% "
        "всего веса. Это и есть первый слой на полу, с него начинается "
        "погрузка и им определяется, влезет груз или нет.",
        f"Ду100-150 — {mid / 1000:.2f} т, {mid / total * 100:.0f}%. Идет "
        "рядом на пол или в низ ящиков.",
        f"Ду80 и меньше — {small / 1000:.2f} т, {small / total * 100:.0f}%. "
        "Вся эта мелочь уходит в ящики, и ее место в контейнере сверху.",
        "Поперек контейнера шириной 2,352 м встает: Ду350 и Ду300 по 2, "
        "Ду250 по 2, Ду200 по 3. Глубина ряда равна строительной длине "
        "задвижки: 500, 450, 400 и 350 мм.",
        "Габариты справочные. Замерить по три штуки каждого диаметра до "
        "погрузки: справочные массы уже оказались завышены почти вдвое.",
    ]:
        cell = ws.cell(row=row, column=1, value="— " + text)
        cell.font = BASE
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[row].height = 32
        row += 1

    widths(ws, [12, 12, 10, 12, 12, 18, 16, 14, 24, 18])
    return ws


def sheet_by_type(wb, items):
    ws = wb.create_sheet("По типам и диаметрам")
    ws["A1"] = "Тип плюс диаметр: что искать на складе"
    ws["A1"].font = TITLE
    ws["A2"] = ("По этой вкладке удобно отбирать: один тип, один диаметр, "
                "сколько всего штук и где лежит.")
    ws["A2"].font = NOTE
    row = 4
    cols = ["Тип", "Ду", "Позиций", "Штук", "Вес, кг", "Локации"]
    for i, h in enumerate(cols, start=1):
        ws.cell(row=row, column=i, value=h)
    head(ws, row, len(cols), height=24)

    agg = defaultdict(lambda: [0, 0.0, 0.0, set()])
    for r in items:
        a = agg[(r["Тип"], r["Ду"])]
        a[0] += 1
        a[1] += r["Штук"]
        a[2] += r["вес"]
        a[3].add(r["Откуда"])
    for (t, d) in sorted(agg, key=lambda k: (-k[1], k[0])):
        row += 1
        pos, qty, kg, locs = agg[(t, d)]
        for i, v in enumerate([t, f"Ду{d}" if d else "", pos, qty,
                               round(kg), ", ".join(sorted(locs))], start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=(i == 6))
            if i in (4, 5):
                cell.number_format = "#,##0"
    widths(ws, [28, 9, 10, 10, 12, 54])
    return ws


def main():
    items = load()

    wb = Workbook()
    wb.remove(wb.active)
    sheet_list(wb, items)
    sheet_by_du(wb, items)
    sheet_by_type(wb, items)

    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / "Арматура_по_размеру.xlsx"
    wb.save(path)

    total = sum(r["вес"] for r in items)
    print(f"готово: {path}\n")
    print(f"{len(items)} позиций, {sum(r['Штук'] for r in items):.0f} шт, "
          f"{total / 1000:.2f} т\n")
    agg = defaultdict(lambda: [0, 0.0, 0.0])
    for r in items:
        a = agg[r["Ду"]]
        a[0] += 1
        a[1] += r["Штук"]
        a[2] += r["вес"]
    run = 0.0
    print(f"{'Ду':>8}{'поз':>6}{'шт':>8}{'кг':>9}{'доля':>7}"
          f"{'нарастающим':>13}")
    for d in sorted(agg, reverse=True):
        pos, qty, kg = agg[d]
        run += kg
        print(f"{('Ду' + str(d)) if d else 'без Ду':>8}{pos:>6}{qty:>8.0f}"
              f"{kg:>9.0f}{kg / total * 100:>6.0f}%{run / total * 100:>12.0f}%")


if __name__ == "__main__":
    main()
