#!/usr/bin/env python3
"""Первый контейнер в Питер: 22 тонны арматуры.

Двадцатифутовый берет 22 тонны. Согласованный список менеджера
(приложения 2 и 3) весит 15,6 т — этого мало, контейнер уйдет
недогруженным на треть. Остальное добираем со склада во Владивостоке.

Веса здесь настоящие. Менеджер взвесила позиции в своем списке, и
они оказались заметно легче наших справочных: клинкет Ду250 — 125 кг
против расчетных 220, Ду300 — 192 против 300. Эти массы и берем, а
где замера нет — пересчитываем от ближайшего диаметра того же типа
по кубу размера.

Чтобы не посчитать один и тот же товар дважды, по каждому сочетанию
«тип плюс диаметр» из владивостокского учета вычитается то, что уже
стоит в списке менеджера.

Результат: outputs/marine_equipment/Контейнер_22т_Питер.xlsx
"""

import pathlib
import re
import sys
from collections import defaultdict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
from build_catalog_xlsx import group_of as catalog_group  # noqa: E402
from build_container_plan import unit_volume  # noqa: E402
from build_marine_registry import SOURCES, parse  # noqa: E402
from build_shipment_spb import (  # noqa: E402
    FLOOR, PACKING, TIERS, is_bronze, is_titanium, item_kind, pack_type,
    tier,
)
from build_spb_list_xlsx import read as read_spb  # noqa: E402
from build_weight_xlsx import du, unit_weight as ref_weight  # noqa: E402

OUT = pathlib.Path(__file__).resolve().parents[2] / "outputs" / "marine_equipment"

FONT = "Arial"
HEAD_FILL = PatternFill("solid", fgColor="1F3864")
HEAD_FONT = Font(name=FONT, size=11, bold=True, color="FFFFFF")
TITLE = Font(name=FONT, size=14, bold=True, color="1F3864")
SUB = Font(name=FONT, size=11, bold=True, color="1F3864")
BOLD = Font(name=FONT, size=10, bold=True)
BASE = Font(name=FONT, size=10)
NOTE = Font(name=FONT, size=9, italic=True, color="666666")
A_FILL = PatternFill("solid", fgColor="EAF3EA")
B_FILL = PatternFill("solid", fgColor="FDF2E3")
MARK_FILL = PatternFill("solid", fgColor="FFF2CC")
NODRAW_FILL = PatternFill("solid", fgColor="FFFF00")
GROUP_FILL = PatternFill("solid", fgColor="D9E2F3")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

LIMIT_KG = 22_000
LIMIT_M3 = 33

FILTERS = re.compile(r"фильтр.*(заборт|заборн)|сетка на фильтр", re.I)

BLOCK_A = "А. Согласованный список"
BLOCK_B = "Б. Добор со склада"


def asrow(r):
    """Строка в том виде, в каком ее ждут функции разбора."""
    return {"Наименование": r["Наименование"],
            "Группа в файле": "",
            "Группа продажи": "Арматура",
            "вес_ед": r.get("вес_ед", 0)}


def key(name):
    """Тип плюс диаметр — по этому ключу сводим два учета."""
    return (item_kind({"Наименование": name, "Группа в файле": ""}),
            du(name) or 0)


def real_weights(spb):
    """Массы по замерам менеджера."""
    acc = defaultdict(list)
    for r in spb:
        if r["Вес 1 шт, кг"]:
            acc[key(r["Наименование"])].append(r["Вес 1 шт, кг"])
    return {k: sum(v) / len(v) for k, v in acc.items()}


def weight_of(name, real):
    """Масса единицы: замер, пересчет от ближайшего диаметра или
    справочник."""
    k = key(name)
    if k in real:
        return real[k]
    same = {kk[1]: v for kk, v in real.items() if kk[0] == k[0] and kk[1]}
    if same and k[1]:
        near = min(same, key=lambda x: abs(x - k[1]))
        # масса арматуры растет примерно как куб размера, но корпус
        # и фланцы тянут медленнее — показатель 2,2 ближе к замерам
        return same[near] * (k[1] / near) ** 2.2
    return ref_weight(name)


def priority(r):
    """Бронза и титан вперед, дальше по цене за килограмм."""
    ppk = r["Сумма"] / r["вес"] if r["вес"] else 0
    first = 0 if (r["бронза"] or r["титан"]) else 1
    return (first, -ppk)


def collect():
    spb = read_spb()
    real = real_weights(spb)

    # блок А: все, что менеджер отобрал на Питер
    block_a = []
    for r in spb:
        qty = r["Кол-во факт"]
        if not qty or r["Отгружено"]:
            # то, что уехало 24 и 29 сентября, второй раз не грузим
            continue
        w = r["Вес 1 шт, кг"] or weight_of(r["Наименование"], real)
        block_a.append({
            "Блок": BLOCK_A,
            "Наименование": r["Наименование"],
            "Чертеж": r["Обозначение"],
            "Ду": r["Ду"],
            "Ру": r["Ру"],
            "Состояние": r["Состояние"],
            "Штук": qty,
            "вес_ед": w,
            "вес": qty * w,
            "объем": qty * unit_volume({"Наименование": r["Наименование"],
                                       "Группа продажи": "Арматура"}),
            "Цена": r["Цена по договору"],
            "Ориентир": r["Ориентир реализации"],
            "Сумма": 0,
            "бронза": "бронз" in r["Наименование"].lower(),
            "титан": "титан" in r["Наименование"].lower(),
            "Выделено": r["Выделено"],
            "Откуда": f"список менеджера, {r['Лист']}",
        })

    # сколько каждого типа и диаметра уже взято
    taken = defaultdict(float)
    for r in spb:
        if not r["Отгружено"]:
            taken[key(r["Наименование"])] += r["Кол-во факт"] or 0

    rows = []
    for spec in SOURCES:
        rows.extend(parse(spec))
    for r in rows:
        r["Сумма"] = (r["Наличие"] or 0) * (r["Цена за ед., руб"] or 0)
    live = [r for r in rows if (r["Наличие"] or 0) > 0]
    pool = [r for r in live
            if catalog_group(r) == "Арматура"
            or FILTERS.search(r["Наименование"])]

    # остаток по каждому ключу после списка менеджера
    have = defaultdict(float)
    for r in pool:
        have[key(r["Наименование"])] += r["Наличие"]
    free = {k: max(0.0, have[k] - taken[k]) for k in have}

    candidates = []
    for r in pool:
        k = key(r["Наименование"])
        if free.get(k, 0) <= 0:
            continue
        qty = min(r["Наличие"], free[k])
        free[k] -= qty
        w = weight_of(r["Наименование"], real)
        price = r["Цена за ед., руб"] or 0
        candidates.append({
            "Блок": BLOCK_B,
            "Наименование": r["Наименование"],
            "Чертеж": r["Чертеж"],
            "Ду": du(r["Наименование"]),
            "Ру": None,
            "Состояние": r["Состояние"],
            "Штук": qty,
            "вес_ед": w,
            "вес": qty * w,
            "объем": qty * unit_volume(dict(r, **{"Группа продажи": "Арматура"})),
            "Цена": price or "",
            "Ориентир": "",
            "Сумма": qty * price,
            "бронза": is_bronze(r),
            "титан": is_titanium(r),
            "Выделено": ("бронза/латунь" if is_bronze(r)
                         else "титан" if is_titanium(r) else ""),
            "Откуда": r["Локация"],
        })

    candidates.sort(key=priority)
    kg = sum(r["вес"] for r in block_a)
    m3 = sum(r["объем"] for r in block_a)
    block_b, rest = [], []
    for r in candidates:
        if kg + r["вес"] <= LIMIT_KG and m3 + r["объем"] <= LIMIT_M3:
            block_b.append(r)
            kg += r["вес"]
            m3 += r["объем"]
        else:
            rest.append(r)
    return block_a, block_b, rest, kg, m3


def head(ws, row, ncols, height=40):
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


def sheet_summary(wb, a, b, rest, kg, m3):
    ws = wb.create_sheet("Сводка", 0)
    ws["A1"] = "Первый контейнер в Петербург: 22 тонны"
    ws["A1"].font = TITLE
    ws["A2"] = ("Двадцатифутовый, 22 т и 33 куб. м. Веса настоящие — "
                "по замерам менеджера, а не по справочнику. Справочные "
                "массы задвижек были завышены почти вдвое.")
    ws["A2"].font = NOTE

    facts = [
        ("Загрузка по весу", f"{kg / 1000:.1f} т",
         f"{kg / LIMIT_KG * 100:.0f}% от 22 т"),
        ("Загрузка по объему", f"{m3:.1f} куб. м",
         f"{m3 / LIMIT_M3 * 100:.0f}% от 33 куб. м"),
        ("Позиций", f"{len(a) + len(b)}",
         f"{sum(r['Штук'] for r in a + b):.0f} штук"),
        ("Блок А, согласованный список",
         f"{sum(r['вес'] for r in a) / 1000:.1f} т",
         f"{len(a)} позиций, {sum(r['Штук'] for r in a):.0f} шт"),
        ("Блок Б, добор со склада",
         f"{sum(r['вес'] for r in b) / 1000:.1f} т",
         f"{len(b)} позиций, {sum(r['Штук'] for r in b):.0f} шт"),
        ("Остается во Владивостоке",
         f"{sum(r['вес'] for r in rest) / 1000:.1f} т",
         f"{len(rest)} позиций"),
    ]
    row = 4
    for label, value, note in facts:
        ws.cell(row=row, column=1, value=label).font = BOLD
        cell = ws.cell(row=row, column=2, value=value)
        cell.font = BOLD
        ws.cell(row=row, column=3, value=note).font = NOTE
        row += 1

    row += 1
    ws.cell(row=row, column=1, value="Что едет, по типам").font = SUB
    row += 1
    for i, h in enumerate(["Тип", "Позиций", "Штук", "Вес, кг",
                           "Объем, куб. м", "Доля веса"], start=1):
        ws.cell(row=row, column=i, value=h)
    head(ws, row, 6, height=24)

    agg = defaultdict(lambda: [0, 0, 0.0, 0.0])
    for r in a + b:
        x = agg[item_kind(asrow(r))]
        x[0] += 1
        x[1] += r["Штук"]
        x[2] += r["вес"]
        x[3] += r["объем"]
    for k, (p, u, w, v) in sorted(agg.items(), key=lambda kv: -kv[1][2]):
        row += 1
        for i, val in enumerate([k, p, u, round(w), round(v, 1), w / kg],
                                start=1):
            cell = ws.cell(row=row, column=i, value=val)
            cell.font = BASE
            cell.border = BORDER
            if i == 4:
                cell.number_format = "#,##0"
            if i == 6:
                cell.number_format = "0.0%"

    row += 2
    ws.cell(row=row, column=1, value="Укладка снизу вверх").font = SUB
    row += 1
    for i, h in enumerate(["Ярус", "Что кладем", "Чем перекладываем",
                           "Вес яруса, кг"], start=1):
        ws.cell(row=row, column=i, value=h)
    head(ws, row, 4, height=24)
    tiers = defaultdict(float)
    for r in a + b:
        tiers[tier(asrow(r))] += r["вес"]
    row += 1
    for i, v in enumerate([FLOOR[0], FLOOR[1], FLOOR[2], "—"], start=1):
        cell = ws.cell(row=row, column=i, value=v)
        cell.font = BASE
        cell.border = BORDER
        cell.fill = GROUP_FILL
        cell.alignment = Alignment(vertical="top", wrap_text=(i in (2, 3)))
    for t in sorted(TIERS):
        row += 1
        label, what, board = TIERS[t]
        for i, v in enumerate([label, what, board, round(tiers[t])],
                              start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=(i in (2, 3)))
            if i == 4:
                cell.number_format = "#,##0"

    row += 2
    for text in [
        "Блок А — это то, что менеджер уже отобрала и согласовала по "
        "приложениям 2 и 3. Он едет целиком.",
        "Блок Б добран со склада по цене за килограмм: бронза и титан "
        "первыми, дальше чем дороже килограмм, тем раньше. За перевозку "
        "платят за вес, поэтому дешевое и тяжелое ждет второго борта.",
        "Чтобы один и тот же товар не уехал дважды, по каждому "
        "сочетанию типа и диаметра из складского учета вычтено то, что "
        "уже стоит в списке менеджера.",
        "Веса блока А — замеры менеджера. В блоке Б замеров нет, там "
        "масса пересчитана от ближайшего диаметра того же типа. Перед "
        "погрузкой характерные места надо взвесить.",
        "Остатки недостоверны: сверка была 16.06.2025. Количества блока "
        "Б подтверждаются только пересчетом на складе.",
    ]:
        cell = ws.cell(row=row, column=1, value="— " + text)
        cell.font = BASE
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[row].height = 30
        row += 1

    widths(ws, [30, 16, 34, 16, 14, 12])
    return ws


COLS = [
    ("№", 5),
    ("Блок", 24),
    ("Ярус", 15),
    ("Тип", 22),
    ("Ду", 7),
    ("Ру", 7),
    ("Наименование", 50),
    ("Чертеж", 22),
    ("Состояние", 12),
    ("Выделено", 22),
    ("Штук", 8),
    ("Вес ед., кг", 11),
    ("Вес всего, кг", 12),
    ("Объем, куб. м", 12),
    ("Цена", 14),
    ("Ориентир реализации", 16),
    ("Как паковать", 32),
    ("Откуда", 22),
]


def sheet_items(wb, name, title, items, note=""):
    ws = wb.create_sheet(name)
    ws["A1"] = title
    ws["A1"].font = TITLE
    if note:
        ws["A2"] = note
        ws["A2"].font = NOTE

    start = 4
    for i, (h, _) in enumerate(COLS, start=1):
        ws.cell(row=start, column=i, value=h)
    head(ws, start, len(COLS))

    items = sorted(items, key=lambda r: (tier(asrow(r)), r["Блок"],
                                         -r["вес"]))
    row = start
    for n, r in enumerate(items, start=1):
        row += 1
        fake = asrow(r)
        values = [n, r["Блок"], TIERS[tier(fake)][0], item_kind(fake),
                  r["Ду"], r["Ру"], r["Наименование"], r["Чертеж"],
                  r["Состояние"], r["Выделено"], r["Штук"],
                  round(r["вес_ед"], 1), round(r["вес"]),
                  round(r["объем"], 3), r["Цена"], r["Ориентир"],
                  pack_type(fake), r["Откуда"]]
        for i, v in enumerate(values, start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top",
                                       wrap_text=(i in (2, 7, 17)))
            if i in (12, 13):
                cell.number_format = "#,##0.0" if i == 12 else "#,##0"
            if i == 14:
                cell.number_format = "#,##0.000"
        ws.cell(row=row, column=2).fill = (A_FILL if r["Блок"] == BLOCK_A
                                           else B_FILL)
        if r["Выделено"]:
            ws.cell(row=row, column=10).fill = MARK_FILL
        if not r["Чертеж"]:
            ws.cell(row=row, column=8).fill = NODRAW_FILL

    last = row
    row += 1
    ws.cell(row=row, column=7, value="ИТОГО").font = BOLD
    for col in (11, 13, 14):
        letter = get_column_letter(col)
        cell = ws.cell(row=row, column=col,
                       value=f"=SUM({letter}{start + 1}:{letter}{last})")
        cell.font = BOLD
        cell.border = BORDER
        cell.number_format = "#,##0" if col != 14 else "#,##0.0"

    ws.auto_filter.ref = f"A{start}:{get_column_letter(len(COLS))}{last}"
    ws.freeze_panes = f"G{start + 1}"
    widths(ws, [w for _, w in COLS])
    return ws


def sheet_packing(wb):
    ws = wb.create_sheet("Упаковка и крепление")
    ws["A1"] = "Как паковать, чтобы доехало без помятых фланцев"
    ws["A1"].font = TITLE
    for i, h in enumerate(["№", "Правило", "Почему так"], start=1):
        ws.cell(row=3, column=i, value=h)
    head(ws, 3, 3, height=24)
    for n, (rule, why) in enumerate(PACKING, start=1):
        row = 3 + n
        for i, v in enumerate([n, rule, why], start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=(i > 1))
        ws.row_dimensions[row].height = 56
    widths(ws, [5, 42, 86])
    return ws


def main():
    a, b, rest, kg, m3 = collect()

    wb = Workbook()
    wb.remove(wb.active)
    sheet_summary(wb, a, b, rest, kg, m3)
    sheet_items(wb, "Контейнер 22 т", "Что грузим в первый контейнер",
                a + b,
                "Блок А — согласованный список менеджера. Блок Б — добор "
                "со склада. Желтым помечен пустой чертеж.")
    sheet_items(wb, "Остается", "Что остается во Владивостоке", rest,
                "Это второй борт: тяжелое и дешевое на килограмм.")
    sheet_packing(wb)

    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / "Контейнер_22т_Питер.xlsx"
    wb.save(path)

    print(f"готово: {path}\n")
    print(f"ВСЕГО:    {len(a) + len(b):>3} поз, "
          f"{sum(r['Штук'] for r in a + b):>5.0f} шт, "
          f"{kg / 1000:>5.2f} т ({kg / LIMIT_KG * 100:.0f}%), "
          f"{m3:>4.1f} куб.м ({m3 / LIMIT_M3 * 100:.0f}%)")
    print(f"  блок А: {len(a):>3} поз, {sum(r['Штук'] for r in a):>5.0f} шт, "
          f"{sum(r['вес'] for r in a) / 1000:>5.2f} т")
    print(f"  блок Б: {len(b):>3} поз, {sum(r['Штук'] for r in b):>5.0f} шт, "
          f"{sum(r['вес'] for r in b) / 1000:>5.2f} т")
    print(f"ОСТАЕТСЯ: {len(rest):>3} поз, "
          f"{sum(r['Штук'] for r in rest):>5.0f} шт, "
          f"{sum(r['вес'] for r in rest) / 1000:>5.2f} т")


if __name__ == "__main__":
    main()
