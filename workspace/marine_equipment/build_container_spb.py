#!/usr/bin/env python3
"""Контейнер в Питер: сорокафутовый на 26 тонн.

Выбран сорокафутовый. Двадцатифутовый до своих тонн не догружается:
внутри у него 33 куб. м, и объем кончается на 21,7 тоннах при любой
грузоподъемности. У сорокафутового 67 куб. м, упор идет в вес, и за
один заход уезжает почти вся арматура.

Согласованный список менеджера (приложения 2 и 3) весит 15,6 т —
этого мало на контейнер, остальное добираем со склада во Владивостоке.

Веса здесь настоящие. Менеджер взвесила позиции в своем списке, и
они оказались заметно легче наших справочных: клинкет Ду250 — 125 кг
против расчетных 220, Ду300 — 192 против 300. Эти массы и берем, а
где замера нет — пересчитываем от ближайшего диаметра того же типа
по кубу размера.

Чтобы не посчитать один и тот же товар дважды, по каждому сочетанию
«тип плюс диаметр» из владивостокского учета вычитается то, что уже
стоит в списке менеджера.

Результат: outputs/marine_equipment/Контейнер_40ф_Питер.xlsx
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
from build_spb_list_xlsx import (  # noqa: E402
    GROUPS, group_of, read as read_spb,
)
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

# сорокафутовый: 26 тонн под груз и 67,2 куб. м внутри
LIMIT_KG = 26_000
LIMIT_M3 = 67.2

# Ящик, сбитый из поддонов: поддон 1200х1000 в основании, высота около
# метра. Полезный объем примерно кубометр, но набивать его доверху
# нельзя — доска бэушная, без рамного бруса по углам.
BOX_KG = 600
BOX_M3 = 0.9
BOX_SIZE = "1200 x 1000 x 910 мм, на поддоне"
# пол сорокафутового
FLOOR_L, FLOOR_W = 12.032, 2.352
FLOOR_AREA = FLOOR_L * FLOOR_W
# доска необрезная 25 мм: кубометр на сорок квадратов настила
BOARD_MM = 25
# товарная доска 4 м, средняя ширина обзольной доски около 20 см
BOARD_L, BOARD_W = 4.0, 0.20
# обзол, подгонка и бой: необрезную всегда берут с запасом
BOARD_WASTE = 1.25
# пол сорокафутового 12,03 х 2,35 м: ящиков 1,2 х 1,0 влезает 22 в ярус
BOX_PER_LAYER = 22

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


def collect(limit_kg=LIMIT_KG, limit_m3=LIMIT_M3):
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
        if kg + r["вес"] <= limit_kg and m3 + r["объем"] <= limit_m3:
            block_b.append(r)
            kg += r["вес"]
            m3 += r["объем"]
        else:
            rest.append(r)
    return block_a, block_b, rest, kg, m3


def plural(n, one, few, many):
    """Позиция, позиции, позиций — по правилам русского счета."""
    n = int(n)
    if n % 100 in range(11, 15):
        return many
    last = n % 10
    if last == 1:
        return one
    if last in (2, 3, 4):
        return few
    return many


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
    ws["A1"] = "Контейнер в Петербург: сорокафутовый, 26 тонн"
    ws["A1"].font = TITLE
    ws["A2"] = ("Сорокафутовый, 26 т и 67,2 куб. м. Веса настоящие — "
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
    ws.cell(row=row, column=1, value="Что едет, по группам").font = SUB
    row += 1
    for i, h in enumerate(["Группа", "Позиций", "Штук", "Вес, кг",
                           "Объем, куб. м", "Доля веса"], start=1):
        ws.cell(row=row, column=i, value=h)
    head(ws, row, 6, height=24)

    agg = defaultdict(lambda: [0, 0, 0.0, 0.0])
    for r in a + b:
        x = agg[group_of(r["Наименование"])]
        x[0] += 1
        x[1] += r["Штук"]
        x[2] += r["вес"]
        x[3] += r["объем"]
    first = row + 1
    ordered = [g for g in GROUPS if g in agg]
    ordered += [g for g in sorted(agg) if g not in GROUPS]
    for k in ordered:
        p, u, w, v = agg[k]
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
    row += 1
    ws.cell(row=row, column=1, value="ИТОГО").font = BOLD
    for col in range(2, 7):
        letter = get_column_letter(col)
        cell = ws.cell(row=row, column=col,
                       value=f"=SUM({letter}{first}:{letter}{row - 1})")
        cell.font = BOLD
        cell.border = BORDER
        cell.number_format = "0.0%" if col == 6 else "#,##0"
        if col == 5:
            cell.number_format = "#,##0.0"

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
    first = row + 1
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
    row += 1
    ws.cell(row=row, column=1, value="ИТОГО").font = BOLD
    cell = ws.cell(row=row, column=4,
                   value=f"=SUM(D{first}:D{row - 1})")
    cell.font = BOLD
    cell.border = BORDER
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
        "Сорокафутовый упирается в вес, а не в объем: 26 тонн занимают "
        "около 38 куб. м из 67. Половина кузова свободна, поэтому "
        "ошибка в расчете упаковки ничего не ломает.",
        "Двадцатифутовый за 315 тысяч взял бы 21,7 т и оставил 5,6 т: "
        "его 33 куб. м кончаются раньше веса. Разница в цене "
        "контейнеров 120 тысяч, а вывезти остаток транспортной "
        "компанией стоило бы 452 тысячи.",
        "Не влезли три позиции без цены: коробка 2-х клапановая Ду80, "
        "задвижки желтые около крышек и блины к клинкету Ду100. По "
        "договору товар без согласованной цены все равно не считается "
        "проданным, так что оставить их правильно.",
    ]:
        cell = ws.cell(row=row, column=1, value="— " + text)
        cell.font = BASE
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[row].height = 30
        row += 1

    widths(ws, [30, 16, 34, 16, 14, 12])
    return ws


# группы и их порядок берем те же, что в списке остатков: от мелкой
# арматуры к крупной. Один разрез на оба файла, чтобы не сверять
# вручную два разных деления
COLS = [
    ("№", 5),
    ("Наименование", 50),
    ("Чертеж", 22),
    ("Ду", 7),
    ("Ру", 7),
    ("Состояние", 12),
    ("ШТУК", 9),
    ("Вес ед., кг", 11),
    ("Вес всего, кг", 12),
    ("Объем, куб. м", 12),
    ("Выделено", 22),
    ("Блок", 24),
    ("Ярус", 15),
    ("Как паковать", 32),
    ("Откуда", 22),
    ("Цена", 14),
    ("Ориентир реализации", 16),
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

    by_group = defaultdict(list)
    for r in items:
        by_group[group_of(r["Наименование"])].append(r)
    order = [g for g in GROUPS if g in by_group]
    order += [g for g in sorted(by_group) if g not in GROUPS]

    row = start
    n = 0
    for group in order:
        rows_in = sorted(by_group[group],
                         key=lambda r: (r["Ду"] or 9999,
                                        r["Наименование"]))
        row += 1
        qty = sum(r["Штук"] for r in rows_in)
        kg = sum(r["вес"] for r in rows_in)
        # разделитель тысяч ставим отдельно: в названии группы есть
        # запятая, и общий replace ее бы съел
        kg_text = f"{kg:,.0f}".replace(",", " ")
        ws.cell(row=row, column=1,
                value=f"{group} — {len(rows_in)} "
                      f"{plural(len(rows_in), 'позиция', 'позиции', 'позиций')}"
                      f", {qty:.0f} шт, {kg_text} кг")
        for c in range(1, len(COLS) + 1):
            cell = ws.cell(row=row, column=c)
            cell.fill = GROUP_FILL
            cell.font = SUB
            cell.border = BORDER
        for r in rows_in:
            row += 1
            n += 1
            fake = asrow(r)
            values = [n, r["Наименование"], r["Чертеж"], r["Ду"], r["Ру"],
                      r["Состояние"], r["Штук"], round(r["вес_ед"], 1),
                      round(r["вес"]), round(r["объем"], 3),
                      r["Выделено"], r["Блок"], TIERS[tier(fake)][0],
                      pack_type(fake), r["Откуда"], r["Цена"],
                      r["Ориентир"]]
            for i, v in enumerate(values, start=1):
                cell = ws.cell(row=row, column=i, value=v)
                cell.font = BASE
                cell.border = BORDER
                cell.alignment = Alignment(vertical="top",
                                           wrap_text=(i in (2, 12, 14)))
                if i == 7:
                    cell.font = BOLD
                if i == 8:
                    cell.number_format = "#,##0.0"
                if i == 9:
                    cell.number_format = "#,##0"
                if i == 10:
                    cell.number_format = "#,##0.000"
            ws.cell(row=row, column=12).fill = (A_FILL
                                                if r["Блок"] == BLOCK_A
                                                else B_FILL)
            if r["Выделено"]:
                ws.cell(row=row, column=11).fill = MARK_FILL
            if not r["Чертеж"]:
                ws.cell(row=row, column=3).fill = NODRAW_FILL

    last = row
    row += 1
    ws.cell(row=row, column=2, value="ИТОГО").font = BOLD
    for col in (7, 9, 10):
        letter = get_column_letter(col)
        cell = ws.cell(row=row, column=col,
                       value=f"=SUM({letter}{start + 1}:{letter}{last})")
        cell.font = BOLD
        cell.border = BORDER
        cell.number_format = "#,##0" if col != 10 else "#,##0.0"

    ws.auto_filter.ref = f"A{start}:{get_column_letter(len(COLS))}{last}"
    ws.freeze_panes = f"C{start + 1}"
    widths(ws, [w for _, w in COLS])
    return ws


def sheet_boxes(wb, items):
    """Сколько ящиков надо сбить и что в каждый кладем.

    В ящики идет то, что лежит верхними ярусами: мелкая арматура и
    фильтры. Задвижки и крупные фланцевые клапаны в ящики не кладем,
    они встают на настил и крепятся к полу.
    """
    ws = wb.create_sheet("Ящики и места")
    ws["A1"] = "Сколько ящиков сбить и что в них класть"
    ws["A1"].font = TITLE
    ws["A2"] = (f"Ящик {BOX_SIZE}. Считаем не больше {BOX_KG} кг и "
                f"{BOX_M3} куб. м на ящик: доска бэушная, рамного бруса "
                "по углам нет. Позиции одной группы стараемся не "
                "разносить по разным ящикам.")
    ws["A2"].font = NOTE

    boxed = [r for r in items if tier(asrow(r)) in (3, 4)]
    floored = [r for r in items if tier(asrow(r)) in (1, 2)]

    # раскладываем по ящикам: группа за группой, не превышая лимиты
    by_group = defaultdict(list)
    for r in boxed:
        by_group[group_of(r["Наименование"])].append(r)
    order = [g for g in GROUPS if g in by_group]
    order += [g for g in sorted(by_group) if g not in GROUPS]

    boxes = []
    cur = {"вес": 0.0, "объем": 0.0, "строки": []}
    for group in order:
        for r in sorted(by_group[group], key=lambda r: (r["Ду"] or 9999,
                                                        r["Наименование"])):
            # позицию, которая одна тяжелее ящика, делим по штукам:
            # 798 клапанов Ду10 это почти тонна, в один ящик нельзя
            left = r["Штук"]
            w1 = r["вес"] / r["Штук"] if r["Штук"] else 0
            v1 = r["объем"] / r["Штук"] if r["Штук"] else 0
            while left > 0:
                room_kg = BOX_KG - cur["вес"]
                room_m3 = BOX_M3 - cur["объем"]
                fit = left
                if w1:
                    fit = min(fit, int(room_kg // w1))
                if v1:
                    fit = min(fit, int(room_m3 // v1))
                if fit < 1:
                    if not cur["строки"]:
                        # одна штука не влезает ни по весу, ни по объему
                        fit = 1
                    else:
                        boxes.append(cur)
                        cur = {"вес": 0.0, "объем": 0.0, "строки": []}
                        continue
                cur["вес"] += fit * w1
                cur["объем"] += fit * v1
                cur["строки"].append(dict(r, Штук=fit, вес=fit * w1,
                                          объем=fit * v1))
                left -= fit
    if cur["строки"]:
        boxes.append(cur)

    # тяжелое вниз, легкое наверх: самые тяжелые ящики в нижний ярус,
    # фильтры всегда сверху — они и легкие, и мятся первыми
    def box_weight(box):
        filters = any(group_of(r["Наименование"]) == "Фильтры забортной воды"
                      for r in box["строки"])
        return (1 if filters else 0, -box["вес"])

    boxes.sort(key=box_weight)
    for i, box in enumerate(boxes):
        if i < BOX_PER_LAYER:
            box["ярус"] = "Нижний ярус ящиков"
        elif i < 2 * BOX_PER_LAYER:
            box["ярус"] = "Верхний ярус ящиков"
        else:
            box["ярус"] = "МЕСТА НЕТ — разложить или оставить"

    row = 4
    facts = [
        ("Ящиков собрать", f"{len(boxes)} шт", BOX_SIZE),
        ("В ящики идет", f"{sum(r['вес'] for r in boxed) / 1000:.1f} т",
         f"{len(boxed)} позиций, "
         f"{sum(r['Штук'] for r in boxed):.0f} штук"),
        ("На настил и поддоны",
         f"{sum(r['вес'] for r in floored) / 1000:.1f} т",
         f"{len(floored)} позиций — задвижки и крупные клапаны, "
         "в ящики не кладем"),
        ("Поддонов под ящики", f"{len(boxes)} шт",
         "по одному на ящик, плюс запас на бой"),
        ("Мест под ящики в контейнере", f"{2 * BOX_PER_LAYER} шт",
         f"{BOX_PER_LAYER} в ярус, два яруса поверх настила"),
    ]
    extra = len(boxes) - 2 * BOX_PER_LAYER
    if extra > 0:
        facts.append(("Ящиков сверх мест", f"{extra} шт",
                      "содержимое разложить по остальным или оставить "
                      "на второй борт"))
    for label, value, note in facts:
        ws.cell(row=row, column=1, value=label).font = BOLD
        ws.cell(row=row, column=2, value=value).font = BOLD
        ws.cell(row=row, column=3, value=note).font = NOTE
        row += 1

    row += 1
    cols = ["Ящик №", "Куда ставить", "Что внутри", "Наименование",
            "Чертеж", "Ду", "Штук", "Вес, кг", "Объем, куб. м"]
    for i, h in enumerate(cols, start=1):
        ws.cell(row=row, column=i, value=h)
    head(ws, row, len(cols), height=28)

    for n, box in enumerate(boxes, start=1):
        groups = sorted({group_of(r["Наименование"]) for r in box["строки"]})
        row += 1
        ws.cell(row=row, column=1, value=f"ЯЩИК {n}")
        ws.cell(row=row, column=2, value=box["ярус"])
        ws.cell(row=row, column=3, value=", ".join(groups))
        ws.cell(row=row, column=7,
                value=sum(r["Штук"] for r in box["строки"]))
        ws.cell(row=row, column=8, value=round(box["вес"]))
        ws.cell(row=row, column=9, value=round(box["объем"], 2))
        for c in range(1, len(cols) + 1):
            cell = ws.cell(row=row, column=c)
            cell.fill = GROUP_FILL
            cell.font = SUB
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=(c == 3))
            if c == 8:
                cell.number_format = "#,##0"
        for r in box["строки"]:
            row += 1
            values = [None, None, group_of(r["Наименование"]),
                      r["Наименование"], r["Чертеж"], r["Ду"], r["Штук"],
                      round(r["вес"]), round(r["объем"], 3)]
            for i, v in enumerate(values, start=1):
                cell = ws.cell(row=row, column=i, value=v)
                cell.font = BASE
                cell.border = BORDER
                cell.alignment = Alignment(vertical="top",
                                           wrap_text=(i in (3, 4)))
                if i == 8:
                    cell.number_format = "#,##0"
                if i == 9:
                    cell.number_format = "#,##0.000"
            if not r["Чертеж"]:
                ws.cell(row=row, column=5).fill = NODRAW_FILL

    row += 2
    ws.cell(row=row, column=1, value="Что учесть по таре").font = SUB
    row += 1
    for text in [
        f"Ящик из бэушной доски на {BOX_KG} кг держит только при брусе "
        "50х50 по четырем углам и по верху. Без него стенка сложится "
        "при первом же ударе на маневрировании.",
        "Дно ящика набиваем доской вплотную, без щелей шире ладони: "
        "мелкий клапан проваливается и бьет соседей.",
        "На каждый ящик — бирка: номер, группа, вес, строки реестра и "
        "стрелки «верх». Принимать груз будут по упаковочному листу.",
        "Задвижки и крупные фланцевые клапаны в ящики не кладем: они "
        "встают прямо на настил, с прокладкой доской между рядами и "
        "креплением к рымам.",
        "Фильтры идут верхним ярусом на ящиках, поэтому ящики под ними "
        "должны быть самыми крепкими — туда ставим те, что с брусом.",
    ]:
        cell = ws.cell(row=row, column=1, value="— " + text)
        cell.font = BASE
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[row].height = 30
        row += 1

    widths(ws, [12, 22, 28, 46, 20, 8, 9, 11, 13])
    ws.freeze_panes = "D5"
    return ws


def sheet_timber(wb, items):
    """Сколько необрезной доски 25 мм купить на перекладку.

    Считаем по площадям настилов, а не на глаз. Настил идет под
    клинкеты, поверх клинкетов под ящики, прокладками внутри ящиков
    между рядами арматуры, плюс распорки и клинья.
    """
    ws = wb.create_sheet("Лес на перекладку")
    ws["A1"] = "Сколько доски купить"
    ws["A1"].font = TITLE
    ws["A2"] = (f"Доска необрезная {BOARD_MM} мм. Пол контейнера "
                f"{FLOOR_L:.2f} x {FLOOR_W:.2f} = {FLOOR_AREA:.1f} кв. м. "
                "Количество считано по площади настилов с запасом "
                f"{(BOARD_WASTE - 1) * 100:.0f}% на обзол и подгонку.")
    ws["A2"].font = NOTE

    gates = [r for r in items
             if item_kind(asrow(r)) == "клинкет (задвижка)"
             and (r["Ду"] or 0) >= 200]
    gate_area = sum((((r["Ду"] or 100) / 1000 + 0.20) ** 2) * r["Штук"]
                    for r in gates)
    boxed = [r for r in items if tier(asrow(r)) in (3, 4)]
    boxes = int(-(-sum(r["вес"] for r in boxed) // BOX_KG))

    row = 4
    for label, value, note in [
        ("Клинкеты Ду200 и выше",
         f"{sum(r['Штук'] for r in gates):.0f} шт",
         f"{sum(r['вес'] for r in gates) / 1000:.2f} т"),
        ("Их пятно на полу", f"{gate_area:.1f} кв. м",
         f"{gate_area / FLOOR_AREA * 100:.0f}% пола — ложатся одним слоем, "
         "второй ярус клинкетов не нужен"),
        ("Ящиков", f"{boxes} шт", f"{BOX_SIZE}"),
    ]:
        ws.cell(row=row, column=1, value=label).font = BOLD
        ws.cell(row=row, column=2, value=value).font = BOLD
        ws.cell(row=row, column=3, value=note).font = NOTE
        row += 1

    row += 1
    cols = ["Куда идет доска", "Площадь, кв. м", "Пояснение"]
    for i, h in enumerate(cols, start=1):
        ws.cell(row=row, column=i, value=h)
    head(ws, row, 3, height=24)

    # прокладки внутри ящика: два ряда, доской закрыто около трети
    inside = boxes * 1.2 * 1.0 * 0.3 * 2
    lines = [
        ("Настил на пол под клинкеты", gate_area,
         "сплошной, чтобы задвижка стояла на доске, а не на рифленом "
         "железе"),
        ("Настил поверх клинкетов", gate_area,
         "по нему встают ящики, доска ложится над корпусами, а не над "
         "маховиками"),
        ("Прокладки внутри ящиков", inside,
         f"{boxes} ящиков, по два прокладочных ряда, доской закрыта "
         "примерно треть площади"),
    ]
    total = sum(x[1] for x in lines)
    lines.append(("Распорки, клинья, подгонка", total * 0.10,
                  "десятая часть от настилов"))
    total = sum(x[1] for x in lines)

    first = row + 1
    for label, area, note in lines:
        row += 1
        for i, v in enumerate([label, round(area, 1), note], start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=(i == 3))
            if i == 2:
                cell.number_format = "#,##0.0"
    row += 1
    ws.cell(row=row, column=1, value="ИТОГО площадь настилов").font = BOLD
    cell = ws.cell(row=row, column=2, value=f"=SUM(B{first}:B{row - 1})")
    cell.font = BOLD
    cell.border = BORDER
    cell.number_format = "#,##0.0"

    row += 2
    ws.cell(row=row, column=1, value="Сколько это в лесе").font = SUB
    row += 1
    volume = total * BOARD_MM / 1000
    with_waste = volume * BOARD_WASTE
    per_board = BOARD_L * BOARD_W
    boards = int(-(-total * BOARD_WASTE // per_board))
    for label, value, note in [
        ("Чистый объем доски, куб. м", round(volume, 2),
         f"{total:.1f} кв. м настила толщиной {BOARD_MM} мм"),
        ("С запасом на обзол и подгонку, куб. м", round(with_waste, 2),
         f"коэффициент {BOARD_WASTE}"),
        ("Доска 4 м, средняя ширина 20 см", f"{boards} шт",
         f"одна доска закрывает {per_board:.1f} кв. м"),
        ("Погонных метров", f"{boards * BOARD_L:.0f} м", ""),
    ]:
        ws.cell(row=row, column=1, value=label).font = BOLD
        cell = ws.cell(row=row, column=2, value=value)
        cell.font = BOLD
        if isinstance(value, float):
            cell.number_format = "#,##0.00"
        ws.cell(row=row, column=3, value=note).font = NOTE
        row += 1

    row += 1
    ws.cell(row=row, column=1, value="Что еще купить").font = SUB
    row += 1
    for text in [
        f"Поддоны под ящики — {boxes} штук плюс запас на бой.",
        "Брус 50х50 на каркас ящиков: по четырем углам и по верхней "
        "обвязке. Без него ящик на 600 кг складывается от первого же "
        "удара на маневрировании.",
        "Стяжные ремни для крепления к рымам и фанера на заглушки "
        "фланцев.",
        "Необрезную доску берут нестроганой и невысушенной — для "
        "перекладки это нормально, но сырая доска тяжелее: "
        f"{with_waste:.1f} куб. м это около "
        f"{with_waste * 700:.0f} кг к весу контейнера. Это надо учесть "
        "в лимите.",
    ]:
        cell = ws.cell(row=row, column=1, value="— " + text)
        cell.font = BASE
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[row].height = 30
        row += 1

    widths(ws, [40, 20, 60])
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
    sheet_items(wb, "Контейнер 40 футов", "Что грузим в контейнер",
                a + b,
                "Блок А — согласованный список менеджера. Блок Б — добор "
                "со склада. Желтым помечен пустой чертеж.")
    sheet_items(wb, "Остается", "Что остается во Владивостоке", rest,
                "Это второй борт: тяжелое и дешевое на килограмм.")
    sheet_boxes(wb, a + b)
    sheet_timber(wb, a + b)
    sheet_packing(wb)

    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / "Контейнер_40ф_Питер.xlsx"
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
