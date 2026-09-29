#!/usr/bin/env python3
"""Два двадцатифутовых контейнера в Санкт-Петербург.

Сорокафутовый упирается в вес, а не в объем: арматура плотная, 25 тонн
набираются на половине объема, и 7,6 тонны остаются во Владивостоке.
Два двадцатифутовых дают 40 тонн вместо 25 и забирают всю арматуру за
один заход, с запасом по весу и по объему.

Контейнер 1 грузим первым: бронза, латунь, титан и все, что дороже на
килограмм. Контейнер 2 забирает остаток — тяжелую сталь, клапанные
коробки и мелочь в ящиках. Если второй уходит позже, первый уже везет
самое ценное.

Результат: outputs/marine_equipment/План_2_контейнера.xlsx
"""

import pathlib
import sys
from collections import defaultdict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
from build_shipment_spb import (  # noqa: E402
    FLOOR, LIMIT_KG, LIMIT_M3, PACKING, QUEUE, TIERS, hold, is_bronze,
    is_titanium, item_kind, material, pack_type, priority, queue, select,
    tier,
)
from build_weight_xlsx import du  # noqa: E402

OUT = pathlib.Path(__file__).resolve().parents[2] / "outputs" / "marine_equipment"

FONT = "Arial"
HEAD_FILL = PatternFill("solid", fgColor="1F3864")
HEAD_FONT = Font(name=FONT, size=11, bold=True, color="FFFFFF")
TITLE = Font(name=FONT, size=14, bold=True, color="1F3864")
SUB = Font(name=FONT, size=11, bold=True, color="1F3864")
BOLD = Font(name=FONT, size=10, bold=True)
BASE = Font(name=FONT, size=10)
NOTE = Font(name=FONT, size=9, italic=True, color="666666")
ONE_FILL = PatternFill("solid", fgColor="EAF3EA")
TWO_FILL = PatternFill("solid", fgColor="FDF2E3")
WARN_FILL = PatternFill("solid", fgColor="FCE4D6")
SUB_FILL = PatternFill("solid", fgColor="D9E2F3")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# двадцатифутовый: 33 куб. м внутри, по весу берем 20 тонн, как
# и решили. По паспорту в него входит больше, но 20 тонн — это то,
# что спокойно едет по дорогам без разрешений
BOX_KG = 20_000
BOX_M3 = 33
BOXES = 2

NAMES = {1: "Контейнер 1", 2: "Контейнер 2"}


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


def first_first(r):
    """Порядок загрузки для двух контейнеров.

    В одном контейнере мелочь стояла последней: она вытесняла тяжелое
    и оставалась во Владивостоке. Здесь уезжает все, вопрос только в
    том, что поедет первым бортом. Поэтому мелочь больше не отодвигаем:
    штуцерный клапан дает 5 758 руб с килограмма против 784 у задвижки,
    и в первом контейнере ему место.
    """
    q = queue(r)
    if q == 5:
        q = 4
    price_per_kg = r["Сумма"] / r["вес"] if r["вес"] else 0
    return (q, -price_per_kg)


def split():
    """Разложить всю арматуру по двум контейнерам.

    Сначала то, что назвал покупатель, дальше по цене за килограмм.
    Первый контейнер набирается доверху самым дорогим, остальное идет
    во второй.
    """
    go, stay, held, _, _ = select()
    pool = sorted(go + stay + held, key=first_first)

    boxes = {n: [] for n in range(1, BOXES + 1)}
    load = {n: [0.0, 0.0] for n in range(1, BOXES + 1)}
    rest = []
    for r in pool:
        for n in range(1, BOXES + 1):
            kg, m3 = load[n]
            if kg + r["вес"] <= BOX_KG and m3 + r["объем"] <= BOX_M3:
                r["Контейнер"] = n
                boxes[n].append(r)
                load[n] = [kg + r["вес"], m3 + r["объем"]]
                break
        else:
            r["Контейнер"] = 0
            rest.append(r)
    return boxes, load, rest


def sheet_summary(wb, boxes, load, rest):
    ws = wb.create_sheet("Сводка", 0)
    ws["A1"] = "Два двадцатифутовых контейнера в Петербург"
    ws["A1"].font = TITLE
    ws["A2"] = (f"По {BOX_KG // 1000} тонн и {BOX_M3} куб. м на контейнер. "
                "Арматура плотная: вес упирается раньше объема, поэтому "
                "два двадцатифутовых выгоднее одного сорокафутового.")
    ws["A2"].font = NOTE

    row = 4
    ws.cell(row=row, column=1, value="Сравнение вариантов").font = SUB
    row += 1
    for i, h in enumerate(["Вариант", "Лимит веса", "Лимит объема",
                           "Увезет арматуры", "Останется во Владивостоке"],
                          start=1):
        ws.cell(row=row, column=i, value=h)
    head(ws, row, 5, height=24)

    total_kg = sum(load[n][0] for n in boxes) + sum(r["вес"] for r in rest)
    total_val = (sum(r["Сумма"] for n in boxes for r in boxes[n])
                 + sum(r["Сумма"] for r in rest))
    packed_kg = sum(load[n][0] for n in boxes)
    packed_val = sum(r["Сумма"] for n in boxes for r in boxes[n])

    variants = [
        ("Один сорокафутовый", f"{LIMIT_KG / 1000:.0f} т",
         f"{LIMIT_M3} куб. м", f"{LIMIT_KG / 1000:.0f} т",
         f"{(total_kg - LIMIT_KG) / 1000:.1f} т"),
        (f"Два двадцатифутовых", f"{BOXES * BOX_KG / 1000:.0f} т",
         f"{BOXES * BOX_M3} куб. м", f"{packed_kg / 1000:.1f} т",
         f"{sum(r['вес'] for r in rest) / 1000:.1f} т"),
    ]
    for v in variants:
        row += 1
        for i, val in enumerate(v, start=1):
            cell = ws.cell(row=row, column=i, value=val)
            cell.font = BOLD if "двадцати" in v[0] else BASE
            cell.border = BORDER
            if "двадцати" in v[0]:
                cell.fill = ONE_FILL

    row += 2
    ws.cell(row=row, column=1, value="Что в каком контейнере").font = SUB
    row += 1
    for i, h in enumerate(["Контейнер", "Позиций", "Штук", "Вес, т",
                           "Загрузка по весу", "Объем, куб. м",
                           "Загрузка по объему", "Стоимость, руб"],
                          start=1):
        ws.cell(row=row, column=i, value=h)
    head(ws, row, 8, height=28)

    for n in sorted(boxes):
        row += 1
        kg, m3 = load[n]
        items = boxes[n]
        values = [NAMES[n], len(items), sum(r["Наличие"] for r in items),
                  round(kg / 1000, 2), kg / BOX_KG, round(m3, 1),
                  m3 / BOX_M3, sum(r["Сумма"] for r in items)]
        for i, v in enumerate(values, start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            cell.fill = ONE_FILL if n == 1 else TWO_FILL
            if i in (5, 7):
                cell.number_format = "0%"
            if i == 8:
                cell.number_format = "#,##0"
    if rest:
        row += 1
        values = ["Не влезло", len(rest), sum(r["Наличие"] for r in rest),
                  round(sum(r["вес"] for r in rest) / 1000, 2), None,
                  round(sum(r["объем"] for r in rest), 1), None,
                  sum(r["Сумма"] for r in rest)]
        for i, v in enumerate(values, start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            cell.fill = WARN_FILL
            if i == 8:
                cell.number_format = "#,##0"

    row += 2
    ws.cell(row=row, column=1, value="Что это дает").font = SUB
    row += 1
    for text in [
        f"Вся арматура и фильтры — это {total_kg / 1000:.1f} т и "
        f"{total_val / 1e6:.1f} млн руб. В два двадцатифутовых входит "
        f"{packed_kg / 1000:.1f} т на {packed_val / 1e6:.1f} млн, то есть "
        "практически все за один заход.",
        f"Сорокафутовый брал {LIMIT_KG / 1000:.0f} т и оставлял "
        f"{(total_kg - LIMIT_KG) / 1000:.1f} т во Владивостоке. Разница "
        "в пользу двух двадцатифутовых — целый второй рейс, которого "
        "не надо делать.",
        "Объем при этом занят чуть больше чем наполовину: груз плотный, "
        "контейнеры упираются в вес. Значит, добрать объемом ничего "
        "не выйдет, зато остается место под правильную укладку и "
        "прокладки.",
        "Двадцатифутовый на 20 тонн спокойно едет по дорогам без "
        "спецразрешений. Сорокафутовый на 25 тонн на автоплече часто "
        "требует согласований.",
        "Контейнеры можно отправить в разное время: первый везет "
        "бронзу, латунь и титан, второй ждет хоть до инвентаризации.",
        "Минус один: два контейнера — это два подъема краном, две "
        "подачи и обычно дороже по ставке, чем один сорокафутовый. "
        "Ставку надо запросить у перевозчика на оба варианта.",
    ]:
        cell = ws.cell(row=row, column=1, value="— " + text)
        cell.font = BASE
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[row].height = 30
        row += 1

    widths(ws, [30, 14, 12, 12, 16, 14, 16, 18])
    return ws


COLS = ["№", "Очередь", "Ярус", "Тип", "Диаметр", "Наименование",
        "Материал", "Штук", "Вес, кг", "Объем, м³", "Стоимость, руб",
        "Руб за кг", "Как паковать", "Локация", "Отметка"]


def sheet_box(wb, n, items, load):
    ws = wb.create_sheet(NAMES[n])
    kg, m3 = load
    ws["A1"] = f"{NAMES[n]}: что грузим"
    ws["A1"].font = TITLE
    ws["A2"] = (f"{len(items)} позиций, {sum(r['Наличие'] for r in items):.0f}"
                f" штук, {kg / 1000:.1f} т из {BOX_KG // 1000} т, "
                f"{m3:.1f} куб. м из {BOX_M3}.")
    ws["A2"].font = NOTE

    start = 4
    for i, h in enumerate(COLS, start=1):
        ws.cell(row=start, column=i, value=h)
    head(ws, start, len(COLS))

    items = sorted(items, key=lambda r: (tier(r), queue(r),
                                         -(r["Сумма"] / r["вес"]
                                           if r["вес"] else 0)))
    for k, r in enumerate(items, start=1):
        row = start + k
        d = du(r["Наименование"])
        values = [k, QUEUE[queue(r)], TIERS[tier(r)][0], item_kind(r),
                  f"Ду{d}" if d else "", r["Наименование"], material(r),
                  r["Наличие"], round(r["вес"]), round(r["объем"], 3),
                  r["Сумма"] or None,
                  round(r["Сумма"] / r["вес"]) if r["вес"] and r["Сумма"]
                  else None,
                  pack_type(r), r["Локация"], hold(r)]
        for i, v in enumerate(values, start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top",
                                       wrap_text=(i in (2, 6, 13)))
            if i in (9, 11, 12):
                cell.number_format = "#,##0"
            if i == 10:
                cell.number_format = "#,##0.000"
        if is_bronze(r) or is_titanium(r):
            ws.cell(row=row, column=7).fill = ONE_FILL
        if hold(r):
            ws.cell(row=row, column=15).fill = WARN_FILL

    last = start + len(items)
    ws.cell(row=last + 1, column=6, value="ИТОГО").font = BOLD
    for col in (8, 9, 11):
        letter = get_column_letter(col)
        cell = ws.cell(row=last + 1, column=col,
                       value=f"=SUM({letter}{start + 1}:{letter}{last})")
        cell.font = BOLD
        cell.border = BORDER
        cell.number_format = "#,##0"

    # ярусы именно этого контейнера
    row = last + 3
    ws.cell(row=row, column=1, value="Укладка в этом контейнере").font = SUB
    row += 1
    for i, h in enumerate(["Ярус", "Что кладем", "Чем перекладываем сверху",
                           "Вес яруса, кг"], start=1):
        ws.cell(row=row, column=i, value=h)
    head(ws, row, 4, height=24)

    tiers = defaultdict(float)
    for r in items:
        tiers[tier(r)] += r["вес"]
    row += 1
    for i, v in enumerate([FLOOR[0], FLOOR[1], FLOOR[2], "—"], start=1):
        cell = ws.cell(row=row, column=i, value=v)
        cell.font = BASE
        cell.border = BORDER
        cell.fill = SUB_FILL
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

    ws.auto_filter.ref = f"A{start}:O{last}"
    ws.freeze_panes = f"F{start + 1}"
    widths(ws, [5, 26, 15, 22, 10, 48, 20, 8, 11, 12, 15, 11, 32, 14, 30])
    return ws


def sheet_packing(wb):
    ws = wb.create_sheet("Упаковка и крепление")
    ws["A1"] = "Как паковать, чтобы доехало без помятых фланцев"
    ws["A1"].font = TITLE
    ws["A2"] = ("Правила одинаковы для обоих контейнеров. Дорога "
                "занимает около месяца, с продольными ударами при "
                "маневрировании.")
    ws["A2"].font = NOTE
    for i, h in enumerate(["№", "Правило", "Почему так"], start=1):
        ws.cell(row=4, column=i, value=h)
    head(ws, 4, 3, height=24)
    for n, (rule, why) in enumerate(PACKING, start=1):
        row = 4 + n
        for i, v in enumerate([n, rule, why], start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=(i > 1))
        ws.row_dimensions[row].height = 58
    widths(ws, [5, 42, 86])
    return ws


def main():
    boxes, load, rest = split()

    wb = Workbook()
    wb.remove(wb.active)
    sheet_summary(wb, boxes, load, rest)
    for n in sorted(boxes):
        sheet_box(wb, n, boxes[n], load[n])
    sheet_packing(wb)

    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / "План_2_контейнера.xlsx"
    wb.save(path)

    print(f"готово: {path}\n")
    for n in sorted(boxes):
        kg, m3 = load[n]
        items = boxes[n]
        print(f"{NAMES[n]}: {len(items):>3} поз, "
              f"{sum(r['Наличие'] for r in items):>5.0f} шт, "
              f"{kg / 1000:>5.1f} т ({kg / BOX_KG * 100:>3.0f}%), "
              f"{m3:>4.1f} куб.м ({m3 / BOX_M3 * 100:>3.0f}%), "
              f"{sum(r['Сумма'] for r in items) / 1e6:>5.1f} млн")
    if rest:
        print(f"не влезло: {len(rest)} поз, "
              f"{sum(r['вес'] for r in rest) / 1000:.1f} т, "
              f"{sum(r['Сумма'] for r in rest) / 1e6:.1f} млн")
    else:
        print("не влезло: ничего, вся арматура уезжает")


if __name__ == "__main__":
    main()
