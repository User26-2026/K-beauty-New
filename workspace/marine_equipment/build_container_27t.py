#!/usr/bin/env python3
"""План загрузки усиленного двадцатифутового контейнера на 27 тонн.

Решение по контейнеру принято: усиленный двадцатифутовый, 27 т,
315 000 руб ставка. Лимит здесь по весу: 26 350 кг груза плюс 650 кг
сырой необрезной доски — ровно 27 т брутто.

Что остается, остается не случайно: это позиции без согласованной цены.
По договору такой товар не считается проданным, и везти его нельзя, пока
цена не закрыта письменно.

Объемы у нас расчетные, по габаритам из справочной таблицы, а не замеры.
По ним груз в 33 куб. м не помещается, и это главное, что надо проверить
в первый день погрузки: если справочник завышает объем, как он завысил
массы задвижек, все войдет. Поэтому в файле есть вкладка с контролем
загрузки по ходу — сколько набрано и сколько осталось по весу и месту.

Результат: outputs/marine_equipment/Контейнер_20ф_усиленный_27т.xlsx
"""

import pathlib
import sys
from collections import defaultdict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
from build_container_spb import COLS, asrow, plural  # noqa: E402
from build_logistics_compare import (  # noqa: E402
    AUTO_SPB, AUTO_VL, BOARD_PIECE, PRICE, board, by_queue, loading, pack,
    rows_all, HOURS_BOX,
)
from build_shipment_spb import (  # noqa: E402
    FLOOR, PACKING, TIERS, pack_type, tier,
)
from build_spb_list_xlsx import GROUPS, group_of  # noqa: E402
from build_shipment_spb import item_kind  # noqa: E402
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
GOOD = PatternFill("solid", fgColor="EAF3EA")
WARN = PatternFill("solid", fgColor="FCE4D6")
GROUP_FILL = PatternFill("solid", fgColor="D9E2F3")
MARK_FILL = PatternFill("solid", fgColor="FFF2CC")
NODRAW_FILL = PatternFill("solid", fgColor="FFFF00")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

GROSS = 27_000          # паспортный вес усиленного двадцатифутового
BOX_M3 = 33.0           # внутренний объем
FLOOR_L, FLOOR_W, HEIGHT = 5.898, 2.352, 2.393
FLOOR_AREA = FLOOR_L * FLOOR_W
CRATE_KG = 600          # больше в ящик на поддоне не набираем
CRATE_L, CRATE_W, CRATE_H = 1.2, 1.0, 0.91


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


def collect27():
    """Что грузим в усиленный: лимит по весу, объем не ограничиваем."""
    pool = rows_all()
    taken, kg, m3 = pack(pool, GROSS, 10 ** 9, by_queue)
    items = []
    for r, qty in taken:
        c = dict(r)
        c["Штук"] = qty
        c["вес"] = qty * r["вес_ед"]
        c["объем"] = qty * r["объем_ед"]
        c["Сумма"] = qty * (r["Сумма"] / (r["ост"] + qty)) if r["Сумма"] else 0
        items.append(c)
    left = []
    for r in pool:
        if r["ост"] > 0:
            c = dict(r)
            c["Штук"] = r["ост"]
            c["вес"] = r["ост"] * r["вес_ед"]
            c["объем"] = r["ост"] * r["объем_ед"]
            left.append(c)
    return items, left, kg, m3


def crates(kg):
    return max(1, int(kg / CRATE_KG) + (1 if kg % CRATE_KG else 0))


def sheet_summary(wb, items, left, kg, m3):
    ws = wb.create_sheet("Сводка", 0)
    bw, bm3 = board(kg)
    ws["A1"] = "Усиленный двадцатифутовый, 27 тонн: что грузим"
    ws["A1"].font = TITLE
    ws["A2"] = ("Лимит по весу. Груз плюс доска дают ровно 27 т брутто. "
                "Очередность погрузки прежняя: согласованный список "
                "менеджера целиком, потом добор со склада по цене за "
                "килограмм.")
    ws["A2"].font = NOTE

    row = 4
    ws.cell(row=row, column=1, value="Загрузка").font = SUB
    row += 1
    tkg = defaultdict(float)
    tm3 = defaultdict(float)
    for r in items:
        t = tier(asrow(r))
        tkg[t] += r["вес"]
        tm3[t] += r["объем"]
    for label, value, note in [
        ("Позиций", len(items), ""),
        ("Штук", sum(r["Штук"] for r in items), ""),
        ("Груз, кг", round(kg), ""),
        ("Доска сырая необрезная 25 мм, кг", round(bw),
         f"{bm3:.2f} куб. м, {bm3 / BOARD_PIECE:.0f} досок 4 м х 200 х 25"),
        ("БРУТТО, кг", round(kg + bw), "паспортный предел 27 000"),
        ("Объем груза, куб. м", round(m3, 1),
         f"плюс доска {bm3:.2f}, внутри контейнера {BOX_M3:.0f}"),
        ("Ящиков с мелочью", crates(tkg[3]),
         f"по {CRATE_KG} кг, {CRATE_L} x {CRATE_W} x {CRATE_H} м"),
        ("Остается во Владивостоке, кг", round(sum(r["вес"] for r in left)),
         f"{len(left)} "
         f"{plural(len(left), 'позиция', 'позиции', 'позиций')} — все без "
         "согласованной цены"),
    ]:
        ws.cell(row=row, column=1, value=label).font = BASE
        cell = ws.cell(row=row, column=2, value=value)
        cell.font = BOLD
        cell.border = BORDER
        cell.number_format = "#,##0.00" if isinstance(value, float) else "#,##0"
        if label == "БРУТТО, кг":
            cell.fill = GOOD
        ws.cell(row=row, column=3, value=note).font = NOTE
        row += 1

    row += 1
    ws.cell(row=row, column=1, value="Расходы по отгрузке").font = SUB
    row += 1
    total = (PRICE["20 усиленный"] + AUTO_VL[20] + AUTO_SPB[20]
             + loading(HOURS_BOX))
    for label, value, note in [
        ("Контейнер усиленный 20 футов", PRICE["20 усиленный"], "ставка"),
        ("Авто во Владивостоке", AUTO_VL[20], "подача под погрузку"),
        ("Авто в Петербурге", AUTO_SPB[20], "вывоз с терминала"),
        ("Погрузка: погрузчик и грузчики", loading(HOURS_BOX),
         "подача 12 000 + 4 ч x 2 000 + 4 грузчика x 4 ч x 750"),
        ("ИТОГО", total, ""),
        ("Доставка на килограмм груза", round(total / kg, 1),
         "руб за кг, против 81 у транспортной компании"),
    ]:
        ws.cell(row=row, column=1, value=label).font = BASE
        cell = ws.cell(row=row, column=2, value=value)
        cell.font = BOLD
        cell.border = BORDER
        cell.number_format = "#,##0.0" if label.startswith("Доставка") else "#,##0"
        if label == "ИТОГО":
            cell.fill = GOOD
        ws.cell(row=row, column=3, value=note).font = NOTE
        row += 1

    row += 1
    ws.cell(row=row, column=1, value="Ярусы").font = SUB
    row += 1
    for i, h in enumerate(["Ярус", "Что лежит", "Вес, кг", "Объем, куб. м",
                           "Доля веса"], start=1):
        ws.cell(row=row, column=i, value=h)
    head(ws, row, 5, height=24)
    for t in sorted(TIERS):
        row += 1
        for i, v in enumerate([TIERS[t][0], TIERS[t][1], round(tkg[t]),
                               round(tm3[t], 1), tkg[t] / kg], start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=(i == 2))
            if i == 3:
                cell.number_format = "#,##0"
            if i == 4:
                cell.number_format = "#,##0.0"
            if i == 5:
                cell.number_format = "0%"

    row += 2
    ws.cell(row=row, column=1, value="Что проверяем в день погрузки").font = SUB
    row += 1
    for text in [
        f"Объем. По нашему справочнику габаритов груз занимает {m3:.1f} "
        f"куб. м плюс {bm3:.2f} на доску, а внутри контейнера "
        f"{BOX_M3:.0f}. Справочник уже один раз завысил: массы задвижек "
        "были выше реальных почти вдвое. Если и объемы завышены, все "
        "войдет. Проверяем на первых ярусах и сразу считаем остаток.",
        f"Вес. Паспортные 27 т — это брутто с доской. Доска съедает "
        f"{bw:.0f} кг, то есть {bw:.0f} кг неувезенной арматуры. "
        "Характерные места взвешиваем: перегруз ловят на весах в порту.",
        "Высота пирога. В этом варианте грузим под крышу, запаса нет. "
        "Задвижки Ду200-350 кладем одним слоем на настил и ничем "
        "тяжелым не накрываем — только доска и ящики.",
        "Позиции без цены не грузим. По договору товар без согласованной "
        "цены не считается проданным. Если цену успеваем закрыть "
        "письменно до погрузки — грузим, и тогда пересчитаем лимит.",
        "Состояние каждой позиции (новое или б/у) пишем при погрузке в "
        "Отгрузка_образец.xlsx, иначе в Питере не разнести расходы на "
        "единицу.",
    ]:
        cell = ws.cell(row=row, column=1, value="— " + text)
        cell.font = BASE
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[row].height = 32
        row += 1

    widths(ws, [40, 16, 60, 16, 14])
    return ws


def sheet_items(wb, name, title, items, note):
    ws = wb.create_sheet(name)
    ws["A1"] = title
    ws["A1"].font = TITLE
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
    k = 0
    for group in order:
        rows_in = sorted(by_group[group],
                         key=lambda r: (r["Ду"] or 9999, r["Наименование"]))
        row += 1
        qty = sum(r["Штук"] for r in rows_in)
        kg_group = sum(r["вес"] for r in rows_in)
        kg_text = f"{kg_group:,.0f}".replace(",", " ")
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
            k += 1
            fake = asrow(r)
            values = [k, r["Наименование"], r["Чертеж"], r["Ду"], r["Ру"],
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
            if r["Выделено"]:
                ws.cell(row=row, column=11).fill = MARK_FILL
            if not r["Чертеж"]:
                ws.cell(row=row, column=3).fill = NODRAW_FILL

    row += 1
    ws.cell(row=row, column=2, value="ИТОГО").font = BOLD
    for col, key, fmt in ((7, "Штук", "#,##0"), (9, "вес", "#,##0"),
                          (10, "объем", "#,##0.0")):
        cell = ws.cell(row=row, column=col, value=sum(r[key] for r in items))
        cell.font = BOLD
        cell.border = BORDER
        cell.number_format = fmt

    ws.freeze_panes = f"C{start + 1}"
    widths(ws, [w for _, w in COLS])
    return ws


def sheet_stowage(wb, items, kg):
    ws = wb.create_sheet("Укладка и контроль")
    bw, bm3 = board(kg)
    ws["A1"] = "Укладка: снизу вверх, с контролем по ходу погрузки"
    ws["A1"].font = TITLE
    ws["A2"] = ("Контейнер грузим под крышу, запаса по высоте нет. Поэтому "
                "после каждого яруса сверяем, сколько набрано по весу и "
                "сколько осталось места.")
    ws["A2"].font = NOTE

    tkg = defaultdict(float)
    tm3 = defaultdict(float)
    for r in items:
        t = tier(asrow(r))
        tkg[t] += r["вес"]
        tm3[t] += r["объем"]

    row = 4
    cols = ["Шаг", "Что кладем", "Чем перекладываем сверху", "Вес, кг",
            "Объем, куб. м", "Высота слоя по полу, м",
            "Набрано нарастающим, кг", "Осталось до 27 т, кг"]
    for i, h in enumerate(cols, start=1):
        ws.cell(row=row, column=i, value=h)
    head(ws, row, len(cols), height=40)

    row += 1
    for i, v in enumerate([FLOOR[0], FLOOR[1], FLOOR[2], round(bw),
                           round(bm3, 2), 0.03, None, None], start=1):
        cell = ws.cell(row=row, column=i, value=v)
        cell.font = BASE
        cell.border = BORDER
        cell.fill = GROUP_FILL
        cell.alignment = Alignment(vertical="top", wrap_text=(i in (2, 3)))
        if i == 4:
            cell.number_format = "#,##0"

    run = bw
    for t in sorted(TIERS):
        row += 1
        run += tkg[t]
        label, what, brd = TIERS[t]
        values = [label, what, brd, round(tkg[t]), round(tm3[t], 1),
                  round(tm3[t] / FLOOR_AREA, 2), round(run),
                  round(GROSS - run)]
        for i, v in enumerate(values, start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=(i in (2, 3)))
            if i in (4, 7, 8):
                cell.number_format = "#,##0"
            if i in (5, 6):
                cell.number_format = "#,##0.00"

    # раскладка клинкетов по рядам и места под ящики
    counts = defaultdict(int)
    for r in items:
        if (tier(asrow(r)) == 1
                and item_kind(asrow(r)) == "клинкет (задвижка)"):
            counts[int(r["Ду"] or 0)] += int(r["Штук"])
    grows = gg.rows_of(counts)
    layers = gg.place(grows)
    gate_len = gg.gate_length(counts)
    free, over, h_gates, free_l = gg.crate_places(counts)
    gate_area = min(gate_len, FLOOR_L) * FLOOR_W
    places = over + free
    nbox = crates(tkg[3])

    row += 2
    ws.cell(row=row, column=1, value="Клинкеты и места под ящики").font = SUB
    row += 1
    for label, value, note in [
        ("Рядов клинкетов", len(grows),
         "задвижка лежит на боку, шток горизонтально поперек"),
        ("Длина пола под клинкетами, м", round(gate_len, 2),
         f"пол контейнера {FLOOR_L} м"),
        ("ЯРУСОВ КЛИНКЕТОВ", layers,
         "больше одного — значит задвижка ляжет на задвижку"),
        ("Высота под клинкеты с досками, м", round(h_gates, 2),
         f"до потолка остается {HEIGHT - h_gates:.2f} м"),
        ("Свободного пола, м", round(free_l, 2),
         f"ряд ящиков просит {CRATE_L} м"),
        ("Мест над клинкетами", over, "сколько ярусов ящиков встает сверху"),
        ("Мест на свободном полу", free, "два яруса ящиков"),
        ("ВСЕГО мест", places, ""),
        ("Нужно под мелочь", nbox, f"ящиков по {CRATE_KG} кг"),
        ("Не хватает мест", max(0, nbox - places),
         "если не ноль — столько ящиков в корпус не встает, это и есть "
         "нехватка объема"),
        ("Ящиков встанет на клинкеты", max(0, nbox - free),
         "все, чему не хватило свободного пола"),
        ("Вес поверх клинкетов, кг",
         round(max(0, nbox - free) * (tkg[3] / nbox) + tkg[4]),
         "ящики плюс фильтры и титан на них"),
        ("Нагрузка на клинкеты, кг/кв. м",
         round((max(0, nbox - free) * (tkg[3] / nbox) + tkg[4]) / gate_area),
         "в двух двадцатифутовых выходит вдвое меньше, и там клинкеты "
         "лежат в один ярус"),
    ]:
        ws.cell(row=row, column=1, value=label).font = BASE
        cell = ws.cell(row=row, column=2, value=value)
        cell.font = BOLD
        cell.border = BORDER
        cell.number_format = "#,##0.0" if isinstance(value, float) else "#,##0"
        if label == "Не хватает мест":
            cell.fill = WARN if nbox > places else GOOD
        ws.cell(row=row, column=3, value=note).font = NOTE
        row += 1

    row += 1
    ws.cell(row=row, column=1, value="Чтобы не подавить арматуру").font = SUB
    row += 1
    floor_kg = tkg[1] + tkg[2]
    for text in [
        f"На полу лежит {floor_kg:,.0f} кг — это "
        f"{floor_kg / FLOOR_AREA:,.0f} кг на квадратный метр. Задвижки "
        "Ду200-350 кладем одним слоем, плашмя, фланцем вниз и никогда "
        "фланцем на фланец.".replace(",", " "),
        f"Ящиков с мелочью {nbox}, по {CRATE_KG} кг. Они стоят на полу "
        "рядом с задвижками и в два уровня, а не поверх задвижек.",
        f"На ящики идет только верхний ярус — {tkg[4]:,.0f} кг фильтров и "
        f"титана, это {tkg[4] / nbox:,.0f} кг на ящик при пределе "
        f"{CRATE_KG}.".replace(",", " "),
        "Между рядами задвижек Ду250-350 шаг прокладок 300-400 мм, а не "
        "600: длинный корпус на двух опорах прогибается и дает трещину "
        "по приливу.",
        "Маховики снимаем, где снимаются, и кладем в ящик с биркой по "
        "номеру места. Нагрузку они не несут.",
        "Пустоты забиваем распорками из той же доски: в полном "
        "контейнере пустот мало, но продольные удары при маневрировании "
        "никуда не денутся.",
        "Влагопоглотитель под крышу: месяц дороги дает конденсат.",
        "Погрузчиком берем только поддоны и обрешетку. Голую задвижку "
        "вилами не поднимаем.",
    ]:
        cell = ws.cell(row=row, column=1, value="— " + text)
        cell.font = BASE
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[row].height = 30
        row += 1

    widths(ws, [26, 44, 40, 12, 14, 16, 18, 16])
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
        ws.row_dimensions[row].height = 58
    widths(ws, [5, 42, 86])
    return ws


def main():
    items, left, kg, m3 = collect27()
    bw, bm3 = board(kg)

    wb = Workbook()
    wb.remove(wb.active)
    sheet_summary(wb, items, left, kg, m3)
    sheet_items(wb, "Контейнер по группам",
                "Усиленный двадцатифутовый: перечень на погрузку",
                items,
                "Порядок — по группам, от мелкой арматуры к крупной. "
                "Желтым помечены позиции без чертежа, их надо дописать.")
    sheet_stowage(wb, items, kg)
    sheet_items(wb, "Остается", "Что не грузим", left,
                "Это позиции без согласованной цены. По договору такой "
                "товар не считается проданным: грузить его нельзя, пока "
                "цена не закрыта письменно.")
    sheet_packing(wb)

    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / "Контейнер_20ф_усиленный_27т.xlsx"
    wb.save(path)

    print(f"готово: {path}\n")
    print(f"груз {kg:,.0f} кг + доска {bw:,.0f} кг = {kg + bw:,.0f} кг брутто"
          .replace(",", " "))
    print(f"объем {m3:.1f} + {bm3:.2f} куб. м при {BOX_M3:.0f} внутри")
    print(f"позиций {len(items)}, штук {sum(r['Штук'] for r in items):.0f}")
    tkg = defaultdict(float)
    for r in items:
        tkg[tier(asrow(r))] += r["вес"]
    for t in sorted(TIERS):
        print(f"  {TIERS[t][0]:<22}{tkg[t]:>9,.0f} кг".replace(",", " "))
    print(f"ящиков с мелочью {crates(tkg[3])}")
    print(f"остается {len(left)} поз, {sum(r['вес'] for r in left):,.0f} кг"
          .replace(",", " "))


if __name__ == "__main__":
    main()
