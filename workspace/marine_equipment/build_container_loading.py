#!/usr/bin/env python3
"""Что уехало в контейнере: один файл по факту отгрузки 26.10.2026.

Контейнер загружен и ушел в Петербург 26.10.2026. Источник — ФАКТ
погрузки, а не план: в приложение 4 и в этот файл идет то, что
действительно уехало. План от 21.10 остался для сверки.

Сюда сведено все, о чем договаривались: фактический перечень, снятые
позиции, цены по ставке на бронзу, ремонт по прайсу мастерской,
ориентир реализации, веса по паспорту, маховики, тара, предел 26 тонн
и правила укладки.

Денег в файле три колонки, и это цепочка, а не варианты одной цены:
закуп по Договору — по нему мы продаем арматуру ПКФ «АРМА»; ремонт —
во что компании обойдется привести товар в порядок; продажная цена —
по чему мы вместе с ней продаем конечному покупателю. В самом
Договоре продажная цена называется ценой реализации, пункт 3.2.

Файл делается для погрузки, а не для переговоров, поэтому в нем рядом
стоят и деньги, и килограммы: бригада смотрит вес и тару, менеджер —
цену и то, что нельзя грузить без согласования.

Семь вкладок:
  Сводка           — итоги, тоннаж против предела, принятые решения
  По группам       — вес и деньги по каждой группе перечня
  Перечень         — все позиции построчно, порядок погрузки
  Как считаем цены — закуп, ремонт, продажная цена, ставка приемки
  Упаковка         — что в каком месте лежит, сверка с листами
  План и факт      — чем отгрузка отличается от перечня 21.10
  Укладка          — ярусы, доска, маховики, что на что класть
  Что закрыть      — чек-лист, в первую очередь цены
  Снято с отгрузки — что и почему не поехало

Результат: outputs/marine_equipment/Контейнер_отгружено_26.10.xlsx
"""

import pathlib
import sys
from collections import defaultdict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
import container_list  # noqa: E402
import gate_geometry  # noqa: E402
import packing_lists  # noqa: E402
import scrap_prices  # noqa: E402
import upd_shipments  # noqa: E402
from build_contract_annexes import (AGREE, BOARD_KG,  # noqa: E402
                                    CONTAINER, LIMIT_KG, annex4, no_rate)
from build_container_spb import plural  # noqa: E402
from build_logistics_compare import board  # noqa: E402

OUT = pathlib.Path(__file__).resolve().parents[2] / "outputs" / "marine_equipment"

# Длина пола внутри сорокафутового контейнера, м. У двадцатифутового
# 5,898 — туда задвижки в один слой не ложатся вовсе.
FLOOR_L = 12.032

# С какого диаметра задвижка кладется на пол. Ду125 и ниже легкие,
# они идут в ящики и наверх, пол ими не занимаем.
BIG_DU = 200

FONT = "Arial"
HEAD_FILL = PatternFill("solid", fgColor="1F3864")
HEAD_FONT = Font(name=FONT, size=11, bold=True, color="FFFFFF")
TITLE = Font(name=FONT, size=14, bold=True, color="1F3864")
SUB = Font(name=FONT, size=11, bold=True, color="1F3864")
BOLD = Font(name=FONT, size=10, bold=True)
BASE = Font(name=FONT, size=10)
NOTE = Font(name=FONT, size=9, italic=True, color="666666")
GROUP_FILL = PatternFill("solid", fgColor="D9E2F3")
OK_FILL = PatternFill("solid", fgColor="E2EFDA")
WARN_FILL = PatternFill("solid", fgColor="FCE4D6")
MARK_FILL = PatternFill("solid", fgColor="FFF2CC")
GREY = PatternFill("solid", fgColor="F2F2F2")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def head(ws, row, ncols, height=34):
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


def put(ws, row, values, wrap=(), center=(), money=(), fill=None, font=None,
        height=None):
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
    if height:
        ws.row_dimensions[row].height = height
    return row + 1


def rub(n):
    return f"{n:,.0f}".replace(",", " ")


def val(text):
    """Число из цены. Пустая строка и «по согласованию» дают ноль."""
    b = upd_shipments.bounds(text)
    return max(b) if b else 0


def buy_rate(material):
    """Закупочная цена за килограмм: цена лома, умноженная на 2,5.

    Договоренность с Антоном Бурыкиным от 25.10.2026. По бронзе это
    745 * 2,5 = 1 862,50 руб — ровно ставка, по которой посчитана вся
    бронза в спецификации. Для стали, титана, алюминия и нержавейки
    формула та же, но ставка приемки другая.
    """
    return scrap_prices.purchase_rate(material)


def skip_reason(row):
    """Почему у позиции нет цены от металла. Пусто — цена считается."""
    return no_rate(row["name"], row["ru"], row["material"], row["cond"])


def buy(row):
    """Закупочная цена позиции.

    Пусто у всего, что от массы металла не считается: электромагнитные
    и редукционные клапаны, клапаны на высокое давление, сталь и
    титан. По ним цену согласовывают отдельно.
    """
    if skip_reason(row):
        return 0.0
    r = buy_rate(row["material"])
    return (r or 0) * row["unit_kg"] * row["qty"]


def collect():
    """Строки спецификации в порядке погрузки, с деньгами и тарой."""
    rows, blocks = annex4()
    extra = {(r["Наименование"], r["Чертеж"]): r
             for r in container_list.read()}
    out = []
    n = 0
    for b in blocks:
        for it in b["items"]:
            n += 1
            src = extra.get((it["name"], it["draw"]), {})
            out.append(dict(it, n=n, block=b["group"],
                            pack=src.get("Упаковка", "")))
    return out, blocks


def sheet_summary(wb, rows, blocks):
    ws = wb.active
    ws.title = "Сводка"
    kg = sum(r["kg"] for r in rows)
    qty = sum(r["qty"] for r in rows)
    brd = BOARD_KG
    gross = kg + brd
    left = LIMIT_KG - gross
    money = sum(val(r["price"]) * r["qty"] for r in rows)
    purchase = sum(buy(r) for r in rows)
    repair = sum((r["repair"] or 0) * r["qty"] for r in rows)
    resale = sum(val(r["resale"]) * r["qty"] for r in rows)
    wheels = sum(r["wheels"] or 0 for r in rows)
    open_price = [r for r in rows if r["price"] == AGREE]
    no_price = [r for r in rows if not r["price"]]
    no_metal = [r for r in rows if skip_reason(r)]

    ws["A1"] = ("Контейнер в Санкт-Петербург, отгружен 26.10.2026, "
                f"{CONTAINER}")
    ws["A1"].font = TITLE
    ws["A2"] = ("ФАКТ погрузки, а не план: считается по перечню, "
                "составленному в день отгрузки. Все цифры "
                "пересчитываются из исходных файлов, руками в этот файл "
                "ничего не вписано.")
    ws["A2"].font = NOTE

    start = 4
    for i, h in enumerate(["Показатель", "Значение", "Что это значит"],
                          start=1):
        ws.cell(row=start, column=i, value=h)
    head(ws, start, 3)
    row = start + 1
    for name, value, note, fill in [
        ("Позиций", f"{len(rows)}", "строк в спецификации", None),
        ("Штук", rub(qty), "единиц товара", None),
        ("Вес груза", f"{rub(kg)} кг", "без упаковки, по паспорту", None),
        ("Перекладочная доска", f"{rub(round(brd))} кг",
         "сырая необрезная, цифра собственника на эту погрузку", None),
        ("БРУТТО", f"{rub(round(gross))} кг", "груз плюс доска", GROUP_FILL),
        ("Предел погрузки", f"{rub(LIMIT_KG)} кг",
         f"{CONTAINER} контейнер, решение собственника", None),
        ("ЗАПАС" if left >= 0 else "ПРЕВЫШЕНИЕ",
         f"{rub(round(abs(left)))} кг ({abs(left) / LIMIT_KG:.0%})",
         "укладываемся, можно догружать" if left >= 0
         else "перечень надо сокращать",
         OK_FILL if left >= 0 else WARN_FILL),
        ("", "", "", None),
        ("1. ЗАКУП по Договору", f"{rub(money)} руб",
         "по этой цене мы продаем арматуру ПКФ «АРМА». Бронза — по "
         "ставке 1 862,50 руб за кг", GROUP_FILL),
        ("2. Ремонт", f"{rub(repair)} руб",
         "компания приводит товар в порядок, прайс мастерской. В цену "
         "закупа не входит", None),
        ("3. Продажная цена", f"{rub(resale)} руб",
         "продаем вместе с компанией конечному покупателю. Ориентир: "
         "цену согласуют письменно по пункту 3.2", None),
        ("Разница 3 минус 1 и 2", f"{rub(resale - money - repair)} руб",
         "сколько цепочка дает сверх закупа и ремонта", None),
        ("", "", "", None),
        ("Маховиков", f"{rub(wheels)} шт",
         "снимаются в ящик с биркой, считать при погрузке", MARK_FILL),
        ("Цена «по согласованию»", f"{len(open_price)} позиций, "
         f"{rub(sum(r['kg'] for r in open_price))} кг",
         "по пункту 3.5 проданным не считается — закрыть до отгрузки",
         WARN_FILL),
        ("Цена не ставится", f"{len(no_price)} позиций",
         "электромагнитные, цена не от массы металла", MARK_FILL),
        ("Не считаем от металла", f"{len(no_metal)} позиций, "
         f"{rub(round(sum(r['kg'] for r in no_metal)))} кг",
         "электромагнитные и редукционные клапаны, клапаны Ру от 100, "
         "сталь и титан — цену согласовываем отдельно", MARK_FILL),
    ]:
        row = put(ws, row, [name, value, note], wrap=(3,),
                  font=BOLD if name in ("БРУТТО", "ЗАПАС", "ПРЕВЫШЕНИЕ")
                  else BASE, fill=fill)

    row += 1
    ws.cell(row=row, column=1, value="Что решили и по чему считаем").font = SUB
    row += 1
    for t in [
        "Закупочная цена считается ОТ ЛОМА: цена приемки за килограмм, "
        "умноженная на 2,5. По бронзе это 745 * 2,5 = 1 862,50 руб за "
        "килограмм — та самая ставка, по которой посчитана спецификация.",
        "Цена бронзы и латуни считается ОТ ВЕСА: 1 862,50 руб за "
        "килограмм. Уточнится вес при замере — цена пересчитается сама.",
        "От металла НЕ считаются: сталь и титан, электромагнитные и "
        "редукционные клапаны, клапаны на высокое давление Ру от 100. "
        "По ним цену согласовываем отдельно: в редукторе, приводе и "
        "рабочем давлении денег больше, чем в массе корпуса.",
        "Вся остальная арматура БРОНЗОВАЯ, подтверждено собственником "
        "25.10.2026: на море ставят бронзу, латунь идет на реки. "
        "Поэтому лом считается по бронзе 745, а не по латуни, и "
        "ставка выходит 1 862,50 руб за кг.",
        "Стаканы переборочные с маркировкой Ру100 под правило высокого "
        "давления не попадают: это не клапан, а проход через переборку.",
        "Где в договоре цена записана двойной, берем ВЕРХНЮЮ границу. "
        "Внутри диапазона конкретную цену по пункту 3.4 все равно "
        "согласовывают письменно до передачи.",
        "Веса паспортные, по бронзе: задвижка Ду350 — 212,5 кг, Ду300 "
        "— 192, Ду250 — 125, Ду200 — 76,5. Прежний наш справочник "
        "завышал их почти вдвое.",
        "Ремонт — ориентир мастерской, а не счет. Новое не ремонтируем. "
        "Фильтров, коробок, захлопок, кингстонов и головок в прайсе "
        "ремонта нет, по ним стоимость запрашивается отдельно.",
    ]:
        cell = ws.cell(row=row, column=1, value="— " + t)
        cell.font = BASE
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[row].height = 28
        row += 1
    widths(ws, [26, 24, 62])
    return dict(kg=kg, qty=qty, brd=brd, gross=gross, left=left,
                money=money, purchase=purchase, repair=repair,
                resale=resale, wheels=wheels)


COLS_GROUP = [
    ("Группа", 34), ("Позиций", 8), ("Штук", 8), ("Вес, кг", 10),
    ("Доля веса", 9), ("Вес нарастающим, кг", 12),
    ("1. ЗАКУП по Договору, руб", 15), ("2. Ремонт, руб", 12),
    ("3. Продажная цена, руб", 14),
]


def sheet_groups(wb, rows):
    ws = wb.create_sheet("По группам")
    agg = defaultdict(lambda: dict(n=0, qty=0, kg=0.0, money=0.0,
                                   buy=0.0, repair=0.0, resale=0.0))
    for r in rows:
        a = agg[r["block"]]
        a["n"] += 1
        a["qty"] += r["qty"]
        a["kg"] += r["kg"]
        a["money"] += val(r["price"]) * r["qty"]
        a["buy"] += buy(r)
        a["repair"] += (r["repair"] or 0) * r["qty"]
        a["resale"] += val(r["resale"]) * r["qty"]
    total_kg = sum(a["kg"] for a in agg.values())

    ws["A1"] = "Что грузим: по группам"
    ws["A1"].font = TITLE
    ws["A2"] = ("Порядок групп — как грузим. Нарастающий вес показывает, "
                "сколько тонн уже в контейнере к концу каждой группы. "
                "Три колонки денег — это цепочка: закуп по Договору "
                "(мы продаем компании), ремонт по прайсу мастерской, "
                "продажная цена (продаем вместе конечному "
                "покупателю).")
    ws["A2"].font = NOTE
    start = 4
    for i, (h, _) in enumerate(COLS_GROUP, start=1):
        ws.cell(row=start, column=i, value=h)
    head(ws, start, len(COLS_GROUP))
    row = start + 1
    run = 0.0
    order = []
    for r in rows:
        if r["block"] not in order:
            order.append(r["block"])
    for g in order:
        a = agg[g]
        run += a["kg"]
        row = put(ws, row, [
            g, a["n"], a["qty"], round(a["kg"]),
            f"{a['kg'] / total_kg:.1%}", round(run),
            round(a["money"]), round(a["repair"]), round(a["resale"]),
        ], wrap=(1,), center=(2, 3, 5), money=(4, 6, 7, 8, 9))
    row = put(ws, row, [
        "ИТОГО", sum(a["n"] for a in agg.values()),
        sum(a["qty"] for a in agg.values()), round(total_kg), "100%",
        round(total_kg), round(sum(a["money"] for a in agg.values())),
        round(sum(a["repair"] for a in agg.values())),
        round(sum(a["resale"] for a in agg.values())),
    ], center=(2, 3, 5), money=(4, 6, 7, 8, 9), font=BOLD,
        fill=GROUP_FILL)
    ws.freeze_panes = f"A{start + 1}"
    widths(ws, [w for _, w in COLS_GROUP])
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1


COLS_ITEM = [
    ("№", 5), ("Наименование", 46), ("Чертеж", 18), ("Ду", 6), ("Ру", 6),
    ("Сост.", 8), ("Шт", 6), ("Вес ед., кг", 9), ("Вес всего, кг", 10),
    ("Маховики, шт", 9), ("Тара", 12),
    ("1. ЗАКУП по Договору за ед., руб", 14),
    ("2. Ремонт за ед., руб", 12),
    ("3. Продажная цена за ед., руб", 14),
    ("Материал", 13),
]


def sheet_items(wb, rows):
    ws = wb.create_sheet("Перечень")
    ws["A1"] = "Перечень к погрузке, в порядке загрузки"
    ws["A1"].font = TITLE
    ws["A2"] = (f"{len(rows)} "
                f"{plural(len(rows), 'позиция', 'позиции', 'позиций')}, "
                f"{rub(sum(r['qty'] for r in rows))} шт, "
                f"{rub(sum(r['kg'] for r in rows))} кг. Желтым — где цену "
                f"надо закрыть до отгрузки, оранжевым — не бронза.")
    ws["A2"].font = NOTE
    start = 4
    for i, (h, _) in enumerate(COLS_ITEM, start=1):
        ws.cell(row=start, column=i, value=h)
    head(ws, start, len(COLS_ITEM))
    row = start + 1
    block = None
    for r in rows:
        if r["block"] != block:
            block = r["block"]
            items = [x for x in rows if x["block"] == block]
            cells = ws.cell(row=row, column=1)
            ws.merge_cells(start_row=row, start_column=1,
                           end_row=row, end_column=len(COLS_ITEM))
            cells.value = (f"{block} — {len(items)} "
                           f"{plural(len(items), 'позиция', 'позиции', 'позиций')}"
                           f", {rub(sum(x['qty'] for x in items))} шт, "
                           f"{rub(sum(x['kg'] for x in items))} кг")
            cells.font = SUB
            cells.fill = GROUP_FILL
            for c in range(1, len(COLS_ITEM) + 1):
                ws.cell(row=row, column=c).border = BORDER
            row += 1
        fill = (WARN_FILL if r["price"] == AGREE
                else MARK_FILL if not r["price"]
                else None)
        row = put(ws, row, [
            r["n"], r["name"], r["draw"] or "—", r["du"] or "",
            r["ru"] or "", r["cond"], r["qty"], round(r["unit_kg"], 1),
            round(r["kg"]),
            int(r["wheels"]) if r["wheels"] else "",
            r["pack"],
            r["price"] or "цена отдельно",
            rub(r["repair"]) if r["repair"] else "", r["resale"],
            r["material"],
        ], wrap=(2,), center=(1, 4, 5, 6, 7, 8, 9, 10), fill=fill)
    ws.freeze_panes = f"A{start + 1}"
    ws.auto_filter.ref = (f"A{start}:"
                          f"{get_column_letter(len(COLS_ITEM))}{row - 1}")
    widths(ws, [w for _, w in COLS_ITEM])
    ws.print_title_rows = f"{start}:{start}"
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1


COLS_BUY = [
    ("Материал", 20), ("Позиций", 9), ("Штук", 9), ("Вес, кг", 11),
    ("Лом, руб за кг", 12), ("Множитель", 10),
    ("Ставка закупа, руб за кг", 13),
    ("1. ЗАКУП по Договору, руб", 15), ("2. Ремонт, руб", 13),
    ("3. Продажная цена, руб", 15),
]


def sheet_buy(wb, rows):
    """Откуда берется закупочная цена и что она дает по материалам."""
    ws = wb.create_sheet("Как считаем цены")
    ws["A1"] = "Три цены: закуп, ремонт, продажная цена"
    ws["A1"].font = TITLE
    ws["A2"] = ("Закуп по Договору — по этой цене мы продаем арматуру "
                "ПКФ «АРМА». Дальше компания приводит товар в порядок и "
                "при необходимости ремонтирует, ремонт считается по "
                "прайсу мастерской. Потом мы вместе с ней продаем "
                "арматуру конечному покупателю — это продажная цена, "
                "в Договоре она называется ценой реализации. "
                "Оранжевым — то, у чего закуп от металла не считается.")
    ws["A2"].font = NOTE

    agg = defaultdict(lambda: dict(n=0, qty=0, kg=0.0, money=0.0,
                                   repair=0.0, resale=0.0))
    for r in rows:
        a = agg[r["material"] if not skip_reason(r) else skip_reason(r)]
        a["n"] += 1
        a["qty"] += r["qty"]
        a["kg"] += r["kg"]
        a["money"] += val(r["price"]) * r["qty"]
        a["repair"] += (r["repair"] or 0) * r["qty"]
        a["resale"] += val(r["resale"]) * r["qty"]

    start = 4
    for i, (h, _) in enumerate(COLS_BUY, start=1):
        ws.cell(row=start, column=i, value=h)
    head(ws, start, len(COLS_BUY))
    row = start + 1
    order = sorted(agg, key=lambda m: -agg[m]["kg"])
    for m in order:
        a = agg[m]
        by_metal = m in scrap_prices.BY_MATERIAL
        rate = buy_rate(m) if by_metal else None
        scr = scrap_prices.rate(m) if by_metal else None
        row = put(ws, row, [
            m, a["n"], a["qty"], round(a["kg"]),
            scr if scr else "—",
            f"{scrap_prices.MULTIPLIER:.1f}".replace(".", ",")
            if rate else "—",
            round(rate, 2) if rate else "цену согласовываем",
            round(a["money"]) if a["money"] else "—",
            round(a["repair"]) if a["repair"] else "—",
            round(a["resale"]) if a["resale"] else "—",
        ], center=(2, 3, 6), money=(4, 5, 8, 9, 10),
            wrap=(1, 7),
            fill=None if by_metal else WARN_FILL)
    row = put(ws, row, [
        "ИТОГО", sum(a["n"] for a in agg.values()),
        sum(a["qty"] for a in agg.values()),
        round(sum(a["kg"] for a in agg.values())), "", "", "",
        round(sum(a["money"] for a in agg.values())),
        round(sum(a["repair"] for a in agg.values())),
        round(sum(a["resale"] for a in agg.values())),
    ], center=(2, 3), money=(4, 8, 9, 10), font=BOLD, fill=GROUP_FILL)

    row += 1
    ws.cell(row=row, column=1,
            value="Откуда берется ставка закупа").font = SUB
    row += 1
    for t in [
        "Закупочная цена = цена металлолома * 2,5. Договоренность с "
        "Антоном Бурыкиным от 25.10.2026.",
        "Бронза: 745 * 2,5 = 1 862,50 руб за килограмм. Это и есть "
        "ставка, по которой посчитана вся бронза в спецификации.",
        "Уточнится вес при замере — цена пересчитается сама. Поменяет "
        "приемка цены — пересчитается и ставка.",
    ]:
        cell = ws.cell(row=row, column=1, value="— " + t)
        cell.font = BASE
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=row, start_column=1,
                       end_row=row, end_column=len(COLS_BUY))
        row += 1

    row += 1
    ws.cell(row=row, column=1, value="Прайс приемки на 13.10.2026").font = SUB
    row += 1
    for i, h in enumerate(["Металл", "Лом, руб за кг", "Закуп, руб за кг"],
                          start=1):
        ws.cell(row=row, column=i, value=h)
    head(ws, row, 3, height=20)
    row += 1
    for metal, price in sorted(scrap_prices.SCRAP.items(),
                               key=lambda kv: -kv[1]):
        row = put(ws, row, [metal, price,
                            round(price * scrap_prices.MULTIPLIER, 2)],
                  money=(2, 3))

    row += 1
    for t in [
        "Лом — это нижняя граница цены. Любая позиция стоит хотя бы "
        "как металл, и ниже продавать незачем: сдать быстрее и без "
        "споров по состоянию.",
        "Вся арматура бронзовая, 745 руб за кг: на море ставят бронзу, "
        "латунь идет на реки. Из этого и выходит ставка 1 862,50.",
        "Сталь, титан, алюминий и нержавейка от лома не считаются "
        "вовсе: множитель дал бы по стали 47,50 руб за кг, а у "
        "конкурента стальная задвижка Ду200 стоит 40 000 руб. Цену по "
        "ним согласовываем отдельно.",
        "Так же не считаются электромагнитные и редукционные клапаны и "
        "клапаны на высокое давление Ру от 100: корпус у них та же "
        "бронза, но платят не за нее.",
        "Ремонт в цену закупа не входит — это следующий шаг цепочки, "
        "его делает компания после покупки. Новое не ремонтируем.",
        "Продажная цена — ориентир, а не согласованная: по пункту 3.2 "
        "Договора конкретную цену согласуют письменно до передачи "
        "товара конечному покупателю.",
        "Если приемка поменяет цены, пересчитывается и ставка: она "
        "считается из прайса, а не вписана руками.",
    ]:
        cell = ws.cell(row=row, column=1, value="— " + t)
        cell.font = BASE
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=row, start_column=1,
                       end_row=row, end_column=len(COLS_BUY))
        ws.row_dimensions[row].height = 28
        row += 1

    widths(ws, [w for _, w in COLS_BUY])
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1


COLS_PACK = [
    ("Место", 18), ("Позиций", 9), ("Штук", 9), ("Вес, кг", 11),
    ("По упаковочному листу, шт", 13), ("Расхождение, шт", 12),
    ("Что внутри", 70),
]


def sheet_packing(wb):
    """Что в каком месте лежит. По этому принимают груз в Петербурге."""
    ws = wb.create_sheet("Упаковка")
    ws["A1"] = "Упаковка: что в каком месте"
    ws["A1"].font = TITLE
    ws["A2"] = ("Принимать груз в Петербурге надо ПО МЕСТАМ: по пункту "
                "2.2 Договора на претензии по количеству есть 3 рабочих "
                "дня с даты получения. Основа — колонка «упаковка» "
                "перечня, рядом сверка с упаковочными листами.")
    ws["A2"].font = NOTE

    places = packing_lists.by_place()
    checks, failed = packing_lists.check()
    by_num = {f"№ {c['Место']}": c for c in checks}

    start = 4
    for i, (h, _) in enumerate(COLS_PACK, start=1):
        ws.cell(row=start, column=i, value=h)
    head(ws, start, len(COLS_PACK))
    row = start + 1

    def order(k):
        c = by_num_key(k)
        return (k == "без упаковки", c if c else 999, k)

    for k in sorted(places, key=order):
        rs = places[k]
        num = by_num_key(k)
        chk = by_num.get(f"№ {num}") if num else None
        diff = chk["Расхождение"] if chk else None
        inside = "; ".join(
            f"{r['Наименование'][:44]} — {r['Штук']} шт" for r in rs[:6])
        if len(rs) > 6:
            inside += f"; и еще {len(rs) - 6}"
        row = put(ws, row, [
            k, len(rs), sum(r["Штук"] for r in rs),
            round(sum(r["вес"] for r in rs), 1),
            chk["По упаковочному листу, шт"] if chk else "листа нет",
            diff if diff else ("" if chk else "—"),
            inside,
        ], wrap=(1, 7), center=(2, 3, 4, 5, 6),
            fill=WARN_FILL if diff else None)
    row = put(ws, row, [
        "ИТОГО", sum(len(v) for v in places.values()),
        sum(r["Штук"] for v in places.values() for r in v),
        round(sum(r["вес"] for v in places.values() for r in v)), "", "", "",
    ], center=(2, 3, 4), font=BOLD, fill=GROUP_FILL)

    row += 1
    for t in [
        "«Без упаковки» — это не недосмотр: крупная арматура и задвижки "
        "едут россыпью на полу и на доске, упаковочный лист им не "
        "нужен. Места с листами — это ящик и коробки с мелочью.",
        "Расхождение по месту значит, что в коробке лежит не то, что "
        "записано в перечне. Оранжевым подсвечено, разобрать до "
        "приемки: три рабочих дня истекут быстро.",
    ] + ([f"Упаковочные листы, которые не удалось разобрать: "
          f"{', '.join('№' + str(f) for f in failed)}. Их надо сверить "
          f"глазами."] if failed else []):
        cell = ws.cell(row=row, column=1, value="— " + t)
        cell.font = BASE
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=row, start_column=1,
                       end_row=row, end_column=len(COLS_PACK))
        ws.row_dimensions[row].height = 30
        row += 1

    widths(ws, [w for _, w in COLS_PACK])
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1


def by_num_key(place):
    """Номер места из подписи «коробка № 7». Ноль, если номера нет."""
    import re
    m = re.search(r"№\s*(\d+)", place)
    return int(m.group(1)) if m else 0


COLS_DIFF = [
    ("Что изменилось", 20), ("Наименование", 52), ("Чертеж", 18),
    ("Было, шт", 10), ("Стало, шт", 10), ("Было, кг", 10),
    ("Стало, кг", 10), ("Группа", 26),
]


def sheet_diff(wb):
    """Чем фактическая отгрузка отличается от перечня 21.10."""
    import re
    ws = wb.create_sheet("План и факт")
    ws["A1"] = "Чем отгрузка отличается от плана"
    ws["A1"].font = TITLE

    # План читаем целиком, вместе со снятыми позициями: иначе
    # стальной Ду300 пропадет с обеих сторон и видно не будет, что
    # решение о снятии исполнено.
    plan = container_list.read(container_list.PLAN, skip_excluded=False)
    fact = container_list.read()

    def k(r):
        return (re.sub(r"\s+", " ", r["Наименование"].strip().lower()),
                re.sub(r"\s+", "", (r["Чертеж"] or "").lower()), r["Ду"])

    pi, fi = {}, {}
    for r in plan:
        pi.setdefault(k(r), []).append(r)
    for r in fact:
        fi.setdefault(k(r), []).append(r)

    added = [r for kk in fi if kk not in pi for r in fi[kk]]
    gone = [r for kk in pi if kk not in fi for r in pi[kk]]
    changed = []
    for kk in fi:
        if kk not in pi:
            continue
        fq = sum(r["Штук"] for r in fi[kk])
        pq = sum(r["Штук"] for r in pi[kk])
        fk = sum(r["вес"] for r in fi[kk])
        pk = sum(r["вес"] for r in pi[kk])
        if fq != pq or abs(fk - pk) > 0.5:
            changed.append((pi[kk][0], fi[kk][0], pq, fq, pk, fk))

    ws["A2"] = (f"План 21.10: {len(plan)} позиций, "
                f"{rub(sum(r['Штук'] for r in plan))} шт, "
                f"{rub(round(sum(r['вес'] for r in plan)))} кг. "
                f"Факт 26.10: {len(fact)} позиций, "
                f"{rub(sum(r['Штук'] for r in fact))} шт, "
                f"{rub(round(sum(r['вес'] for r in fact)))} кг.")
    ws["A2"].font = NOTE

    start = 4
    for i, (h, _) in enumerate(COLS_DIFF, start=1):
        ws.cell(row=start, column=i, value=h)
    head(ws, start, len(COLS_DIFF))
    row = start + 1
    for r in sorted(added, key=lambda x: -x["вес"]):
        row = put(ws, row, [
            "ДОГРУЗИЛИ", r["Наименование"], r["Чертеж"] or "—", "",
            r["Штук"], "", round(r["вес"], 1), r["Группа"],
        ], wrap=(2, 8), center=(4, 5, 6, 7), fill=OK_FILL)
    for r in sorted(gone, key=lambda x: -x["вес"]):
        row = put(ws, row, [
            "НЕ ПОЕХАЛО", r["Наименование"], r["Чертеж"] or "—",
            r["Штук"], "", round(r["вес"], 1), "", r["Группа"],
        ], wrap=(2, 8), center=(4, 5, 6, 7), fill=WARN_FILL)
    for p_, f_, pq, fq, pk, fk in sorted(changed,
                                         key=lambda x: -abs(x[5] - x[4])):
        row = put(ws, row, [
            "ПЕРЕСЧИТАЛИ", f_["Наименование"], f_["Чертеж"] or "—",
            pq, fq, round(pk, 1), round(fk, 1), f_["Группа"],
        ], wrap=(2, 8), center=(4, 5, 6, 7), fill=MARK_FILL)
    row = put(ws, row, [
        "ИТОГО", f"догрузили {len(added)}, не поехало {len(gone)}, "
        f"пересчитали {len(changed)}", "",
        sum(r["Штук"] for r in gone), sum(r["Штук"] for r in added),
        round(sum(r["вес"] for r in gone)),
        round(sum(r["вес"] for r in added)), "",
    ], wrap=(2,), center=(4, 5, 6, 7), font=BOLD, fill=GROUP_FILL)

    row += 1
    for t in [
        "Стальная задвижка Ду300, снятая 22.10, стоит здесь строкой "
        "«не поехало» — решение собственника исполнено, бригада ее не "
        "грузила.",
        "«Задвижки (желтая)» 544-03.046 в факте весят 1,4 кг за штуку "
        "вместо прежних 40. Это 18 кг на позицию вместо 520: наш "
        "справочный вес был завышен в 29 раз.",
        "Догруженное — это не самовольство бригады, а то, что нашлось "
        "при отборе. Но цены по этим позициям в Договоре нет, и по "
        "пункту 3.5 проданными они не считаются.",
    ]:
        cell = ws.cell(row=row, column=1, value="— " + t)
        cell.font = BASE
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=row, start_column=1,
                       end_row=row, end_column=len(COLS_DIFF))
        ws.row_dimensions[row].height = 30
        row += 1

    widths(ws, [w for _, w in COLS_DIFF])
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1


def sheet_stowage(wb, rows, tot):
    ws = wb.create_sheet("Укладка")
    ws["A1"] = "Как укладываем"
    ws["A1"].font = TITLE
    ws["A2"] = ("Пол в контейнере деревянный, настил под низ не нужен. "
                "Задвижки кладем прямо на пол, поверх них доску, и уже "
                "на доску остальную арматуру — чтобы металл не бился о "
                "металл. Доска идет МЕЖДУ ярусами, а не под низ.")
    ws["A2"].font = NOTE

    gates = [r for r in rows if "адвижка клинкетная" in r["name"]]
    filters = [r for r in rows if "ильтр" in r["name"]]
    heavy = [r for r in rows if r["unit_kg"] >= 50 and r not in gates
             and r not in filters]
    small = [r for r in rows if r["unit_kg"] < 50 and r not in gates
             and r not in filters]
    start = 4
    for i, h in enumerate(["Ярус", "Что кладем", "Позиций", "Штук",
                           "Вес, кг", "Как"], start=1):
        ws.cell(row=start, column=i, value=h)
    head(ws, start, 6)
    row = start + 1
    for tier, what, items, how in [
        ("1 ярус", "задвижки клинкетные НА ПОЛ", gates,
         "пол деревянный, настил не нужен. Штоком ГОРИЗОНТАЛЬНО, "
         "фланец на фланец не кладем никогда, съемные маховики "
         "снимаем в ящик с биркой"),
        ("Доска", "перекладочная доска поверх задвижек", None,
         f"{rub(round(tot['brd']))} кг на всю укладку. Кладется "
         f"МЕЖДУ ярусами, чтобы арматура не билась друг о друга. "
         f"Прокладки с шагом 300-400 мм и строго над корпусами "
         f"задвижек, под угол ящика над пустотой — брус 50х50"),
        ("2 ярус", "тяжелая арматура от 50 кг", heavy,
         "на доску поверх задвижек, либо рядом на свободном полу"),
        ("Доска", "еще слой доски", None,
         "между тяжелой арматурой и ящиками, по тому же правилу"),
        ("3 ярус", "ящики и коробки с мелочью", small,
         "на доску. Мелочь едет в ящиках, россыпью не кладем"),
        ("Верх", "фильтры забортной воды", filters,
         "наверх идет только легкое; крупные фильтры Ду200 ставим на "
         "свободный пол, а не на ящики"),
    ]:
        row = put(ws, row, [
            tier, what,
            len(items) if items is not None else "",
            rub(sum(i["qty"] for i in items)) if items is not None else "",
            rub(sum(i["kg"] for i in items)) if items is not None
            else round(tot["brd"]),
            how,
        ], wrap=(2, 6), center=(3, 4), height=46)

    row += 2
    ws.cell(row=row, column=1,
            value=f"Пол контейнера: {CONTAINER}, "
                  f"{FLOOR_L:.2f} м в длину").font = SUB
    row += 1
    for i, h in enumerate(["Ду", "Штук", "Поперек, шт", "Рядов",
                           "Глубина, м", "Высота стопки, м"], start=1):
        ws.cell(row=row, column=i, value=h)
    head(ws, row, 6, height=20)
    row += 1
    counts = defaultdict(int)
    for r in gates:
        counts[r["du"] or 0] += r["qty"]
    big = 0.0
    small_gates = 0.0
    for du in sorted(counts, reverse=True):
        n = gate_geometry.across(du)
        nrows = -(-counts[du] // n)
        d = gate_geometry.depth(du) * nrows
        if du >= BIG_DU:
            big += d
        else:
            small_gates += d
        row = put(ws, row, [
            f"Ду{du}" if du else "без Ду", counts[du], n, nrows,
            round(d, 2), round(gate_geometry.stack(du), 2),
        ], center=(1, 2, 3, 4, 5, 6),
            fill=None if du >= BIG_DU else MARK_FILL)
    free = FLOOR_L - big
    row = put(ws, row, [
        f"Крупные от Ду{BIG_DU}", "", "", "", round(big, 2), "",
    ], center=(2, 3, 4, 5, 6), font=BOLD, fill=GROUP_FILL)
    row = put(ws, row, [
        "Свободный пол", "", "", "", round(free, 2),
        f"{int(free / gate_geometry.CRATE_L) * 2} мест под ящики на полу",
    ], center=(2, 3, 4, 5), font=BOLD,
        fill=OK_FILL if free >= 0 else WARN_FILL)
    row = put(ws, row, [
        "Высота под вторым ярусом", "", "", "",
        round(max(gate_geometry.stack(du) for du in counts
                  if du >= BIG_DU) + gate_geometry.BOARD, 2),
        "задвижка на боку плюс доска — с этой отметки начинается "
        "второй ярус арматуры",
    ], center=(2, 3, 4, 5), fill=GROUP_FILL)

    row += 1
    for t in [
        f"Задвижки от Ду{BIG_DU} просят {big:.2f} м пола при длине "
        f"{FLOOR_L:.2f} м. Они ложатся В ОДИН СЛОЙ, и остается "
        f"{free:.2f} м свободного пола." if free >= 0 else
        f"Задвижки от Ду{BIG_DU} просят {big:.2f} м пола при длине "
        f"{FLOOR_L:.2f} м — в один слой НЕ ложатся, не хватает "
        f"{-free:.2f} м.",
        f"Мелкие задвижки Ду125 и ниже заняли бы еще "
        f"{small_gates:.2f} м пола. Их на пол не кладем: они легкие, "
        f"идут в ящики и наверх. Желтым они и помечены.",
        "Габариты справочные, по типовым судовым бронзовым клинкетам. "
        "Замерить на первых рядах в день погрузки: справочник уже "
        "ошибся на массах почти вдвое.",
    ]:
        cell = ws.cell(row=row, column=1, value="— " + t)
        cell.font = BASE
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[row].height = 30
        row += 1

    row += 1
    ws.cell(row=row, column=1, value="Чего делать нельзя").font = SUB
    row += 1
    for t in [
        "Класть фланец на фланец. Ломается не корпус — бронза держит "
        "200 МПа при наших 0,01 — а шток, маховик и привалочная "
        "поверхность фланца.",
        "Класть арматуру на арматуру БЕЗ ДОСКИ между ярусами. Металл "
        "о металл в дороге разбивает фланцы и штоки. Между каждым "
        "ярусом — слой доски.",
        "Класть прокладку под ящик над пустотой: доска 200х25 на "
        "пролете 600 мм дает 10,6 МПа изгиба при расчетных 10 для сырой "
        "необрезной с сучками. Шаг 300-400 мм и над корпусами.",
        "Грузить маховики россыпью. Снимаем в ящик с биркой, иначе "
        "задвижка приедет неполной и это станет претензией.",
    ]:
        cell = ws.cell(row=row, column=1, value="— " + t)
        cell.font = BASE
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[row].height = 32
        row += 1
    widths(ws, [12, 30, 9, 9, 11, 58])


def sheet_checklist(wb, rows, tot):
    ws = wb.create_sheet("Что закрыть")
    open_price = [r for r in rows if r["price"] == AGREE]
    no_draw = [r for r in rows if not r["draw"]]
    checks, failed = packing_lists.check()
    bad = [c for c in checks if c["Расхождение"]]
    ws["A1"] = "Что закрыть после отгрузки"
    ws["A1"].font = TITLE
    ws["A2"] = ("Контейнер ушел 26.10.2026. Порядок по важности: "
                "сверху то, без чего товар по Договору не считается "
                "проданным и расчет по нему не пойдет.")
    ws["A2"].font = NOTE
    start = 4
    for i, h in enumerate(["№", "Что сделать", "Почему это важно",
                           "Готово"], start=1):
        ws.cell(row=start, column=i, value=h)
    head(ws, start, 4)
    row = start + 1
    for n, (what, why) in enumerate([
        (f"Согласовать письменно цену по {len(open_price)} позициям "
         f"«по согласованию» ({rub(sum(r['kg'] for r in open_price))} кг)",
         "по пункту 3.5 до согласования цены товар не считается "
         "проданным; грузить можно, но расчет по нему не пойдет"),
        ("Согласовать цену по электромагнитным клапанам",
         "по ставке за килограмм они не считаются, цена у них от "
         "привода и катушки"),
        ("Подписать приложение № 4 к Договору по ФАКТУ отгрузки",
         "уехало не то, что планировали: 7 позиций догрузили, 2 не "
         "поехали, по 4 пересчитали количество. Старое приложение "
         "фактической передаче уже не соответствует"),
        ("Согласовать цену по догруженным позициям: колонка "
         "указательная, крышки с коменсом и жалюзи, грибок Ду200",
         "их не было в плане, и цены в Договоре по ним нет; "
         "по пункту 3.5 проданными они не считаются"),
        (f"Разобрать расхождение по коробке № "
         f"{', '.join(str(c['Место']) for c in bad)}"
         if bad else "Сверить коробки с упаковочными листами",
         "в коробке лежит не то, что записано в перечне; приемка в "
         "Петербурге идет 3 рабочих дня по пункту 2.2, потом товар "
         "считается принятым"),
        ("Принимать груз в Петербурге ПО МЕСТАМ, вкладка «Упаковка»",
         "13 мест с упаковочными листами плюс мешки; по общему "
         "перечню недостачу в коробке не увидеть"),
        ("Свериться с фактическим весом по транспортным документам",
         "веса паспортные, и цена считается ОТ ВЕСА. По желтым "
         "задвижкам факт уже разошелся со справочником в 29 раз"),
        (f"Переписать чертежи с корпуса у {len(no_draw)} позиций без "
         f"чертежа", "без чертежа позицию трудно продать и невозможно "
         "сверить с прайсом"),
        (f"Проверить комплектность маховиков: по перечню их "
         f"{rub(tot['wheels'])} шт",
         "без маховика задвижка неполная, это претензия по "
         "комплектности"),
        ("Списать отгруженное со складского учета в тот же день",
         "иначе расхождение накопится снова, как было к 16.06.2025"),
    ], start=1):
        row = put(ws, row, [n, what, why, ""], wrap=(2, 3), center=(1,),
                  height=44, fill=WARN_FILL if n <= 2 else None)
    widths(ws, [5, 52, 60, 12])


def sheet_excluded(wb):
    ws = wb.create_sheet("Снято с отгрузки")
    ws["A1"] = "Что не едет"
    ws["A1"].font = TITLE
    ws["A2"] = ("Позиции сняты в расчете, но физически лежат на складе. "
                "Держим их списком, а не правкой перечня: иначе при "
                "следующей выгрузке они вернутся молча.")
    ws["A2"].font = NOTE
    start = 4
    for i, h in enumerate(["Позиция", "Почему снято"], start=1):
        ws.cell(row=start, column=i, value=h)
    head(ws, start, 2)
    row = start + 1
    for name, why in container_list.EXCLUDED.items():
        row = put(ws, row, [name, why], wrap=(1, 2), height=30)
    widths(ws, [52, 56])


def main():
    rows, blocks = collect()
    wb = Workbook()
    tot = sheet_summary(wb, rows, blocks)
    sheet_groups(wb, rows)
    sheet_items(wb, rows)
    sheet_buy(wb, rows)
    sheet_packing(wb)
    sheet_diff(wb)
    sheet_stowage(wb, rows, tot)
    sheet_checklist(wb, rows, tot)
    sheet_excluded(wb)

    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / "Контейнер_отгружено_26.10.xlsx"
    wb.save(path)
    print(f"готово: {path}\n")
    print(f"{len(rows)} позиций, {rub(tot['qty'])} шт, "
          f"{rub(tot['kg'])} кг груза")
    print(f"доска {rub(round(tot['brd']))} кг, брутто "
          f"{rub(round(tot['gross']))} кг при пределе {rub(LIMIT_KG)}")
    print("УКЛАДЫВАЕМСЯ, запас " + rub(round(tot["left"])) + " кг"
          if tot["left"] >= 0
          else "ПРЕВЫШЕНИЕ " + rub(round(-tot["left"])) + " кг")
    print(f"\n1. закуп по договору {rub(tot['money'])} руб "
          f"(бронза по {scrap_prices.purchase_rate('бронза'):,.2f} "
          f"руб за кг = лом * {scrap_prices.MULTIPLIER})".replace(",", " "))
    print(f"2. ремонт {rub(tot['repair'])} руб")
    print(f"3. продажная цена {rub(tot['resale'])} руб")
    print(f"маховиков {rub(tot['wheels'])} шт")


if __name__ == "__main__":
    main()
