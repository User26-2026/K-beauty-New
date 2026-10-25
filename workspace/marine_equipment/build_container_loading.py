#!/usr/bin/env python3
"""Что грузим в контейнер: один файл для дня погрузки.

Сюда сведено все, о чем договаривались: перечень от 21.10.2026, снятые
позиции, цены по ставке на бронзу, ремонт по прайсу мастерской,
ориентир реализации, веса по паспорту, маховики, тара, предел 25 тонн
и правила укладки.

Файл делается для погрузки, а не для переговоров, поэтому в нем рядом
стоят и деньги, и килограммы: бригада смотрит вес и тару, менеджер —
цену и то, что нельзя грузить без согласования.

Шесть вкладок:
  Сводка           — итоги, тоннаж против предела, принятые решения
  По группам       — вес и деньги по каждой группе перечня
  Перечень         — все позиции построчно, порядок погрузки
  Укладка          — ярусы, доска, маховики, что на что класть
  Перед погрузкой  — чек-лист, что закрыть до машины
  Снято с отгрузки — что и почему не едет

Результат: outputs/marine_equipment/Контейнер_погрузка.xlsx
"""

import pathlib
import sys
from collections import defaultdict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
import container_list  # noqa: E402
import upd_shipments  # noqa: E402
from build_contract_annexes import AGREE, LIMIT_KG, annex4  # noqa: E402
from build_container_spb import plural  # noqa: E402
from build_logistics_compare import board  # noqa: E402

OUT = pathlib.Path(__file__).resolve().parents[2] / "outputs" / "marine_equipment"

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
    brd = board(kg)[0]
    gross = kg + brd
    left = LIMIT_KG - gross
    money = sum(val(r["price"]) * r["qty"] for r in rows)
    repair = sum((r["repair"] or 0) * r["qty"] for r in rows)
    resale = sum(val(r["resale"]) * r["qty"] for r in rows)
    wheels = sum(r["wheels"] or 0 for r in rows)
    open_price = [r for r in rows if r["price"] == AGREE]
    no_price = [r for r in rows if not r["price"]]

    ws["A1"] = "Погрузка контейнера в Санкт-Петербург"
    ws["A1"].font = TITLE
    ws["A2"] = ("Перечень от 21.10.2026 со снятыми позициями. Все цифры "
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
         "сырая необрезная, 0,035 куб. м на тонну груза", None),
        ("БРУТТО", f"{rub(round(gross))} кг", "груз плюс доска", GROUP_FILL),
        ("Предел погрузки", f"{rub(LIMIT_KG)} кг", "решение собственника",
         None),
        ("ЗАПАС" if left >= 0 else "ПРЕВЫШЕНИЕ",
         f"{rub(round(abs(left)))} кг ({abs(left) / LIMIT_KG:.0%})",
         "укладываемся, можно догружать" if left >= 0
         else "перечень надо сокращать",
         OK_FILL if left >= 0 else WARN_FILL),
        ("", "", "", None),
        ("Сумма по Договору", f"{rub(money)} руб",
         "по ставке 1 862,50 руб за кг на бронзу и латунь", None),
        ("Ремонт", f"{rub(repair)} руб",
         "прайс мастерской, в цену по Договору не входит", None),
        ("Ориентир реализации", f"{rub(resale)} руб",
         "информационно, цену согласуют письменно по пункту 3.2", None),
        ("", "", "", None),
        ("Маховиков", f"{rub(wheels)} шт",
         "снимаются в ящик с биркой, считать при погрузке", MARK_FILL),
        ("Цена «по согласованию»", f"{len(open_price)} позиций, "
         f"{rub(sum(r['kg'] for r in open_price))} кг",
         "по пункту 3.5 проданным не считается — закрыть до отгрузки",
         WARN_FILL),
        ("Цена не ставится", f"{len(no_price)} позиций",
         "электромагнитные, цена не от массы металла", MARK_FILL),
    ]:
        row = put(ws, row, [name, value, note], wrap=(3,),
                  font=BOLD if name in ("БРУТТО", "ЗАПАС", "ПРЕВЫШЕНИЕ")
                  else BASE, fill=fill)

    row += 1
    ws.cell(row=row, column=1, value="Что решили и по чему считаем").font = SUB
    row += 1
    for t in [
        "Цена бронзы и латуни считается ОТ ВЕСА: 1 862,50 руб за "
        "килограмм. Уточнится вес при замере — цена пересчитается сама.",
        "Сталь, титан и алюминий по ставке не считаются, у них цена "
        "«по согласованию». Электромагнитные клапаны тоже: там платят "
        "за привод и катушку, а не за металл.",
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
                money=money, repair=repair, resale=resale, wheels=wheels)


COLS_GROUP = [
    ("Группа", 34), ("Позиций", 8), ("Штук", 8), ("Вес, кг", 10),
    ("Доля веса", 9), ("Вес нарастающим, кг", 12),
    ("По Договору, руб", 14), ("Ремонт, руб", 12),
    ("Реализация, руб", 14), ("Руб за кг", 10),
]


def sheet_groups(wb, rows):
    ws = wb.create_sheet("По группам")
    agg = defaultdict(lambda: dict(n=0, qty=0, kg=0.0, money=0.0,
                                   repair=0.0, resale=0.0))
    for r in rows:
        a = agg[r["block"]]
        a["n"] += 1
        a["qty"] += r["qty"]
        a["kg"] += r["kg"]
        a["money"] += val(r["price"]) * r["qty"]
        a["repair"] += (r["repair"] or 0) * r["qty"]
        a["resale"] += val(r["resale"]) * r["qty"]
    total_kg = sum(a["kg"] for a in agg.values())

    ws["A1"] = "Что грузим: по группам"
    ws["A1"].font = TITLE
    ws["A2"] = ("Порядок групп — как грузим. Нарастающий вес показывает, "
                "сколько тонн уже в контейнере к концу каждой группы.")
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
            round(a["money"] / a["kg"]) if a["kg"] else "",
        ], wrap=(1,), center=(2, 3, 5), money=(4, 6, 7, 8, 9, 10))
    row = put(ws, row, [
        "ИТОГО", sum(a["n"] for a in agg.values()),
        sum(a["qty"] for a in agg.values()), round(total_kg), "100%",
        round(total_kg), round(sum(a["money"] for a in agg.values())),
        round(sum(a["repair"] for a in agg.values())),
        round(sum(a["resale"] for a in agg.values())),
        round(sum(a["money"] for a in agg.values()) / total_kg),
    ], center=(2, 3, 5), money=(4, 6, 7, 8, 9, 10), font=BOLD,
        fill=GROUP_FILL)
    ws.freeze_panes = f"A{start + 1}"
    widths(ws, [w for _, w in COLS_GROUP])
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1


COLS_ITEM = [
    ("№", 5), ("Наименование", 46), ("Чертеж", 18), ("Ду", 6), ("Ру", 6),
    ("Сост.", 8), ("Шт", 6), ("Вес ед., кг", 9), ("Вес всего, кг", 10),
    ("Маховики, шт", 9), ("Тара", 12), ("По Договору за ед., руб", 13),
    ("Ремонт за ед., руб", 11), ("Реализация за ед., руб", 13),
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
            r["pack"], r["price"] or "цена отдельно",
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


def sheet_stowage(wb, rows, tot):
    ws = wb.create_sheet("Укладка")
    ws["A1"] = "Как укладываем"
    ws["A1"].font = TITLE
    ws["A2"] = ("Порядок ярусов и правила сохранности. Главный риск не "
                "вес сам по себе, а что лежит поверх нижнего ряда "
                "задвижек.")
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
        ("Настил", "сырая необрезная доска по полу", None,
         f"{rub(round(tot['brd']))} кг доски на всю укладку, "
         f"0,035 куб. м на тонну груза"),
        ("1 ярус", "задвижки клинкетные", gates,
         "штоком ГОРИЗОНТАЛЬНО, фланец на фланец не кладем никогда, "
         "съемные маховики снимаем в ящик с биркой"),
        ("2 ярус", "тяжелая арматура от 50 кг", heavy,
         "рядом с задвижками на полу, а не сверху"),
        ("3 ярус", "ящики и коробки с мелочью", small,
         "на доску поверх задвижек, прокладки с шагом 300-400 мм и "
         "строго над корпусами, под угол ящика над пустотой — брус 50х50"),
        ("Верх", "фильтры забортной воды", filters,
         "наверх идет только легкое; крупные фильтры Ду200 ставим на "
         "пол, а не на ящики"),
    ]:
        row = put(ws, row, [
            tier, what,
            len(items) if items is not None else "",
            rub(sum(i["qty"] for i in items)) if items is not None else "",
            rub(sum(i["kg"] for i in items)) if items is not None
            else round(tot["brd"]),
            how,
        ], wrap=(2, 6), center=(3, 4), height=46)

    row += 1
    ws.cell(row=row, column=1, value="Чего делать нельзя").font = SUB
    row += 1
    for t in [
        "Класть фланец на фланец. Ломается не корпус — бронза держит "
        "200 МПа при наших 0,01 — а шток, маховик и привалочная "
        "поверхность фланца.",
        "Ставить арматуру на арматуру ярусами. Задвижки и тяжелое лежат "
        "на полу в один слой, ящики стоят рядом, на ящики идут только "
        "фильтры.",
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
    ws = wb.create_sheet("Перед погрузкой")
    open_price = [r for r in rows if r["price"] == AGREE]
    no_draw = [r for r in rows if not r["draw"]]
    ws["A1"] = "Что закрыть до того, как придет машина"
    ws["A1"].font = TITLE
    ws["A2"] = ("Порядок по важности: сверху то, без чего товар по "
                "договору не считается проданным.")
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
        ("Снять и уложить маховики в ящик с биркой, "
         f"по перечню их {rub(tot['wheels'])} шт",
         "без маховика задвижка неполная; россыпью они теряются и "
         "превращаются в претензию по комплектности"),
        ("Взвесить на погрузке хотя бы по одной задвижке каждого "
         "диаметра и один крупный фильтр",
         "веса паспортные; прошлый справочник врал вдвое, и цена "
         "считается ОТ ВЕСА — ошибка в весе это ошибка в деньгах"),
        (f"Переписать чертежи с корпуса у {len(no_draw)} позиций без "
         f"чертежа", "без чертежа позицию трудно продать и невозможно "
         "сверить с прайсом"),
        ("Проверить, что снятые позиции физически не поехали",
         "они исключены в расчете, но на складе лежат рядом"),
        ("Заказать доску: "
         f"{rub(round(tot['brd']))} кг сырой необрезной",
         "перекладка между ярусами; сухую не берем, дороже и не найти"),
        ("Заказать погрузчик и бригаду: подача 12 000, 2 000 в час, "
         "четверо грузчиков по 750 в час, около 4 часов",
         "контейнер выходит примерно в 32 000 руб"),
        ("Сфотографировать укладку по ярусам до закрытия дверей",
         "приемка в Петербурге идет 3 рабочих дня по пункту 2.2, "
         "потом товар считается принятым"),
        ("Отдать бригаде вкладку «Перечень», менеджеру — «По группам»",
         "бригаде нужны вес и тара, менеджеру цена и что нельзя "
         "грузить без согласования"),
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
    sheet_stowage(wb, rows, tot)
    sheet_checklist(wb, rows, tot)
    sheet_excluded(wb)

    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / "Контейнер_погрузка.xlsx"
    wb.save(path)
    print(f"готово: {path}\n")
    print(f"{len(rows)} позиций, {rub(tot['qty'])} шт, "
          f"{rub(tot['kg'])} кг груза")
    print(f"доска {rub(round(tot['brd']))} кг, брутто "
          f"{rub(round(tot['gross']))} кг при пределе {rub(LIMIT_KG)}")
    print("УКЛАДЫВАЕМСЯ, запас " + rub(round(tot["left"])) + " кг"
          if tot["left"] >= 0
          else "ПРЕВЫШЕНИЕ " + rub(round(-tot["left"])) + " кг")
    print(f"\nпо договору {rub(tot['money'])}, ремонт {rub(tot['repair'])}, "
          f"реализация {rub(tot['resale'])} руб")
    print(f"маховиков {rub(tot['wheels'])} шт")


if __name__ == "__main__":
    main()
