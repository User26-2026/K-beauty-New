#!/usr/bin/env python3
"""Сверка цен покупателя с нашими по отгруженному контейнеру.

Покупатель вернул наш файл Контейнер_отгружено_26.10.xlsx, дописав
колонку «Рыночная закупка» — свою цену закупа за единицу. Заполнил 191
позицию из 222, в том числе 73 из тех, где у нас стояло «по
согласованию». Файл сверяет его цену с нашей позиция к позиции.

ЧТО ВИДНО СРАЗУ. Итоги почти совпадают: его предложение 27,26 млн
против нашего закупа 26,83 млн. Но это обманчиво, потому что набраны
они по-разному: по сравнимым позициям он урезал нашу цену на 8,97 млн,
а добавил 9,46 млн на позициях, где цены у нас не было. То есть он не
поднял цену, он переложил ее с крупного на мелкое.

ПОЧЕМУ ТАК. По крупной арматуре он ставит ВЕРХНЮЮ ГРАНИЦУ СТАРОЙ ЦЕНЫ
ДОГОВОРА, а не ставку за килограмм: задвижка Ду350 — 300 000 вместо
395 781, Ду300 — 200 000 вместо 357 600, Ду250 — 150 000 вместо
232 812, Ду200 — 90 000 вместо 142 481. Из 104 позиций, где старая
цена договора есть, 61 совпадает с его ценой до рубля. Ставка за
килограмм крупное дорожила, и именно там он от нее отказался.

ГЛАВНАЯ ЦИФРА ДЛЯ РАЗГОВОРА. Договоренность от 25.10.2026 — закуп
равен цене лома на 2,5. Его предложение по 191 позиции дает 2,00 лома,
а не 2,5. По ставке те же позиции стоили бы 34,12 млн против его
27,26 млн, то есть он просит скидку 6,86 млн, или 20%.

Сильнее всего он отходит от договоренности на тяжелом: коробка 2-х
клапановая Ду80 — 1,28 лома, задвижка Ду300 — 1,40, коробка 3-х
клапановая Ду80 — 1,41, фильтр Ду200 — 1,54, задвижка Ду200 — 1,58,
Ду250 — 1,61. На мелочи он наоборот выше 2,5: там ставка давала
копейки, и его цена ближе к рыночной.

Ниже лома он опустился один раз: клапан штуцерный проходной
бессальниковый, 3 шт, 3 500 руб при ломе 4 246.

Он поправил материал двух позиций на сталь — фильтр Ду100
(427-03.192-2) и паровой клапан Ду65Ру10 (581-03.034). Владельцем это
не подтверждено, поэтому в расчет не идет: сталь вместо бронзы меняет
лом в 39 раз и цену продажи втрое, такое подтверждают, а не принимают
на слово. Позиции помечены на отдельной вкладке.

Результат: outputs/marine_equipment/Сверка_цен_покупателя.xlsx
"""

import pathlib
import sys
from collections import defaultdict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
import buyer_prices  # noqa: E402
import scrap_prices  # noqa: E402
import upd_shipments  # noqa: E402
from build_container_loading import buy_rate, collect, skip_reason  # noqa: E402
from build_container_spb import plural  # noqa: E402
from build_contract_annexes import (clean, norm_draw,  # noqa: E402
                                    normalize_price, price_index)

OUT = (pathlib.Path(__file__).resolve().parents[2] / "outputs"
       / "marine_equipment")

MULT = scrap_prices.MULTIPLIER

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
BAD_FILL = PatternFill("solid", fgColor="FCE4D6")
RED_FILL = PatternFill("solid", fgColor="F8CBAD")
MARK_FILL = PatternFill("solid", fgColor="FFF2CC")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def rub(n):
    return f"{n:,.0f}".replace(",", " ")


def val(text):
    """Число из нашей цены. «По согласованию» и пусто дают ноль."""
    b = upd_shipments.bounds(str(text or ""))
    return max(b) if b else 0


def head(ws, row, ncols, height=38):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEAD_FILL
        cell.font = HEAD_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center",
                                   horizontal="center")
        cell.border = BORDER
    ws.row_dimensions[row].height = height


def put(ws, row, values, wrap=(), center=(), money=(), ratio=(), fill=None,
        font=None):
    for i, v in enumerate(values, start=1):
        cell = ws.cell(row=row, column=i, value=v)
        cell.font = font or BASE
        cell.border = BORDER
        cell.alignment = Alignment(
            vertical="top", wrap_text=(i in wrap),
            horizontal="center" if i in center else None)
        if i in money:
            cell.number_format = "#,##0"
        if i in ratio:
            cell.number_format = "0.00"
        if fill:
            cell.fill = fill
    return row + 1


def widths(ws, values):
    for i, w in enumerate(values, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def notes(ws, row, ncols, texts, height=30):
    for t in texts:
        cell = ws.cell(row=row, column=1, value="— " + t)
        cell.font = BASE
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=row, start_column=1, end_row=row,
                       end_column=ncols)
        ws.row_dimensions[row].height = height
        row += 1
    return row


def landscape(ws):
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1


def old_contract_price(row, by_draw, by_name):
    """Верхняя граница старой цены договора, если она там есть."""
    p = by_draw.get(norm_draw(row["draw"])) or by_name.get(clean(row["name"]))
    if not p or "согласован" in str(p).lower():
        return None
    out, _ = normalize_price(str(p))
    b = upd_shipments.bounds(out)
    return max(b) if b else None


def build():
    """Строки сверки: наша цена, его цена, лом, ставка, старый договор."""
    spec, _ = collect()
    his = buyer_prices.read()
    by_draw, by_name = price_index()
    rows = []
    for r in spec:
        h = his.get(r["n"], {})
        ours = val(r["price"])
        scrap = scrap_prices.BY_MATERIAL.get(
            r["material"], scrap_prices.SCRAP["бронза"]) * r["unit_kg"]
        rate = (buy_rate(r["material"]) or 0) * r["unit_kg"]
        theirs = h.get("Цена")
        rows.append(dict(
            n=r["n"], group=r["block"], name=r["name"], draw=r["draw"],
            du=r["du"], ru=r["ru"], cond=r["cond"], qty=r["qty"],
            kg=r["unit_kg"], material=r["material"],
            ours=ours, ours_txt=str(r["price"] or ""),
            why=skip_reason(r), repair=r["repair"] or 0,
            resale=val(r["resale"]), resale_txt=str(r["resale"] or ""),
            theirs=theirs, note=h.get("Пометка", ""),
            scrap=scrap, rate=rate,
            old=old_contract_price(r, by_draw, by_name),
        ))
    for r in rows:
        r["diff"] = (r["theirs"] - r["ours"]) * r["qty"] \
            if r["theirs"] and r["ours"] else 0
        r["k"] = r["theirs"] / r["scrap"] if r["theirs"] and r["scrap"] else 0
        r["eq_old"] = bool(r["theirs"] and r["old"]
                           and abs(r["theirs"] - r["old"]) < 1)
    return rows


def totals(rows):
    both = [r for r in rows if r["theirs"] and r["ours"]]
    closed = [r for r in rows if r["theirs"] and not r["ours"]]
    ours_only = [r for r in rows if r["ours"] and not r["theirs"]]
    neither = [r for r in rows if not r["ours"] and not r["theirs"]]
    priced = [r for r in rows if r["theirs"]]
    return dict(
        both=both, closed=closed, ours_only=ours_only, neither=neither,
        priced=priced,
        s_both_our=sum(r["ours"] * r["qty"] for r in both),
        s_both_his=sum(r["theirs"] * r["qty"] for r in both),
        s_closed=sum(r["theirs"] * r["qty"] for r in closed),
        s_his=sum(r["theirs"] * r["qty"] for r in priced),
        s_our=sum(r["ours"] * r["qty"] for r in rows if r["ours"]),
        scrap_priced=sum(r["scrap"] * r["qty"] for r in priced),
        rate_priced=sum(r["rate"] * r["qty"] for r in priced),
        resale_priced=sum(r["resale"] * r["qty"] for r in priced),
    )


def sheet_summary(wb, rows, t):
    ws = wb.active
    ws.title = "Сводка"
    ws["A1"] = "Цены покупателя против наших: контейнер от 26.10.2026"
    ws["A1"].font = TITLE
    ws["A2"] = (f"{buyer_prices.WHO} вернул наш файл с дописанной колонкой "
                f"«Рыночная закупка»: цена проставлена у "
                f"{len(t['priced'])} "
                f"{plural(len(t['priced']), 'позиции', 'позиций', 'позиций')}"
                f" из {len(rows)}. Все наши цифры здесь — из нашего "
                f"же расчета, а не из его файла.")
    ws["A2"].font = NOTE

    k_his = t["s_his"] / t["scrap_priced"] if t["scrap_priced"] else 0
    want = t["scrap_priced"] * MULT
    row = 4
    ws.cell(row=row, column=1, value="Главное").font = SUB
    row += 1
    cols = [("Что считаем", 54), ("Позиций", 9), ("Штук", 9),
            ("Наша цена, руб", 16), ("Его цена, руб", 16),
            ("Разница, руб", 16), ("Его к нашей", 11)]
    for i, (h, _) in enumerate(cols, start=1):
        ws.cell(row=row, column=i, value=h)
    head(ws, row, len(cols))
    row += 1
    both, closed, only = t["both"], t["closed"], t["ours_only"]
    row = put(ws, row, [
        "Сравнимые: цена есть и у нас, и у него", len(both),
        sum(r["qty"] for r in both), round(t["s_both_our"]),
        round(t["s_both_his"]), round(t["s_both_his"] - t["s_both_our"]),
        f"{t['s_both_his'] / t['s_both_our']:.0%}" if t["s_both_our"] else "",
    ], wrap=(1,), center=(2, 3, 7), money=(4, 5, 6), fill=BAD_FILL)
    row = put(ws, row, [
        "Он закрыл то, что у нас шло «по согласованию»", len(closed),
        sum(r["qty"] for r in closed), "", round(t["s_closed"]),
        round(t["s_closed"]), "",
    ], wrap=(1,), center=(2, 3, 7), money=(4, 5, 6), fill=OK_FILL)
    row = put(ws, row, [
        "Цена есть у нас, у него пусто", len(only),
        sum(r["qty"] for r in only),
        round(sum(r["ours"] * r["qty"] for r in only)), "", "", "",
    ], wrap=(1,), center=(2, 3, 7), money=(4, 5, 6))
    row = put(ws, row, [
        "Цены нет ни у нас, ни у него", len(t["neither"]),
        sum(r["qty"] for r in t["neither"]), "", "", "", "",
    ], wrap=(1,), center=(2, 3, 7), money=(4, 5, 6), fill=MARK_FILL)
    row = put(ws, row, [
        "ИТОГО", len(rows), sum(r["qty"] for r in rows), round(t["s_our"]),
        round(t["s_his"]), round(t["s_his"] - t["s_our"]),
        f"{t['s_his'] / t['s_our']:.0%}" if t["s_our"] else "",
    ], center=(2, 3, 7), money=(4, 5, 6), font=BOLD, fill=GROUP_FILL)

    row += 1
    row = notes(ws, row, len(cols), [
        "Итоги почти совпадают, и это обманчиво. По сравнимым позициям он "
        "урезал нашу цену на "
        f"{rub(t['s_both_our'] - t['s_both_his'])} руб, а добавил "
        f"{rub(t['s_closed'])} руб там, где цены у нас не было. Он не "
        "поднял цену, он переложил ее с крупного на мелкое.",
        "Строка ИТОГО по нашей цене — это весь закуп по Договору, 120 "
        "позиций. Остальные у нас шли «по согласованию»: по пункту 3.5 "
        "такой товар проданным не считается.",
        "На вкладке «По группам» наша сумма на "
        f"{rub(sum(r['ours'] * r['qty'] for r in t['ours_only']))} руб "
        "меньше: там считаются только позиции, которые он оценил, а три "
        "позиции с нашей ценой он оставил пустыми.",
    ], height=42)

    row += 1
    ws.cell(row=row, column=1,
            value="Договоренность от 25.10.2026: закуп = лом × 2,5"
            ).font = SUB
    row += 1
    cols2 = [("Показатель", 54), ("Руб", 18), ("Кратность лому", 14)]
    for i, (h, _) in enumerate(cols2, start=1):
        ws.cell(row=row, column=i, value=h)
    head(ws, row, len(cols2), height=22)
    row += 1
    for label, v, k, fill in [
        (f"Лом по {len(t['priced'])} позициям, которые он оценил",
         round(t["scrap_priced"]), 1.0, None),
        (f"Должно быть по договоренности: лом × {MULT:g}".replace(".", ","),
         round(want), MULT, OK_FILL),
        ("Его предложение", round(t["s_his"]),
         k_his, BAD_FILL),
        ("Разница с договоренностью", round(t["s_his"] - want),
         0, RED_FILL),
        ("Ориентир реализации по тем же позициям",
         round(t["resale_priced"]),
         t["resale_priced"] / t["scrap_priced"] if t["scrap_priced"] else 0,
         None),
    ]:
        row = put(ws, row, [label, v, k or ""], wrap=(1,), center=(3,),
                  money=(2,), ratio=(3,), fill=fill)

    row += 1
    row = notes(ws, row, len(cols2) + 4, [
        f"Его предложение — это {k_his:.2f} лома".replace(".", ",")
        + f" вместо {MULT:g}".replace(".", ",")
        + f". В рублях это {rub(want - t['s_his'])} руб скидки, или "
        f"{(want - t['s_his']) / want:.0%} от суммы по договоренности.",
        "Лом считается по прайсу приемки на 13.10.2026: бронза 745 руб за "
        "кг, латунь 635, сталь 19, титан 130, алюминий 260. Материал берем "
        "наш, а не его.",
        "Ориентир реализации — третья цена цепочки, по ней товар "
        "продается конечному покупателю. Она здесь для масштаба: закуп "
        "обсуждается не в вакууме.",
    ], height=42)
    widths(ws, [w for _, w in cols] + [12])
    landscape(ws)


COLS = [
    ("№", 5), ("Группа", 24), ("Наименование", 46), ("Чертеж", 17),
    ("Ду", 6), ("Шт", 6), ("Вес ед., кг", 9), ("Материал", 11),
    ("Наша цена за ед., руб", 13), ("Его цена за ед., руб", 13),
    ("Разница на позицию, руб", 14), ("Его к нашей", 10),
    ("Лом за ед., руб", 11), ("Кратность лому", 10),
    ("По ставке 1 862,50 за ед., руб", 13),
    ("Старая цена договора, руб", 13), ("Совпала со старой", 10),
    ("Почему у нас нет цены", 30), ("Его пометка", 18),
]


def sheet_rows(wb, rows):
    ws = wb.create_sheet("Построчно")
    ws["A1"] = "Все 222 позиции контейнера: наша цена против его"
    ws["A1"].font = TITLE
    ws["A2"] = ("Оранжевым — где он ниже нашей цены. Зеленым — где выше "
                "или равно. Желтым — где у нас цены не было и он ее "
                "предложил.")
    ws["A2"].font = NOTE
    start = 4
    for i, (h, _) in enumerate(COLS, start=1):
        ws.cell(row=start, column=i, value=h)
    head(ws, start, len(COLS))
    row = start + 1
    for r in rows:
        if r["theirs"] and r["ours"]:
            fill = BAD_FILL if r["theirs"] < r["ours"] else OK_FILL
        elif r["theirs"]:
            fill = MARK_FILL
        else:
            fill = None
        row = put(ws, row, [
            r["n"], r["group"], r["name"], r["draw"], r["du"] or "",
            r["qty"], r["kg"], r["material"],
            round(r["ours"]) or r["ours_txt"],
            round(r["theirs"]) if r["theirs"] else "",
            round(r["diff"]) if r["diff"] else "",
            f"{r['theirs'] / r['ours']:.0%}"
            if r["theirs"] and r["ours"] else "",
            round(r["scrap"]) or "", r["k"] or "",
            round(r["rate"]) or "", round(r["old"]) if r["old"] else "",
            "да" if r["eq_old"] else "", r["why"], r["note"],
        ], wrap=(3, 18), center=(1, 5, 6, 7, 12, 14, 17),
            money=(9, 10, 11, 13, 15, 16), ratio=(14,), fill=fill)
    ws.freeze_panes = f"A{start + 1}"
    ws.auto_filter.ref = f"A{start}:{get_column_letter(len(COLS))}{row - 1}"
    widths(ws, [w for _, w in COLS])
    ws.print_title_rows = f"{start}:{start}"
    landscape(ws)


def sheet_cuts(wb, rows):
    ws = wb.create_sheet("Где он срезал")
    sel = sorted([r for r in rows if r["diff"] < 0], key=lambda r: r["diff"])
    ws["A1"] = "Позиции, где его цена ниже нашей"
    ws["A1"].font = TITLE
    ws["A2"] = (f"{len(sel)} "
                f"{plural(len(sel), 'позиция', 'позиции', 'позиций')}, "
                f"{rub(sum(r['qty'] for r in sel))} шт, "
                f"минус {rub(-sum(r['diff'] for r in sel))} руб. По "
                f"убыванию потери на позиции.")
    ws["A2"].font = NOTE
    cols = [("№", 5), ("Наименование", 46), ("Ду", 6), ("Шт", 6),
            ("Вес ед., кг", 9), ("Наша, руб", 12), ("Его, руб", 12),
            ("Минус на позиции, руб", 14), ("Его к нашей", 10),
            ("Кратность лому", 11), ("Старая цена договора", 13),
            ("Совпала со старой", 10)]
    start = 4
    for i, (h, _) in enumerate(cols, start=1):
        ws.cell(row=start, column=i, value=h)
    head(ws, start, len(cols))
    row = start + 1
    for r in sel:
        row = put(ws, row, [
            r["n"], r["name"], r["du"] or "", r["qty"], r["kg"],
            round(r["ours"]), round(r["theirs"]), round(r["diff"]),
            f"{r['theirs'] / r['ours']:.0%}", r["k"],
            round(r["old"]) if r["old"] else "",
            "да" if r["eq_old"] else "",
        ], wrap=(2,), center=(1, 3, 4, 5, 9, 10, 12),
            money=(6, 7, 8, 11), ratio=(10,),
            fill=BAD_FILL if r["k"] and r["k"] < MULT else None)
    row = put(ws, row, ["ИТОГО", f"{len(sel)} позиций", "",
                        sum(r["qty"] for r in sel), "",
                        round(sum(r["ours"] * r["qty"] for r in sel)),
                        round(sum(r["theirs"] * r["qty"] for r in sel)),
                        round(sum(r["diff"] for r in sel)), "", "", "", ""],
              center=(4,), money=(6, 7, 8), font=BOLD, fill=GROUP_FILL)
    row += 1
    row = notefill(ws, row, len(cols), rows)
    ws.freeze_panes = f"A{start + 1}"
    widths(ws, [w for _, w in cols])
    landscape(ws)


def notefill(ws, row, ncols, rows):
    eq = [r for r in rows if r["eq_old"]]
    cmp_old = [r for r in rows if r["theirs"] and r["old"]]
    return notes(ws, row, ncols, [
        f"Из {len(cmp_old)} позиций, где старая цена договора есть, его "
        f"цена совпадает с ее верхней границей у {len(eq)}. То есть он "
        f"вернулся к старой цене договора вместо ставки за килограмм.",
        "Ставка за килограмм крупное дорожила, а мелочь дешевила. Он "
        "отказался от нее именно на крупном: задвижка Ду350 — 300 000 "
        "вместо 395 781, Ду300 — 200 000 вместо 357 600, Ду250 — 150 000 "
        "вместо 232 812, Ду200 — 90 000 вместо 142 481.",
        "Оранжевым отмечены строки, где его цена ниже 2,5 лома, то есть "
        "ниже договоренности от 25.10.2026.",
    ], height=42)


def sheet_closed(wb, rows):
    ws = wb.create_sheet("Что он закрыл")
    sel = sorted([r for r in rows if r["theirs"] and not r["ours"]],
                 key=lambda r: -r["theirs"] * r["qty"])
    ws["A1"] = "Позиции, где у нас стояло «по согласованию»"
    ws["A1"].font = TITLE
    ws["A2"] = (f"{len(sel)} "
                f"{plural(len(sel), 'позиция', 'позиции', 'позиций')}, "
                f"{rub(sum(r['qty'] for r in sel))} шт — "
                f"он предложил по ним "
                f"{rub(sum(r['theirs'] * r['qty'] for r in sel))} руб. По "
                f"пункту 3.5 Договора товар без согласованной цены "
                f"проданным не считается, так что это и есть то, что надо "
                f"закрыть письменно.")
    ws["A2"].font = NOTE
    cols = [("№", 5), ("Наименование", 46), ("Ду", 6), ("Ру", 6), ("Шт", 6),
            ("Вес ед., кг", 9), ("Сост.", 8), ("Материал", 11),
            ("Его цена за ед., руб", 12), ("Сумма по его цене, руб", 14),
            ("Лом за ед., руб", 11), ("Кратность лому", 11),
            ("По ставке за ед., руб", 12),
            ("Почему у нас не было цены", 34)]
    start = 4
    for i, (h, _) in enumerate(cols, start=1):
        ws.cell(row=start, column=i, value=h)
    head(ws, start, len(cols))
    row = start + 1
    for r in sel:
        row = put(ws, row, [
            r["n"], r["name"], r["du"] or "", r["ru"], r["qty"], r["kg"],
            r["cond"], r["material"], round(r["theirs"]),
            round(r["theirs"] * r["qty"]), round(r["scrap"]) or "",
            r["k"] or "", round(r["rate"]) or "", r["why"],
        ], wrap=(2, 14), center=(1, 3, 4, 5, 6, 7, 12),
            money=(9, 10, 11, 13), ratio=(12,),
            fill=BAD_FILL if r["k"] and r["k"] < MULT else OK_FILL)
    row = put(ws, row, [
        "ИТОГО", f"{len(sel)} позиций", "", "", sum(r["qty"] for r in sel),
        "", "", "", "", round(sum(r["theirs"] * r["qty"] for r in sel)),
        "", "", round(sum(r["rate"] * r["qty"] for r in sel)), "",
    ], center=(5,), money=(10, 13), font=BOLD, fill=GROUP_FILL)
    row += 1
    by = defaultdict(lambda: [0, 0, 0.0])
    for r in sel:
        a = by[r["why"] or "цена была закрыта"]
        a[0] += 1
        a[1] += r["qty"]
        a[2] += r["theirs"] * r["qty"]
    row += 1
    ws.cell(row=row, column=1, value="По причинам, почему цены не было").font = SUB
    row += 1
    for i, h in enumerate(["Причина", "Позиций", "Штук", "Его сумма, руб"],
                          start=1):
        ws.cell(row=row, column=i, value=h)
    head(ws, row, 4, height=22)
    row += 1
    for k, v in sorted(by.items(), key=lambda kv: -kv[1][2]):
        row = put(ws, row, [k, v[0], v[1], round(v[2])], wrap=(1,),
                  center=(2, 3), money=(4,))
    ws.freeze_panes = f"A{start + 1}"
    widths(ws, [w for _, w in cols])
    landscape(ws)


def sheet_floor(wb, rows, t):
    ws = wb.create_sheet("Ниже договоренности")
    sel = sorted([r for r in rows if r["theirs"] and r["k"] and r["k"] < MULT],
                 key=lambda r: r["k"])
    ws["A1"] = f"Где его цена ниже лома × {MULT:g}".replace(".", ",")
    ws["A1"].font = TITLE
    lost = sum((r["scrap"] * MULT - r["theirs"]) * r["qty"] for r in sel)
    ws["A2"] = (f"{len(sel)} "
                f"{plural(len(sel), 'позиция', 'позиции', 'позиций')}, "
                f"{rub(sum(r['qty'] for r in sel))} шт. "
                f"Чтобы выйти на договоренность от 25.10.2026, по ним надо "
                f"добрать {rub(lost)} руб. По возрастанию кратности: "
                f"сверху то, где отход от договоренности самый большой.")
    ws["A2"].font = NOTE
    cols = [("№", 5), ("Наименование", 46), ("Ду", 6), ("Шт", 6),
            ("Вес ед., кг", 9), ("Материал", 11), ("Его цена, руб", 12),
            ("Лом за ед., руб", 11), ("Кратность лому", 11),
            (f"Лом × {MULT:g}".replace(".", ",") + " за ед., руб", 12),
            ("Недобор на позиции, руб", 14), ("Наша цена, руб", 12)]
    start = 4
    for i, (h, _) in enumerate(cols, start=1):
        ws.cell(row=start, column=i, value=h)
    head(ws, start, len(cols))
    row = start + 1
    for r in sel:
        gap = (r["scrap"] * MULT - r["theirs"]) * r["qty"]
        row = put(ws, row, [
            r["n"], r["name"], r["du"] or "", r["qty"], r["kg"],
            r["material"], round(r["theirs"]), round(r["scrap"]), r["k"],
            round(r["scrap"] * MULT), round(gap),
            round(r["ours"]) if r["ours"] else "",
        ], wrap=(2,), center=(1, 3, 4, 5, 9), money=(7, 8, 10, 11, 12),
            ratio=(9,), fill=RED_FILL if r["k"] < 1 else BAD_FILL)
    row = put(ws, row, [
        "ИТОГО", f"{len(sel)} позиций", "", sum(r["qty"] for r in sel), "",
        "", round(sum(r["theirs"] * r["qty"] for r in sel)), "", "",
        round(sum(r["scrap"] * MULT * r["qty"] for r in sel)), round(lost),
        round(sum(r["ours"] * r["qty"] for r in sel)),
    ], center=(4,), money=(7, 10, 11, 12), font=BOLD, fill=GROUP_FILL)
    row += 1
    row = notes(ws, row, len(cols), [
        "Красным — где он опустился ниже самого лома. Такую позицию "
        "выгоднее сдать в приемку, чем продать ему: быстрее и без споров "
        "по состоянию.",
        "Договоренность от 25.10.2026: закупочная цена равна цене лома, "
        "умноженной на 2,5. По бронзе это 1 862,50 руб за килограмм — та "
        "самая ставка, по которой посчитана вся спецификация.",
        "Тяжелое он срезает сильнее всего: чем больше в позиции металла, "
        "тем дальше его цена от ставки. На мелочи он наоборот выше 2,5 "
        "лома, потому что там ставка давала копейки.",
        f"Недобор по этим {len(sel)} позициям — {rub(lost)} руб, а разрыв "
        f"по всему предложению меньше, "
        f"{rub(t['scrap_priced'] * MULT - t['s_his'])} руб: на мелочи он "
        f"платит выше договоренности и часть недобора перекрывает. "
        f"Торговаться имеет смысл по этому списку, а не по итогу.",
    ], height=42)
    ws.freeze_panes = f"A{start + 1}"
    widths(ws, [w for _, w in cols])
    landscape(ws)


def sheet_material(wb, rows):
    ws = wb.create_sheet("Материал под вопросом")
    ws["A1"] = "Материал, который покупатель поправил сам"
    ws["A1"].font = TITLE
    ws["A2"] = ("Владельцем это НЕ подтверждено, поэтому в расчет не идет. "
                "Сталь вместо бронзы меняет лом в 39 раз и цену продажи "
                "втрое — такое подтверждают, а не принимают на слово.")
    ws["A2"].font = NOTE
    cols = [("Чертеж", 18), ("Наименование", 46), ("Было", 12),
            ("Стало по его словам", 14), ("Откуда", 34),
            ("Шт", 6), ("Вес ед., кг", 10),
            ("Лом как бронза, руб", 13), ("Лом как сталь, руб", 13),
            ("Наша цена, руб", 13), ("Его цена, руб", 13)]
    start = 4
    for i, (h, _) in enumerate(cols, start=1):
        ws.cell(row=start, column=i, value=h)
    head(ws, start, len(cols))
    row = start + 1
    index = {norm_draw(r["draw"]): r for r in rows}
    for c in buyer_prices.CLAIMED_STEEL:
        r = index.get(norm_draw(c["draw"]), {})
        kg = r.get("kg", 0)
        row = put(ws, row, [
            c["draw"], c["name"], c["was"], c["now"], c["how"],
            r.get("qty", ""), kg,
            round(kg * scrap_prices.SCRAP["бронза"] * r.get("qty", 1)),
            round(kg * scrap_prices.SCRAP["сталь"] * r.get("qty", 1)),
            round(r.get("ours", 0) * r.get("qty", 1)) or "",
            round((r.get("theirs") or 0) * r.get("qty", 1)) or "",
        ], wrap=(2, 5), center=(3, 4, 6, 7), money=(8, 9, 10, 11),
            fill=MARK_FILL)
    row += 1
    row = notes(ws, row, len(cols), [
        "Подтвержденные владельцем материалы лежат в materials.py и бьют "
        "определение по названию. Эти две позиции туда не внесены: "
        "источник — покупатель, а он заинтересованная сторона.",
        "Если сталь подтвердится, по этим позициям ставка на бронзу не "
        "применяется вовсе, и цену надо согласовывать отдельно — как по "
        "всей стали, титану, алюминию и нержавейке.",
        "Спросить у владельца по чертежу: 427-03.192-2 (фильтр забортной "
        "воды Ду100) и 581-03.034 (клапан паровой Ду65Ру10).",
    ], height=42)
    widths(ws, [w for _, w in cols])
    landscape(ws)


def sheet_groups(wb, rows):
    ws = wb.create_sheet("По группам")
    ws["A1"] = "Сверка по товарным группам"
    ws["A1"].font = TITLE
    ws["A2"] = ("Видно, где он срезал, а где добавил: группа к группе. "
                "Считаются только позиции, которые он оценил.")
    ws["A2"].font = NOTE
    agg = defaultdict(lambda: [0, 0, 0.0, 0.0, 0.0, 0.0])
    for r in rows:
        if not r["theirs"]:
            continue
        a = agg[r["group"]]
        a[0] += 1
        a[1] += r["qty"]
        a[2] += r["ours"] * r["qty"]
        a[3] += r["theirs"] * r["qty"]
        a[4] += r["scrap"] * r["qty"]
        a[5] += r["rate"] * r["qty"]
    cols = [("Группа", 34), ("Позиций", 9), ("Штук", 9),
            ("Наша цена, руб", 15), ("Его цена, руб", 15),
            ("Разница, руб", 15), ("Лом, руб", 14),
            ("Кратность лому", 11), ("По ставке, руб", 15)]
    start = 4
    for i, (h, _) in enumerate(cols, start=1):
        ws.cell(row=start, column=i, value=h)
    head(ws, start, len(cols))
    row = start + 1
    for g, a in sorted(agg.items(), key=lambda kv: kv[1][3] - kv[1][2]):
        k = a[3] / a[4] if a[4] else 0
        row = put(ws, row, [
            g, a[0], a[1], round(a[2]), round(a[3]), round(a[3] - a[2]),
            round(a[4]), k or "", round(a[5]),
        ], wrap=(1,), center=(2, 3, 8), money=(4, 5, 6, 7, 9), ratio=(8,),
            fill=BAD_FILL if k and k < MULT else OK_FILL)
    tot = [sum(a[i] for a in agg.values()) for i in range(6)]
    row = put(ws, row, [
        "ИТОГО", tot[0], tot[1], round(tot[2]), round(tot[3]),
        round(tot[3] - tot[2]), round(tot[4]),
        tot[3] / tot[4] if tot[4] else "", round(tot[5]),
    ], center=(2, 3, 8), money=(4, 5, 6, 7, 9), ratio=(8,), font=BOLD,
        fill=GROUP_FILL)
    row += 1
    row = notes(ws, row, len(cols), [
        "Оранжевым — группы, где его цена ниже 2,5 лома, то есть ниже "
        "договоренности. Зеленым — где он ее выдерживает.",
        "Наша цена по группе считается только по тем позициям, которые он "
        "оценил, иначе сравнение было бы нечестным: у нас по многим "
        "позициям стояло «по согласованию».",
    ], height=32)
    ws.freeze_panes = f"A{start + 1}"
    widths(ws, [w for _, w in cols])
    landscape(ws)


def main():
    rows = build()
    t = totals(rows)

    wb = Workbook()
    sheet_summary(wb, rows, t)
    sheet_groups(wb, rows)
    sheet_rows(wb, rows)
    sheet_cuts(wb, rows)
    sheet_closed(wb, rows)
    sheet_floor(wb, rows, t)
    sheet_material(wb, rows)

    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / "Сверка_цен_покупателя.xlsx"
    wb.save(path)

    print(f"готово: {path}\n")
    print(f"позиций {len(rows)}, он оценил {len(t['priced'])}")
    print(f"сравнимых {len(t['both'])}: наша {rub(t['s_both_our'])} -> "
          f"его {rub(t['s_both_his'])} "
          f"({t['s_both_his'] / t['s_both_our']:.0%}), "
          f"минус {rub(t['s_both_our'] - t['s_both_his'])}")
    print(f"закрыл наших пустых {len(t['closed'])}: "
          f"{rub(t['s_closed'])} руб")
    print(f"всего его предложение {rub(t['s_his'])} против нашего закупа "
          f"{rub(t['s_our'])}")
    print(f"\nлом по его позициям {rub(t['scrap_priced'])}, "
          f"по договоренности лом x{MULT:g} = "
          f"{rub(t['scrap_priced'] * MULT)}")
    print(f"его предложение это {t['s_his'] / t['scrap_priced']:.2f} лома, "
          f"недобор {rub(t['scrap_priced'] * MULT - t['s_his'])} руб "
          f"({(t['scrap_priced'] * MULT - t['s_his']) / (t['scrap_priced'] * MULT):.0%})")
    low = [r for r in rows if r["theirs"] and r["k"] and r["k"] < MULT]
    print(f"ниже договоренности {len(low)} позиций, "
          f"{sum(r['qty'] for r in low)} шт, недобор "
          f"{rub(sum((r['scrap'] * MULT - r['theirs']) * r['qty'] for r in low))}")
    eq = [r for r in rows if r["eq_old"]]
    cmp_old = [r for r in rows if r["theirs"] and r["old"]]
    print(f"совпало со старой ценой договора {len(eq)} из {len(cmp_old)}")


if __name__ == "__main__":
    main()
