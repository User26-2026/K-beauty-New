#!/usr/bin/env python3
"""Сдать в лом или продавать по позициям: сравнение в деньгах.

Лом — это не запасной вариант, а полноценный сценарий: деньги за
несколько дней, без инвентаризации, без контейнера, без переговоров и
без отсрочки платежа. Поэтому его надо посчитать честно и положить
рядом с продажей, а не держать в уме как «ну в крайнем случае».

Считаем три сценария: сдать все, продать все и смешанный — продать то,
что заметно дороже лома, и сдать то, что около него.

Цены приемки на 13.10.2026 из scrap_prices.py, наши продажные — из
договора с ПКФ «АРМА», по диапазону берем верхнюю границу.

Результат: outputs/marine_equipment/Сравнение_лом_и_продажа.xlsx
"""

import pathlib
import sys
from collections import defaultdict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
from build_competitor_compare import our_price  # noqa: E402
from build_competitor_valuation import competitor_index, pick  # noqa: E402
from build_container_spb import asrow, plural  # noqa: E402
from build_logistics_compare import (  # noqa: E402
    AUTO_SPB, AUTO_VL, HOURS_BOX, PRICE, loading, rows_all,
)
from build_shipment_spb import material  # noqa: E402
from build_spb_list_xlsx import GROUPS, group_of  # noqa: E402
import price_fixes  # noqa: E402
import scrap_prices as scrap  # noqa: E402

OUT = pathlib.Path(__file__).resolve().parents[2] / "outputs" / "marine_equipment"

FONT = "Arial"
HEAD_FILL = PatternFill("solid", fgColor="1F3864")
HEAD_FONT = Font(name=FONT, size=11, bold=True, color="FFFFFF")
TITLE = Font(name=FONT, size=14, bold=True, color="1F3864")
SUB = Font(name=FONT, size=11, bold=True, color="1F3864")
BOLD = Font(name=FONT, size=10, bold=True)
BASE = Font(name=FONT, size=10)
NOTE = Font(name=FONT, size=9, italic=True, color="666666")
SELL = PatternFill("solid", fgColor="EAF3EA")
SCRAP_FILL = PatternFill("solid", fgColor="FCE4D6")
MIX = PatternFill("solid", fgColor="FFF2CC")
GROUP_FILL = PatternFill("solid", fgColor="D9E2F3")
INPUT_FILL = PatternFill("solid", fgColor="FFFF00")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# ниже этого отношения к лому возить через всю страну незачем
MIX_LIMIT = 1.5

# доставка арматуры до приемки во Владивостоке: своя машина по городу
SCRAP_HAUL = 30_000


def load():
    """Перечень с тремя оценками: лом, наша цена, цена конкурента."""
    by_draw, by_type, _ = competitor_index()
    rows = []
    for r in rows_all():
        m = material(asrow(r))
        price = r["Сумма"] / r["Штук"] if r["Сумма"] and r["Штук"] else 0
        if not price:
            lo, hi, flag = our_price(r.get("Цена"))
            fixed = price_fixes.fix(r["Наименование"])
            price = ((fixed or lo) if flag else hi) or 0
        p, how = pick(by_draw, by_type, r["Чертеж"],
                      group_of(r["Наименование"]), r["Ду"], m)
        rows.append({
            "Наименование": r["Наименование"],
            "Чертеж": r["Чертеж"],
            "Ду": r["Ду"],
            "Материал": m,
            "Штук": r["Штук"],
            "вес": r["вес"],
            "Группа": group_of(r["Наименование"]),
            "Лом": scrap.value(m, r["вес"]) or 0,
            "Наша": price * r["Штук"],
            "Конкурент": (p or 0) * r["Штук"],
        })
    return rows


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


def split(rows):
    """Что продаем, что сдаем в смешанном сценарии."""
    sell, give = [], []
    for r in rows:
        if not r["Наша"] or (r["Лом"] and r["Наша"] / r["Лом"] < MIX_LIMIT):
            give.append(r)
        else:
            sell.append(r)
    return sell, give


def sheet_summary(wb, rows):
    ws = wb.create_sheet("Сравнение", 0)
    kg = sum(r["вес"] for r in rows)
    lom = sum(r["Лом"] for r in rows)
    our = sum(r["Наша"] for r in rows)
    sell, give = split(rows)
    sell_kg = sum(r["вес"] for r in sell)
    give_kg = sum(r["вес"] for r in give)
    ship = (PRICE["40"] + AUTO_VL[40] + AUTO_SPB[40] + loading(HOURS_BOX))

    ws["A1"] = "Сдать в лом или продавать: сравнение в деньгах"
    ws["A1"].font = TITLE
    ws["A2"] = (f"Весь перечень на отгрузку — {len(rows)} "
                f"{plural(len(rows), 'позиция', 'позиции', 'позиций')}, "
                f"{kg / 1000:.2f} т. Лом считаем по бронзе "
                f"{scrap.SCRAP['бронза']} руб за кг: по ней же "
                "собственник назначил закупочную цену, лом на 2,5.")
    ws["A2"].font = NOTE

    row = 4
    ws.cell(row=row, column=1, value="Ставки, с которыми считаем").font = SUB
    row += 1
    for label, value, note in [
        ("Бронза, руб за кг", scrap.SCRAP["бронза"],
         "по ней и считаем, класс «бронза/латунь» идет сюда"),
        ("Латунь, руб за кг", scrap.SCRAP["латунь"],
         "нижняя граница, если материал разделить"),
        ("Черный металл, руб за кг", scrap.SCRAP["сталь"], ""),
        ("Закупочная цена, лом на", scrap.MULTIPLIER,
         "договоренность с Антоном от 25.10.2026: по бронзе "
         f"{scrap.purchase_rate('бронза'):,.2f} руб за кг".replace(",", " ")),
        ("Довезти до приемки во Владивостоке", SCRAP_HAUL,
         "своя машина по городу, несколько рейсов"),
        ("Отгрузка в Петербург, сорокафутовый", ship,
         "контейнер, авто на обоих концах, погрузка"),
    ]:
        ws.cell(row=row, column=1, value=label).font = BASE
        cell = ws.cell(row=row, column=2, value=value)
        cell.font = BOLD
        cell.border = BORDER
        cell.fill = INPUT_FILL
        cell.number_format = "#,##0"
        ws.cell(row=row, column=3, value=note).font = NOTE
        row += 1

    row += 1
    ws.cell(row=row, column=1, value="Три сценария").font = SUB
    row += 1
    cols = ["Сценарий", "Что уходит", "Вес, т", "Выручка, руб",
            "Расходы, руб", "На руки, руб", "Руб за кг", "Срок",
            "Когда деньги"]
    for i, h in enumerate(cols, start=1):
        ws.cell(row=row, column=i, value=h)
    head(ws, row, len(cols), height=32)

    mix_rev = sum(r["Наша"] for r in sell) + sum(r["Лом"] for r in give)
    rows_data = [
        ("Сдать все в лом", f"{len(rows)} позиций целиком", kg, lom,
         SCRAP_HAUL, lom - SCRAP_HAUL, "несколько дней",
         "сразу, по факту сдачи", SCRAP_FILL),
        ("Продать все по позициям",
         f"{len(rows)} позиций в Петербург", kg, our, ship, our - ship,
         "месяцы", "по мере реализации, отсрочка по договору", SELL),
        ("Смешанный: продать дорогое, сдать остальное",
         f"продать {len(sell)}, сдать {len(give)}", kg, mix_rev,
         ship + SCRAP_HAUL, mix_rev - ship - SCRAP_HAUL,
         "месяцы, но груза меньше", "часть сразу, часть по реализации",
         MIX),
    ]
    for name, what, w, rev, exp, net, term, when, fill in rows_data:
        row += 1
        values = [name, what, round(w / 1000, 2), round(rev), round(exp),
                  round(net), round(net / w) if w else 0, term, when]
        for i, v in enumerate(values, start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BOLD if "лом" in name.lower() else BASE
            cell.border = BORDER
            cell.fill = fill
            cell.alignment = Alignment(vertical="top", wrap_text=(i in (2, 9)))
            if i == 3:
                cell.number_format = "#,##0.00"
            if i in (4, 5, 6, 7):
                cell.number_format = "#,##0"
        ws.row_dimensions[row].height = 30

    row += 2
    ws.cell(row=row, column=1, value="Что показывает сравнение").font = SUB
    row += 1
    diff = (our - ship) - (lom - SCRAP_HAUL)
    for text in [
        f"Разница между продажей и ломом — {diff / 1e6:.1f} млн руб. "
        f"Лом дает {(lom - SCRAP_HAUL) / 1e6:.1f} млн за несколько дней, "
        f"продажа {(our - ship) / 1e6:.1f} млн за месяцы и с отсрочкой "
        "платежа. Это и есть цена нашей работы: инвентаризация, "
        "контейнер, хранение, переговоры и ожидание.",
        f"В пересчете на килограмм: лом дает "
        f"{(lom - SCRAP_HAUL) / kg:.0f} руб, продажа "
        f"{(our - ship) / kg:.0f} руб. Разница "
        f"{(our - lom) / kg:.0f} руб с килограмма.",
        f"Доставка в Петербург на этом фоне мелочь: {ship:,.0f} руб — "
        f"{ship / our * 100:.1f}% выручки от продажи. Экономить на "
        "контейнере бессмысленно, а вот сдать в лом вместо продажи — "
        f"потерять {(our - ship - lom + SCRAP_HAUL) / (our - ship):.0%} "
        "денег.".replace(",", " "),
        f"Смешанный сценарий почти не отличается от продажи: позиций "
        f"около лома всего {len(give)}, и весят они "
        f"{give_kg:,.0f} кг. Отдельно возиться с ними смысла нет, "
        "но и везти через всю страну незачем.".replace(",", " "),
        f"Лом считаем по бронзе {scrap.SCRAP['бронза']} руб за кг. Если "
        f"часть окажется латунью по {scrap.SCRAP['латунь']}, лом "
        f"просядет до "
        f"{sum(r['вес'] for r in rows if r['Материал'] == 'бронза/латунь') * scrap.SCRAP['латунь'] / 1e6:.1f}"
        " млн. Это еще один довод проставить материал в учете.",
        f"Цена по договору — это тот же лом, умноженный на "
        f"{scrap.MULTIPLIER}: бронза {scrap.SCRAP['бронза']} на "
        f"{scrap.MULTIPLIER} дает "
        f"{scrap.purchase_rate('бронза'):,.2f} руб за кг. То есть "
        "продажа по спецификации и есть лом с коэффициентом, и запас "
        "над приемкой зашит в саму цену.".replace(",", " "),
    ]:
        cell = ws.cell(row=row, column=1, value="— " + text)
        cell.font = BASE
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[row].height = 44
        row += 1

    widths(ws, [44, 30, 11, 16, 14, 16, 12, 22, 34])
    return ws


def sheet_groups(wb, rows):
    ws = wb.create_sheet("По группам")
    ws["A1"] = "Лом против продажи по группам арматуры"
    ws["A1"].font = TITLE
    ws["A2"] = ("Колонка «во сколько раз дороже лома» показывает, где "
                "продажа действительно имеет смысл, а где разница "
                "незначительна.")
    ws["A2"].font = NOTE

    row = 4
    cols = ["Группа", "Позиций", "Вес, кг", "В ломе, руб",
            "По нашей цене, руб", "Во сколько раз дороже лома",
            "Лом, руб за кг", "Наша цена, руб за кг", "Разница на кг, руб"]
    for i, h in enumerate(cols, start=1):
        ws.cell(row=row, column=i, value=h)
    head(ws, row, len(cols), height=44)

    by_group = defaultdict(list)
    for r in rows:
        by_group[r["Группа"]].append(r)
    order = sorted(by_group, key=lambda g: -sum(x["Наша"] for x in by_group[g]))
    for g in order:
        row += 1
        items = by_group[g]
        kg = sum(r["вес"] for r in items)
        lom = sum(r["Лом"] for r in items)
        our = sum(r["Наша"] for r in items)
        values = [g, len(items), round(kg), round(lom), round(our),
                  round(our / lom, 1) if lom else None,
                  round(lom / kg) if kg else None,
                  round(our / kg) if kg else None,
                  round((our - lom) / kg) if kg else None]
        for i, v in enumerate(values, start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            if i in (3, 4, 5, 7, 8, 9):
                cell.number_format = "#,##0"
            if i == 6:
                cell.number_format = "#,##0.0"
                if isinstance(v, (int, float)):
                    cell.fill = SELL if v >= 2 else SCRAP_FILL
    widths(ws, [34, 10, 12, 16, 18, 16, 14, 16, 16])
    return ws


def sheet_give(wb, rows):
    ws = wb.create_sheet("Кандидаты в лом")
    sell, give = split(rows)
    ws["A1"] = "Позиции, которые проще сдать, чем везти"
    ws["A1"].font = TITLE
    ws["A2"] = (f"{len(give)} позиций, {sum(r['вес'] for r in give):,.0f} кг. "
                f"Сюда попадает то, где наша цена ниже {MIX_LIMIT} лома, и "
                "то, где цены нет вообще. Везти такое через всю страну и "
                "потом хранить в Петербурге смысла мало."
                .replace(",", " "))
    ws["A2"].font = NOTE

    row = 4
    cols = ["№", "Наименование", "Чертеж", "Ду", "Материал", "Штук",
            "Вес, кг", "Лом, руб", "Наша цена, руб", "Отношение к лому",
            "Почему здесь"]
    for i, h in enumerate(cols, start=1):
        ws.cell(row=row, column=i, value=h)
    head(ws, row, len(cols), height=32)

    give.sort(key=lambda r: -r["вес"])
    for k, r in enumerate(give, start=1):
        row += 1
        ratio = r["Наша"] / r["Лом"] if r["Наша"] and r["Лом"] else None
        why = ("цены нет" if not r["Наша"]
               else "цена ниже лома" if ratio and ratio < 1
               else "цена почти как лом")
        values = [k, r["Наименование"], r["Чертеж"], r["Ду"], r["Материал"],
                  r["Штук"], round(r["вес"]), round(r["Лом"]),
                  round(r["Наша"]) or None, round(ratio, 2) if ratio else None,
                  why]
        for i, v in enumerate(values, start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=(i == 2))
            if i in (6, 7, 8, 9):
                cell.number_format = "#,##0"
            if i == 10:
                cell.number_format = "#,##0.00"
                if isinstance(v, (int, float)) and v < 1:
                    cell.fill = SCRAP_FILL
    widths(ws, [5, 50, 20, 7, 16, 8, 11, 14, 16, 16, 22])
    return ws


def main():
    rows = load()
    wb = Workbook()
    wb.remove(wb.active)
    sheet_summary(wb, rows)
    sheet_groups(wb, rows)
    sheet_give(wb, rows)

    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / "Сравнение_лом_и_продажа.xlsx"
    wb.save(path)

    kg = sum(r["вес"] for r in rows)
    lom = sum(r["Лом"] for r in rows)
    our = sum(r["Наша"] for r in rows)
    sell, give = split(rows)
    ship = PRICE["40"] + AUTO_VL[40] + AUTO_SPB[40] + loading(HOURS_BOX)
    print(f"готово: {path}\n")
    print(f"{len(rows)} позиций, {kg / 1000:.2f} т\n")
    print(f"{'сценарий':<44}{'выручка':>12}{'расходы':>10}{'на руки':>12}"
          f"{'руб/кг':>9}")
    for name, rev, exp in [
        ("Сдать все в лом", lom, SCRAP_HAUL),
        ("Продать все по позициям", our, ship),
        ("Смешанный",
         sum(r["Наша"] for r in sell) + sum(r["Лом"] for r in give),
         ship + SCRAP_HAUL),
    ]:
        net = rev - exp
        print(f"{name:<44}{rev:>12,.0f}{exp:>10,.0f}{net:>12,.0f}"
              f"{net / kg:>9,.0f}".replace(",", " "))
    print(f"\nкандидатов в лом: {len(give)} позиций, "
          f"{sum(r['вес'] for r in give):,.0f} кг".replace(",", " "))


if __name__ == "__main__":
    main()
