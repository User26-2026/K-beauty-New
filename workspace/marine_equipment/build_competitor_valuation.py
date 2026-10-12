#!/usr/bin/env python3
"""Сколько стоит наша арматура по ценам конкурента с Камчатки.

Прайс конкурента от 03.10.2026 — это рынок Дальнего Востока, тот самый,
где мы продаем. Поэтому его цены годятся как оценка нашего склада:
не «сколько записано в учете», а «почем такое же стоит у соседа».

Цену подбираем в два прохода. Сначала по номеру чертежа: один чертеж —
одно и то же изделие, это надежно. Потом по группе, диаметру и
материалу: так закрываются позиции без чертежа. Бронзу со сталью не
смешиваем, у конкурента стальная задвижка втрое дешевле бронзовой.

Что не удалось сопоставить напрямую, оцениваем по цене килограмма
своей же группы — это грубее, и в файле такие строки отделены.

Отдельно считаем, какую долю от стоимости съедает доставка в Петербург.
Расходы берем из build_logistics_compare: контейнер, авто на обоих
концах и погрузка.

Результат: outputs/marine_equipment/Оценка_по_конкуренту.xlsx
"""

import pathlib
import re
import statistics
import sys
from collections import defaultdict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
from build_catalog_xlsx import group_of as catalog_group  # noqa: E402
from build_competitor_compare import (  # noqa: E402
    norm_draw, our_price, read_competitor,
)
from build_container_spb import (  # noqa: E402
    asrow, plural, real_weights, weight_of,
)
from build_logistics_compare import (  # noqa: E402
    AUTO_SPB, AUTO_VL, HOURS_BOX, PRICE, loading, rows_all,
)
from build_marine_registry import SOURCES, parse  # noqa: E402
from build_shipment_spb import FILTERS, material  # noqa: E402
from build_spb_list_xlsx import GROUPS, group_of, read as read_spb  # noqa: E402
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
HARD = PatternFill("solid", fgColor="EAF3EA")    # цена по чертежу
SOFT = PatternFill("solid", fgColor="FFF2CC")    # цена по типу
GUESS = PatternFill("solid", fgColor="FDF2E3")   # оценка по килограмму
NONE_FILL = PatternFill("solid", fgColor="FCE4D6")
GROUP_FILL = PatternFill("solid", fgColor="D9E2F3")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

BY_DRAW = "по чертежу"
BY_TYPE = "по типу и диаметру"
BY_KG = "оценка по цене килограмма"
NO_PRICE = "аналога нет"


def competitor_index():
    """Три индекса цен конкурента: по чертежу, по группе с диаметром."""
    rows = [r for r in read_competitor() if r["Цена"]]
    by_draw = defaultdict(list)
    by_type = defaultdict(list)
    for r in rows:
        mat = "сталь" if r["Материал"] == "сталь" else "бронза"
        if r["Чертеж"]:
            by_draw[r["Чертеж"]].append((r["Цена"], mat))
        if r["Группа"] and r["Ду"]:
            by_type[(r["Группа"], r["Ду"], mat)].append(r["Цена"])
    return by_draw, by_type, rows


def pick(by_draw, by_type, draw, group, d, mat):
    """Цена за штуку по конкуренту и то, откуда она взята."""
    mat = "сталь" if mat == "сталь" else "бронза"
    key = norm_draw(draw)
    if key and key in by_draw:
        same = [p for p, m in by_draw[key] if m == mat]
        vals = same or [p for p, _ in by_draw[key]]
        return statistics.median(vals), BY_DRAW
    if group and d and (group, d, mat) in by_type:
        return statistics.median(by_type[(group, d, mat)]), BY_TYPE
    return None, NO_PRICE


def our_stock():
    """Вся арматура и фильтры в наличии.

    Берем тот же свод, что и планы отгрузки: согласованный список
    менеджера плюс склад во Владивостоке, с вычетом задвоений по
    сочетанию «тип плюс диаметр». Иначе позиции из списка менеджера
    считались бы дважды или терялись, и склад с отгрузкой не сходились
    бы между собой.
    """
    out = []
    for r in rows_all():
        price = r["Сумма"] / r["Штук"] if r["Сумма"] and r["Штук"] else 0
        if not price:
            # у списка менеджера цена записана текстом и диапазоном
            lo, hi, _ = our_price(r.get("Цена"))
            price = (lo + hi) / 2 if lo and hi else 0
        out.append({
            "Наименование": r["Наименование"],
            "Чертеж": r.get("Чертеж") or "",
            "Ду": r.get("Ду") or du(r["Наименование"]),
            "Группа": group_of(r["Наименование"]),
            "Материал": material(asrow(r)),
            "Штук": r["Штук"],
            "вес_ед": r["вес_ед"],
            "вес": r["вес"],
            "Наша цена": price,
            "Локация": r["Откуда"],
            "Блок": r["Блок"],
        })
    return out


def value(items, by_draw, by_type):
    """Проставить цены конкурента, потом добить оценкой по килограмму."""
    for r in items:
        p, how = pick(by_draw, by_type, r["Чертеж"], r["Группа"], r["Ду"],
                      r["Материал"])
        r["Цена конкурента"] = p
        r["Источник цены"] = how
        r["Сумма конкурента"] = (p or 0) * r["Штук"]

    # цена килограмма по каждой группе, считаем только по сопоставленному
    per_kg = {}
    for g in {r["Группа"] for r in items}:
        got = [r for r in items
               if r["Группа"] == g and r["Цена конкурента"] and r["вес"]]
        if got:
            per_kg[g] = (sum(r["Сумма конкурента"] for r in got)
                         / sum(r["вес"] for r in got))
    all_kg = (sum(r["Сумма конкурента"] for r in items if r["Цена конкурента"])
              / sum(r["вес"] for r in items if r["Цена конкурента"] and r["вес"]))

    for r in items:
        if r["Цена конкурента"] or not r["вес_ед"]:
            continue
        rate = per_kg.get(r["Группа"], all_kg)
        r["Цена конкурента"] = rate * r["вес_ед"]
        r["Источник цены"] = BY_KG
        r["Сумма конкурента"] = r["Цена конкурента"] * r["Штук"]
    return per_kg, all_kg


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


def money(cell):
    cell.number_format = "#,##0"


def sheet_summary(wb, items, ship, per_kg, all_kg, comp):
    ws = wb.create_sheet("Сводка", 0)
    ws["A1"] = "Наша арматура в деньгах по прайсу конкурента с Камчатки"
    ws["A1"].font = TITLE
    ws["A2"] = ("Прайс от 03.10.2026, тот же рынок, где мы продаем. "
                "Цена подбирается по номеру чертежа, где его нет — по "
                "группе, диаметру и материалу. Остальное оценено по цене "
                "килограмма своей группы.")
    ws["A2"].font = NOTE

    hard = [r for r in items if r["Источник цены"] == BY_DRAW]
    soft = [r for r in items if r["Источник цены"] == BY_TYPE]
    guess = [r for r in items if r["Источник цены"] == BY_KG]
    none = [r for r in items if r["Источник цены"] == NO_PRICE]
    total = sum(r["Сумма конкурента"] for r in items)
    ours = sum(r["Наша цена"] * r["Штук"] for r in items)

    row = 4
    ws.cell(row=row, column=1, value="Весь склад, арматура и фильтры").font = SUB
    row += 1
    for i, h in enumerate(["Откуда цена", "Позиций", "Штук", "Вес, кг",
                           "По конкуренту, руб", "Доля суммы"], start=1):
        ws.cell(row=row, column=i, value=h)
    head(ws, row, 6, height=30)
    for label, rows_, fill in [
        (BY_DRAW, hard, HARD), (BY_TYPE, soft, SOFT),
        (BY_KG, guess, GUESS), (NO_PRICE, none, NONE_FILL),
    ]:
        row += 1
        s = sum(r["Сумма конкурента"] for r in rows_)
        for i, v in enumerate([label, len(rows_),
                               sum(r["Штук"] for r in rows_),
                               round(sum(r["вес"] for r in rows_)), s,
                               s / total if total else 0], start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            cell.fill = fill
            if i in (3, 4, 5):
                money(cell)
            if i == 6:
                cell.number_format = "0%"
    row += 1
    for i, v in enumerate(["ИТОГО", len(items), sum(r["Штук"] for r in items),
                           round(sum(r["вес"] for r in items)), total, 1],
                          start=1):
        cell = ws.cell(row=row, column=i, value=v)
        cell.font = BOLD
        cell.border = BORDER
        if i in (3, 4, 5):
            money(cell)
        if i == 6:
            cell.number_format = "0%"

    row += 2
    ws.cell(row=row, column=1, value="Сравнение с нашим учетом").font = SUB
    row += 1
    for label, v, note in [
        ("По нашим учетным ценам, руб", round(ours), "то, что стоит в файлах"),
        ("По ценам конкурента, руб", round(total), ""),
        ("Разница, руб", round(total - ours), ""),
        ("Отношение", round(total / ours, 2) if ours else 0,
         "больше единицы — конкурент дороже нас"),
        ("Средняя цена килограмма по конкуренту, руб", round(all_kg),
         "по всей сопоставленной арматуре"),
        ("Позиций без согласованной цены у нас",
         sum(1 for r in items if not r["Наша цена"]),
         "вес у них есть, а рублей в учете нет — наша сумма занижена "
         "именно на них"),
    ]:
        ws.cell(row=row, column=1, value=label).font = BASE
        cell = ws.cell(row=row, column=2, value=v)
        cell.font = BOLD
        cell.border = BORDER
        cell.number_format = "#,##0.00" if label == "Отношение" else "#,##0"
        ws.cell(row=row, column=3, value=note).font = NOTE
        row += 1

    row += 1
    ws.cell(row=row, column=1,
            value="Отгрузка в Петербург: сколько съедает логистика").font = SUB
    row += 1
    ship_val = sum(r["Сумма конкурента"] for r in ship)
    ship_kg = sum(r["вес"] for r in ship)
    variants = [
        ("Усиленный двадцатифутовый, 27 т",
         PRICE["20 усиленный"] + AUTO_VL[20] + AUTO_SPB[20]
         + loading(HOURS_BOX)),
        ("Два двадцатифутовых по 22 т",
         2 * (PRICE["20 обычный"] + AUTO_VL[20] + AUTO_SPB[20])
         + loading(2 * HOURS_BOX)),
        ("Один сорокафутовый",
         PRICE["40"] + AUTO_VL[40] + AUTO_SPB[40] + loading(HOURS_BOX)),
    ]
    for i, h in enumerate(["Вариант", "Расходы на доставку, руб",
                           "Везем товара на, руб", "Доля логистики",
                           "Руб на килограмм товара"], start=1):
        ws.cell(row=row, column=i, value=h)
    head(ws, row, 5, height=30)
    for name, cost in variants:
        row += 1
        for i, v in enumerate([name, cost, round(ship_val),
                               cost / ship_val if ship_val else 0,
                               round(cost / ship_kg, 1) if ship_kg else 0],
                              start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BOLD if "усиленный" in name.lower() else BASE
            cell.border = BORDER
            if i in (2, 3):
                money(cell)
            if i == 4:
                cell.number_format = "0.0%"
            if i == 5:
                cell.number_format = "#,##0.0"
            if "усиленный" in name.lower():
                cell.fill = HARD
    row += 1
    cell = ws.cell(row=row, column=1,
                   value=f"За базу взят весь перечень — {ship_kg / 1000:.2f} т "
                         f"на {ship_val / 1e6:.1f} млн. В усиленный контейнер "
                         "входит 26,35 т из них, на долю логистики эта "
                         "разница не влияет.")
    cell.font = NOTE

    row += 2
    ws.cell(row=row, column=1, value="Что из этого следует").font = SUB
    row += 1
    best = min(c for _, c in variants)
    for text in [
        f"Вся арматура и фильтры на складе по ценам конкурента — "
        f"{total / 1e6:.1f} млн руб при весе "
        f"{sum(r['вес'] for r in items) / 1000:.1f} т. По нашим учетным "
        f"ценам те же позиции стоят {ours / 1e6:.1f} млн, но у "
        f"{sum(1 for r in items if not r['Наша цена'])} позиций цена в "
        "договоре записана как «по согласованию», и в нашу сумму они "
        "входят нулем. Настоящий разрыв меньше.",
        f"Доставка в Петербург стоит от {best:,.0f} руб и съедает "
        f"{best / ship_val * 100:.1f}% стоимости отгружаемого товара. "
        "Для товара, который едет через всю страну, это немного: "
        "транспортная компания на том же весе взяла бы в несколько раз "
        "больше.".replace(",", " "),
        f"Килограмм арматуры по конкуренту стоит в среднем "
        f"{all_kg:,.0f} руб, а довезти его до Петербурга стоит "
        f"{best / ship_kg:.1f} руб. Отсюда и доля."
        .replace(",", " "),
        f"Твердая часть оценки — {len(hard) + len(soft)} "
        f"{plural(len(hard) + len(soft), 'позиция', 'позиции', 'позиций')} "
        f"на {sum(r['Сумма конкурента'] for r in hard + soft) / 1e6:.1f} млн, "
        "это цены по чертежу и по типу. Остальное оценено по цене "
        "килограмма и годится для прикидки, а не для переговоров.",
        "Цены конкурента — это его прайс, а не сделки. Реальная цена "
        "продажи обычно ниже прайса, и до инвентаризации эти суммы "
        "покупателю не называем.",
    ]:
        cell = ws.cell(row=row, column=1, value="— " + text)
        cell.font = BASE
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[row].height = 32
        row += 1

    widths(ws, [46, 22, 20, 18, 22, 16])
    return ws


def sheet_groups(wb, items, per_kg):
    ws = wb.create_sheet("По группам")
    ws["A1"] = "По группам арматуры"
    ws["A1"].font = TITLE
    row = 3
    cols = ["Группа", "Позиций", "Штук", "Вес, кг", "По конкуренту, руб",
            "По нашему учету, руб", "Отношение", "Руб за кг у конкурента",
            "Цена по чертежу, позиций"]
    for i, h in enumerate(cols, start=1):
        ws.cell(row=row, column=i, value=h)
    head(ws, row, len(cols), height=40)

    by_group = defaultdict(list)
    for r in items:
        by_group[r["Группа"]].append(r)
    order = [g for g in GROUPS if g in by_group]
    order += [g for g in sorted(by_group) if g not in GROUPS]
    for g in order:
        row += 1
        rs = by_group[g]
        comp = sum(r["Сумма конкурента"] for r in rs)
        ours = sum(r["Наша цена"] * r["Штук"] for r in rs)
        kg = sum(r["вес"] for r in rs)
        values = [g, len(rs), sum(r["Штук"] for r in rs), round(kg),
                  round(comp), round(ours),
                  round(comp / ours, 2) if ours else None,
                  round(comp / kg) if kg else None,
                  sum(1 for r in rs if r["Источник цены"] == BY_DRAW)]
        for i, v in enumerate(values, start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            if i in (3, 4, 5, 6, 8):
                money(cell)
            if i == 7:
                cell.number_format = "#,##0.00"
    widths(ws, [34, 12, 10, 12, 20, 20, 12, 20, 20])
    return ws


COLS = ["№", "Наименование", "Чертеж", "Ду", "Группа", "Материал", "Штук",
        "Вес ед., кг", "Цена конкурента за шт", "Сумма по конкуренту",
        "Наша цена за шт", "Наша сумма", "Отношение", "Откуда цена",
        "Локация"]


def sheet_items(wb, name, title, items, note):
    ws = wb.create_sheet(name)
    ws["A1"] = title
    ws["A1"].font = TITLE
    ws["A2"] = note
    ws["A2"].font = NOTE
    start = 4
    for i, h in enumerate(COLS, start=1):
        ws.cell(row=start, column=i, value=h)
    head(ws, start, len(COLS))

    rows = sorted(items, key=lambda r: -r["Сумма конкурента"])
    for k, r in enumerate(rows, start=1):
        row = start + k
        ours = r["Наша цена"] * r["Штук"]
        values = [k, r["Наименование"], r["Чертеж"], r["Ду"], r["Группа"],
                  r["Материал"], r["Штук"], round(r["вес_ед"], 1),
                  round(r["Цена конкурента"]) if r["Цена конкурента"] else None,
                  round(r["Сумма конкурента"]) or None,
                  r["Наша цена"] or None, ours or None,
                  round(r["Сумма конкурента"] / ours, 2) if ours else None,
                  r["Источник цены"], r["Локация"]]
        for i, v in enumerate(values, start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=(i == 2))
            if i in (7, 9, 10, 11, 12):
                money(cell)
            if i == 13:
                cell.number_format = "#,##0.00"
        fill = {BY_DRAW: HARD, BY_TYPE: SOFT, BY_KG: GUESS,
                NO_PRICE: NONE_FILL}[r["Источник цены"]]
        ws.cell(row=row, column=14).fill = fill

    ws.auto_filter.ref = f"A{start}:O{start + len(rows)}"
    ws.freeze_panes = f"C{start + 1}"
    widths(ws, [5, 52, 20, 7, 28, 16, 9, 12, 20, 20, 16, 16, 12, 26, 18])
    return ws


def main():
    by_draw, by_type, comp = competitor_index()
    items = our_stock()
    per_kg, all_kg = value(items, by_draw, by_type)

    # в Петербург едет весь этот перечень, кроме позиций без цены
    # в Петербург едет весь этот перечень: в усиленный контейнер
    # входит 26,35 т из 27,30, на долю логистики разница не влияет
    ship = items

    wb = Workbook()
    wb.remove(wb.active)
    sheet_summary(wb, items, ship, per_kg, all_kg, comp)
    sheet_groups(wb, items, per_kg)
    sheet_items(wb, "Цена по чертежу",
                "Позиции, где цена конкурента найдена по номеру чертежа",
                [r for r in items if r["Источник цены"] == BY_DRAW],
                "Самая надежная часть оценки: один чертеж — одно изделие.")
    sheet_items(wb, "Цена по типу",
                "Позиции, где цена подобрана по группе, диаметру и материалу",
                [r for r in items if r["Источник цены"] == BY_TYPE],
                "Совпадение приблизительное: исполнение может отличаться.")
    sheet_items(wb, "Оценка по килограмму",
                "Позиции без аналога: оценены по цене килограмма своей группы",
                [r for r in items if r["Источник цены"] == BY_KG],
                "Для прикидки. В переговорах такие цифры не называем.")

    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / "Оценка_по_конкуренту.xlsx"
    wb.save(path)

    total = sum(r["Сумма конкурента"] for r in items)
    ours = sum(r["Наша цена"] * r["Штук"] for r in items)
    kg = sum(r["вес"] for r in items)
    ship_val = sum(r["Сумма конкурента"] for r in ship)
    ship_kg = sum(r["вес"] for r in ship)
    print(f"готово: {path}\n")
    print(f"склад: {len(items)} позиций, {kg / 1000:.1f} т")
    print(f"  по конкуренту {total / 1e6:>7.1f} млн руб")
    print(f"  по нашему учету {ours / 1e6:>5.1f} млн руб "
          f"(отношение {total / ours:.2f})")
    for how in (BY_DRAW, BY_TYPE, BY_KG, NO_PRICE):
        rs = [r for r in items if r["Источник цены"] == how]
        if rs:
            print(f"  {how:<28}{len(rs):>4} поз, "
                  f"{sum(r['Сумма конкурента'] for r in rs) / 1e6:>6.1f} млн")
    print(f"\nотгрузка в Питер: {ship_kg / 1000:.2f} т на "
          f"{ship_val / 1e6:.1f} млн руб по конкуренту")
    for name, cost in [
        ("усиленный 20ф 27 т",
         PRICE["20 усиленный"] + AUTO_VL[20] + AUTO_SPB[20] + loading(HOURS_BOX)),
        ("два 20ф по 22 т",
         2 * (PRICE["20 обычный"] + AUTO_VL[20] + AUTO_SPB[20])
         + loading(2 * HOURS_BOX)),
        ("один 40ф",
         PRICE["40"] + AUTO_VL[40] + AUTO_SPB[40] + loading(HOURS_BOX)),
    ]:
        print(f"  {name:<22}{cost:>9,.0f} руб = "
              f"{cost / ship_val * 100:>5.2f}% от стоимости, "
              f"{cost / ship_kg:>5.1f} руб/кг".replace(",", " "))


if __name__ == "__main__":
    main()
