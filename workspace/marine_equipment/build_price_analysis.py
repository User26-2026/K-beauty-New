#!/usr/bin/env python3
"""Анализ цен по вкладке «Цены ,2026»: что это за цены и как их читать.

Менеджер собрала цены от семи источников по 175 позициям. Первое, что
показывает расчет: это НЕ семь предложений за один и тот же товар.
Отношение к камчатскому прайсу по медиане: Титов 1,00, Кристина 1,13,
Евротрейд 1,33, Гелион 1,50, Бурыкин 2,89, Ленпромарматура 5,58.

Значит в одной таблице лежат цены разной природы. У Титова 35 цен из 75
совпадают с камчатским прайсом до рубля — это один и тот же источник.
Ленпромарматура в пять-шесть раз выше рынка: так выглядит прайс продавца
новой арматуры, а не предложение купить у нас. Поэтому складывать
максимумы по строкам и называть это выручкой нельзя.

Цены к тому же даны на разных условиях — у кого с НДС, у кого без, у
кого наличными с дисконтом. Каждая приводится к одному виду: сколько
остается на руки. Коэффициенты в желтых ячейках, их надо подтвердить.

Рыночный ориентир — прайс конкурента с Камчатки от 03.10.2026, он
подбирается по номеру чертежа, а где чертежа нет — по типу, диаметру и
материалу.

Источник: data/marine_equipment/quotes/Цены_покупателей_13.10.26.xlsx
Результат: outputs/marine_equipment/Анализ_цен.xlsx
"""

import pathlib
import re
import sys
from collections import defaultdict

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = pathlib.Path(__file__).resolve().parents[2]
SRC = (ROOT / "data" / "marine_equipment" / "quotes"
       / "Цены_покупателей_13.10.26.xlsx")
OUT = ROOT / "outputs" / "marine_equipment"

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
from build_competitor_valuation import (  # noqa: E402
    BY_DRAW, BY_TYPE, NO_PRICE, competitor_index, pick,
)
from build_shipment_spb import item_kind, material  # noqa: E402
from build_spb_list_xlsx import group_of  # noqa: E402
from build_weight_xlsx import du as du_of  # noqa: E402

FONT = "Arial"
HEAD_FILL = PatternFill("solid", fgColor="1F3864")
HEAD_FONT = Font(name=FONT, size=11, bold=True, color="FFFFFF")
TITLE = Font(name=FONT, size=14, bold=True, color="1F3864")
SUB = Font(name=FONT, size=11, bold=True, color="1F3864")
BOLD = Font(name=FONT, size=10, bold=True)
BASE = Font(name=FONT, size=10)
NOTE = Font(name=FONT, size=9, italic=True, color="666666")
BEST = PatternFill("solid", fgColor="EAF3EA")
WARN = PatternFill("solid", fgColor="FCE4D6")
MARK = PatternFill("solid", fgColor="FFF2CC")
INPUT_FILL = PatternFill("solid", fgColor="FFFF00")
GROUP_FILL = PatternFill("solid", fgColor="D9E2F3")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# колонка на вкладке «Цены ,2026»: кто, как считать «на руки», пояснение
# Коэффициент — то, на что умножаем цену из файла.
BUYERS = [
    (10, "Кристина (Веда)", 1 / 1.22,
     "цена с НДС 22%, делим на 1,22"),
    (11, "Гелион", 0.88,
     "цена от ИП наличными минус 12%"),
    (12, "ПромВест", 1.0, "цен пока нет"),
    (13, "Андрей Евротрейд", 1 / 1.22,
     "безнал с НДС 22%, делим на 1,22"),
    (15, "Титов", 1.0,
     "условия в шапке не указаны, берем как есть — уточнить"),
    (16, "Ленпромарматура", 1.0,
     "без НДС, февраль 2026, берем как есть"),
    (17, "Антон Бурыкин", 0.94,
     "цена ИП наличными минус 6%, февраль 2026"),
]
COND_COL = 14          # состояние по Титову
NAME_COL, DRAW_COL, STATE_COL, W_COL, QTY_COL = 2, 3, 4, 5, 7


def norm_draw(d):
    if d is None:
        return ""
    s = str(d).upper().replace(" ", "").replace("…", "")
    return s.replace("–", "-").replace("—", "-").strip(".")


def read_quotes():
    ws = load_workbook(SRC, data_only=True)["Цены ,2026"]
    rows = []
    for r in range(2, ws.max_row + 1):
        name = ws.cell(row=r, column=NAME_COL).value
        if not name or not str(name).strip():
            continue
        qty = ws.cell(row=r, column=QTY_COL).value or 0
        w = ws.cell(row=r, column=W_COL).value or 0
        draw = ws.cell(row=r, column=DRAW_COL).value
        if draw and "без чертеж" in str(draw).lower():
            draw = ""
        item = {
            "Наименование": str(name).strip(),
            "Чертеж": norm_draw(draw),
            "Состояние": str(ws.cell(row=r, column=STATE_COL).value
                             or "").strip(),
            "Вес": float(w) if isinstance(w, (int, float)) else 0.0,
            "Штук": float(qty) if isinstance(qty, (int, float)) else 0.0,
            "Титов состояние": str(ws.cell(row=r, column=COND_COL).value
                                   or "").strip(),
            "Цены": {},
        }
        for col, who, k, _ in BUYERS:
            v = ws.cell(row=r, column=col).value
            if isinstance(v, (int, float)) and v > 0:
                item["Цены"][who] = float(v)
        item["Ду"] = du_of(item["Наименование"])
        item["Группа"] = group_of(item["Наименование"])
        fake = {"Наименование": item["Наименование"], "Группа в файле": ""}
        item["Тип"] = item_kind(fake)
        item["Материал"] = material(fake)
        rows.append(item)
    return rows


def net(who, price):
    """Сколько остается на руки по условиям этого покупателя."""
    for _, name, k, _ in BUYERS:
        if name == who:
            return price * k
    return price


def head(ws, row, ncols, height=44):
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


def enrich(items, by_draw, by_type):
    for it in items:
        p, how = pick(by_draw, by_type, it["Чертеж"], it["Группа"],
                      it["Ду"], it["Материал"])
        it["Конкурент"] = p
        it["Источник"] = how
        nets = {who: net(who, v) for who, v in it["Цены"].items()}
        it["На руки"] = nets
        if nets:
            best = max(nets.items(), key=lambda kv: kv[1])
            it["Лучший"] = best[0]
            it["Лучшая"] = best[1]
            others = sorted((v for k, v in nets.items() if k != best[0]),
                            reverse=True)
            it["Вторая"] = others[0] if others else None
        else:
            it["Лучший"] = it["Лучшая"] = it["Вторая"] = None


def ratio_stat(items, who):
    """Отношение цен источника к камчатскому прайсу."""
    both = [(it["Цены"][who], it["Конкурент"]) for it in items
            if who in it["Цены"] and it["Конкурент"]]
    if not both:
        return None, 0, 0
    same = sum(1 for a, b in both if abs(a - b) / b < 0.01)
    ratios = sorted(a / b for a, b in both)
    return ratios[len(ratios) // 2], len(both), same


def verdict(mid, same, n):
    """Что это за цена, судя по отношению к рынку."""
    if mid is None:
        return "нечего сравнить"
    if n and same >= n * 0.3:
        return "тот же прайс, что у конкурента — не независимая цена"
    if mid < 1.2:
        return "на уровне рынка"
    if mid < 2:
        return "выше рынка, похоже на цену с наценкой"
    return "в разы выше рынка — похоже на прайс продавца или цену реализации"


def sheet_summary(wb, items):
    ws = wb.create_sheet("Сводка", 0)
    ws["A1"] = "Анализ цен по вкладке «Цены ,2026»"
    ws["A1"].font = TITLE
    ws["A2"] = ("Главное: в таблице лежат цены разной природы. Прежде чем "
                "складывать, надо понять, кто покупает у нас, а кто просто "
                "показал свой прайс. Отношение к камчатскому прайсу это и "
                "показывает.")
    ws["A2"].font = NOTE

    row = 4
    ws.cell(row=row, column=1, value="Как приводим цены").font = SUB
    row += 1
    for i, h in enumerate(["Покупатель", "Коэффициент", "Условия из шапки"],
                          start=1):
        ws.cell(row=row, column=i, value=h)
    head(ws, row, 3, height=24)
    for _, who, k, note in BUYERS:
        row += 1
        ws.cell(row=row, column=1, value=who).font = BASE
        cell = ws.cell(row=row, column=2, value=round(k, 4))
        cell.font = BOLD
        cell.fill = INPUT_FILL
        cell.border = BORDER
        cell.number_format = "0.0000"
        ws.cell(row=row, column=3, value=note).font = NOTE

    row += 2
    ws.cell(row=row, column=1,
            value="Что это за цены: сверка с камчатским прайсом").font = SUB
    row += 1
    cols = ["Источник", "Позиций с ценой", "Доля из 175",
            "Пересечений с конкурентом", "Совпало до рубля",
            "Медиана отношения к рынку", "Руб за кг на руки",
            "Что это, судя по цифрам"]
    for i, h in enumerate(cols, start=1):
        ws.cell(row=row, column=i, value=h)
    head(ws, row, len(cols), height=46)

    stat = defaultdict(lambda: [0, 0.0, 0.0, 0.0, 0])
    for it in items:
        for who, v in it["Цены"].items():
            qty = it["Штук"] or 1
            s = stat[who]
            s[0] += 1
            s[1] += v * qty
            s[2] += net(who, v) * qty
            s[3] += it["Вес"] * qty
        if it["Лучший"]:
            stat[it["Лучший"]][4] += 1

    for _, who, _, _ in BUYERS:
        row += 1
        n, gross, net_sum, kg, wins = stat[who]
        mid, cross, same = ratio_stat(items, who)
        values = [who, n, n / len(items) if items else 0, cross, same,
                  round(mid, 2) if mid else None,
                  round(net_sum / kg) if kg else None,
                  verdict(mid, same, cross)]
        for i, v in enumerate(values, start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=(i == 8))
            if i in (2, 4, 5, 7):
                cell.number_format = "#,##0"
            if i == 3:
                cell.number_format = "0%"
            if i == 6:
                cell.number_format = "#,##0.00"
                if v and v > 2:
                    cell.fill = WARN
                elif v and v < 1.2:
                    cell.fill = BEST

    row += 2
    ws.cell(row=row, column=1, value="Итоги").font = SUB
    row += 1
    priced = [it for it in items if it["Лучшая"]]
    multi = [it for it in items if len(it["На руки"]) > 1]
    spreads = sorted(max(it["На руки"].values()) / min(it["На руки"].values())
                     for it in multi) if multi else []
    best_sum = sum(it["Лучшая"] * (it["Штук"] or 1) for it in priced)
    comp_sum = sum(it["Конкурент"] * (it["Штук"] or 1)
                   for it in priced if it["Конкурент"])
    comp_cover = [it for it in priced if it["Конкурент"]]
    kg_priced = sum(it["Вес"] * (it["Штук"] or 1) for it in priced)
    for label, value, note in [
        ("Позиций в файле", len(items), ""),
        ("Из них есть хотя бы одна цена", len(priced),
         f"{len(priced) / len(items) * 100:.0f}% позиций"),
        ("Позиций без единой цены", len(items) - len(priced),
         "никто не оценил — рассылать отдельно"),
        ("Позиций с двумя и более ценами", len(multi),
         "только по ним и можно сравнивать источники"),
        ("Медианный разброс между источниками",
         round(spreads[len(spreads) // 2], 2) if spreads else None,
         "во столько раз максимум выше минимума по одной позиции"),
        ("Максимальный разброс", round(spreads[-1], 2) if spreads else None,
         "такие строки смотреть отдельно: скорее всего разная трактовка"),
        ("Верхняя граница, если взять самую дорогую цену, руб",
         round(best_sum),
         "это НЕ прогноз выручки: часть цен — прайсы продавцов"),
        ("Вес этих позиций, кг", round(kg_priced), ""),
        ("Выходит на килограмм, руб",
         round(best_sum / kg_priced) if kg_priced else 0,
         "у конкурента средняя 1 817 руб за кг"),
        ("Те же позиции по ценам конкурента, руб", round(comp_sum),
         f"совпало {len(comp_cover)} позиций из {len(priced)}"),
    ]:
        ws.cell(row=row, column=1, value=label).font = BASE
        cell = ws.cell(row=row, column=2, value=value)
        cell.font = BOLD
        cell.border = BORDER
        cell.number_format = "#,##0"
        ws.cell(row=row, column=3, value=note).font = NOTE
        row += 1

    row += 1
    ws.cell(row=row, column=1, value="Что из этого следует").font = SUB
    row += 1
    tit = stat["Титов"]
    t_mid, t_cross, t_same = ratio_stat(items, "Титов")
    l_mid, _, _ = ratio_stat(items, "Ленпромарматура")
    b_mid, _, _ = ratio_stat(items, "Антон Бурыкин")
    for text in [
        "Складывать эти цены в одну сумму нельзя: они разной природы. "
        "Отношение к камчатскому прайсу по медиане — от 1,00 у Титова до "
        f"{l_mid:.1f} у Ленпромарматуры. Это разные вещи, а не разные "
        "предложения за одно и то же.",
        f"Титов — это тот же камчатский прайс. Из {t_cross} сравнимых "
        f"позиций {t_same} совпадают до рубля, медиана отношения "
        f"{t_mid:.2f}. Значит независимых источников меньше, чем колонок.",
        f"Ленпромарматура выше рынка в {l_mid:.1f} раза, Бурыкин в "
        f"{b_mid:.1f}. Так выглядит прайс продавца новой арматуры или "
        "цена реализации конечному покупателю, а не цена, по которой "
        "покупают у нас. Подтвердить до того, как на них опираться.",
        f"Цены есть только по {len(priced)} позициям из {len(items)}. "
        f"По остальным {len(items) - len(priced)} не ответил никто — это "
        "и есть первая задача: разослать их повторно.",
        f"Шире всех отработал Титов: {tit[0]} позиций. У остальных от 4 "
        "до 19, ПромВест не дал ни одной.",
        "Условия у всех разные: цена с НДС 22% и цена без НДС — это "
        "разница в 22% на ровном месте. Коэффициенты приведения в желтых "
        "ячейках, поправьте, если трактовка неверна. По Титову условия в "
        "шапке не указаны вообще, а это 80 цен.",
        "Ленпромарматура и Бурыкин давали цены в феврале 2026. Прошло "
        "восемь месяцев, их надо обновлять.",
    ]:
        cell = ws.cell(row=row, column=1, value="— " + text)
        cell.font = BASE
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[row].height = 32
        row += 1

    widths(ws, [54, 18, 56, 22, 18, 22, 16, 46])
    return ws


def sheet_items(wb, items):
    ws = wb.create_sheet("По позициям")
    ws["A1"] = "Цены по каждой позиции, приведенные к «на руки»"
    ws["A1"].font = TITLE
    ws["A2"] = ("Зеленым — лучшее предложение по позиции. В колонках "
                "покупателей цены уже приведены. Рядом цена конкурента "
                "как рыночный ориентир.")
    ws["A2"].font = NOTE

    names = [who for _, who, _, _ in BUYERS]
    cols = (["№", "Наименование", "Чертеж", "Ду", "Состояние", "Вес, кг",
             "Штук"] + names
            + ["Лучшая цена", "Кто дал лучшую", "Вторая цена",
               "Отрыв лучшей", "Конкурент, руб", "Откуда цена конкурента",
               "Руб за кг по лучшей"])
    start = 4
    for i, h in enumerate(cols, start=1):
        ws.cell(row=start, column=i, value=h)
    head(ws, start, len(cols))

    rows = sorted(items, key=lambda it: -(it["Лучшая"] or 0))
    for k, it in enumerate(rows, start=1):
        row = start + k
        base = [k, it["Наименование"], it["Чертеж"], it["Ду"],
                it["Состояние"], round(it["Вес"], 1), it["Штук"]]
        prices = [round(it["На руки"][w]) if w in it["На руки"] else None
                  for w in names]
        tail = [round(it["Лучшая"]) if it["Лучшая"] else None,
                it["Лучший"],
                round(it["Вторая"]) if it["Вторая"] else None,
                (it["Лучшая"] / it["Вторая"] - 1)
                if it["Лучшая"] and it["Вторая"] else None,
                round(it["Конкурент"]) if it["Конкурент"] else None,
                it["Источник"] if it["Конкурент"] else "",
                round(it["Лучшая"] / it["Вес"])
                if it["Лучшая"] and it["Вес"] else None]
        for i, v in enumerate(base + prices + tail, start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=(i == 2))
            if i >= 6:
                cell.number_format = "#,##0"
            if i == 6:
                cell.number_format = "#,##0.0"
        # отметить лучшего
        if it["Лучший"] in names:
            ws.cell(row=row, column=8 + names.index(it["Лучший"])).fill = BEST
        col_gap = 7 + len(names) + 4
        ws.cell(row=row, column=col_gap).number_format = "0%"
        if not it["Цены"]:
            ws.cell(row=row, column=2).fill = WARN

    last = start + len(rows)
    ws.auto_filter.ref = f"A{start}:{get_column_letter(len(cols))}{last}"
    ws.freeze_panes = f"C{start + 1}"
    widths(ws, [5, 50, 18, 7, 13, 10, 8] + [16] * len(names)
           + [14, 20, 14, 12, 15, 24, 16])
    return ws


def sheet_nobody(wb, items):
    ws = wb.create_sheet("Никто не оценил")
    rows = [it for it in items if not it["Цены"]]
    ws["A1"] = "Позиции, по которым нет ни одной цены"
    ws["A1"].font = TITLE
    ws["A2"] = (f"{len(rows)} позиций. Их надо разослать повторно: без "
                "цены позиция не продается и, по договору, не считается "
                "проданной.")
    ws["A2"].font = NOTE
    cols = ["№", "Наименование", "Чертеж", "Ду", "Состояние", "Вес, кг",
            "Штук", "Общий вес, кг", "Цена конкурента, руб",
            "Откуда цена конкурента"]
    for i, h in enumerate(cols, start=1):
        ws.cell(row=4, column=i, value=h)
    head(ws, 4, len(cols), height=32)
    for k, it in enumerate(sorted(rows, key=lambda x: -(x["Вес"] * (x["Штук"] or 1))),
                           start=1):
        row = 4 + k
        values = [k, it["Наименование"], it["Чертеж"], it["Ду"],
                  it["Состояние"], round(it["Вес"], 1), it["Штук"],
                  round(it["Вес"] * (it["Штук"] or 1)),
                  round(it["Конкурент"]) if it["Конкурент"] else None,
                  it["Источник"] if it["Конкурент"] else ""]
        for i, v in enumerate(values, start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=(i == 2))
            if i in (7, 8, 9):
                cell.number_format = "#,##0"
            if i == 6:
                cell.number_format = "#,##0.0"
        if it["Конкурент"]:
            ws.cell(row=row, column=9).fill = MARK
    widths(ws, [5, 52, 18, 7, 14, 10, 8, 14, 18, 24])
    return ws


def sheet_groups(wb, items):
    ws = wb.create_sheet("По группам")
    ws["A1"] = "Срез по группам арматуры"
    ws["A1"].font = TITLE
    row = 3
    cols = ["Группа", "Позиций", "С ценой", "Лучшая сумма, руб",
            "Вес, кг", "Руб за кг", "По конкуренту, руб", "Отношение"]
    for i, h in enumerate(cols, start=1):
        ws.cell(row=row, column=i, value=h)
    head(ws, row, len(cols), height=32)
    agg = defaultdict(lambda: [0, 0, 0.0, 0.0, 0.0])
    for it in items:
        a = agg[it["Группа"]]
        a[0] += 1
        qty = it["Штук"] or 1
        if it["Лучшая"]:
            a[1] += 1
            a[2] += it["Лучшая"] * qty
            a[4] += it["Вес"] * qty
        if it["Конкурент"]:
            a[3] += it["Конкурент"] * qty
    for g in sorted(agg, key=lambda x: -agg[x][2]):
        row += 1
        pos, withp, best, comp, kg = agg[g]
        values = [g, pos, withp, round(best), round(kg),
                  round(best / kg) if kg else None, round(comp),
                  round(best / comp, 2) if comp and best else None]
        for i, v in enumerate(values, start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            if i in (4, 5, 6, 7):
                cell.number_format = "#,##0"
            if i == 8:
                cell.number_format = "#,##0.00"
    widths(ws, [34, 10, 10, 20, 12, 12, 20, 12])
    return ws


def main():
    by_draw, by_type, comp = competitor_index()
    items = read_quotes()
    enrich(items, by_draw, by_type)

    wb = Workbook()
    wb.remove(wb.active)
    sheet_summary(wb, items)
    sheet_items(wb, items)
    sheet_nobody(wb, items)
    sheet_groups(wb, items)

    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / "Анализ_цен.xlsx"
    wb.save(path)

    priced = [it for it in items if it["Лучшая"]]
    print(f"готово: {path}\n")
    print(f"позиций в файле: {len(items)}, с ценой хотя бы от одного: "
          f"{len(priced)}, без цены: {len(items) - len(priced)}\n")
    stat = defaultdict(lambda: [0, 0.0, 0])
    for it in items:
        for who, v in it["Цены"].items():
            stat[who][0] += 1
            stat[who][1] += net(who, v) * (it["Штук"] or 1)
        if it["Лучший"]:
            stat[it["Лучший"]][2] += 1
    print(f"{'источник':<22}{'позиций':>9}{'пересеч':>9}{'совпало':>9}"
          f"{'медиана к рынку':>17}")
    for _, who, _, _ in BUYERS:
        n, s_, w = stat[who]
        mid, cross, same = ratio_stat(items, who)
        print(f"{who:<22}{n:>9}{cross:>9}{same:>9}"
              f"{(f'{mid:.2f}' if mid else '-'):>17}")
    multi = [it for it in items if len(it["На руки"]) > 1]
    spreads = sorted(max(it["На руки"].values()) / min(it["На руки"].values())
                     for it in multi)
    if spreads:
        print(f"\nпозиций с 2+ ценами: {len(multi)}, разброс медиана "
              f"{spreads[len(spreads) // 2]:.2f}, максимум {spreads[-1]:.2f}")
    best_sum = sum(it["Лучшая"] * (it["Штук"] or 1) for it in priced)
    kg = sum(it["Вес"] * (it["Штук"] or 1) for it in priced)
    print(f"верхняя граница по самым дорогим ценам: {best_sum:,.0f} руб за "
          f"{kg:,.0f} кг = {best_sum / kg:,.0f} руб/кг (НЕ прогноз выручки)"
          .replace(",", " "))


if __name__ == "__main__":
    main()
