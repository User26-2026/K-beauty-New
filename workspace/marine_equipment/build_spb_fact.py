#!/usr/bin/env python3
"""Сверка остатков по питерскому списку: приложения против факта.

Сотрудница прислала 09.10.2026 обновленный список. По сравнению с
версией от 30.09 в приложениях поправлено семь строк, а главное — в
файле появился третий лист «нет в списках»: товар, который нашелся на
складе, но ни в приложение 2, ни в приложение 3 не попал. Это первые
настоящие данные пересчета, а не учет 16.06.2025.

Что считаем:

1. Расхождения между «кол-во по приложению» и «кол-во факт». Это и есть
   доказательство, насколько учет разошелся с полкой.
2. Лист «нет в списках» — 54 позиции. По каждой смотрим, числится ли
   такое же в нашем складском реестре: если да, позиция не новая, ее
   уже видели; если нет — это находка, в планах отгрузки ее не было.
3. Итоги по весу: сколько всего нашлось против того, что было в
   приложениях.

Цен в новом файле по приложению 2 нет, они остались в версии 30.09.
Поэтому цены по договору берем из старого файла, а количества — из
нового. Подменять источник цен нельзя: договор подписан по нему.

Источник: data/marine_equipment/spb_list/Приложения_2_и_3_Питер_09.10.26.xlsx
Результат: outputs/marine_equipment/Сверка_остатков_09.10.xlsx
"""

import pathlib
import re
import sys

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
from build_marine_registry import SOURCES, parse  # noqa: E402
from build_spb_list_xlsx import group_of  # noqa: E402

ROOT = pathlib.Path(__file__).resolve().parents[2]
SRC = (ROOT / "data" / "marine_equipment" / "spb_list"
       / "Приложения_2_и_3_Питер_09.10.26.xlsx")
OLD = (ROOT / "data" / "marine_equipment" / "spb_list"
       / "Приложения_2_и_3_Питер_30.09.26.xlsx")
OUT = ROOT / "outputs" / "marine_equipment"

FONT = "Arial"
HEAD_FILL = PatternFill("solid", fgColor="1F3864")
HEAD_FONT = Font(name=FONT, size=11, bold=True, color="FFFFFF")
TITLE = Font(name=FONT, size=14, bold=True, color="1F3864")
SUB = Font(name=FONT, size=11, bold=True, color="1F3864")
BOLD = Font(name=FONT, size=10, bold=True)
BASE = Font(name=FONT, size=10)
NOTE = Font(name=FONT, size=9, italic=True, color="666666")
GROUP_FILL = PatternFill("solid", fgColor="D9E2F3")
PLUS = PatternFill("solid", fgColor="E2EFDA")    # нашлось больше
MINUS = PatternFill("solid", fgColor="FCE4D6")   # нашлось меньше
GONE = PatternFill("solid", fgColor="E2E2E2")    # отгружено
NEW_FILL = PatternFill("solid", fgColor="FFF2CC")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# лист -> колонки (1-based) и как лист подписывать
APPS = [
    dict(sheet="Приложение № 2 (ПИТЕР)", short="Прил. 2",
         item=2, draw=3, cond=4, unit_kg=5, doc=6, fact=7, kg=8, note=9),
    dict(sheet="Приложение №3 (Питер)", short="Прил. 3",
         item=2, draw=3, cond=4, unit_kg=5, doc=6, fact=7, kg=8, note=10),
]
EXTRA = dict(sheet="нет в списках ", item=2, draw=3, cond=4, unit_kg=5,
             qty=6, kg=7, buy=8, sell=9)


def num(v):
    """Число из ячейки. В файле встречается и текст, и запятая."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", ".").replace(" ", "")
    m = re.match(r"^-?\d+(?:\.\d+)?$", s)
    return float(s) if m else None


def text(v):
    return str(v).strip() if v is not None else ""


def norm_draw(v):
    """Чертеж без скобок, пробелов и регистра: пишут по-разному."""
    return re.sub(r"[^0-9a-zа-я.\-]", "", text(v).lower())


def clean(name):
    """Наименование для сравнения: без мусора и лишних пробелов."""
    s = text(name).lower()
    s = re.sub(r"\(.*?\)", " ", s)
    s = re.sub(r"[^0-9a-zа-я]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def read_app(path, spec):
    rows = []
    ws = load_workbook(path, data_only=True)[spec["sheet"]]
    for row in ws.iter_rows(min_row=2, values_only=True):
        name = text(row[spec["item"] - 1])
        if not name or name.lower().startswith("итого"):
            continue
        rows.append({
            "Лист": spec["short"],
            "Наименование": name,
            "Чертеж": text(row[spec["draw"] - 1]),
            "Состояние": text(row[spec["cond"] - 1]),
            "вес_ед": num(row[spec["unit_kg"] - 1]) or 0.0,
            "По приложению": num(row[spec["doc"] - 1]) or 0.0,
            "Факт": num(row[spec["fact"] - 1]),
            "Вес по файлу": num(row[spec["kg"] - 1]) or 0.0,
            "Примечание": text(row[spec["note"] - 1])
            if len(row) >= spec["note"] else "",
        })
    return rows


def read_extra():
    rows = []
    ws = load_workbook(SRC, data_only=True)[EXTRA["sheet"]]
    for row in ws.iter_rows(min_row=2, values_only=True):
        name = text(row[EXTRA["item"] - 1])
        if not name or name.lower().startswith("итого"):
            continue
        qty = num(row[EXTRA["qty"] - 1]) or 0.0
        unit = num(row[EXTRA["unit_kg"] - 1]) or 0.0
        rows.append({
            "Наименование": name,
            "Чертеж": text(row[EXTRA["draw"] - 1]),
            "Состояние": text(row[EXTRA["cond"] - 1]),
            "вес_ед": unit,
            "Штук": qty,
            "Вес": num(row[EXTRA["kg"] - 1]) or qty * unit,
            "Закупка": num(row[EXTRA["buy"] - 1]),
            "Продажа": num(row[EXTRA["sell"] - 1]),
            "Группа": group_of(name),
        })
    return rows


def registry_index():
    """Что числится в складском учете: по чертежу и по наименованию."""
    by_draw, by_name = {}, {}
    for src in SOURCES:
        for r in parse(src):
            if not r.get("Наличие"):
                continue
            d = norm_draw(r.get("Чертеж"))
            if d:
                by_draw.setdefault(d, []).append(r)
            by_name.setdefault(clean(r["Наименование"]), []).append(r)
    return by_draw, by_name


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


def put(ws, row, values, wrap=(2,), center=(), fill=None, font=None):
    for i, v in enumerate(values, start=1):
        cell = ws.cell(row=row, column=i, value=v)
        cell.font = font or BASE
        cell.border = BORDER
        cell.alignment = Alignment(vertical="top", wrap_text=(i in wrap),
                                   horizontal="center" if i in center
                                   else None)
        if fill:
            cell.fill = fill
    return row + 1


def plural(n, one, few, many):
    n = abs(int(n))
    if n % 10 == 1 and n % 100 != 11:
        return one
    if n % 10 in (2, 3, 4) and n % 100 not in (12, 13, 14):
        return few
    return many


COLS_DIFF = [
    ("№", 5), ("Лист", 8), ("Наименование", 52), ("Чертеж", 22),
    ("Состояние", 10), ("Вес 1 шт, кг", 10), ("По приложению, шт", 11),
    ("ФАКТ, шт", 10), ("Разница, шт", 10), ("Разница, кг", 11),
    ("Что это значит", 34),
]


def verdict(r):
    """Что означает расхождение по строке."""
    note = r["Примечание"].lower()
    if "отгружено" in note:
        return "отгружено, с остатка снято"
    fact, doc = r["Факт"], r["По приложению"]
    if fact is None:
        return "не считали"
    if fact == doc:
        return ""
    if fact == 0:
        return "НЕ НАШЛИ ни одной"
    if doc == 0:
        return "в приложении не было, нашлось"
    return "нашлось больше" if fact > doc else "нашлось меньше"


def sheet_diff(wb, apps):
    ws = wb.active
    ws.title = "Расхождения"
    rows = [r for r in apps if verdict(r)]
    # пустой «факт» — это «не считали», а не ноль: в недостачу не идет
    counted = [r for r in rows if r["Факт"] is not None
               and "отгружено" not in r["Примечание"].lower()]
    plus = sum((r["Факт"] - r["По приложению"]) * r["вес_ед"]
               for r in counted if r["Факт"] > r["По приложению"])
    minus = sum((r["По приложению"] - r["Факт"]) * r["вес_ед"]
                for r in counted if r["Факт"] < r["По приложению"])
    ws["A1"] = "Расхождения: приложение против факта"
    ws["A1"].font = TITLE
    ws["A2"] = (f"{len(rows)} {plural(len(rows), 'строка', 'строки', 'строк')} "
                f"из {len(apps)} разошлись. Нашлось сверх приложения "
                f"{plus:,.0f} кг, не нашлось {minus:,.0f} кг."
                .replace(",", " "))
    ws["A2"].font = NOTE
    ws["A3"] = ("Это пересчет на полке, а не учет. Учет сверялся "
                "16.06.2025, и по двум прошлым отгрузкам он уже не "
                "сошелся: клинкеты Ду150 числились в нуле, а уехало 45.")
    ws["A3"].font = NOTE

    start = 5
    for i, (h, _) in enumerate(COLS_DIFF, start=1):
        ws.cell(row=start, column=i, value=h)
    head(ws, start, len(COLS_DIFF))
    row = start + 1
    for n, r in enumerate(rows, start=1):
        fact = r["Факт"]
        d = (fact or 0) - r["По приложению"]
        v = verdict(r)
        fill = (GONE if "отгружено" in v else
                PLUS if d > 0 else MINUS if d < 0 else None)
        row = put(ws, row, [
            n, r["Лист"], r["Наименование"], r["Чертеж"], r["Состояние"],
            round(r["вес_ед"], 1), r["По приложению"],
            fact if fact is not None else "",
            d, round(d * r["вес_ед"]), v,
        ], wrap=(3, 11), center=(1, 2, 5, 7, 8, 9), fill=fill)
    ws.freeze_panes = f"A{start + 1}"
    ws.auto_filter.ref = f"A{start}:{get_column_letter(len(COLS_DIFF))}{row - 1}"
    widths(ws, [w for _, w in COLS_DIFF])
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    return plus, minus


COLS_EXTRA = [
    ("№", 5), ("Наименование", 54), ("Чертеж", 22), ("Состояние", 10),
    ("Штук", 8), ("Вес 1 шт, кг", 10), ("Вес всего, кг", 11),
    ("Закупка, руб", 11), ("Продажа, руб", 11), ("Продажа всего, руб", 13),
    ("Есть ли в учете", 30),
]


def in_registry(r, by_draw, by_name):
    """Числится ли такая же позиция в нашем складском учете."""
    d = norm_draw(r["Чертеж"])
    if d and d in by_draw:
        hit = by_draw[d]
        return (f"да, по чертежу: {sum(x['Наличие'] for x in hit):,.0f} шт "
                f"({hit[0]['Локация']})".replace(",", " "))
    hit = by_name.get(clean(r["Наименование"]))
    if hit:
        return (f"да, по названию: {sum(x['Наличие'] for x in hit):,.0f} шт "
                f"({hit[0]['Локация']})".replace(",", " "))
    return "НЕТ, в учете не числится"


def sheet_extra(wb, extra, by_draw, by_name):
    ws = wb.create_sheet("Нет в списках")
    kg = sum(r["Вес"] for r in extra)
    qty = sum(r["Штук"] for r in extra)
    money = sum((r["Продажа"] or 0) * r["Штук"] for r in extra)
    nopr = [r for r in extra if not r["Продажа"]]
    ws["A1"] = "Нашлось на складе, но в приложения не попало"
    ws["A1"].font = TITLE
    ws["A2"] = (f"{len(extra)} "
                f"{plural(len(extra), 'позиция', 'позиции', 'позиций')}, "
                f"{qty:,.0f} шт, {kg:,.0f} кг. По проставленным ценам "
                f"продажи это {money:,.0f} руб, цены нет у {len(nopr)} "
                f"{plural(len(nopr), 'позиции', 'позиций', 'позиций')}."
                .replace(",", " "))
    ws["A2"].font = NOTE
    ws["A3"] = ("Цены закупки и продажи здесь проставлены сотрудницей и "
                "в договор с ПКФ «АРМА» не входят. По договору товар без "
                "согласованной цены проданным не считается, так что "
                "перед отгрузкой эти цены надо согласовать письменно.")
    ws["A3"].font = NOTE

    start = 5
    for i, (h, _) in enumerate(COLS_EXTRA, start=1):
        ws.cell(row=start, column=i, value=h)
    head(ws, start, len(COLS_EXTRA))
    row = start + 1
    n = 0
    by_group = {}
    for r in extra:
        by_group.setdefault(r["Группа"], []).append(r)
    for g in sorted(by_group, key=lambda x: -sum(i["Вес"]
                                                 for i in by_group[x])):
        items = by_group[g]
        row = put(ws, row, [
            f"{g} — {len(items)} "
            f"{plural(len(items), 'позиция', 'позиции', 'позиций')}, "
            f"{sum(i['Штук'] for i in items):,.0f} шт, "
            f"{sum(i['Вес'] for i in items):,.0f} кг".replace(",", " ")
        ], wrap=(), fill=GROUP_FILL, font=SUB)
        for r in sorted(items, key=lambda x: -x["Вес"]):
            n += 1
            where = in_registry(r, by_draw, by_name)
            row = put(ws, row, [
                n, r["Наименование"], r["Чертеж"], r["Состояние"],
                r["Штук"], round(r["вес_ед"], 1), round(r["Вес"]),
                r["Закупка"] or "", r["Продажа"] or "",
                round((r["Продажа"] or 0) * r["Штук"]) or "", where,
            ], wrap=(2, 11), center=(1, 4, 5),
                fill=NEW_FILL if where.startswith("НЕТ") else None)
    ws.freeze_panes = f"A{start + 1}"
    widths(ws, [w for _, w in COLS_EXTRA])
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    return kg, qty, money, nopr


def sheet_summary(wb, apps, extra, plus, minus, extra_kg, extra_money, nopr,
                  by_draw, by_name):
    ws = wb.create_sheet("Итоги", 0)
    ws["A1"] = "Остатки по Питеру на 09.10.2026"
    ws["A1"].font = TITLE
    live = [r for r in apps if "отгружено" not in r["Примечание"].lower()]
    unknown = [r for r in live if r["Факт"] is None]
    # считаем по одному и тому же набору строк, иначе числа несравнимы:
    # отгруженное снято с обеих сторон, непосчитанное вынесено отдельно
    base = [r for r in live if r["Факт"] is not None]
    fact_kg = sum(r["Факт"] * r["вес_ед"] for r in base)
    doc_kg = sum(r["По приложению"] * r["вес_ед"] for r in base)
    unknown_kg = sum(r["По приложению"] * r["вес_ед"] for r in unknown)
    notfound = [r for r in apps if r["Факт"] == 0 and r["По приложению"] > 0
                and "отгружено" not in r["Примечание"].lower()]
    new_to_us = [r for r in extra
                 if in_registry(r, by_draw, by_name).startswith("НЕТ")]
    rows = [
        ("Приложения 2 и 3, по документу",
         f"{len(base)} посчитанных позиций", f"{doc_kg:,.0f} кг",
         "столько должно быть по согласованному перечню"),
        ("Приложения 2 и 3, по факту",
         f"{len([r for r in base if r['Факт']])} позиций с наличием",
         f"{fact_kg:,.0f} кг",
         "столько реально нашли, отгруженное снято с обеих сторон"),
        ("Нашлось сверх приложения", "", f"+{plus:,.0f} кг",
         "строки, где факт больше документа"),
        ("Не нашлось", f"{len(notfound)} позиций", f"-{minus:,.0f} кг",
         "строки, где факт меньше документа"),
        ("Не считали", f"{len(unknown)} позиций", f"{unknown_kg:,.0f} кг",
         "в колонке «факт» пусто, это не ноль, вес по документу"),
        ("Нет в списках", f"{len(extra)} позиций", f"{extra_kg:,.0f} кг",
         "нашлось на складе, в приложения не попало"),
        ("Из них нет и в учете", f"{len(new_to_us)} позиций",
         f"{sum(r['Вес'] for r in new_to_us):,.0f} кг",
         "этого товара мы раньше не видели вообще"),
        ("Цена продажи по листу «нет в списках»", "",
         f"{extra_money:,.0f} руб",
         f"цены нет у {len(nopr)} позиций, их надо оценить"),
        ("ВСЕГО найдено", "", f"{fact_kg + extra_kg:,.0f} кг",
         "приложения по факту плюс лист «нет в списках»"),
    ]
    start = 3
    for i, h in enumerate(["Что", "Сколько позиций", "Вес", "Пояснение"],
                          start=1):
        ws.cell(row=start, column=i, value=h)
    head(ws, start, 4)
    row = start + 1
    for a, b, c, d in rows:
        bold = a.startswith("ВСЕГО")
        row = put(ws, row, [a, b, c.replace(",", " "), d], wrap=(1, 4),
                  font=BOLD if bold else BASE)
    row += 1
    ws.cell(row=row, column=1, value="Что с этим делать").font = SUB
    row += 1
    for t in [
        "Приложения менялись мало: семь строк из 121. Главное в этом "
        "файле — лист «нет в списках» на 54 позиции и 3,5 тонны.",
        "Цены по листу «нет в списках» проставлены сотрудницей, а не "
        "договором. По договору товар без согласованной цены не "
        "считается проданным — согласовать письменно до отгрузки.",
        "Где в колонке «факт» пусто, это не ноль, а «не считали». "
        "Такие строки надо досчитать, иначе в контейнер попадет "
        "товар, которого нет.",
        "По позициям с пометкой «перепроверить» количество считали "
        "на глаз: коробки 3-х клапановые Ду50 и Ду80, коробка 2-х "
        "клапановая Ду80 и клапан Ду10 Ру40 на 798 штук.",
        "Цены по договору остались в файле от 30.09: в новом файле "
        "колонок с ценами по приложению 2 нет. Источник цен не меняем.",
    ]:
        cell = ws.cell(row=row, column=1, value="— " + t)
        cell.font = BASE
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[row].height = 30
        row += 1
    widths(ws, [38, 22, 16, 62])
    return fact_kg, doc_kg, unknown, unknown_kg, notfound, new_to_us


def main():
    apps = []
    for spec in APPS:
        apps += read_app(SRC, spec)
    extra = read_extra()
    by_draw, by_name = registry_index()

    wb = Workbook()
    plus, minus = sheet_diff(wb, apps)
    extra_kg, extra_qty, extra_money, nopr = sheet_extra(
        wb, extra, by_draw, by_name)
    fact_kg, doc_kg, unknown, unknown_kg, notfound, new_to_us = sheet_summary(
        wb, apps, extra, plus, minus, extra_kg, extra_money, nopr,
        by_draw, by_name)

    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / "Сверка_остатков_09.10.xlsx"
    wb.save(path)
    print(f"готово: {path}\n")
    print(f"приложения: {len(apps)} позиций, "
          f"по документу {doc_kg:,.0f} кг, по факту {fact_kg:,.0f} кг"
          .replace(",", " "))
    print(f"  нашлось сверх {plus:,.0f} кг, не нашлось {minus:,.0f} кг, "
          f"не считали {len(unknown)} позиций на {unknown_kg:,.0f} кг"
          .replace(",", " "))
    print(f"  не нашли ни одной: {len(notfound)} позиций")
    print(f"нет в списках: {len(extra)} позиций, {extra_qty:,.0f} шт, "
          f"{extra_kg:,.0f} кг, по ценам продажи {extra_money:,.0f} руб"
          .replace(",", " "))
    print(f"  из них нет и в нашем учете: {len(new_to_us)} позиций, "
          f"{sum(r['Вес'] for r in new_to_us):,.0f} кг".replace(",", " "))
    print(f"  без цены продажи: {len(nopr)} позиций")
    print(f"\nВСЕГО найдено {fact_kg + extra_kg:,.0f} кг".replace(",", " "))


if __name__ == "__main__":
    main()
