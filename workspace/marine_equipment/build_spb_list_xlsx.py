#!/usr/bin/env python3
"""Остатки по питерскому списку: два приложения в одну таблицу по группам.

Менеджер прислал остатки двумя листами — приложение 2 и приложение 3.
Раскладка колонок у них разная, в одном есть цены, в другом «по
согласованию». Здесь они сведены в один файл и разложены по группам:
штуцерная угловая и проходная, фланцевая угловая и проходная, клинкеты,
фильтра, кингстоны, вентиляционные головки, смотровая колонка, дальше
захлопки, коробки, краны и остальное.

Внутри каждой группы позиции идут от меньшего диаметра к большему.

Отдельно помечены титан, арматура высокого давления (Ру от 100),
педальные и электромагнитные клапаны и котельная арматура — паровой
клапан, сигнальные предохранительные и водоуказательная колонка.

Цены в исходнике записаны по-разному: где-то число, где-то диапазон
«30-35» в тысячах, где-то «55-650» явно с опечаткой. Поэтому цены
переносятся текстом как есть и не суммируются.

Источник: data/marine_equipment/spb_list/Приложения_2_и_3_Питер_30.09.26.xlsx
Результат: outputs/marine_equipment/Остатки_Питер_по_группам.xlsx
"""

import pathlib
import re

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = pathlib.Path(__file__).resolve().parents[2]
SRC = (ROOT / "data" / "marine_equipment" / "spb_list"
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
MARK_FILL = PatternFill("solid", fgColor="FFF2CC")
WARN_FILL = PatternFill("solid", fgColor="FCE4D6")
GONE_FILL = PatternFill("solid", fgColor="E2E2E2")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# лист исходника -> раскладка колонок (1-based)
SHEETS = [
    dict(name="Приложение № 2 (ПИТЕР)", short="Прил. 2", start=2,
         item=2, draw=3, cond=4, unit_kg=5, qty_doc=6, qty_fact=7,
         total_kg=8, price=9, resale=10, note=11),
    dict(name="Приложение №3 (Питер)", short="Прил. 3", start=2,
         item=2, draw=3, cond=4, unit_kg=5, qty_doc=6, qty_fact=7,
         total_kg=8, price=9, resale=None, note=10),
]

# порядок групп задан заказчиком
GROUPS = [
    "Штуцерная угловая",
    "Штуцерная проходная",
    "Штуцерная прочая",
    "Фланцевая угловая",
    "Фланцевая проходная",
    "Фланцевая прочая",
    "Клинкеты",
    "Фильтры забортной воды",
    "Кингстоны",
    "Головки вентиляционные",
    "Колонка смотровая",
    "Захлопки",
    "Коробки клапанные",
    "Краны",
    "Стаканы переборочные",
    "Прочее",
]

COLUMNS = [
    ("№", 5),
    ("Группа", 22),
    ("Ду", 8),
    ("Ру", 8),
    ("Наименование", 54),
    ("Обозначение", 22),
    ("Состояние", 11),
    ("Выделено", 26),
    ("Вес 1 шт, кг", 11),
    ("Кол-во по приложению", 13),
    ("Кол-во факт", 11),
    ("Расхождение", 11),
    ("Общий вес, кг", 12),
    ("Цена по договору", 16),
    ("Ориентир реализации", 18),
    ("Примечание", 24),
    ("Лист", 10),
]


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


def text(v):
    if v is None:
        return ""
    if isinstance(v, str):
        return v.strip()
    return str(v).strip()


def number(v):
    if isinstance(v, (int, float)):
        return float(v)
    return None


def du(name):
    """Условный проход. В названиях встречается «Ду20/10» и «Ду 15/1» —
    берем первое число."""
    m = re.search(r"[ДдD][Ууy]\s*\.?\s*(\d{1,3})", name)
    return int(m.group(1)) if m else None


def ru(name):
    """Условное давление. Пишут и «Ру400», и «Р160», и «Рр150/10»,
    и «Ру200…30» — берем первое число после Р."""
    m = re.search(r"[Рр]{1,2}\s*[уy]?\s*\.?\s*(\d{1,3})", name)
    if not m:
        return None
    # «Ду 10 Ру400» — не спутать с диаметром
    return int(m.group(1))


def group_of(name):
    """Группа по названию. Угловая и проходная разделены отдельно —
    это разные изделия для покупателя."""
    low = name.lower()
    angular = "углов" in low
    straight = "проход" in low

    if re.search(r"клинкет|задвижк", low):
        return "Клинкеты"
    if "фильтр" in low:
        return "Фильтры забортной воды"
    if "кингстон" in low:
        return "Кингстоны"
    if re.search(r"головка возд|грибк", low):
        return "Головки вентиляционные"
    if "колонк" in low:
        return "Колонка смотровая"
    if "захлопк" in low:
        return "Захлопки"
    if re.search(r"коробка \d|коробка 2-х|коробка 3-х", low):
        return "Коробки клапанные"
    if low.startswith("кран") or "кран " in low[:6]:
        return "Краны"
    if "стакан переборочн" in low:
        return "Стаканы переборочные"

    if re.search(r"штуцерн|штущерн", low):
        if angular:
            return "Штуцерная угловая"
        if straight:
            return "Штуцерная проходная"
        return "Штуцерная прочая"
    if "фланц" in low:
        if angular:
            return "Фланцевая угловая"
        if straight:
            return "Фланцевая проходная"
        return "Фланцевая прочая"
    # резьбовые и дюритовые присоединения идут к штуцерной
    if re.search(r"муфтов|цапков|под дюрит|присоед", low):
        if angular:
            return "Штуцерная угловая"
        if straight:
            return "Штуцерная проходная"
        return "Штуцерная прочая"
    return "Прочее"


def marks(name, pressure):
    """Что выделяем отдельно."""
    low = name.lower()
    out = []
    if "титан" in low:
        out.append("титан")
    if pressure and pressure >= 100:
        out.append(f"высокое давление Ру{pressure}")
    if "педальн" in low:
        out.append("педальный")
    if "электромагнит" in low or "саленоид" in low or "соленоид" in low:
        out.append("электромагнитный")
    # котельная арматура: паровые, сигнальные предохранительные и
    # водоуказательная колонка со своим клапаном. Редукционные сюда не
    # берем — в этом списке они воздушные, на Ру200 и Ру400
    if (re.search(r"паров", low)
            or re.search(r"предохранительн.*сигнальн", low)
            or "колонк" in low):
        out.append("котельная")
    return "; ".join(out)


def read():
    wb = load_workbook(SRC, data_only=True)
    rows = []
    for spec in SHEETS:
        ws = wb[spec["name"]]
        for r in range(spec["start"], ws.max_row + 1):
            name = text(ws.cell(row=r, column=spec["item"]).value)
            if not name or name.lower().startswith("итого"):
                continue
            pressure = ru(name)
            rows.append({
                "Наименование": name,
                "Обозначение": text(ws.cell(row=r,
                                            column=spec["draw"]).value),
                "Состояние": text(ws.cell(row=r,
                                          column=spec["cond"]).value),
                "Ду": du(name),
                "Ру": pressure,
                "Вес 1 шт, кг": number(ws.cell(row=r,
                                               column=spec["unit_kg"]).value),
                "Кол-во по приложению": number(
                    ws.cell(row=r, column=spec["qty_doc"]).value),
                "Кол-во факт": number(
                    ws.cell(row=r, column=spec["qty_fact"]).value),
                "Общий вес, кг": number(
                    ws.cell(row=r, column=spec["total_kg"]).value),
                "Цена по договору": text(
                    ws.cell(row=r, column=spec["price"]).value),
                "Ориентир реализации": text(
                    ws.cell(row=r, column=spec["resale"]).value
                ) if spec["resale"] else "",
                "Примечание": text(ws.cell(row=r,
                                           column=spec["note"]).value),
                "Лист": spec["short"],
            })
    for r in rows:
        r["Группа"] = group_of(r["Наименование"])
        r["Выделено"] = marks(r["Наименование"], r["Ру"])
        doc = r["Кол-во по приложению"]
        fact = r["Кол-во факт"]
        r["Расхождение"] = (fact - doc) if (doc is not None
                                            and fact is not None) else None
        r["Отгружено"] = "отгружено" in r["Примечание"].lower()
    return rows


def sort_key(r):
    return (r["Ду"] if r["Ду"] is not None else 9999,
            r["Ру"] or 0, r["Наименование"])


def sheet_main(wb, rows):
    ws = wb.create_sheet("По группам")
    ws["A1"] = "Остатки по питерскому списку на 30.09.2026"
    ws["A1"].font = TITLE
    ws["A2"] = ("Приложения 2 и 3 сведены вместе. Внутри группы — от "
                "меньшего диаметра к большему. Желтым помечены титан, "
                "высокое давление, котельная арматура, педальные и "
                "электромагнитные клапаны.")
    ws["A2"].font = NOTE

    start = 4
    for i, (name, _) in enumerate(COLUMNS, start=1):
        ws.cell(row=start, column=i, value=name)
    head(ws, start, len(COLUMNS))

    row = start
    n = 0
    for group in GROUPS:
        items = sorted([r for r in rows if r["Группа"] == group],
                       key=sort_key)
        if not items:
            continue
        row += 1
        qty = sum(r["Кол-во факт"] or 0 for r in items)
        kg = sum(r["Общий вес, кг"] or 0 for r in items)
        ws.cell(row=row, column=1,
                value=f"{group} — {len(items)} позиций, {qty:.0f} шт, "
                      f"{kg:,.0f} кг".replace(",", " "))
        for c in range(1, len(COLUMNS) + 1):
            cell = ws.cell(row=row, column=c)
            cell.fill = GROUP_FILL
            cell.font = SUB
            cell.border = BORDER
        for r in items:
            row += 1
            n += 1
            values = [n, group, r["Ду"], r["Ру"], r["Наименование"],
                      r["Обозначение"], r["Состояние"], r["Выделено"],
                      r["Вес 1 шт, кг"], r["Кол-во по приложению"],
                      r["Кол-во факт"], r["Расхождение"],
                      r["Общий вес, кг"], r["Цена по договору"],
                      r["Ориентир реализации"], r["Примечание"], r["Лист"]]
            for i, v in enumerate(values, start=1):
                cell = ws.cell(row=row, column=i, value=v)
                cell.font = BASE
                cell.border = BORDER
                cell.alignment = Alignment(vertical="top",
                                           wrap_text=(i in (5, 8, 16)))
                if i in (9, 13):
                    cell.number_format = "#,##0.0"
                if i in (10, 11, 12):
                    cell.number_format = "#,##0"
            if r["Выделено"]:
                ws.cell(row=row, column=8).fill = MARK_FILL
            if r["Расхождение"]:
                ws.cell(row=row, column=12).fill = WARN_FILL
            if r["Отгружено"]:
                for c in range(1, len(COLUMNS) + 1):
                    ws.cell(row=row, column=c).fill = GONE_FILL

    last = row
    row += 1
    ws.cell(row=row, column=5, value="ИТОГО").font = BOLD
    for col in (10, 11, 13):
        letter = get_column_letter(col)
        cell = ws.cell(row=row, column=col,
                       value=f"=SUM({letter}{start + 1}:{letter}{last})")
        cell.font = BOLD
        cell.border = BORDER
        cell.number_format = "#,##0"

    ws.auto_filter.ref = f"A{start}:{get_column_letter(len(COLUMNS))}{last}"
    ws.freeze_panes = f"E{start + 1}"
    widths(ws, [w for _, w in COLUMNS])
    return ws


def sheet_groups(wb, rows):
    ws = wb.create_sheet("Сводка по группам")
    ws["A1"] = "Сколько чего по группам"
    ws["A1"].font = TITLE
    ws["A2"] = ("Считаем по колонке «Кол-во факт». Цены в исходнике "
                "записаны диапазонами и в разных единицах, поэтому "
                "суммы по деньгам здесь нет.")
    ws["A2"].font = NOTE

    cols = ["Группа", "Позиций", "Штук по приложению", "Штук по факту",
            "Расхождение, шт", "Общий вес, кг", "Доля веса"]
    for i, h in enumerate(cols, start=1):
        ws.cell(row=4, column=i, value=h)
    head(ws, 4, len(cols), height=30)

    total_kg = sum(r["Общий вес, кг"] or 0 for r in rows) or 1
    row = 4
    for group in GROUPS:
        items = [r for r in rows if r["Группа"] == group]
        if not items:
            continue
        row += 1
        kg = sum(r["Общий вес, кг"] or 0 for r in items)
        values = [group, len(items),
                  sum(r["Кол-во по приложению"] or 0 for r in items),
                  sum(r["Кол-во факт"] or 0 for r in items),
                  sum(r["Расхождение"] or 0 for r in items),
                  round(kg, 1), kg / total_kg]
        for i, v in enumerate(values, start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            if i in (3, 4, 5, 6):
                cell.number_format = "#,##0"
            if i == 7:
                cell.number_format = "0.0%"
            if i == 5 and v:
                cell.fill = WARN_FILL
    row += 1
    ws.cell(row=row, column=1, value="ИТОГО").font = BOLD
    for col in range(2, 7):
        letter = get_column_letter(col)
        cell = ws.cell(row=row, column=col,
                       value=f"=SUM({letter}5:{letter}{row - 1})")
        cell.font = BOLD
        cell.border = BORDER
        cell.number_format = "#,##0"

    widths(ws, [26, 11, 18, 14, 15, 14, 11])
    return ws


def sheet_marks(wb, rows):
    ws = wb.create_sheet("Выделенные позиции")
    ws["A1"] = ("Титан, высокое давление, котельная арматура, "
                "педальные и электромагнитные")
    ws["A1"].font = TITLE
    ws["A2"] = ("Высоким давлением считаем Ру от 100. Котельная — "
                "паровой клапан, сигнальные предохранительные и "
                "водоуказательная колонка. Эти позиции стоят дороже "
                "обычных и спрашивают их отдельно.")
    ws["A2"].font = NOTE

    cols = ["Признак", "Группа", "Ду", "Ру", "Наименование", "Обозначение",
            "Состояние", "Кол-во факт", "Вес 1 шт, кг", "Общий вес, кг",
            "Цена по договору", "Ориентир реализации"]
    for i, h in enumerate(cols, start=1):
        ws.cell(row=4, column=i, value=h)
    head(ws, 4, len(cols))

    order = ["титан", "высокое давление", "котельная", "педальный",
             "электромагнитный"]
    row = 4
    for key in order:
        items = sorted([r for r in rows if key in r["Выделено"]],
                       key=sort_key)
        if not items:
            continue
        row += 1
        qty = sum(r["Кол-во факт"] or 0 for r in items)
        ws.cell(row=row, column=1,
                value=f"{key.capitalize()} — {len(items)} позиций, "
                      f"{qty:.0f} шт")
        for c in range(1, len(cols) + 1):
            cell = ws.cell(row=row, column=c)
            cell.fill = GROUP_FILL
            cell.font = SUB
            cell.border = BORDER
        for r in items:
            row += 1
            values = [r["Выделено"], r["Группа"], r["Ду"], r["Ру"],
                      r["Наименование"], r["Обозначение"], r["Состояние"],
                      r["Кол-во факт"], r["Вес 1 шт, кг"],
                      r["Общий вес, кг"], r["Цена по договору"],
                      r["Ориентир реализации"]]
            for i, v in enumerate(values, start=1):
                cell = ws.cell(row=row, column=i, value=v)
                cell.font = BASE
                cell.border = BORDER
                cell.alignment = Alignment(vertical="top",
                                           wrap_text=(i in (1, 5)))
                if i in (9, 10):
                    cell.number_format = "#,##0.0"

    widths(ws, [26, 22, 8, 8, 52, 22, 11, 11, 11, 12, 16, 18])
    ws.freeze_panes = "E5"
    return ws


def sheet_issues(wb, rows):
    ws = wb.create_sheet("Расхождения и вопросы")
    ws["A1"] = "Что надо проверить до продажи"
    ws["A1"].font = TITLE
    ws["A2"] = ("Собрано из самого списка: расхождения между приложением "
                "и фактом, пометки менеджера и строки с нулевым весом.")
    ws["A2"].font = NOTE

    cols = ["Что не так", "Группа", "Наименование", "Обозначение",
            "Кол-во по приложению", "Кол-во факт", "Расхождение",
            "Вес 1 шт, кг", "Общий вес, кг", "Примечание"]
    for i, h in enumerate(cols, start=1):
        ws.cell(row=4, column=i, value=h)
    head(ws, 4, len(cols))

    row = 4
    for r in sorted(rows, key=lambda r: (r["Группа"], sort_key(r))):
        why = []
        if r["Расхождение"]:
            why.append("факт не сходится с приложением")
        if r["Отгружено"]:
            why.append("уже отгружено")
        if r["Примечание"] and not r["Отгружено"]:
            why.append(r["Примечание"].lower())
        if (r["Кол-во факт"] or 0) > 0 and not (r["Общий вес, кг"] or 0):
            why.append("общий вес не посчитан")
        if (r["Кол-во факт"] or 0) > 0 and not (r["Вес 1 шт, кг"] or 0):
            why.append("нет веса единицы")
        if "?" in r["Состояние"]:
            why.append("состояние под вопросом")
        if not why:
            continue
        row += 1
        values = ["; ".join(why), r["Группа"], r["Наименование"],
                  r["Обозначение"], r["Кол-во по приложению"],
                  r["Кол-во факт"], r["Расхождение"], r["Вес 1 шт, кг"],
                  r["Общий вес, кг"], r["Примечание"]]
        for i, v in enumerate(values, start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top",
                                       wrap_text=(i in (1, 3, 10)))
            if i in (5, 6, 7):
                cell.number_format = "#,##0"
            if i in (8, 9):
                cell.number_format = "#,##0.0"
        ws.cell(row=row, column=1).fill = WARN_FILL

    widths(ws, [38, 22, 50, 20, 13, 11, 11, 11, 12, 24])
    ws.freeze_panes = "C5"
    return ws


def main():
    rows = read()

    wb = Workbook()
    wb.remove(wb.active)
    sheet_main(wb, rows)
    sheet_groups(wb, rows)
    sheet_marks(wb, rows)
    sheet_issues(wb, rows)

    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / "Остатки_Питер_по_группам.xlsx"
    wb.save(path)

    print(f"готово: {path}")
    print(f"позиций: {len(rows)}, "
          f"штук по факту: {sum(r['Кол-во факт'] or 0 for r in rows):.0f}, "
          f"вес: {sum(r['Общий вес, кг'] or 0 for r in rows):,.0f} кг"
          .replace(",", " "))
    print()
    for group in GROUPS:
        items = [r for r in rows if r["Группа"] == group]
        if not items:
            continue
        print(f"  {group:<26}{len(items):>3} поз "
              f"{sum(r['Кол-во факт'] or 0 for r in items):>6.0f} шт "
              f"{sum(r['Общий вес, кг'] or 0 for r in items):>9,.0f} кг"
              .replace(",", " "))
    print()
    for key in ("титан", "высокое давление", "котельная", "педальный",
                "электромагнитный"):
        items = [r for r in rows if key in r["Выделено"]]
        print(f"  выделено «{key}»: {len(items)} позиций, "
              f"{sum(r['Кол-во факт'] or 0 for r in items):.0f} шт")


if __name__ == "__main__":
    main()
