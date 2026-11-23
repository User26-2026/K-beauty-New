#!/usr/bin/env python3
"""Сравнение наших договорных цен с прайсом конкурента с Камчатки.

Прайс прислан 03.10.2026, это судовая запорная арматура: 1 017 строк
в двух колонках на листе, 744 чертежа с ценами. Новая складского
хранения помечена словом New, остальное восстановленное.

Сверяем двумя проходами. Сначала по номеру чертежа — это надежно,
один чертеж означает одно и то же изделие. Потом по типу и диаметру —
так закрываются наши позиции без чертежа, но совпадение приблизительное:
у конкурента в одной строке может стоять бронза, а у нас сталь.

Наши цены местами записаны диапазоном и в разных единицах: «130-150»
это тысячи, «25-35000» это рубли. Разбираем с приведением к рублям,
сомнительные помечаем.

Источник: data/marine_equipment/competitors/Прайс_Камчатка_03.10.26.xlsx
Результат: outputs/marine_equipment/Сравнение_с_конкурентом.xlsx
"""

import pathlib
import re
import sys

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = pathlib.Path(__file__).resolve().parents[2]
SRC = (ROOT / "data" / "marine_equipment" / "competitors"
       / "Прайс_Камчатка_03.10.26.xlsx")
OUT = ROOT / "outputs" / "marine_equipment"

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
from build_spb_list_xlsx import GROUPS, read as read_spb  # noqa: E402

FONT = "Arial"
HEAD_FILL = PatternFill("solid", fgColor="1F3864")
HEAD_FONT = Font(name=FONT, size=11, bold=True, color="FFFFFF")
TITLE = Font(name=FONT, size=14, bold=True, color="1F3864")
SUB = Font(name=FONT, size=11, bold=True, color="1F3864")
BOLD = Font(name=FONT, size=10, bold=True)
BASE = Font(name=FONT, size=10)
NOTE = Font(name=FONT, size=9, italic=True, color="666666")
LOW_FILL = PatternFill("solid", fgColor="FCE4D6")   # мы дешевле
HIGH_FILL = PatternFill("solid", fgColor="EAF3EA")  # мы дороже
WARN_FILL = PatternFill("solid", fgColor="FFF2CC")
GROUP_FILL = PatternFill("solid", fgColor="D9E2F3")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# раздел прайса конкурента -> наша группа
SECTIONS = [
    (r"задвижка клинкетная", "Клинкеты"),
    (r"фильтр забортной воды", "Фильтры забортной воды"),
    (r"кингстон", "Кингстоны"),
    (r"захлопка", "Захлопки"),
    (r"колонка\s+указательная", "Колонка смотровая"),
    (r"коробка", "Коробки клапанные"),
    (r"кран", "Краны"),
    (r"клапан.*углов.*штуцерн|клапан.*штуцерн.*углов",
     "Штуцерная угловая"),
    (r"клапан.*проходн.*штуцерн|клапан.*штуцерн.*проходн",
     "Штуцерная проходная"),
    (r"клапан.*углов.*фланцев|клапан.*фланцев.*углов",
     "Фланцевая угловая"),
    (r"клапан.*проходн.*фланцев|клапан.*фланцев.*проходн",
     "Фланцевая проходная"),
    (r"манометров", "Штуцерная проходная"),
    (r"предохранительн", "Штуцерная угловая"),
    # эти два раздела висели без нашей группы, хотя группы под них есть:
    # редукционные штуцерные и электромагнитные. Фланцевые редукционные
    # сюда не попадают, их раньше забирает правило фланцевой проходной
    (r"редукционн", "Редукционные клапаны"),
    (r"электромагнитн", "Электромагнитные клапаны"),
]


def norm_draw(d):
    """Чертеж к одному виду: без пробелов, без хвостовых точек."""
    if d is None:
        return ""
    s = str(d).upper().replace(" ", "").replace("…", "")
    s = s.replace("–", "-").replace("—", "-")
    return s.strip(".")


def our_price(text):
    """Наша цена по договору в рубли.

    В приложениях встречается «55000», «30-35» (тысячи), «25-35000»
    (рубли) и «по согласованию». Возвращаем минимум, максимум и
    пометку, если запись сомнительная.
    """
    s = str(text or "").strip().lower()
    if not s or "согласован" in s:
        return None, None, "цена не согласована"
    nums = re.findall(r"\d+", s.replace(" ", ""))
    if not nums:
        return None, None, "цена не читается"
    vals = [int(n) for n in nums]
    vals = [v * 1000 if v < 1000 else v for v in vals]
    lo, hi = min(vals), max(vals)
    flag = ""
    if hi > lo * 3:
        flag = "запись сомнительная, проверить"
    return lo, hi, flag


def section_group(header):
    low = (header or "").lower()
    for pattern, group in SECTIONS:
        if re.search(pattern, low):
            return group
    return ""


def du_of(text):
    m = re.match(r"\s*(\d{1,3})", str(text or ""))
    return int(m.group(1)) if m else None


def read_competitor():
    wb = load_workbook(SRC, data_only=True)
    ws = wb["Перечень СЗА"]
    rows = []
    # лист сверстан в две колонки: позиции идут и слева, и справа
    for block in ((1, 2, 3, 4, 6), (8, 9, 10, 11, 13)):
        header = ""
        for r in range(5, ws.max_row + 1):
            cell = ws.cell(row=r, column=block[0]).value
            if cell is None:
                continue
            s = str(cell).strip()
            if not re.match(r"^\d|^№|^\d+\s*\"|^1 1/2", s) and len(s) > 12:
                header = s
                continue
            price = ws.cell(row=r, column=block[4]).value
            rows.append({
                "Раздел": header,
                "Группа": section_group(header),
                "Ду текст": s,
                "Ду": du_of(s),
                "Ру": ws.cell(row=r, column=block[1]).value,
                "Чертеж": norm_draw(ws.cell(row=r, column=block[2]).value),
                "Наличие": ws.cell(row=r, column=block[3]).value,
                "Цена": price if isinstance(price, (int, float)) else None,
                "Новая": "new" in s.lower(),
                # конкурент помечает стальные «(ст.)», и они втрое
                # дешевле бронзовых того же диаметра
                "Материал": ("сталь" if "ст." in s.lower()
                             else "бронза" if "бронз" in header.lower()
                             else ""),
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


COLS = [
    ("Группа", 24),
    ("Наименование у нас", 46),
    ("Чертеж", 20),
    ("Ду", 7),
    ("Штук", 8),
    ("Наша цена по договору", 18),
    ("Наша цена, руб", 14),
    ("Цена конкурента, руб", 15),
    ("Разница, руб", 13),
    ("Разница", 11),
    ("Позиция конкурента", 30),
    ("Наличие у конкурента", 13),
    ("Пометка", 30),
]


def sheet_match(wb, name, title, pairs, note):
    ws = wb.create_sheet(name)
    ws["A1"] = title
    ws["A1"].font = TITLE
    ws["A2"] = note
    ws["A2"].font = NOTE

    start = 4
    for i, (h, _) in enumerate(COLS, start=1):
        ws.cell(row=start, column=i, value=h)
    head(ws, start, len(COLS))

    row = start
    order = {g: i for i, g in enumerate(GROUPS)}
    pairs = sorted(pairs, key=lambda p: (order.get(p["Группа"], 99),
                                         p["Ду"] or 0))
    for p in pairs:
        row += 1
        values = [p["Группа"], p["Наименование"], p["Чертеж"], p["Ду"],
                  p["Штук"], p["Цена текст"], p["Наша"], p["Конкурент"],
                  p["Разница"], p["Доля"], p["Позиция конкурента"],
                  p["Наличие"], p["Пометка"]]
        for i, v in enumerate(values, start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top",
                                       wrap_text=(i in (2, 11, 13)))
            if i in (7, 8, 9):
                cell.number_format = "#,##0"
            if i == 10:
                cell.number_format = "+0%;-0%"
        if p["Разница"] is not None:
            fill = LOW_FILL if p["Разница"] < 0 else HIGH_FILL
            ws.cell(row=row, column=9).fill = fill
            ws.cell(row=row, column=10).fill = fill
        if p["Пометка"]:
            ws.cell(row=row, column=13).fill = WARN_FILL

    ws.auto_filter.ref = f"A{start}:{get_column_letter(len(COLS))}{row}"
    ws.freeze_panes = f"C{start + 1}"
    widths(ws, [w for _, w in COLS])
    return ws


TYPE_COLS = [
    ("Группа", 24),
    ("Наименование у нас", 46),
    ("Ду", 7),
    ("Штук", 8),
    ("Наша цена по договору", 18),
    ("Наша цена, руб", 14),
    ("Конкурент бронза, от", 15),
    ("Конкурент бронза, до", 15),
    ("Конкурент сталь, от", 14),
    ("Конкурент сталь, до", 14),
    ("Строк у конкурента", 12),
    ("Где мы относительно бронзы", 18),
    ("Пометка", 28),
]


def sheet_by_type(wb, spb, comp):
    """Сверка по типу и диаметру, с разделением бронзы и стали.

    Усреднять тут нельзя: у конкурента задвижка Ду200 бронзовая стоит
    150 000, а стальная того же диаметра 40 000. Среднее между ними
    не значит ничего.
    """
    ws = wb.create_sheet("Совпало по типу")
    ws["A1"] = "Сверка по типу и диаметру"
    ws["A1"].font = TITLE
    ws["A2"] = ("Вся наша арматура бронзовая или латунная, поэтому "
                "сравниваем с бронзовой колонкой конкурента. Стальная "
                "приведена справочно: у него она втрое дешевле, и "
                "путать их нельзя.")
    ws["A2"].font = NOTE

    start = 4
    for i, (h, _) in enumerate(TYPE_COLS, start=1):
        ws.cell(row=start, column=i, value=h)
    head(ws, start, len(TYPE_COLS))

    buckets = {}
    for r in comp:
        if r["Группа"] and r["Ду"] and r["Цена"]:
            buckets.setdefault((r["Группа"], r["Ду"]), []).append(r)

    order = {g: i for i, g in enumerate(GROUPS)}
    rows = []
    for o in spb:
        if not o["Кол-во факт"]:
            continue
        key = (o["Группа"], o["Ду"])
        if key not in buckets:
            continue
        lo, hi, flag = our_price(o["Цена по договору"])
        ours = (lo + hi) / 2 if lo else None
        same = buckets[key]
        bronze = [x["Цена"] for x in same if x["Материал"] != "сталь"]
        steel = [x["Цена"] for x in same if x["Материал"] == "сталь"]
        # вся наша арматура бронзовая или латунная, поэтому сравниваем
        # с бронзовой колонкой конкурента, стальная приведена справочно
        verdict = ""
        if ours and bronze:
            if ours < min(bronze):
                verdict = f"дешевле на {(1 - ours / min(bronze)) * 100:.0f}%"
            elif ours > max(bronze):
                verdict = f"дороже на {(ours / max(bronze) - 1) * 100:.0f}%"
            else:
                verdict = "внутри вилки"
        elif ours and steel:
            verdict = "у конкурента только сталь, сравнивать не с чем"
        rows.append((o, ours, bronze, steel, verdict, flag))

    rows.sort(key=lambda t: (order.get(t[0]["Группа"], 99),
                             -(t[0]["Ду"] or 0)))
    row = start
    for o, ours, bronze, steel, verdict, flag in rows:
        row += 1
        values = [o["Группа"], o["Наименование"], o["Ду"],
                  o["Кол-во факт"], o["Цена по договору"], ours,
                  min(bronze) if bronze else None,
                  max(bronze) if bronze else None,
                  min(steel) if steel else None,
                  max(steel) if steel else None,
                  len(bronze) + len(steel), verdict, flag]
        for i, v in enumerate(values, start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top",
                                       wrap_text=(i in (2, 13)))
            if i in (6, 7, 8, 9, 10):
                cell.number_format = "#,##0"
        if verdict.startswith("дешевле"):
            ws.cell(row=row, column=12).fill = LOW_FILL
        elif verdict.startswith("дороже"):
            ws.cell(row=row, column=12).fill = HIGH_FILL
        if flag:
            ws.cell(row=row, column=13).fill = WARN_FILL

    ws.auto_filter.ref = f"A{start}:{get_column_letter(len(TYPE_COLS))}{row}"
    ws.freeze_panes = f"C{start + 1}"
    widths(ws, [w for _, w in TYPE_COLS])
    return ws, rows


def sheet_price(wb, rows):
    ws = wb.create_sheet("Прайс конкурента")
    ws["A1"] = "Прайс с Камчатки, разобранный в одну таблицу"
    ws["A1"].font = TITLE
    ws["A2"] = ("Исходный лист сверстан в две колонки, здесь они сведены. "
                "New означает новую складского хранения, остальное "
                "восстановленное.")
    ws["A2"].font = NOTE

    cols = ["Раздел", "Наша группа", "Ду", "Ру", "Чертеж", "Материал",
            "Наличие", "Цена, руб", "Новая"]
    for i, h in enumerate(cols, start=1):
        ws.cell(row=4, column=i, value=h)
    head(ws, 4, len(cols), height=28)

    row = 4
    for r in sorted(rows, key=lambda r: (r["Раздел"], -(r["Ду"] or 0))):
        row += 1
        values = [r["Раздел"], r["Группа"], r["Ду"], r["Ру"], r["Чертеж"],
                  r["Материал"], r["Наличие"], r["Цена"],
                  "да" if r["Новая"] else ""]
        for i, v in enumerate(values, start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=(i == 1))
            if i == 8:
                cell.number_format = "#,##0"
    ws.auto_filter.ref = f"A4:{get_column_letter(len(cols))}{row}"
    ws.freeze_panes = "C5"
    widths(ws, [44, 24, 7, 8, 20, 10, 10, 13, 8])
    return ws


def build_pairs(spb, comp, by_draw=True):
    """Сопоставление наших позиций с прайсом конкурента."""
    draws = {}
    for r in comp:
        if r["Чертеж"] and r["Цена"]:
            # берем первую цену по чертежу: дальше в прайсе идут дубли
            draws.setdefault(r["Чертеж"], r)

    by_group = {}
    for r in comp:
        if r["Группа"] and r["Ду"] and r["Цена"]:
            by_group.setdefault((r["Группа"], r["Ду"]), []).append(r)

    pairs = []
    for o in spb:
        if not o["Кол-во факт"]:
            continue
        lo, hi, flag = our_price(o["Цена по договору"])
        ours = (lo + hi) / 2 if lo else None
        match = None
        how = ""
        draw = norm_draw(o["Обозначение"])
        if by_draw and draw and draw in draws:
            match = draws[draw]
            how = "по чертежу"
        elif not by_draw:
            key = (o["Группа"], o["Ду"])
            if key in by_group:
                same = by_group[key]
                prices = [x["Цена"] for x in same]
                match = dict(same[0], Цена=sum(prices) / len(prices))
                how = f"по типу и диаметру, {len(same)} строк у конкурента"
        if not match:
            continue
        theirs = match["Цена"]
        diff = (ours - theirs) if ours else None
        pairs.append({
            "Группа": o["Группа"],
            "Наименование": o["Наименование"],
            "Чертеж": o["Обозначение"],
            "Ду": o["Ду"],
            "Штук": o["Кол-во факт"],
            "Цена текст": o["Цена по договору"],
            "Наша": ours,
            "Конкурент": theirs,
            "Разница": diff,
            "Доля": (diff / theirs) if (diff is not None and theirs) else None,
            "Позиция конкурента": f"{match['Раздел'][:28]} Ду{match['Ду']}",
            "Наличие": match["Наличие"],
            "Пометка": "; ".join(x for x in (flag, how) if x),
        })
    return pairs


def sheet_summary(wb, by_draw, type_rows, comp):
    ws = wb.create_sheet("Итоги", 0)
    ws["A1"] = "Наши цены против прайса с Камчатки"
    ws["A1"].font = TITLE
    ws["A2"] = ("Сравниваем цену по договору — ту, которую покупатель "
                "платит нам. Цена конкурента — его цена продажи, то есть "
                "потолок рынка, а не закупка.")
    ws["A2"].font = NOTE

    priced = [p for p in by_draw if p["Разница"] is not None]
    cheaper = [p for p in priced if p["Разница"] < 0]
    dearer = [p for p in priced if p["Разница"] > 0]

    row = 4
    facts = [
        ("Строк в прайсе конкурента", len(comp), ""),
        ("Из них с ценой", sum(1 for r in comp if r["Цена"]), ""),
        ("Совпало по номеру чертежа", len(by_draw),
         "надежное сравнение"),
        ("Совпало по типу и диаметру", len(type_rows),
         "приблизительное: материал и исполнение могут отличаться"),
        ("", "", ""),
        ("Мы дешевле конкурента", len(cheaper),
         f"из {len(priced)} сравнимых позиций"),
        ("Мы дороже конкурента", len(dearer), ""),
        ("Средняя разница",
         (sum(p["Доля"] for p in priced) / len(priced)) if priced else 0,
         "минус означает, что мы продаем дешевле"),
    ]
    for label, value, note in facts:
        ws.cell(row=row, column=1, value=label).font = BOLD
        cell = ws.cell(row=row, column=2, value=value)
        cell.font = BOLD
        if label == "Средняя разница":
            cell.number_format = "+0%;-0%"
        ws.cell(row=row, column=3, value=note).font = NOTE
        row += 1

    row += 1
    ws.cell(row=row, column=1,
            value="Где мы дешевле всего, по чертежам").font = SUB
    row += 1
    for i, h in enumerate(["Наименование", "Чертеж", "Наша цена",
                           "Конкурент", "Разница"], start=1):
        ws.cell(row=row, column=i, value=h)
    head(ws, row, 5, height=24)
    for p in sorted(cheaper, key=lambda p: p["Доля"])[:12]:
        row += 1
        for i, v in enumerate([p["Наименование"], p["Чертеж"], p["Наша"],
                               p["Конкурент"], p["Доля"]], start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=(i == 1))
            if i in (3, 4):
                cell.number_format = "#,##0"
            if i == 5:
                cell.number_format = "+0%;-0%"
                cell.fill = LOW_FILL

    row += 2
    ws.cell(row=row, column=1, value="Что из этого следует").font = SUB
    row += 1
    for text in [
        "Прайс конкурента — это цены продажи конечному покупателю, а не "
        "закупка. Сравнивать с ними надо не нашу цену по договору, а "
        "цену, по которой покупатель перепродает наш товар.",
        "Наша цена по договору ниже камчатских почти везде. Это нормально: "
        "мы продаем оптом одной партией, он поштучно под заявку. Но "
        "разрыв показывает, сколько зарабатывает покупатель на перепродаже.",
        "По договору цена реализации согласуется с нами письменно "
        "заранее. Этот прайс — готовый аргумент в таком согласовании: "
        "видно, по какой цене то же изделие стоит на рынке.",
        "Конкурент отдельно помечает новое складского хранения словом "
        "New, остальное у него восстановленное. У нас состояние "
        "проставлено не везде, а новое стоит дороже — это надо "
        "зафиксировать при погрузке.",
        "Вся наша арматура бронзовая или латунная, включая все клинкеты. "
        "Стальными идут только паровой и редукционные клапаны. Поэтому "
        "сравниваем с бронзовой колонкой конкурента: его стальные "
        "позиции втрое дешевле и к нам отношения не имеют.",
    ]:
        cell = ws.cell(row=row, column=1, value="— " + text)
        cell.font = BASE
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[row].height = 32
        row += 1

    widths(ws, [46, 22, 16, 16, 14])
    return ws


def main():
    comp = read_competitor()
    spb = read_spb()
    by_draw = build_pairs(spb, comp, by_draw=True)

    wb = Workbook()
    wb.remove(wb.active)
    ws_type, type_rows = sheet_by_type(wb, spb, comp)
    sheet_summary(wb, by_draw, type_rows, comp)
    sheet_match(wb, "Совпало по чертежу",
                "Сверка по номеру чертежа", by_draw,
                "Один чертеж — одно изделие. Это надежное сравнение.")
    wb.move_sheet("Совпало по типу", offset=1)
    sheet_price(wb, comp)

    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / "Сравнение_с_конкурентом.xlsx"
    wb.save(path)

    priced = [p for p in by_draw if p["Разница"] is not None]
    print(f"готово: {path}\n")
    print(f"прайс конкурента: {len(comp)} строк, "
          f"{sum(1 for r in comp if r['Цена'])} с ценой")
    print(f"совпало по чертежу: {len(by_draw)}, из них с нашей ценой "
          f"{len(priced)}")
    print(f"совпало по типу и диаметру: {len(type_rows)}")
    if priced:
        avg = sum(p["Доля"] for p in priced) / len(priced)
        print(f"мы дешевле в {sum(1 for p in priced if p['Разница'] < 0)} "
              f"случаях из {len(priced)}, средняя разница {avg * 100:+.0f}%")


if __name__ == "__main__":
    main()
