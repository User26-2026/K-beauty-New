#!/usr/bin/env python3
"""Перечень запчастей и комплектующих на складе.

Запчасть без указания, к чему она подходит, не продается. Поэтому
главная колонка здесь — «К какому оборудованию»: она берется из
подзаголовков раздела ЗИП в учетных файлах, а где их нет —
из самого наименования.

Результат:
  outputs/marine_equipment/Запчасти_склад.xlsx
  outputs/marine_equipment/zapchasti_sklad.tsv
"""

import pathlib
import re
import sys
from collections import defaultdict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
from build_marine_registry import SOURCES, parse  # noqa: E402

OUT = pathlib.Path(__file__).resolve().parents[2] / "outputs" / "marine_equipment"

FONT = "Arial"
HEAD_FILL = PatternFill("solid", fgColor="1F3864")
HEAD_FONT = Font(name=FONT, size=11, bold=True, color="FFFFFF")
BASE = Font(name=FONT, size=10)
NOTE = Font(name=FONT, size=9, italic=True, color="666666")
NOPRICE_FILL = PatternFill("solid", fgColor="FFF2CC")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# что считаем запчастью или комплектующим
PART_WORDS = re.compile(
    r"колесо|кольцо|манжет|втулк|шайб|поршень|плунжер|седло клапана|"
    r"направляющая|уплотнен|сальников|обойма|фонарь выхлоп|корпус выхлоп|"
    r"тнвд|кулачков|поплавок|крышка цилиндра|коленвал|ротор|"
    r"клапан всасывающ|клапан выхлопн|клапан нагнетат|клапан пусковой|"
    r"клапан компресс|^фильтр$|стекло к|стекло с решеткой|решетка|дно\)|"
    r"\(дно\)|патрон к|резинки для|сетка на фильтр|барабан|диски тормозн|"
    r"вилка к|амортизатор|клинь|шток", re.I)

# не запчасти: целые изделия и арматура, попадающие под слова выше
NOT_PARTS = re.compile(
    r"^светильник сс-?328 ?\(?(220|127)?v? ?(без|с) ?амортизатор|"
    r"^светильник сс-?328 ?\(ок|^светильник сс-?328 ?\(китай|"
    r"^светильник сс-?328 ?\(эра|^светильник сс-?328 б/у|"
    r"^задвижка|^двигатель б/у|данфас|"
    r"клапан.*\bду ?\d|^клапан для манометра", re.I)

# к какому оборудованию относится, если модель не записана отдельной строкой
BY_NAME = [
    (r"сс[- ]?328", "Светильник СС328"),
    (r"сс[- ]?373", "Светильник СС373"),
    (r"сс[- ]?83[89]|сс[- ]?840", "Светильники СС838-840"),
    (r"сц-?\s?3|сц-?3", "Сепаратор СЦ-3"),
    (r"сц-?\s?1,5|сц-?1,5", "Сепаратор СЦ-1,5"),
    (r"пкэ-?300", "Плита ПКЭ-300"),
    (r"кингстон рс-300", "Кингстон РС-300"),
    (r"светильник импортн", "Светильник импортный"),
    (r"тдп-6а|тормозн", "Тормозные диски"),
    (r"8нвд48а2у", "Двигатель 8NVD48A2U"),
    (r"пилстик", "Двигатель Пилстик"),
    (r"коленвал 18/22", "Двигатель 18/22"),
    (r"сетка на фильтр", "Фильтры забортной воды"),
    (r"светильника импортного|светильник импортн", "Светильник импортный"),
]

COLUMNS = ["№", "К какому оборудованию", "Наименование", "Чертеж", "Наличие",
           "Цена за ед., руб", "Сумма, руб", "Состояние", "Локация"]


# один двигатель записан в файлах и латиницей, и кириллицей
NORMALIZE = {
    "8NVD48A2U": "Двигатель 8NVD48A2U",
    "3Д12": "Двигатель 3Д12",
    "6ЧН25/34": "Двигатель 6ЧН25/34",
    "ЭКП 70/25": "Компрессор ЭКП 70/25",
}


def equipment(r):
    if r["Оборудование"]:
        return NORMALIZE.get(r["Оборудование"], r["Оборудование"])
    name = r["Наименование"]
    for pattern, label in BY_NAME:
        if re.search(pattern, name, re.I):
            return label
    if r["Группа в файле"] == "Барабан":
        return "Сепараторы, корпуса барабанов"
    return "Не определено, нужна дефектовка"


def collect():
    rows = []
    for spec in SOURCES:
        rows.extend(parse(spec))
    for r in rows:
        r["Сумма"] = (r["Наличие"] or 0) * (r["Цена за ед., руб"] or 0)

    parts = []
    for r in rows:
        if (r["Наличие"] or 0) <= 0:
            continue
        if NOT_PARTS.search(r["Наименование"]):
            continue
        is_part = (r["Подгруппа"] == "Двигатели и ЗИП"
                   or r["Оборудование"]
                   or r["Группа в файле"] in ("ЗИП", "Барабан")
                   or PART_WORDS.search(r["Наименование"]))
        if is_part:
            r["Оборудование итог"] = equipment(r)
            parts.append(r)
    return parts


def fmt(v):
    if v is None or v == 0:
        return ""
    if isinstance(v, float):
        return str(int(v)) if v == int(v) else f"{v:.2f}"
    return str(v)


def write_tsv(path, parts):
    with path.open("w", encoding="utf-8", newline="") as fh:
        fh.write("\t".join(COLUMNS) + "\n")
        for n, r in enumerate(parts, start=1):
            fh.write("\t".join([
                str(n), r["Оборудование итог"], r["Наименование"], r["Чертеж"],
                fmt(r["Наличие"]), fmt(r["Цена за ед., руб"]), fmt(r["Сумма"]),
                r["Состояние"], r["Локация"],
            ]) + "\n")


def write_xlsx(path, parts):
    wb = Workbook()
    ws = wb.active
    ws.title = "Запчасти"
    ws["A1"] = "Запчасти и комплектующие на складе"
    ws["A1"].font = Font(name=FONT, size=14, bold=True, color="1F3864")
    ws["A2"] = ("Сгруппировано по оборудованию, к которому подходит запчасть. "
                "Желтым отмечены позиции без цены — их надо оценить.")
    ws["A2"].font = NOTE

    start = 4
    for i, h in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=start, column=i, value=h)
        cell.fill = HEAD_FILL
        cell.font = HEAD_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[start].height = 28

    for n, r in enumerate(parts, start=1):
        row = start + n
        values = [n, r["Оборудование итог"], r["Наименование"], r["Чертеж"],
                  r["Наличие"], r["Цена за ед., руб"] or None,
                  r["Сумма"] or None, r["Состояние"], r["Локация"]]
        for i, v in enumerate(values, start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=(i == 3))
            if i in (6, 7):
                cell.number_format = "#,##0"
        if not r["Цена за ед., руб"]:
            ws.cell(row=row, column=6).fill = NOPRICE_FILL

    last = start + len(parts)
    total = last + 1
    ws.cell(row=total, column=3, value="ИТОГО").font = Font(name=FONT, bold=True)
    for col in (5, 7):
        letter = get_column_letter(col)
        cell = ws.cell(row=total, column=col,
                       value=f"=SUM({letter}{start + 1}:{letter}{last})")
        cell.font = Font(name=FONT, bold=True)
        cell.border = BORDER
        cell.number_format = "#,##0"
    ws.cell(row=total + 1, column=3, value="Позиций без цены").font = BASE
    ws.cell(row=total + 1, column=5,
            value=f"=COUNTBLANK(F{start + 1}:F{last})").font = Font(
                name=FONT, bold=True, color="C00000")

    ws.auto_filter.ref = f"A{start}:I{last}"
    ws.freeze_panes = f"A{start + 1}"
    for i, w in enumerate([5, 30, 58, 22, 10, 15, 14, 13, 15], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # сводка по оборудованию
    ws2 = wb.create_sheet("По оборудованию")
    ws2["A1"] = "Сколько запчастей по каждому оборудованию"
    ws2["A1"].font = Font(name=FONT, size=14, bold=True, color="1F3864")
    heads = ["Оборудование", "Позиций", "Единиц", "Сумма, руб", "Из них без цены"]
    for i, h in enumerate(heads, start=1):
        cell = ws2.cell(row=3, column=i, value=h)
        cell.fill = HEAD_FILL
        cell.font = HEAD_FONT
        cell.border = BORDER

    agg = defaultdict(lambda: [0, 0, 0, 0])
    for r in parts:
        a = agg[r["Оборудование итог"]]
        a[0] += 1
        a[1] += r["Наличие"]
        a[2] += r["Сумма"]
        a[3] += 0 if r["Цена за ед., руб"] else 1

    row = 3
    for name, (p, u, s, np_) in sorted(agg.items(), key=lambda kv: -kv[1][2]):
        row += 1
        for i, v in enumerate([name, p, u, s or None, np_ or None], start=1):
            cell = ws2.cell(row=row, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            if i == 4:
                cell.number_format = "#,##0"
    row += 1
    ws2.cell(row=row, column=1, value="ИТОГО").font = Font(name=FONT, bold=True)
    for col in (2, 3, 4):
        letter = get_column_letter(col)
        cell = ws2.cell(row=row, column=col, value=f"=SUM({letter}4:{letter}{row - 1})")
        cell.font = Font(name=FONT, bold=True)
        cell.number_format = "#,##0"
    for i, w in enumerate([36, 10, 10, 16, 16], start=1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    wb.save(path)


def main():
    parts = collect()
    parts.sort(key=lambda r: (r["Оборудование итог"] == "Не определено, нужна дефектовка",
                              r["Оборудование итог"], -r["Сумма"]))
    OUT.mkdir(parents=True, exist_ok=True)
    write_tsv(OUT / "zapchasti_sklad.tsv", parts)
    write_xlsx(OUT / "Запчасти_склад.xlsx", parts)

    units = sum(r["Наличие"] for r in parts)
    total = sum(r["Сумма"] for r in parts)
    no_price = [r for r in parts if not r["Цена за ед., руб"]]
    print(f"позиций: {len(parts)}, единиц: {units:.0f}, сумма: {total:,.0f} руб"
          .replace(",", " "))
    print(f"без цены: {len(no_price)} позиций, "
          f"{sum(r['Наличие'] for r in no_price):.0f} единиц")

    agg = defaultdict(lambda: [0, 0, 0])
    for r in parts:
        a = agg[r["Оборудование итог"]]
        a[0] += 1
        a[1] += r["Наличие"]
        a[2] += r["Сумма"]
    print("\nпо оборудованию:")
    for name, (p, u, s) in sorted(agg.items(), key=lambda kv: -kv[1][2]):
        print(f"  {name:<34} поз {p:>3}  ед {u:>5.0f}  {s:>11,.0f}".replace(",", " "))


if __name__ == "__main__":
    main()
