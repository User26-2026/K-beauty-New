#!/usr/bin/env python3
"""Что осталось на складе: остаток по товарным группам.

Менеджер прислала остатки за сентябрь 2026 по одиннадцати локациям,
десятью файлами. ОТГРУЖЕННОЕ ИЗ НИХ УЖЕ СНЯТО — во всех, а не только
в сводном. Это первое, что надо знать, потому что вычитать отгрузки
второй раз нельзя.

Доказательства. По складу 1А и 1Б в сводном на 24 позиции и 174 шт
меньше, чем в файле с ценами, и это ровно контейнерные задвижка Ду80
36 шт, клапан Ду50 Ру10/25 47 шт, головки воздушно-квадратные, грибки
АМГ и задвижка Ду250 17 шт. По фойе убраны все клинкеты и все фильтры
забортной воды, включая керченские 45 задвижек Ду150 и питерские
фильтры Ду300. Количества при этом не разошлись ни по одной позиции:
позиция либо есть целиком, либо убрана целиком — так выглядит вычет
отгруженного, а не пересчет склада.

По отдельным файлам то же самое, если сверить со старым учетом по
наличию на 16.06.2025: склад 2 был 108 позиций и 688 шт, стал 35 и
203; склад 3 был 235 и 4 260, стал 189 и 3 518; менеджерская была 35
и 1 410, стала 9 и 95. Пропала штуцерная мелочь, переборочные стаканы,
клапанные коробки, клапан фланцевый проходной Ду50 и задвижка
клинкетная Ду100 — то есть содержимое контейнера. 798 клапанов Ду10
Ру40 («СУНДУК») лежали в менеджерской, и их там больше нет: в
контейнер уехало 796 штук этой позиции.

Поэтому файл не вычитает отгрузки, а делает две другие вещи.

Первое: раскладывает остаток по товарным группам так, как их просил
собственник — арматура, электрика, насосы, ЗИП, корпусная часть
(двери, иллюминаторы, стекла к иллюминаторам), сепараторы, камбузное
оборудование, блокформы, прочее. Внутри ЗИП подгруппа — модель
оборудования: без нее запчасть не продать.

Второе: сверяет сентябрьский остаток со старым учетом и показывает,
что пропало ПОМИМО отгрузок. Сошлось или не сошлось считается по
группе, диаметру и материалу, а не по строке: в учете одна строка на
47 клапанов, а в отгрузке те же клапаны разбиты по чертежам, и по
названиям они не совпадут никогда. Сталь с бронзой не сводится —
у конкурента стальная задвижка Ду200 стоит 40 000 против 150 000 за
бронзовую, это разный товар.

Что не сошлось — это не обязательно потеря. Это может быть продажа
до отгрузок, может быть пересчет, а может быть и то, что позицию
просто не нашли. Проверять глазами, а не считать.

Результат: outputs/marine_equipment/Остаток_склада_по_группам.xlsx
"""

import pathlib
import sys
from collections import defaultdict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
import build_marine_registry as reg  # noqa: E402
import container_list  # noqa: E402
import stock_groups as sg  # noqa: E402
import stock_sheets  # noqa: E402
import upd_shipments  # noqa: E402
from build_contract_annexes import norm_draw  # noqa: E402

OUT = (pathlib.Path(__file__).resolve().parents[2] / "outputs"
       / "marine_equipment")

# Локация склада: как называется, учетный файл, откуда взят остаток за
# сентябрь 2026. None в последнем поле — остатка за сентябрь нет.
LOCATIONS = [
    ("Фойе", "foye.xls", "сводный, лист «Фойе,цех»"),
    ("Цех", "foye.xls", "сводный, лист «Фойе,цех»"),
    ("Склад 1А, 1Б", "sklad_1a_1b.xlsx", "сводный, лист «Склад 1А, 1Б»"),
    ("Гараж 2", "garazh_2.xlsx", "сводный, лист «Гараж 2»"),
    ("Контейнер 2", "konteiner_2.xls", "сводный, лист «Контейнер 2»"),
    ("Улица", "нет учетного файла", "сводный, лист «Улица»"),
    ("Склад 2", "sklad_2.xls", "Склад_2_сент2026.xls"),
    ("Склад 3", "sklad_3.xls", "Склад_3_сент2026.xls"),
    ("Склад электрики", "sklad_elektriki.xls",
     "Склад_электрики_сент2026.xls"),
    ("Гараж 1", "garazh_1.xls", "Гараж_1_сент2026.xls"),
    ("Менеджерская", "menedzherskaya.xls", "Менеджерская_сент2026.xls"),
    ("Контейнер 1", "konteiner_1.xls", None),
    ("ЗИП и насосы", "zip_nasosy.csv", None),
]

# Как лист сентябрьского файла называется в старом учете.
TO_OLD = {
    "Фойе,цех": "Фойе",
    "Склад 1А, 1Б": "Склад 1А/1Б",
    "Гараж 2": "Гараж 2",
    "Контейнер 2": "Контейнер 2",
    "Склад 2": "Склад 2",
    "Склад3": "Склад 3",
    "Склад электрики": "Склад электрики",
    "гараж 1": "Гараж 1",
    "склад менеджерская": "Менеджерская",
    "Улица ": "ЗИП и насосы",
}

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
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def rub(n):
    return f"{n:,.0f}".replace(",", " ")


def head(ws, row, ncols, height=36):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEAD_FILL
        cell.font = HEAD_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center",
                                   horizontal="center")
        cell.border = BORDER
    ws.row_dimensions[row].height = height


def put(ws, row, values, wrap=(), center=(), money=(), fill=None, font=None):
    for i, v in enumerate(values, start=1):
        cell = ws.cell(row=row, column=i, value=v)
        cell.font = font or BASE
        cell.border = BORDER
        cell.alignment = Alignment(
            vertical="top", wrap_text=(i in wrap),
            horizontal="center" if i in center else None)
        if i in money:
            cell.number_format = "#,##0"
        if fill:
            cell.fill = fill
    return row + 1


def widths(ws, values):
    for i, w in enumerate(values, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def notes(ws, row, ncols, texts, height=28):
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


def shipped():
    """Все отгруженное: контейнер плюс три ранние партии."""
    out = []
    for r in container_list.read():
        out.append({"Откуда": "Контейнер 26.10.2026",
                    "Наименование": r["Наименование"],
                    "Чертеж": r["Чертеж"], "Штук": r["Штук"]})
    for s in upd_shipments.SHIPMENTS:
        for it in s["items"]:
            out.append({"Откуда": f"{s['title']}, {s['shipped']}",
                        "Наименование": it["name"],
                        "Чертеж": it.get("draw", ""), "Штук": it["qty"]})
    for r in out:
        r["Группа"], r["Подгруппа"] = sg.classify(r["Наименование"])
        r["Ду"] = sg.du_of(r["Наименование"])
        r["Сталь"] = sg.not_bronze(r["Наименование"])
    return out


def enrich(stock):
    """Группа, подгруппа, диаметр, цена и пометки по каждой позиции."""
    rows = []
    for r in stock:
        g, sub = sg.classify(r["Наименование"], r.get("Раздел", ""),
                             r["Подгруппа"])
        du = sg.du_of(r["Наименование"]) if g == sg.ARMATURE else 0
        price = r["Цена"] or 0
        marks = []
        if sg.broken(r["Наименование"]):
            marks.append("брак")
        elif sg.flawed(r["Наименование"]):
            marks.append("некомплект или в разборе")
        if r.get("Пересчет"):
            marks.append(r["Пересчет"])
        if not price:
            marks.append("цены нет")
        rows.append(dict(r, Группа=g, Подгруппа2=sub, Ду=du, Цена1=price,
                         Сумма1=round(price * r["Штук"]),
                         Пометки=", ".join(marks),
                         Брак=sg.broken(r["Наименование"]),
                         Дефект=sg.flawed(r["Наименование"])))
    return rows


def key_of(name, sub, du, steel):
    return (sub, du, steel)


def reconcile(rows, ship):
    """Сверка сентябрьского остатка со старым учетом.

    По каждой локации: сколько было по наличию на 16.06.2025, сколько
    стало, сколько пропало и сколько из пропавшего объясняется
    отгрузками. Сопоставление по группе, диаметру и материалу: по
    названиям строки учета и строки отгрузки не совпадают.
    """
    old = []
    for spec in reg.SOURCES:
        for r in reg.parse(spec):
            if not r["Наличие"]:
                continue
            g, sub = sg.classify(r["Наименование"], r.get("Группа в файле", ""),
                                 r.get("Оборудование", ""))
            old.append({
                "Локация": r["Локация"], "Наименование": r["Наименование"],
                "Чертеж": r["Чертеж"] or "", "Штук": int(r["Наличие"]),
                "Группа": g, "Подгруппа": sub,
                "Ду": sg.du_of(r["Наименование"]) if g == sg.ARMATURE else 0,
                "Сталь": sg.not_bronze(r["Наименование"]),
                "Цена": r["Цена за ед., руб"] or 0,
            })

    # сентябрьский остаток в тех же ключах
    now = defaultdict(int)
    now_loc = defaultdict(int)
    now_sub = defaultdict(int)
    for r in rows:
        loc = TO_OLD.get(r["Лист"], r["Лист"])
        now[(loc,) + key_of(r["Наименование"], r["Подгруппа2"], r["Ду"],
                            sg.not_bronze(r["Наименование"]))] += r["Штук"]
        now_loc[loc] += r["Штук"]
        now_sub[(loc, r["Подгруппа2"])] += r["Штук"]
    sent = {name for name, _, got in LOCATIONS if got}
    sent |= {"Склад 1А/1Б"} if "Склад 1А, 1Б" in sent else set()
    sent |= {"ЗИП и насосы"} if "Улица" in sent else set()

    # отгруженное в тех же ключах, без локации
    gone = defaultdict(int)
    for s in ship:
        gone[key_of(s["Наименование"], s["Подгруппа"], s["Ду"],
                    s["Сталь"])] += s["Штук"]

    # что пропало по каждому ключу локации
    missing = []
    by_loc = defaultdict(lambda: [0, 0, 0, 0, 0, 0])
    old_loc = defaultdict(lambda: [0, 0])
    for r in old:
        k = (r["Локация"],) + key_of(r["Наименование"], r["Подгруппа"],
                                     r["Ду"], r["Сталь"])
        old_loc[r["Локация"]][0] += 1
        old_loc[r["Локация"]][1] += r["Штук"]

    # Три прохода. Первый: та же позиция в той же локации. Второй:
    # отгрузки. Третий: та же позиция, но в ДРУГОЙ локации — товар по
    # складу переставляли, блокформы из перечня ЗИП теперь числятся в
    # контейнере 2, и без этого прохода они выглядели бы недостачей.
    seen = defaultdict(int)
    stage1 = []
    for r in sorted(old, key=lambda x: -x["Штук"]):
        k = (r["Локация"],) + key_of(r["Наименование"], r["Подгруппа"],
                                     r["Ду"], r["Сталь"])
        left = now[k] - seen[k]
        keep = max(0, min(r["Штук"], left))
        seen[k] += keep
        lost = r["Штук"] - keep
        if lost <= 0:
            continue
        gk = k[1:]
        took = min(lost, gone[gk])
        gone[gk] -= took
        stage1.append((r, lost, took, lost - took))

    # свободный остаток по каждому ключу, где бы он ни лежал
    free = defaultdict(int)
    for k, v in now.items():
        rest = v - seen[k]
        if rest > 0:
            free[k[1:]] += rest

    for r, lost, took, rest in stage1:
        gk = key_of(r["Наименование"], r["Подгруппа"], r["Ду"], r["Сталь"])
        moved = min(rest, free[gk])
        free[gk] -= moved
        rest -= moved
        a = by_loc[r["Локация"]]
        a[0] += 1
        a[1] += lost
        a[2] += took
        a[3] += rest
        a[5] += moved
        if rest <= 0:
            continue
        a[4] += 1
        if r["Локация"] not in sent:
            why = "локация за сентябрь не прислана"
        elif now_sub.get((r["Локация"], r["Подгруппа"]), 0) == 0:
            why = "вся подгруппа в сентябрьский файл не попала"
        else:
            why = "подгруппа в файле есть, этой позиции нет"
        missing.append(dict(r, Пропало=lost, Отгружено=took,
                            Переехало=moved, Непонятно=rest, Почему=why,
                            Сумма=round((r["Цена"] or 0) * rest)))
    return old, old_loc, now_loc, by_loc, missing


def sheet_groups(wb, rows):
    ws = wb.active
    ws.title = "По группам"
    agg = defaultdict(lambda: defaultdict(lambda: [0, 0, 0.0]))
    for r in rows:
        a = agg[r["Группа"]][r["Подгруппа2"]]
        a[0] += 1
        a[1] += r["Штук"]
        a[2] += r["Сумма1"]
    total_sum = sum(v[2] for g in agg.values() for v in g.values())

    ws["A1"] = "Что осталось на складе: по группам товара"
    ws["A1"].font = TITLE
    have = sum(1 for *_, g in LOCATIONS if g)
    priced = sum(1 for r in rows if r["Цена1"])
    ws["A2"] = (
        f"Остатки менеджера за сентябрь 2026 по {have} локациям из "
        f"{len(LOCATIONS)}. Отгруженное из них УЖЕ СНЯТО — во всех "
        f"файлах, а не только в сводном, поэтому второй раз отгрузки не "
        f"вычитались. Не присланы контейнер 1 и перечень ЗИП насосов. "
        f"Цена есть у {priced} позиций из {len(rows)}, по остальным "
        f"сумма не считается.")
    ws["A2"].font = NOTE

    cols = [("Группа товара", 30), ("Подгруппа", 40), ("Позиций", 9),
            ("Штук", 10), ("Сумма, руб", 16), ("Доля суммы", 10),
            ("Позиций без цены", 11)]
    start = 4
    for i, (h, _) in enumerate(cols, start=1):
        ws.cell(row=start, column=i, value=h)
    head(ws, start, len(cols))
    row = start + 1
    noprice = defaultdict(lambda: defaultdict(int))
    for r in rows:
        if not r["Цена1"]:
            noprice[r["Группа"]][r["Подгруппа2"]] += 1
    for g in sorted(agg, key=lambda x: -sum(v[2] for v in agg[x].values())):
        subs = agg[g]
        tot = [sum(v[i] for v in subs.values()) for i in range(2)]
        tots = sum(v[2] for v in subs.values())
        row = put(ws, row, [
            g, "", tot[0], tot[1], round(tots),
            f"{tots / total_sum:.1%}" if total_sum else "",
            sum(noprice[g].values()),
        ], center=(3, 4, 6, 7), money=(5,), font=BOLD, fill=GROUP_FILL)
        for sub, v in sorted(subs.items(), key=lambda kv: -kv[1][2]):
            row = put(ws, row, ["", sub, v[0], v[1], round(v[2]), "",
                                noprice[g].get(sub, 0) or ""],
                      wrap=(2,), center=(3, 4, 7), money=(5,))
    row = put(ws, row, [
        "ИТОГО", "", len(rows), sum(r["Штук"] for r in rows),
        round(total_sum), "100%", sum(1 for r in rows if not r["Цена1"]),
    ], center=(3, 4, 6, 7), money=(5,), font=BOLD, fill=OK_FILL)

    row += 1
    row = notes(ws, row, len(cols), [
        "Группы заданы так, как их просил собственник: арматура, "
        "электрика, насосы, ЗИП, корпусная часть, сепараторы, камбузное "
        "оборудование, блокформы, прочее. Корпусная часть — это двери, "
        "иллюминаторы и стекла к иллюминаторам.",
        "Внутри ЗИП подгруппа — модель оборудования. Запчасть без "
        "модели не продать, поэтому разметка листов менеджера сохранена "
        "как есть.",
        "Стекла к иллюминаторам идут в корпусную часть, а стекла к "
        "светильникам — в электрику. Слово одно, товар разный.",
        "Сумма считается по цене из файлов менеджера. В сводном файле "
        "цен нет вообще, они подставлены по наименованию из четырех "
        "файлов с ценами — поэтому сумма есть только у тех позиций, "
        "которые в них нашлись.",
        "Брак, некомплект и разобранное входят в сумму по той же цене, "
        "хотя стоят дешевле. Они помечены на вкладке «Остаток».",
    ])
    ws.freeze_panes = f"A{start + 1}"
    widths(ws, [w for _, w in cols])
    landscape(ws)


COLS_REST = [
    ("№", 5), ("Группа", 24), ("Подгруппа", 30), ("Наименование", 58),
    ("Локация", 16), ("Ду", 6), ("Чертеж", 18), ("Штук", 8),
    ("Цена, руб", 11), ("Сумма, руб", 13), ("Пометки", 34),
]


def sheet_rest(wb, rows):
    ws = wb.create_sheet("Остаток")
    ws["A1"] = "Остаток построчно"
    ws["A1"].font = TITLE
    ws["A2"] = ("Желтым — брак, некомплект и разобранное: такой товар "
                "продается отдельно и дешевле. Оранжевым — позиции без "
                "цены.")
    ws["A2"].font = NOTE
    start = 4
    for i, (h, _) in enumerate(COLS_REST, start=1):
        ws.cell(row=start, column=i, value=h)
    head(ws, start, len(COLS_REST))
    row = start + 1
    n = 0
    for r in sorted(rows, key=lambda x: (x["Группа"], x["Подгруппа2"],
                                         -x["Сумма1"], x["Наименование"])):
        n += 1
        fill = (MARK_FILL if r["Брак"] or r["Дефект"]
                else WARN_FILL if not r["Цена1"] else None)
        row = put(ws, row, [
            n, r["Группа"], r["Подгруппа2"], r["Наименование"],
            r["Локация"], r["Ду"] or "", r["Чертеж"] or "", r["Штук"],
            round(r["Цена1"]) or "", r["Сумма1"] or "", r["Пометки"],
        ], wrap=(4, 11), center=(1, 6, 8), money=(9, 10), fill=fill)
    ws.freeze_panes = f"A{start + 1}"
    ws.auto_filter.ref = (f"A{start}:"
                          f"{get_column_letter(len(COLS_REST))}{row - 1}")
    widths(ws, [w for _, w in COLS_REST])
    ws.print_title_rows = f"{start}:{start}"
    landscape(ws)


def sheet_check(wb, old_loc, now_loc, by_loc):
    ws = wb.create_sheet("Сверка с учетом")
    ws["A1"] = "Сентябрьский остаток против учета на 16.06.2025"
    ws["A1"].font = TITLE
    ws["A2"] = ("Это проверка того, что отгруженное из файлов снято. "
                "Сошлось или не сошлось считается по группе, диаметру и "
                "материалу: по названиям строки учета и строки отгрузки "
                "не совпадают никогда.")
    ws["A2"].font = NOTE
    cols = [("Локация", 34), ("Было по учету, поз", 11),
            ("Было по учету, шт", 11), ("Осталось в сентябре, шт", 12),
            ("Пропало, шт", 11), ("Из них объяснено отгрузками, шт", 14),
            ("Числится в другой локации, шт", 13),
            ("Не объяснено, шт", 12), ("Позиций не объяснено", 12)]
    start = 4
    for i, (h, _) in enumerate(cols, start=1):
        ws.cell(row=start, column=i, value=h)
    head(ws, start, len(cols))
    row = start + 1
    alias = {"Склад 1А, 1Б": "Склад 1А/1Б", "Улица": "ЗИП и насосы"}
    title = {"Улица": "Улица (в учете — перечень ЗИП и насосов)"}
    seen = set()
    tot = [0, 0, 0, 0, 0, 0, 0, 0]
    for name, _, got in LOCATIONS:
        loc = alias.get(name, name)
        if name in ("Цех", "ЗИП и насосы") or loc in seen:
            continue
        seen.add(loc)
        name = title.get(name, name)
        o = old_loc.get(loc, [0, 0])
        a = by_loc.get(loc, [0, 0, 0, 0, 0, 0])
        vals = [o[0], o[1], now_loc.get(loc, 0), a[1], a[2], a[5], a[3],
                a[4]]
        for i, v in enumerate(vals):
            tot[i] += v
        fill = WARN_FILL if not got else (OK_FILL if a[3] == 0 else None)
        row = put(ws, row, [name] + [v or "" for v in vals],
                  center=tuple(range(2, 10)), money=(3, 4, 5, 6, 7, 8),
                  fill=fill)
    row = put(ws, row, ["ИТОГО"] + tot, center=tuple(range(2, 10)),
              money=(3, 4, 5, 6, 7, 8), font=BOLD, fill=GROUP_FILL)
    row += 1
    row = notes(ws, row, len(cols), [
        "«Пропало» — разница между наличием по учету на 16.06.2025 и "
        "остатком за сентябрь 2026. Учет с тех пор не вели, продажи в "
        "него не заносили, поэтому разница накопилась за 15 месяцев.",
        "«Объяснено отгрузками» — это керченские задвижки, питерские "
        "фильтры, запчасти от 29.09 и контейнер от 26.10. Там, где "
        "объяснено все, строка зеленая.",
        "«Числится в другой локации» — товар по складу переставляли. "
        "Блокформы из перечня ЗИП и насосов теперь стоят в контейнере "
        "2, поэтому по перечню они выглядят пропавшими, а по складу "
        "никуда не девались.",
        "«Не объяснено» не равно «потеряно». Это может быть продажа до "
        "отгрузок, может быть пересчет, а может быть и то, что позицию "
        "не нашли. Построчно — на вкладке «Пропало без отгрузки».",
        "Улица в учете покрыта перечнем ЗИП и насосов, там нет цен и "
        "нет части позиций, поэтому сверка по ней условная.",
        "Контейнер 1 и перечень ЗИП насосов за сентябрь не присланы: "
        "по ним остаток посчитать не из чего, и сверка пустая.",
    ])
    ws.freeze_panes = f"A{start + 1}"
    widths(ws, [w for _, w in cols])
    landscape(ws)


def sheet_missing(wb, missing):
    ws = wb.create_sheet("Нет в файлах за сентябрь")
    ws["A1"] = "Было в учете, в сентябрьских файлах нет, отгрузками не объясняется"
    ws["A1"].font = TITLE
    ws["A2"] = ("Главное по этой вкладке: сентябрьские файлы — не полный "
                "пересчет склада. Целыми блоками отсутствует ЗИП "
                "двигателя Пилстик 6ЧН40/46 со склада электрики (1 221 "
                "шт) и большая часть светильников СС328 из фойе "
                "(около 2 000 шт), а по контейнеру 1 и перечню ЗИП "
                "насосов файлов нет вовсе. Колонка «почему» это и "
                "разделяет: сначала надо дослать файлы, и только потом "
                "считать недостачу.")
    ws["A2"].font = NOTE
    cols = [("№", 5), ("Локация", 18), ("Группа", 24), ("Подгруппа", 28),
            ("Наименование", 56), ("Чертеж", 18), ("Ду", 6),
            ("Было, шт", 9), ("Объяснено отгрузкой, шт", 12),
            ("НЕ объяснено, шт", 12), ("Почему", 32), ("Цена, руб", 11),
            ("Сумма неясного, руб", 14)]
    start = 4
    for i, (h, _) in enumerate(cols, start=1):
        ws.cell(row=start, column=i, value=h)
    head(ws, start, len(cols))
    row = start + 1
    for n, r in enumerate(sorted(missing, key=lambda x: -x["Сумма"]
                                 or -x["Непонятно"]), start=1):
        row = put(ws, row, [
            n, r["Локация"], r["Группа"], r["Подгруппа"], r["Наименование"],
            r["Чертеж"], r["Ду"] or "", r["Пропало"], r["Отгружено"] or "",
            r["Непонятно"], r["Почему"], round(r["Цена"]) or "",
            r["Сумма"] or "",
        ], wrap=(5, 11), center=(1, 7, 8, 9, 10), money=(12, 13),
            fill=WARN_FILL if "не прислана" in r["Почему"] else None)
    row = put(ws, row, ["ИТОГО", "", "", "", f"{len(missing)} позиций", "",
                        "", sum(r["Пропало"] for r in missing),
                        sum(r["Отгружено"] for r in missing),
                        sum(r["Непонятно"] for r in missing), "", "",
                        sum(r["Сумма"] for r in missing)],
              center=(8, 9, 10), money=(12, 13), font=BOLD, fill=GROUP_FILL)
    row += 1
    row = notes(ws, row, len(cols), [
        "«Вся подгруппа в сентябрьский файл не попала» — это не "
        "недостача, а неполный файл: ЗИП Пилстика, светильники СС328, "
        "блокформы. Запросить у менеджера.",
        "«Подгруппа в файле есть, этой позиции нет» — вот это надо "
        "проверять по складу: соседние позиции той же подгруппы "
        "менеджер пересчитала, а эту нет.",
        "Сумма считается по цене из старого учета и только по той "
        "части, которая не объяснена отгрузкой. Цена внутренняя, "
        "сделками не проверена.",
        "Три цифры учета под проверкой и в разговоре с покупателем не "
        "называются: клапан Ду10 Ру40 «СУНДУК» 798 шт по 9 000, "
        "иллюминатор Ф300 створчатый и регулятор РТНД-150.",
    ])
    ws.freeze_panes = f"A{start + 1}"
    ws.auto_filter.ref = f"A{start}:{get_column_letter(len(cols))}{row - 3}"
    widths(ws, [w for _, w in cols])
    ws.print_title_rows = f"{start}:{start}"
    landscape(ws)


def sheet_sources(wb, rows, ship, parse_notes):
    ws = wb.create_sheet("Источники")
    ws["A1"] = "Откуда цифры"
    ws["A1"].font = TITLE
    row = 3
    ws.cell(row=row, column=1,
            value="Остатки менеджера, сентябрь 2026").font = SUB
    row += 1
    for i, h in enumerate(["Файл", "Лист", "Позиций", "Штук", "Сумма, руб"],
                          start=1):
        ws.cell(row=row, column=i, value=h)
    head(ws, row, 5, height=20)
    row += 1
    agg = defaultdict(lambda: [0, 0, 0.0])
    for r in rows:
        a = agg[(r["Файл"], r["Лист"])]
        a[0] += 1
        a[1] += r["Штук"]
        a[2] += r["Сумма1"]
    for (f, s), v in agg.items():
        row = put(ws, row, [f, s, v[0], v[1], round(v[2])],
                  wrap=(1,), center=(3, 4), money=(5,))
    row = put(ws, row, ["ИТОГО", "", sum(v[0] for v in agg.values()),
                        sum(v[1] for v in agg.values()),
                        round(sum(v[2] for v in agg.values()))],
              center=(3, 4), money=(5,), font=BOLD, fill=GROUP_FILL)

    row += 2
    ws.cell(row=row, column=1,
            value="Отгрузки, которые менеджер уже сняла").font = SUB
    row += 1
    for i, h in enumerate(["Отгрузка", "Позиций", "Штук"], start=1):
        ws.cell(row=row, column=i, value=h)
    head(ws, row, 3, height=20)
    row += 1
    byship = defaultdict(lambda: [0, 0])
    for s in ship:
        a = byship[s["Откуда"]]
        a[0] += 1
        a[1] += s["Штук"]
    for k, v in byship.items():
        row = put(ws, row, [k, v[0], v[1]], wrap=(1,), center=(2, 3))
    row = put(ws, row, ["ИТОГО", sum(v[0] for v in byship.values()),
                        sum(v[1] for v in byship.values())],
              center=(2, 3), font=BOLD, fill=GROUP_FILL)

    row += 2
    ws.cell(row=row, column=1,
            value="Локации склада: что пришло за сентябрь").font = SUB
    row += 1
    for i, h in enumerate(["Локация", "Учетный файл",
                           "Остаток за сентябрь 2026"], start=1):
        ws.cell(row=row, column=i, value=h)
    head(ws, row, 3, height=20)
    row += 1
    for name, fname, got in LOCATIONS:
        row = put(ws, row, [name, fname, got or "НЕ ПРИСЛАН"], wrap=(3,),
                  fill=None if got else WARN_FILL)

    row += 2
    ws.cell(row=row, column=1, value="Замечания по разбору файлов").font = SUB
    row += 1
    for i, h in enumerate(["Лист", "Строка", "Наименование", "Замечание"],
                          start=1):
        ws.cell(row=row, column=i, value=h)
    head(ws, row, 4, height=20)
    row += 1
    for n in parse_notes:
        row = put(ws, row, [n["Лист"], n["Строка"] or "",
                            n["Наименование"], n["Замечание"]],
                  wrap=(3, 4), center=(2,), fill=MARK_FILL)

    row += 1
    row = notes(ws, row, 5, [
        "Отгруженное снято менеджером во всех файлах. По складу 1А и 1Б "
        "в сводном на 24 позиции и 174 шт меньше, чем в файле с ценами, "
        "и это ровно контейнерные задвижка Ду80 36 шт, клапан Ду50 "
        "Ру10/25 47 шт, головки, грибки и задвижка Ду250 17 шт. По фойе "
        "убраны все клинкеты и все фильтры забортной воды.",
        "Количества в сводном не разошлись ни по одной позиции: позиция "
        "либо есть целиком, либо убрана целиком. Так выглядит вычет "
        "отгруженного, а не пересчет склада.",
        "798 клапанов Ду10 Ру40 («СУНДУК») лежали в менеджерской, и их "
        "там больше нет: в контейнер уехало 796 штук этой позиции.",
        "Цены в сводном файле не проставлены. Они подставлены по "
        "наименованию из четырех файлов с ценами, поэтому сумма есть "
        "только у части позиций.",
        "По менеджерской в файле написано «склад не пересчитывался», а "
        "числа в нем сдвинуты на строку вверх относительно "
        "наименований — сопоставлены по порядку, проверить по складу.",
        "ОСТАТКИ ОСТАЮТСЯ НЕДОСТОВЕРНЫМИ. До инвентаризации не "
        "публикуем объявления, не рассылаем КП с количествами и не "
        "называем стоимость склада покупателю.",
    ], height=40)
    widths(ws, [26, 24, 32, 46, 16])


def main():
    stock, parse_notes = stock_sheets.read_with_notes()
    rows = enrich(stock)
    ship = shipped()
    old, old_loc, now_loc, by_loc, missing = reconcile(rows, ship)

    wb = Workbook()
    sheet_groups(wb, rows)
    sheet_rest(wb, rows)
    sheet_check(wb, old_loc, now_loc, by_loc)
    sheet_missing(wb, missing)
    sheet_sources(wb, rows, ship, parse_notes)

    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / "Остаток_склада_по_группам.xlsx"
    wb.save(path)

    print(f"готово: {path}\n")
    print(f"остаток: {len(rows)} позиций, "
          f"{rub(sum(r['Штук'] for r in rows))} шт, "
          f"{rub(sum(r['Сумма1'] for r in rows))} руб")
    agg = defaultdict(lambda: [0, 0, 0.0])
    for r in rows:
        a = agg[r["Группа"]]
        a[0] += 1
        a[1] += r["Штук"]
        a[2] += r["Сумма1"]
    for g, v in sorted(agg.items(), key=lambda kv: -kv[1][2]):
        print(f"  {g:<26}{v[0]:>4} поз, {v[1]:>5} шт, {rub(v[2]):>12} руб")
    print(f"\nучет 16.06.2025: {sum(v[1] for v in old_loc.values()):,} шт"
          .replace(",", " "))
    print(f"пропало {sum(v[1] for v in by_loc.values()):,} шт, из них "
          f"объяснено отгрузками {sum(v[2] for v in by_loc.values()):,}, "
          f"не объяснено {sum(v[3] for v in by_loc.values()):,}"
          .replace(",", " "))
    print(f"позиций без объяснения: {len(missing)}")
    why = defaultdict(lambda: [0, 0])
    for r in missing:
        why[r["Почему"]][0] += 1
        why[r["Почему"]][1] += r["Непонятно"]
    for k, v in sorted(why.items(), key=lambda kv: -kv[1][1]):
        print(f"  {k:<46}{v[0]:>4} поз, {rub(v[1]):>6} шт")


if __name__ == "__main__":
    main()
