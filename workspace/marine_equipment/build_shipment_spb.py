#!/usr/bin/env python3
"""Отгрузка арматуры покупателю в Санкт-Петербург.

Контейнер едет из Владивостока, а не стоит на площадке, поэтому лимит
по весу жесткий и нужна раскрепежка: на железной дороге груз получает
сильные продольные удары при маневрировании.

Под отгрузку идут клапаны, задвижки и клинкеты, фильтры забортной
воды, бронза и латунь. Всего этого около 31 тонны, в контейнер войдет
меньше, поэтому позиции отбираются по деньгам за килограмм: за
перевозку платим за вес, значит первым едет то, что дороже на кило.

Первыми грузим клинкеты Ду250 и Ду350 бронзовые и латунные: их
покупатель назвал отдельно. Материала в учете по ним нет, поэтому
при отборе их смотрят глазами — везем только бронзу и латунь.

Результат: outputs/marine_equipment/Отгрузка_Питер.xlsx
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
from build_catalog_xlsx import group_of  # noqa: E402
from build_weight_xlsx import unit_weight, du  # noqa: E402
from build_container_plan import unit_volume  # noqa: E402

OUT = pathlib.Path(__file__).resolve().parents[2] / "outputs" / "marine_equipment"

FONT = "Arial"
HEAD_FILL = PatternFill("solid", fgColor="1F3864")
HEAD_FONT = Font(name=FONT, size=11, bold=True, color="FFFFFF")
TITLE = Font(name=FONT, size=14, bold=True, color="1F3864")
BASE = Font(name=FONT, size=10)
NOTE = Font(name=FONT, size=9, italic=True, color="666666")
GO_FILL = PatternFill("solid", fgColor="EAF3EA")
STAY_FILL = PatternFill("solid", fgColor="FDF2E3")
BRONZE_FILL = PatternFill("solid", fgColor="FCE4D6")
SUB_FILL = PatternFill("solid", fgColor="D9E2F3")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# 25 тонн — предел загрузки 40-футового контейнера. Массы у нас
# справочные, с погрешностью около трети, поэтому перед отправкой
# характерные места надо взвесить: перегруз ловят на весах в порту.
LIMIT_KG = 25_000
LIMIT_M3 = 67

FILTERS = re.compile(r"фильтр.*(заборт|заборн)|сетка на фильтр", re.I)
BRONZE = re.compile(r"бронз|латун", re.I)
GATES = re.compile(r"клинкет|задвижк", re.I)
TITANIUM = re.compile(r"титан", re.I)
STEEL = re.compile(r"сталь|стальн", re.I)
# мелочь: то, что едет в ящиках и чем добиваем контейнер до предела
SMALL = re.compile(r"штуцерн|штущерн|муфтов|цапков|под дюрит|педальн|"
                   r"манометр|^кран", re.I)

# покупатель назвал их первыми: клинкеты этих диаметров, бронза и латунь
FIRST_DU = (250, 300, 350)

QUEUE = {
    1: "1. Клинкет Ду250, Ду300, Ду350, бронза и латунь",
    2: "2. Титановая арматура",
    3: "3. Остальная бронза и латунь",
    4: "4. Прочее, по цене за килограмм",
    5: "5. Мелочь — ей добиваем контейнер до 25 тонн",
}

# что в этот контейнер не грузим
HOLD_STEEL = "сталь, в этот раз не грузим"
HOLD_BOXES = "клапанная коробка, следующей отгрузкой"


def is_bronze(r):
    """Бронза и латунь: в учете материал указан то в названии,
    то только в заголовке раздела."""
    return bool(BRONZE.search(r["Наименование"] + " "
                              + (r["Группа в файле"] or "")))


def is_titanium(r):
    return bool(TITANIUM.search(r["Наименование"] + " "
                                + (r["Группа в файле"] or "")))


def queue(r):
    """Очередь погрузки задана покупателем, а не расчетом: сначала
    крупные клинкеты из бронзы и латуни, потом титан — он легкий,
    потом остальная цветнина. Мелкая арматура может уехать позже."""
    n = r["Наименование"]
    d = du(n) or 0
    if GATES.search(n) and d in FIRST_DU and not STEEL.search(n):
        return 1
    if is_titanium(r):
        return 2
    if is_bronze(r):
        return 3
    if SMALL.search(n) or (0 < d < 50):
        # мелочью добиваем остаток: что не влезло, уедет следующим бортом
        return 5
    return 4


def hold(r):
    """Почему позиция не едет этим контейнером. Сталь отложена целиком,
    клапанные коробки уйдут следующей отгрузкой.

    Материал в учете проставлен у 33 позиций из 271: отсеять всю сталь
    по названию нельзя, ее ловят при отборе на складе."""
    if STEEL.search(r["Наименование"] + " " + (r["Группа в файле"] or "")):
        return HOLD_STEEL
    if item_kind(r) == "коробка клапанная":
        return HOLD_BOXES
    return ""


def material(r):
    """В учете материал проставлен не везде. Там, где его нет, а позиция
    идет первой очередью, отбор делается на складе глазами."""
    if is_titanium(r):
        return "титан"
    if is_bronze(r):
        return "бронза/латунь"
    if queue(r) == 1:
        return "проверить: везем бронзу и латунь"
    if STEEL.search(r["Наименование"]):
        return "сталь"
    return ""


def item_kind(r):
    n = r["Наименование"].lower()
    if FILTERS.search(n):
        return "фильтр забортной воды"
    if re.search(r"клинкет|задвижк", n):
        return "клинкет (задвижка)"
    if re.search(r"коробка \d|коробка 2-х|коробка 3-х", n):
        return "коробка клапанная"
    if re.search(r"захлопк", n):
        return "захлопка"
    if re.search(r"кингстон", n):
        return "кингстон"
    if re.search(r"^кран", n):
        return "кран"
    if re.search(r"штуцерн|штущерн|муфтов|цапков|под дюрит|педальн|манометр",
                 n):
        return "клапан штуцерный"
    if re.search(r"фланц", n):
        return "клапан фланцевый"
    return "прочая арматура"


# пол контейнера тоже идет настилом: задвижка Ду350 стоит на доске,
# а не на рифленом железе
FLOOR = ("Пол контейнера", "Сплошной настил доской 40 мм под первый ярус",
         "Настил кладем до погрузки, по всей площади под задвижки")

TIERS = {
    1: ("Ярус 1, на настил", "Задвижки и клинкеты Ду200, 250, 300, 350",
        "Доска 40 мм поперек, на всю ширину"),
    2: ("Ярус 2", "Остальные клинкеты, задвижки и арматура Ду100-150",
        "Доска 25-40 мм поперек"),
    3: ("Ярус 3", "Ящики с мелкой арматурой до Ду80",
        "Сплошной настил доской по верху ящиков, под верхний ярус"),
    4: ("Ярус 4, верх", "Фильтры забортной воды и титановая арматура",
        "Сверху ничего не кладем"),
}


def tier(r):
    """Порядок укладки снизу вверх: тяжелые задвижки на пол, ящики с
    мелочью выше, фильтры сверху. Схема: настил, ярус, доска, ярус."""
    d = du(r["Наименование"]) or 0
    k = item_kind(r)
    if k == "фильтр забортной воды" or is_titanium(r):
        # титан легкий и дорогой, его кладем наверх
        return 4
    if k == "клинкет (задвижка)" and d >= 200:
        return 1
    if k == "клинкет (задвижка)" or d >= 100:
        return 2
    return 3


def pack_type(r):
    """Как паковать: от этого зависит, что заказывать заранее."""
    d = du(r["Наименование"]) or 0
    k = item_kind(r)
    if k == "фильтр забортной воды":
        # фильтры едут верхним ярусом, на ящиках
        return "Обрешетка, ставится верхним ярусом на ящики"
    if d >= 200:
        return "Поддон с обрешеткой, крепление к полу"
    if d >= 100:
        return "Поддон, прокладки между рядами"
    if d >= 50:
        return "Обрешетка на поддоне"
    return "Ящик сплошной, с перегородками"


def priority(r):
    """Сначала то, что покупатель назвал сам, потом чем дороже
    килограмм, тем выгоднее его везти через всю страну."""
    price_per_kg = r["Сумма"] / r["вес"] if r["вес"] else 0
    return (queue(r), -price_per_kg)


def head(ws, row, ncols, height=30):
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


PACKING = [
    ("Фланцы закрыть и стянуть",
     "Привалочная поверхность фланца — рабочая: забой или задир по ней "
     "делает арматуру некондиционной. На каждый фланец кладем кружок "
     "фанеры или плотного картона по диаметру и стягиваем стрейчем. "
     "Две детали фланец к фланцу — через прокладку, иначе побьют друг друга."),
    ("Шток не должен нести нагрузку",
     "Клапаны и клинкеты укладываем штоком вверх или вбок, но никогда "
     "не опираем на шток и не кладем на него ничего сверху. Погнутый шток "
     "означает, что арматуру надо разбирать и править."),
    ("Маховики снять или защитить",
     "Маховик — самая выступающая и самая хрупкая часть. Чугунный "
     "маховик колется от удара. Где снимается — снять и уложить отдельно, "
     "подписав, к какой позиции."),
    ("Бронзу отдельно от стали",
     "Бронза и латунь мягче: в одном ящике со стальной арматурой они "
     "сминаются на кромках и фланцах. Бронзовые позиции — в свои ящики, "
     "с прокладкой между рядами."),
    ("Прокладки между рядами",
     "Между рядами кладем доску 25-40 мм на всю ширину. Штабель без "
     "прокладок за дорогу садится, и нижний ряд принимает вес всего "
     "штабеля точечно."),
    ("Пустоты заполнять",
     "Все, что может ездить внутри ящика или по контейнеру, за месяц "
     "дороги разобьет соседей. Пустоты забиваем обрезками доски, "
     "воздушно-пузырчатой пленкой или ветошью."),
    ("Крепление к контейнеру",
     "Поддоны крепим стяжными ремнями к рымам, между рядами ставим "
     "распорные бруски. На железной дороге при маневрировании продольный "
     "удар доходит до двух весов груза — незакрепленный поддон поедет."),
    ("Грузим ярусами: настил, ярус, доска, ярус",
     "Сначала настил доской 40 мм на пол контейнера. На него задвижки "
     "Ду350, Ду300, Ду250 и Ду200, ближе к середине длины. Сверху "
     "доска, на нее остальная арматура Ду100-150, снова доска, на нее "
     "ящики с мелкими клапанами, и только сверху фильтры. Доска "
     "обязательна и под нижним ярусом, и между каждыми двумя ярусами: "
     "без нее фланец верхнего ряда стоит на маховике нижнего."),
    ("Маркировка каждого места",
     "На ящике: номер места, что внутри, вес, строки реестра, стрелки "
     "«верх». Покупатель должен принять груз по упаковочному листу, "
     "не вскрывая все подряд."),
    ("Упаковочный лист и фото",
     "На каждое место — свой лист, общий реестр отправки и фотографии "
     "загрузки. Это и документ приемки, и доказательство при споре "
     "о повреждениях."),
]


def sheet_summary(wb, go, stay, held):
    ws = wb.create_sheet("Сводка отгрузки", 0)
    ws["A1"] = "Отгрузка арматуры в Санкт-Петербург"
    ws["A1"].font = TITLE
    ws["A2"] = ("Контейнер едет, значит лимит по весу жесткий. Взято "
                f"{LIMIT_KG / 1000:.0f} тонн: оценка массы справочная, "
                "запас на погрешность обязателен. Первая очередь — "
                "клинкеты Ду250, Ду300 и Ду350 из бронзы и латуни, потом "
                "титан, потом остальная бронза и латунь, потом прочее по "
                "цене за килограмм, и мелочью добиваем до предела.")
    ws["A2"].font = NOTE
    ws["A3"] = ("Сталь в этот раз не грузим, клапанные коробки уходят "
                "следующей отгрузкой. Но материал в учете проставлен "
                "только у 33 позиций из 271: по названию отсеяна лишь "
                "явная сталь, остальное смотрим глазами при отборе. "
                "Освободившийся вес добираем мелочью с листа «Не влезло» "
                "сверху вниз.")
    ws["A3"].font = NOTE

    facts = [
        ("Едет в Петербург", f"{len(go)} позиций",
         f"{sum(r['Наличие'] for r in go):.0f} штук"),
        ("Вес отгрузки", f"{sum(r['вес'] for r in go) / 1000:.1f} т",
         f"лимит {LIMIT_KG / 1000:.0f} т"),
        ("Объем отгрузки", f"{sum(r['объем'] for r in go):.0f} куб. м",
         f"объем контейнера {LIMIT_M3} куб. м"),
        ("Стоимость отгрузки", sum(r["Сумма"] for r in go), "по нашим ценам"),
        ("Не влезло по весу", f"{len(stay)} позиций",
         f"{sum(r['вес'] for r in stay) / 1000:.1f} т, "
         f"{sum(r['Сумма'] for r in stay) / 1e6:.1f} млн руб"),
        ("Отложено: сталь и коробки", f"{len(held)} позиций",
         f"{sum(r['вес'] for r in held) / 1000:.1f} т, "
         f"{sum(r['Сумма'] for r in held) / 1e6:.1f} млн руб"),
    ]
    for i, (label, value, note) in enumerate(facts):
        row = 4 + i
        ws.cell(row=row, column=1, value=label).font = Font(name=FONT, size=10,
                                                            bold=True)
        cell = ws.cell(row=row, column=2, value=value)
        cell.font = Font(name=FONT, size=10, bold=True)
        if isinstance(value, (int, float)):
            cell.number_format = "#,##0"
        ws.cell(row=row, column=3, value=note).font = NOTE

    start = 11
    ws.cell(row=start - 1, column=1, value="Что едет, по типам").font = TITLE
    for i, h in enumerate(["Тип", "Позиций", "Штук", "Вес, т", "Объем, м³",
                           "Стоимость, руб", "Руб за кг"], start=1):
        ws.cell(row=start, column=i, value=h)
    head(ws, start, 7, height=24)

    agg = defaultdict(lambda: [0, 0, 0.0, 0.0, 0.0])
    for r in go:
        a = agg[item_kind(r)]
        a[0] += 1
        a[1] += r["Наличие"]
        a[2] += r["вес"]
        a[3] += r["объем"]
        a[4] += r["Сумма"]
    row = start
    for k, (p, u, w, v, s) in sorted(agg.items(), key=lambda kv: -kv[1][4]):
        row += 1
        values = [k, p, u, round(w / 1000, 2), round(v, 1), s or None,
                  round(s / w) if w else None]
        for i, val in enumerate(values, start=1):
            cell = ws.cell(row=row, column=i, value=val)
            cell.font = BASE
            cell.border = BORDER
            if i in (6, 7):
                cell.number_format = "#,##0"
    widths(ws, [28, 11, 10, 11, 12, 18, 12])
    return ws


ITEM_COLS = ["№", "Едет", "Очередь", "Ярус", "Тип", "Диаметр",
             "Наименование", "Материал", "Штук", "Вес, кг", "Объем, м³",
             "Стоимость, руб", "Руб за кг", "Как паковать", "Локация",
             "Почему не едет"]


def sheet_items(wb, rows, title, name):
    ws = wb.create_sheet(name)
    ws["A1"] = title
    ws["A1"].font = TITLE
    start = 3
    for i, h in enumerate(ITEM_COLS, start=1):
        ws.cell(row=start, column=i, value=h)
    head(ws, start, len(ITEM_COLS))

    for n, r in enumerate(rows, start=1):
        row = start + n
        d = du(r["Наименование"])
        values = [n, r["Едет"], QUEUE[queue(r)], TIERS[tier(r)][0],
                  item_kind(r), f"Ду{d}" if d else "",
                  r["Наименование"], material(r),
                  r["Наличие"], round(r["вес"]), round(r["объем"], 3),
                  r["Сумма"] or None,
                  round(r["Сумма"] / r["вес"]) if r["вес"] and r["Сумма"]
                  else None,
                  pack_type(r), r["Локация"], r.get("Почему не едет", "")]
        for i, v in enumerate(values, start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top",
                                       wrap_text=(i in (3, 7, 14)))
            if i in (10, 12, 13):
                cell.number_format = "#,##0"
            if i == 11:
                cell.number_format = "#,##0.000"
        ws.cell(row=row, column=2).fill = (GO_FILL if r["Едет"] == "да"
                                           else STAY_FILL)
        if queue(r) < 3:
            ws.cell(row=row, column=8).fill = BRONZE_FILL

    last = start + len(rows)
    ws.cell(row=last + 1, column=7, value="ИТОГО").font = Font(name=FONT,
                                                               bold=True)
    for col in (9, 10, 12):
        letter = get_column_letter(col)
        cell = ws.cell(row=last + 1, column=col,
                       value=f"=SUM({letter}{start + 1}:{letter}{last})")
        cell.font = Font(name=FONT, bold=True)
        cell.border = BORDER
        cell.number_format = "#,##0"

    ws.auto_filter.ref = f"A{start}:P{last}"
    ws.freeze_panes = f"G{start + 1}"
    widths(ws, [5, 8, 26, 15, 22, 11, 50, 22, 8, 11, 12, 15, 11, 34, 15, 32])
    return ws


def sheet_packing(wb, go):
    ws = wb.create_sheet("Упаковка и крепление")
    ws["A1"] = "Как паковать, чтобы доехало без помятых фланцев"
    ws["A1"].font = TITLE
    ws["A2"] = ("Дорога Владивосток — Петербург по железной дороге занимает "
                "около месяца, с маневрированием и продольными ударами.")
    ws["A2"].font = NOTE

    for i, h in enumerate(["№", "Правило", "Почему так"], start=1):
        ws.cell(row=4, column=i, value=h)
    head(ws, 4, 3, height=24)
    for n, (rule, why) in enumerate(PACKING, start=1):
        row = 4 + n
        for i, v in enumerate([n, rule, why], start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=(i > 1))
        ws.row_dimensions[row].height = 58

    # схема укладки: как это выглядит в контейнере снизу вверх
    row = 4 + len(PACKING) + 2
    ws.cell(row=row, column=1, value="Схема укладки снизу вверх").font = TITLE
    row += 1
    for i, h in enumerate(["Ярус", "Что кладем", "Чем перекладываем сверху",
                           "Вес яруса, кг"], start=1):
        ws.cell(row=row, column=i, value=h)
    head(ws, row, 4, height=24)

    tiers = defaultdict(float)
    for r in go:
        tiers[tier(r)] += r["вес"]
    row += 1
    for i, v in enumerate([FLOOR[0], FLOOR[1], FLOOR[2], "—"], start=1):
        cell = ws.cell(row=row, column=i, value=v)
        cell.font = BASE
        cell.border = BORDER
        cell.fill = SUB_FILL
        cell.alignment = Alignment(vertical="top", wrap_text=(i in (2, 3)))
    for t in sorted(TIERS):
        row += 1
        label, what, board = TIERS[t]
        for i, v in enumerate([label, what, board, round(tiers[t])], start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=(i in (2, 3)))
            if i == 4:
                cell.number_format = "#,##0"
    row += 1
    for text in [
        "Ярусы идут снизу вверх. Порядок погрузки обратный порядку "
        "выгрузки: что нужно покупателю первым, ставим ближе к дверям.",
        f"Фильтры ({tiers[4]:.0f} кг) стоят на ящиках, значит ящики "
        "должны быть каркасные: брус по углам и по верху. Фильтр Ду200 "
        "весит 120 кг, Ду300 — 230 кг, тонкая обрешетка под ним сложится. "
        "Крупные фильтры ставим над углами ящиков, а не над серединой "
        "крышки.",
        "Между ярусами доска ложится не абы как, а над стенками ящиков и "
        "над корпусами, а не над маховиками и штоками.",
    ]:
        cell = ws.cell(row=row, column=1, value=text)
        cell.font = NOTE
        cell.alignment = Alignment(wrap_text=True)
        ws.row_dimensions[row].height = 28
        row += 1

    row += 2
    ws.cell(row=row, column=1, value="Сколько тары заказать").font = TITLE
    row += 1
    for i, h in enumerate(["Тип упаковки", "Позиций", "Штук", "Вес, кг"],
                          start=1):
        ws.cell(row=row, column=i, value=h)
    head(ws, row, 4, height=24)

    agg = defaultdict(lambda: [0, 0, 0.0])
    for r in go:
        a = agg[pack_type(r)]
        a[0] += 1
        a[1] += r["Наличие"]
        a[2] += r["вес"]
    for k, (p, u, w) in sorted(agg.items(), key=lambda kv: -kv[1][2]):
        row += 1
        for i, v in enumerate([k, p, u, round(w)], start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=(i == 1))
            if i == 4:
                cell.number_format = "#,##0"

    row += 2
    ws.cell(row=row, column=1,
            value="Что купить до начала упаковки").font = TITLE
    for text in [
        "Поддоны: считать по весу, на один поддон берем не больше 700-800 кг",
        "Доска 25-40 мм на прокладки между рядами и распорные бруски",
        "Фанера или плотный картон на заглушки фланцев",
        "Стрейч-пленка и стальная или полиэстеровая обвязочная лента",
        "Стяжные ремни для крепления поддонов к рымам контейнера",
        "Влагопоглотитель под крышу контейнера: месяц дороги дает конденсат",
        "Маркер и бирки на каждое место",
        "Доска 40 мм на настил по полу и под первый ярус: 25 мм под "
        "задвижками Ду250-350 ломается",
        "Брус 50х50 в углы и по верху ящиков с мелочью: на них встанут "
        "фильтры",
    ]:
        row += 1
        cell = ws.cell(row=row, column=1, value="— " + text)
        cell.font = BASE
        cell.alignment = Alignment(wrap_text=True)

    widths(ws, [46, 40, 40, 14])
    return ws


def select():
    """Что едет и что остается. Отбор нужен и здесь, и в файле учета
    отгрузки, поэтому вынесен отдельно."""
    rows = []
    for spec in SOURCES:
        rows.extend(parse(spec))
    for r in rows:
        r["Сумма"] = (r["Наличие"] or 0) * (r["Цена за ед., руб"] or 0)
    live = [r for r in rows if (r["Наличие"] or 0) > 0]
    for r in live:
        r["Группа продажи"] = group_of(r)

    # что покупатель просил: арматура и фильтры забортной воды
    sel = [r for r in live
           if r["Группа продажи"] == "Арматура"
           or FILTERS.search(r["Наименование"])]
    for r in sel:
        r["вес_ед"] = unit_weight(r["Наименование"])
        r["вес"] = r["Наличие"] * r["вес_ед"]
        r["объем"] = r["Наличие"] * unit_volume(r)

    held = []
    pool = []
    for r in sel:
        r["Почему не едет"] = hold(r)
        if r["Почему не едет"]:
            r["Едет"] = "позже"
            held.append(r)
        else:
            pool.append(r)

    pool.sort(key=priority)
    held.sort(key=priority)
    go, stay, kg, m3 = [], [], 0.0, 0.0
    for r in pool:
        if kg + r["вес"] <= LIMIT_KG and m3 + r["объем"] <= LIMIT_M3:
            r["Едет"] = "да"
            go.append(r)
            kg += r["вес"]
            m3 += r["объем"]
        else:
            r["Едет"] = "нет"
            r["Почему не едет"] = "не влезло по весу"
            stay.append(r)
    return go, stay, held, kg, m3


def main():
    go, stay, held, kg, m3 = select()

    wb = Workbook()
    wb.remove(wb.active)
    sheet_summary(wb, go, stay, held)
    sheet_items(wb, go, "Что едет в Петербург", "Едет")
    sheet_items(wb, stay, "Не влезло по весу", "Не влезло")
    sheet_items(wb, held, "Отложено: сталь и клапанные коробки",
                "Не грузим сейчас")
    sheet_packing(wb, go)

    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / "Отгрузка_Питер.xlsx"
    wb.save(path)

    print(f"готово: {path}\n")
    print(f"ЕДЕТ:      {len(go):>3} поз, {sum(r['Наличие'] for r in go):>5.0f} шт, "
          f"{kg / 1000:>5.1f} т, {m3:>4.1f} куб.м, "
          f"{sum(r['Сумма'] for r in go) / 1e6:>5.1f} млн")
    print(f"НЕ ВЛЕЗЛО: {len(stay):>3} поз, {sum(r['Наличие'] for r in stay):>5.0f} шт, "
          f"{sum(r['вес'] for r in stay) / 1000:>5.1f} т, "
          f"{sum(r['Сумма'] for r in stay) / 1e6:>5.1f} млн")
    print(f"ОТЛОЖЕНО:  {len(held):>3} поз, {sum(r['Наличие'] for r in held):>5.0f} шт, "
          f"{sum(r['вес'] for r in held) / 1000:>5.1f} т, "
          f"{sum(r['Сумма'] for r in held) / 1e6:>5.1f} млн  "
          f"(сталь и клапанные коробки)")

    agg = defaultdict(lambda: [0, 0, 0.0, 0.0])
    for r in go:
        a = agg[item_kind(r)]
        a[0] += 1
        a[1] += r["Наличие"]
        a[2] += r["вес"]
        a[3] += r["Сумма"]
    print("\nчто едет по типам:")
    for k, (p, u, w, s) in sorted(agg.items(), key=lambda kv: -kv[1][3]):
        print(f"  {k:<24} {u:>5.0f} шт  {w / 1000:>5.2f} т  "
              f"{s / 1e6:>5.1f} млн  {s / max(w, 1):>6,.0f} руб/кг"
              .replace(",", " "))


if __name__ == "__main__":
    main()
