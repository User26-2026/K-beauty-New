#!/usr/bin/env python3
"""План загрузки контейнеров.

Ограничение по весу у 40-футового — 25 тонн, объем около 67 м³.
У имеющегося 20-футового объем около 33 м³ при сопоставимой
грузоподъемности. Всего товара примерно 103 тонны, поэтому в контейнеры
уходит не все: часть остается во дворе под навесом, часть продается.

Два принципа распределения:
  1. В контейнер идет то, что боится улицы и легко пропадает: мелочь,
     приборы, ЗИП, электрика. Под навес — чугун и сталь крупных
     размеров, которым зимовка не вредит.
  2. Внутри контейнера ходовой товар грузится последним и стоит
     у дверей. Иначе за одним клинкетом придется разгружать половину.

Веса и объемы оценочные, по типовым размерам. Перед заказом контейнера
взвесить 5-10 характерных позиций.

Результат: outputs/marine_equipment/План_загрузки_контейнера.xlsx
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
from build_weight_xlsx import unit_weight as armature_weight, du  # noqa: E402

OUT = pathlib.Path(__file__).resolve().parents[2] / "outputs" / "marine_equipment"

FONT = "Arial"
HEAD_FILL = PatternFill("solid", fgColor="1F3864")
HEAD_FONT = Font(name=FONT, size=11, bold=True, color="FFFFFF")
TITLE = Font(name=FONT, size=14, bold=True, color="1F3864")
BASE = Font(name=FONT, size=10)
NOTE = Font(name=FONT, size=9, italic=True, color="666666")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

FILLS = {
    "40-футовый контейнер": PatternFill("solid", fgColor="E8EEF7"),
    "20-футовый контейнер": PatternFill("solid", fgColor="EAF3EA"),
    "Двор под навесом": PatternFill("solid", fgColor="FDF2E3"),
    "Продать лотом": PatternFill("solid", fgColor="FCE4E4"),
}

CONTAINER_40 = {"вес": 25_000, "объем": 67}
CONTAINER_20 = {"вес": 25_000, "объем": 33}


def other_weight(r):
    """Оценка массы единицы для неарматурных групп, кг."""
    n = r["Наименование"].lower()
    d = du(r["Наименование"]) or 0
    table = [
        (r"брашпил", 2500), (r"лебедка лэ|лебедк(?!.*тралов)", 1800),
        (r"подрулив", 1500), (r"блок цилиндр", 800), (r"кран-балк", 800),
        (r"компрессор", 700), (r"сепаратор", 600), (r"^трал ", 600),
        (r"коленвал", 500), (r"плита камбузн", 460), (r"^насос|^нцв", 250),
        (r"^ротор", 200), (r"крышка цилиндра", 120), (r"барабан сц", 120),
        (r"головка блока|головка поршн", 40), (r"электропривод", 40),
        (r"трансформатор", 30), (r"весы", 30), (r"поршень", 25),
        (r"иллюминатор", 25), (r"^дверь|^двери", 90),
        (r"крышка вгн|крышка легкая", 70), (r"блокформ", 12),
        (r"головка воздушно|грибк", 12), (r"колесо", 8),
        (r"диски тормозн", 8), (r"фланец|фланцы", 6), (r"ствол пожарн", 3),
        (r"гайка|гайки|стакан переборочн", 2), (r"светильник", 2),
        (r"автомат|контактор|^км ?\d|реле|пускател|кнопка|звонок", 1),
        (r"вкладыш|болт|винт|толкател|шарнир|ударник|штанга|пружин|"
         r"распылител|плунжер|седло|направляющ|втулк|ось |ротокап|шайб|"
         r"манжет|палец|шатун|сухар|кулачк", 1.5),
        (r"лампоч|лампа|предохранит|резинк|стекло|кольцо|патрон", 0.1),
    ]
    if re.search(r"фильтр.*(заборт|заборн)", n):
        return {50: 25, 100: 60, 150: 110, 200: 180, 300: 320}.get(d, 60)
    for pattern, w in table:
        if re.search(pattern, n):
            return w
    return 3


def unit_volume(r):
    """Оценка объема единицы, м³: по габаритам типовых изделий."""
    n = r["Наименование"].lower()
    d = du(r["Наименование"]) or 0
    if re.search(r"брашпил|подрулив|кран-балк", n):
        return 4.0
    if re.search(r"лебедк(?!.*тралов)|компрессор|сепаратор|^трал ", n):
        return 2.0
    if re.search(r"плита камбузн", n):
        return 1.2
    if re.search(r"^насос|^нцв|блок цилиндр|коленвал", n):
        return 0.8
    if re.search(r"^дверь|^двери|крышка вгн|крышка легкая", n):
        return 0.15
    if re.search(r"иллюминатор", n):
        return 0.03
    if re.search(r"светильник", n):
        return 0.0125
    if r["Группа продажи"] == "Арматура" and d:
        return {6: 0.0005, 10: 0.001, 15: 0.0015, 20: 0.002, 25: 0.003,
                32: 0.005, 40: 0.008, 50: 0.012, 65: 0.02, 70: 0.022,
                80: 0.03, 100: 0.045, 125: 0.07, 150: 0.1, 200: 0.16,
                250: 0.25, 300: 0.35, 350: 0.45}.get(
                    min([6, 10, 15, 20, 25, 32, 40, 50, 65, 70, 80, 100, 125,
                         150, 200, 250, 300, 350], key=lambda x: abs(x - d)),
                    0.02)
    if re.search(r"лампоч|лампа|предохранит|резинк|стекло|кольцо|патрон|"
                 r"вкладыш|болт|винт|шайб|пружин", n):
        return 0.0005
    return 0.01


def destination(r):
    """Куда едет позиция."""
    g = r["Группа продажи"]
    n = r["Наименование"].lower()
    d = du(r["Наименование"]) or 0

    if re.search(r"светильник сс328", n):
        return "Продать лотом"
    if g in ("Брашпили, лебедки, трал, подруливающие",
             "Сепараторы и компрессоры", "Камбузное оборудование"):
        return "Двор под навесом"
    if g == "Насосы" and not re.search(r"колесо|кольцо|манжет|втулк|шайб|"
                                       r"поплавок|уплотнен|обойма|рти", n):
        return "Двор под навесом"
    # тяжелая арматура не боится улицы, а вес контейнера ограничен.
    # клинкеты уходят во двор с Ду100: они втрое тяжелее клапана того же
    # диаметра, а чугуну и бронзе зимовка под тентом не вредит
    if g == "Арматура":
        heavy_gate = re.search(r"клинкет|задвижк", n) and d >= 100
        # клапанные коробки весят 25-60 кг и улицы не боятся
        boxes = re.search(r"коробка \d|коробка 2-х|коробка 3-х", n)
        if d >= 125 or heavy_gate or boxes:
            return "Двор под навесом"
    if g == "Двери и иллюминаторы":
        return "20-футовый контейнер"
    if g == "Электрика":
        return "20-футовый контейнер"
    return "40-футовый контейнер"


ORDER = ["40-футовый контейнер", "20-футовый контейнер", "Двор под навесом",
         "Продать лотом"]

LOADING_RULES = [
    ("Основание под контейнер",
     "Груженый контейнер нельзя ставить на грунт: просядет неравномерно, "
     "раму перекосит, двери перестанут закрываться. Под все четыре угла — "
     "бетонные блоки или плиты на выровненной площадке."),
    ("Тяжелое на поддоны",
     "Фанерный пол рассчитан на распределенную нагрузку. Клинкет Ду300 "
     "весом три центнера на маленькой опорной площади его продавит. "
     "Все, что тяжелее 50 кг, ставим на поддоны."),
    ("Распределять по всей длине",
     "Груз раскладываем равномерно от дальней стенки к дверям, а не "
     "сваливаем в одном конце. Перекос по весу гнет раму."),
    ("Ходовой товар — к дверям",
     "Грузим в обратном порядке спроса: сначала вглубь то, что спрашивают "
     "редко, последней — ходовую позицию. Иначе за одним клинкетом надо "
     "разгружать половину контейнера."),
    ("Проход вдоль одной стены",
     "Оставляем проход шириной хотя бы полметра вдоль одной стены. "
     "Контейнер под завязку — это склад, из которого ничего не достать."),
    ("Мелочь в ящики с маркировкой",
     "Каждый ящик подписываем: что внутри и номера строк реестра. "
     "Без описи контейнер превращается в свалку."),
    ("Тяжелое не ставить на легкое",
     "Штабель собираем снизу вверх по убыванию веса. Ящик с арматурой "
     "на коробке светильников раздавит коробку."),
    ("Вентиляция и конденсат",
     "В закрытом контейнере зимой выпадает конденсат. Электрику и приборы "
     "кладем в закрытые ящики, под крышку контейнера — влагопоглотитель."),
]


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


def sheet_plan(wb, rows):
    ws = wb.create_sheet("Что куда", 0)
    ws["A1"] = "План загрузки: что в контейнеры, что во двор"
    ws["A1"].font = TITLE
    ws["A2"] = ("40-футовый: 25 тонн и около 67 куб. м. Имеющийся "
                "20-футовый: около 33 куб. м. Веса и объемы оценочные, "
                "перед заказом контейнера взвесить характерные позиции.")
    ws["A2"].font = NOTE

    headers = ["Куда", "Группа", "Позиций", "Штук", "Вес, тонн",
               "Объем, куб. м", "Стоимость, руб"]
    start = 4
    for i, h in enumerate(headers, start=1):
        ws.cell(row=start, column=i, value=h)
    head(ws, start, len(headers))

    agg = defaultdict(lambda: [0, 0, 0.0, 0.0, 0.0])
    for r in rows:
        a = agg[(r["Куда"], r["Группа продажи"])]
        a[0] += 1
        a[1] += r["Наличие"]
        a[2] += r["вес"]
        a[3] += r["объем"]
        a[4] += r["Сумма"]

    row = start
    for dest in ORDER:
        items = [(k, v) for k, v in agg.items() if k[0] == dest]
        if not items:
            continue
        for (_, group), (p, u, w, v, s) in sorted(items,
                                                  key=lambda kv: -kv[1][2]):
            row += 1
            values = [dest, group, p, u, round(w / 1000, 1), round(v),
                      s or None]
            for i, val in enumerate(values, start=1):
                cell = ws.cell(row=row, column=i, value=val)
                cell.font = BASE
                cell.border = BORDER
                cell.alignment = Alignment(vertical="top", wrap_text=(i == 2))
                cell.fill = FILLS[dest]
                if i == 5:
                    cell.number_format = "#,##0.0"
                if i == 7:
                    cell.number_format = "#,##0"
        # итог по направлению
        row += 1
        tot = [sum(v[i] for k, v in items) for i in range(5)]
        for i, val in enumerate([f"ИТОГО {dest}", "", tot[0], tot[1],
                                 round(tot[2] / 1000, 1), round(tot[3]),
                                 tot[4] or None], start=1):
            cell = ws.cell(row=row, column=i, value=val)
            cell.font = Font(name=FONT, size=10, bold=True)
            cell.border = BORDER
            cell.fill = FILLS[dest]
            if i == 5:
                cell.number_format = "#,##0.0"
            if i == 7:
                cell.number_format = "#,##0"
        row += 1

    widths(ws, [24, 42, 10, 10, 13, 15, 18])
    return ws


def sheet_order(wb, rows):
    ws = wb.create_sheet("Порядок загрузки")
    ws["A1"] = "В каком порядке грузить 40-футовый"
    ws["A1"].font = TITLE
    ws["A2"] = ("Грузим от дальней стенки к дверям. Первым едет то, что "
                "спрашивают реже всего, последним — ходовое.")
    ws["A2"].font = NOTE

    zones = [
        ("1. Дальняя стенка", "Штуцерная мелочь Ду6-32 в ящиках",
         "2 000 с лишним штук, спрос редкий и по одной. "
         "Ящики маркируем по строкам реестра."),
        ("2. Дальняя треть, низ", "ЗИП двигателей в ящиках",
         "Пилстик, VD26/20, 6ЧН25/34. Спрос узкий: ищут под свой "
         "двигатель. Ящики по моделям, подпись крупно."),
        ("3. Середина, пол", "Фланцевая арматура Ду50-150 на поддонах",
         "Самое тяжелое из того, что едет в контейнер. Ставим по центру "
         "длины, чтобы не перекосило."),
        ("4. Середина, верх", "Коробки клапанные, захлопки, кингстоны",
         "Средний вес, кладем поверх поддонов с фланцевой арматурой."),
        ("5. Ближняя треть", "Головки, грибки, фланцы, гайки, стволы, стаканы",
         "Мелочь обвязки, спрос средний."),
        ("6. У дверей", "Ходовая арматура: клинкеты Ду65-100, фильтры",
         "Спрашивают чаще всего. Должно доставаться без разгрузки."),
    ]

    headers = ["Зона контейнера", "Что грузим", "Почему так"]
    start = 4
    for i, h in enumerate(headers, start=1):
        ws.cell(row=start, column=i, value=h)
    head(ws, start, len(headers))

    for n, (zone, what, why) in enumerate(zones, start=1):
        row = start + n
        for i, v in enumerate([zone, what, why], start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BASE if i == 3 else Font(name=FONT, size=10, bold=True)
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        ws.row_dimensions[row].height = 44

    row = start + len(zones) + 2
    ws.cell(row=row, column=1, value="Правила загрузки").font = TITLE
    row += 1
    for i, h in enumerate(["№", "Правило", "Почему"], start=1):
        ws.cell(row=row, column=i, value=h)
    head(ws, row, 3, height=22)
    for n, (rule, why) in enumerate(LOADING_RULES, start=1):
        row += 1
        for i, v in enumerate([n, rule, why], start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=(i > 1))
        ws.row_dimensions[row].height = 44

    widths(ws, [26, 44, 86])
    return ws


def sheet_items(wb, rows):
    ws = wb.create_sheet("По позициям")
    ws["A1"] = "Каждая позиция: куда едет, сколько весит"
    ws["A1"].font = TITLE
    cols = ["№", "Куда", "Группа", "Наименование", "Штук", "Вес ед., кг",
            "Вес, кг", "Объем, куб. м", "Локация"]
    start = 3
    for i, h in enumerate(cols, start=1):
        ws.cell(row=start, column=i, value=h)
    head(ws, start, len(cols))

    ordered = sorted(rows, key=lambda r: (ORDER.index(r["Куда"]), -r["вес"]))
    for n, r in enumerate(ordered, start=1):
        row = start + n
        values = [n, r["Куда"], r["Группа продажи"], r["Наименование"],
                  r["Наличие"], r["вес_ед"], round(r["вес"]),
                  round(r["объем"], 3), r["Локация"]]
        for i, v in enumerate(values, start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=(i == 4))
            if i in (6, 7):
                cell.number_format = "#,##0.0"
            if i == 8:
                cell.number_format = "#,##0.000"
        ws.cell(row=row, column=2).fill = FILLS[r["Куда"]]

    last = start + len(ordered)
    ws.cell(row=last + 1, column=4, value="ИТОГО").font = Font(name=FONT,
                                                               bold=True)
    for col in (5, 7, 8):
        letter = get_column_letter(col)
        cell = ws.cell(row=last + 1, column=col,
                       value=f"=SUM({letter}{start + 1}:{letter}{last})")
        cell.font = Font(name=FONT, bold=True)
        cell.border = BORDER

    ws.auto_filter.ref = f"A{start}:I{last}"
    ws.freeze_panes = f"D{start + 1}"
    widths(ws, [5, 22, 34, 54, 8, 13, 12, 14, 16])
    return ws


def main():
    rows = []
    for spec in SOURCES:
        rows.extend(parse(spec))
    for r in rows:
        r["Сумма"] = (r["Наличие"] or 0) * (r["Цена за ед., руб"] or 0)
    live = [r for r in rows if (r["Наличие"] or 0) > 0]

    for r in live:
        r["Группа продажи"] = group_of(r)
        r["вес_ед"] = (armature_weight(r["Наименование"])
                       if r["Группа продажи"] == "Арматура"
                       else other_weight(r))
        r["вес"] = r["Наличие"] * r["вес_ед"]
        r["объем"] = r["Наличие"] * unit_volume(r)
        r["Куда"] = destination(r)

    wb = Workbook()
    wb.remove(wb.active)
    sheet_plan(wb, live)
    sheet_order(wb, live)
    sheet_items(wb, live)

    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / "План_загрузки_контейнера.xlsx"
    wb.save(path)
    print(f"готово: {path}\n")

    agg = defaultdict(lambda: [0, 0, 0.0, 0.0, 0.0])
    for r in live:
        a = agg[r["Куда"]]
        a[0] += 1
        a[1] += r["Наличие"]
        a[2] += r["вес"]
        a[3] += r["объем"]
        a[4] += r["Сумма"]
    for dest in ORDER:
        p, u, w, v, s = agg[dest]
        print(f"{dest:<24} поз {p:>4}  штук {u:>6.0f}  "
              f"{w / 1000:>5.1f} т  {v:>5.0f} куб.м  {s / 1e6:>5.1f} млн")
    print(f"\n40-футовый: лимит {CONTAINER_40['вес'] / 1000:.0f} т "
          f"и {CONTAINER_40['объем']} куб. м")
    p, u, w, v, s = agg["40-футовый контейнер"]
    print(f"  загрузка {w / 1000:.1f} т ({w / CONTAINER_40['вес'] * 100:.0f}% "
          f"лимита), {v:.0f} куб. м "
          f"({v / CONTAINER_40['объем'] * 100:.0f}% объема)")
    p, u, w, v, s = agg["20-футовый контейнер"]
    print(f"20-футовый: {w / 1000:.1f} т, {v:.0f} куб. м "
          f"при объеме {CONTAINER_20['объем']} куб. м")


if __name__ == "__main__":
    main()
