#!/usr/bin/env python3
"""Остатки склада по локациям, файлы менеджера за сентябрь 2026.

Десять файлов, двенадцать локаций. Раскладка колонок у каждого файла
своя, и общего правила нет: в одном количество в D, в другом в C, в
третьем в E. Поэтому раскладка задана ТАБЛИЦЕЙ по листам, а не ищется
по шапке — шапки есть только у двух файлов из десяти, а перепутать
колонку количества с ценой или с чертежом значит испортить расчет.

СВОДНЫЙ ФАЙЛ — ЭТО ОСТАТОК ПОСЛЕ ОТГРУЗКИ. Пять его листов (Фойе и
цех, Гараж 2, Контейнер 2, Склад 1А и 1Б, Улица) дублируют четыре
файла с ценами, но отгруженное из них уже вычтено: по Складу 1А и 1Б
убраны 24 позиции и 174 шт, и это ровно то, что уехало контейнером —
задвижка Ду80 36 шт, клапан Ду50 Ру10/25 47 шт, головки воздушно-
квадратные, грибки АМГ, задвижка Ду250 17 шт. По Фойе убраны все
клинкеты и все фильтры забортной воды, включая керченские 45 задвижек
Ду150 и питерские фильтры Ду300. Количества при этом не разошлись ни
по одной позиции: позиция либо есть целиком, либо убрана целиком —
так выглядит вычет отгруженного, а не пересчет.

Поэтому по этим пяти локациям остаток берется из СВОДНОГО, а файлы с
ценами служат только источником цен: цена подставляется по
наименованию. В сводном вдобавок есть 33 позиции, которых в файлах с
ценами нет совсем — стекла иллюминаторные, сепараторы в разборе,
запчасти 18/22, плиты. Он новее.

Остальные пять локаций (Склад 2, Склад 3, Склад электрики,
Менеджерская, Гараж 1) пришли отдельными файлами, и в них отгруженное
тоже снято. Сверка со старым учетом по наличию на 16.06.2025:
склад 2 был 108 позиций и 688 шт, стал 35 и 203; склад 3 был 235 и
4 260, стал 189 и 3 518; менеджерская была 35 и 1 410, стала 9 и 95.
Пропало ровно то, чем набивали контейнер, — штуцерная мелочь,
переборочные стаканы, клапанные коробки, клапан фланцевый проходной
Ду50 и задвижка клинкетная Ду100. Отдельно: 798 клапанов Ду10 Ру40
(«СУНДУК») лежали в менеджерской, и их там больше нет — в контейнер
уехало 796 штук этой позиции.

Поэтому вычитать отгруженное из этих файлов НЕ НУЖНО: оно уже снято.
Сверку со старым учетом считает build_warehouse_rest.py, он же
показывает, что пропало помимо отгрузок.

Отдельные случаи:

  Склад 2     — колонка количества меняется по ходу листа: у арматуры
                заполнены D и E, у электрики C и E. Единственная
                колонка, заполненная у каждой позиции, — E, по ней и
                считаем.
  Гараж 1     — наименование в A, а не в B; количество в D («наличие
                24.08.23»), именно по нему сходится колонка «сумма».
                Колонка E помечена «нет в наличии, отгружен».
  Менеджерская — числа сдвинуты на строку ВВЕРХ относительно
                наименований: подзаголовок «Судовая запорная арматура»
                стоит с количеством 2, а позиция под ним пустая. Числа
                и наименования пересобираются по порядку, а не по
                строке. Лист помечен «Склад не пересчитывался».

Старый формат .xls читается через xlrd, новый через openpyxl.
"""

import pathlib
import re

ROOT = pathlib.Path(__file__).resolve().parents[2]
SRC = ROOT / "data" / "marine_equipment" / "stock_09.2026"

SVOD = "Сводный_5_локаций_сент2026.xlsx"

# Остаток: (файл, лист) -> раскладка. Индексы колонок с нуля.
#   name  — наименование, qty — количество в наличии
#   draw  — номер чертежа отдельной колонкой
#   price — цена за единицу, total — сумма
#   shift — числа стоят на строку выше наименования
#   clean — отгруженное из листа уже вычтено менеджером
STOCK = {
    (SVOD, "Фойе,цех"):      dict(name=1, qty=2, clean=True),
    (SVOD, "Гараж 2"):       dict(name=1, qty=3, clean=True),
    (SVOD, "Контейнер 2"):   dict(name=1, qty=2, clean=True),
    (SVOD, "Склад 1А, 1Б"):  dict(name=1, qty=3, clean=True),
    (SVOD, "Улица "):        dict(name=1, qty=2, clean=True),
    ("Склад_2_сент2026.xls", "Склад 2"):
        dict(name=1, qty=4, draw=2, clean=True),
    ("Склад_3_сент2026.xls", "Склад3"):
        dict(name=1, qty=3, draw=2, clean=True),
    ("Склад_электрики_сент2026.xls", "Склад электрики"):
        dict(name=1, qty=2, clean=True),
    ("Гараж_1_сент2026.xls", "гараж 1"):
        dict(name=0, qty=3, price=5, total=6, clean=True),
    ("Менеджерская_сент2026.xls", "склад менеджерская"):
        dict(name=1, qty=3, shift=True, clean=True),
}

# Только цены: эти файлы старше сводного, остаток из них не берем.
PRICES = {
    ("Склад_1А_1Б_сент2026.xlsx", "Склад 1А, 1Б"):
        dict(name=1, qty=3, price=4, total=5),
    ("Фойе_и_цех_сент2026.xls", "Фойе"):
        dict(name=1, qty=2, price=3, total=4),
    ("Фойе_и_цех_сент2026.xls", "цех "):
        dict(name=1, qty=2, price=3, total=4),
    ("Гараж_2_сент2026.xlsx", "Гараж 2"):
        dict(name=1, qty=3, price=4, total=5),
    ("Контейнер_2_сент2026.xls", "Контейнер 2"):
        dict(name=1, qty=2, price=3, total=4),
}

LABEL = {
    "Фойе,цех": "Фойе и цех",
    "Гараж 2": "Гараж 2",
    "Контейнер 2": "Контейнер 2",
    "Склад 1А, 1Б": "Склад 1А, 1Б",
    "Улица ": "Улица",
    "Склад 2": "Склад 2",
    "Склад3": "Склад 3",
    "Склад электрики": "Склад электрики",
    "гараж 1": "Гараж 1",
    "склад менеджерская": "Менеджерская",
}

# Листы, по которым менеджер прямо написала, что склад не пересчитан
NOT_COUNTED = {"склад менеджерская"}

SKIP = re.compile(r"^\s*(итого|всего|сумма|наименование|наименованик|"
                  r"примечание|цена)\b", re.I)

# шапка листа: название локации, а не группы товара. «СКДАД 3» —
# опечатка менеджера в самом файле.
HEADING = re.compile(
    r"^(судовое оборудование|фойе|цех|гараж|контейнер|скл[ад]{2}|улица|"
    r"менеджерская|скдад)\b", re.I)

# Подзаголовки групп, которые встречаются с количеством против себя.
GROUP_WORDS = re.compile(
    r"^(арматура|судовая запорная арматура|электрика|насосы|зип|прочее|"
    r"контакторы|барабан|фланец приварной|кор\w*сная часть|"
    r"сепараторы?|блокформы|камбуз)\b\s*(\(.*\))?\s*$", re.I)

# Разделы менеджера: верхний уровень разметки листа. Внутри раздела
# «ЗИП» идут подзаголовки с моделями оборудования (ЦВС 10/40, VD26/20,
# 6ЧН25/34), и терять раздел нельзя: иначе запчасть выглядит просто
# позицией, а без модели оборудования запчасть не продать.
SECTION_WORDS = re.compile(
    r"^(арматура|судовая запорная арматура|электрика|насосы|зип|прочее|"
    r"кор\w*сная часть|сепараторы?|блокформы|камбуз)\b", re.I)


def text(v):
    return re.sub(r"\s+", " ", str(v)).strip() if v is not None else ""


def num(v):
    if isinstance(v, (int, float)):
        return float(v)
    s = text(v).replace(",", ".").replace(" ", "")
    return float(s) if re.match(r"^-?\d+(\.\d+)?$", s) else None


def key(s):
    """Ключ для сопоставления наименований между файлами."""
    return re.sub(r"[^a-zа-я0-9]", "", text(s).lower())


def read_xlsx(path):
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True)
    for ws in wb.worksheets:
        yield ws.title, [list(row) for row in ws.iter_rows(values_only=True)]


def read_xls(path):
    import xlrd
    wb = xlrd.open_workbook(path)
    for sh in wb.sheets():
        yield sh.name, [[sh.cell_value(r, c) for c in range(sh.ncols)]
                        for r in range(sh.nrows)]


def sheets(fname):
    path = SRC / fname
    reader = read_xls if path.suffix == ".xls" else read_xlsx
    return reader(path)


def cell(row, col):
    return row[col] if col is not None and 0 <= col < len(row) else None


def parse(fname, sheet, rows, col):
    """Позиции одного листа. Возвращает (позиции, замечания).

    Товар или подзаголовок различаются так: подзаголовок — это строка
    с названием, у которой нет ни количества, ни локации, ни денег.
    Группы вроде «ЗИП» или «НЦВ 100/30» размечают разделы внутри листа,
    и терять их нельзя: по ним менеджер и считает склад, а запчасть без
    модели оборудования не продать.
    """
    out, notes = [], []
    group = section = ""
    loc_col = 1 if col["name"] == 0 else 0
    pending, names = [], []

    for i, row in enumerate(rows, start=1):
        name = text(cell(row, col["name"]))
        qty = num(cell(row, col.get("qty")))
        place = text(cell(row, loc_col))
        money = num(cell(row, col.get("price"))) or num(
            cell(row, col.get("total")))

        if col.get("shift"):
            if qty is not None:
                pending.append(qty)
            if name and not SKIP.match(name):
                if HEADING.match(name):
                    continue
                if GROUP_WORDS.match(name):
                    if SECTION_WORDS.match(name):
                        section, group = name, ""
                    else:
                        group = name
                    continue
                names.append((i, name, section, group))
            continue

        if not name or SKIP.match(name) or HEADING.match(name):
            continue

        if GROUP_WORDS.match(name) or (qty is None and not place
                                       and money is None):
            if SECTION_WORDS.match(name):
                section, group = name, ""
            else:
                group = name
            if qty:
                notes.append(dict(
                    Файл=fname, Лист=sheet, Строка=i, Наименование=name,
                    Замечание=f"подзаголовок группы с количеством "
                              f"{qty:g} шт — чье это количество, уточнить"))
            continue

        if qty is None:
            notes.append(dict(
                Файл=fname, Лист=sheet, Строка=i, Наименование=name,
                Замечание="количество не проставлено (пусто, а не ноль)"))

        out.append(dict(
            Файл=fname, Лист=sheet, Строка=i,
            Локация=place if place and col["name"] != 0
            else LABEL.get(sheet, sheet),
            Раздел=section, Подгруппа=group, Наименование=name,
            Чертеж=text(cell(row, col.get("draw"))),
            Штук=int(qty or 0),
            Цена=num(cell(row, col.get("price"))) or 0,
            Сумма=num(cell(row, col.get("total"))) or 0,
            Вычтено=bool(col.get("clean")),
        ))

    if col.get("shift"):
        if len(pending) != len(names):
            notes.append(dict(
                Файл=fname, Лист=sheet, Строка=0, Наименование="",
                Замечание=f"чисел {len(pending)}, наименований "
                          f"{len(names)} — сопоставить вручную"))
        for qty, (nrow, name, sec, grp) in zip(pending, names):
            out.append(dict(
                Файл=fname, Лист=sheet, Строка=nrow,
                Локация=LABEL.get(sheet, sheet), Раздел=sec, Подгруппа=grp,
                Наименование=name, Чертеж="", Штук=int(qty or 0),
                Цена=0, Сумма=0, Вычтено=False))
        notes.append(dict(
            Файл=fname, Лист=sheet, Строка=0, Наименование="",
            Замечание="числа в файле сдвинуты на строку вверх "
                      "относительно наименований, сопоставлены по "
                      "порядку — проверить по складу"))
    return out, notes


def price_index():
    """Цена за единицу по наименованию, из файлов с ценами."""
    idx = {}
    for (fname, sheet), col in PRICES.items():
        for name, rws in sheets(fname):
            if name != sheet:
                continue
            rows, _ = parse(fname, sheet, rws, col)
            for r in rows:
                if r["Цена"]:
                    idx.setdefault(key(r["Наименование"]), r["Цена"])
    return idx


def read_with_notes():
    out, notes = [], []
    for (fname, sheet), col in STOCK.items():
        found = False
        for name, rws in sheets(fname):
            if name != sheet:
                continue
            found = True
            a, b = parse(fname, sheet, rws, col)
            out += a
            notes += b
        if not found:
            notes.append(dict(Файл=fname, Лист=sheet, Строка=0,
                              Наименование="",
                              Замечание="лист не найден в файле"))
    idx = price_index()
    for r in out:
        if not r["Цена"]:
            p = idx.get(key(r["Наименование"]))
            if p:
                r["Цена"] = p
                r["Откуда цена"] = "по наименованию из файла с ценами"
            else:
                r["Откуда цена"] = ""
        else:
            r["Откуда цена"] = "в файле локации"
        r["Сумма"] = r["Цена"] * r["Штук"]
        r["Пересчет"] = ("склад не пересчитывался"
                         if r["Лист"] in NOT_COUNTED else "")
    return out, notes


def read():
    return read_with_notes()[0]


if __name__ == "__main__":
    rows, notes = read_with_notes()
    from collections import defaultdict
    by = defaultdict(lambda: [0, 0, 0.0, 0])
    for r in rows:
        a = by[LABEL.get(r["Лист"], r["Лист"])]
        a[0] += 1
        a[1] += r["Штук"]
        a[2] += r["Сумма"]
        a[3] += 1 if r["Вычтено"] else 0
    print(f"позиций {len(rows)}, штук "
          f"{sum(r['Штук'] for r in rows):,}".replace(",", " "))
    print()
    print(f"  {'локация':<20}{'поз':>5}{'штук':>8}{'руб':>16}  отгруженное")
    for k, (n, q, s, c) in sorted(by.items(), key=lambda kv: -kv[1][2]):
        money = f"{s:>16,.0f}".replace(",", " ") if s else " " * 15 + "—"
        print(f"  {k:<20}{n:>5}{q:>8}{money}  "
              f"{'вычтено' if c else 'НЕ вычтено'}")
    print(f"\nитого по указанным ценам: "
          f"{sum(r['Сумма'] for r in rows):,.0f} руб".replace(",", " "))
    have = sum(1 for r in rows if r["Цена"])
    print(f"цена есть у {have} позиций из {len(rows)}")
    print(f"\nзамечаний: {len(notes)}")
    for n in notes:
        print(f"  {LABEL.get(n['Лист'], n['Лист']):<18} стр {n['Строка']:>3} "
              f"{n['Наименование'][:38]:<40}{n['Замечание']}")
