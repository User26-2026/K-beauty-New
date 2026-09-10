#!/usr/bin/env python3
"""Единый реестр склада судового оборудования с ценами.

Источник — девять учетных файлов по локациям в data/marine_equipment/source/.
У каждого своя раскладка колонок, поэтому раскладка задана явно в SOURCES.

На выходе в outputs/marine_equipment/:
  sudovoe_registry_<дата>.tsv    — все позиции, включая нулевые остатки
  sudovoe_price_list_<дата>.tsv  — прайс на продажу: наличие > 0 и цена есть
  sudovoe_no_price_<дата>.tsv    — наличие есть, цены нет: заполнить вручную
"""

import csv
import datetime as dt
import pathlib
import re

import openpyxl
import xlrd

ROOT = pathlib.Path(__file__).resolve().parents[2]
SRC = ROOT / "data" / "marine_equipment" / "source"
OUT = ROOT / "outputs" / "marine_equipment"

# file, локация, первая строка данных, индексы колонок
# qty  — количество по учету, stock — фактическое наличие,
# name/draw/price/total/note — наименование, чертеж, цена, сумма, примечание
SOURCES = [
    dict(file="garazh_1.xls", loc="Гараж 1", start=4,
         name=0, draw=None, qty=2, stock=3, price=5, total=6, note=None),
    dict(file="garazh_2.xlsx", loc="Гараж 2", start=2,
         name=2, draw=None, qty=4, stock=4, price=6, total=7, note=8),
    dict(file="konteiner_1.xls", loc="Контейнер 1", start=4,
         name=1, draw=None, qty=2, stock=2, price=5, total=6, note=None),
    dict(file="konteiner_2.xls", loc="Контейнер 2", start=4,
         name=0, draw=None, qty=2, stock=3, price=5, total=6, note=7),
    dict(file="sklad_1a_1b.xlsx", loc="Склад 1А/1Б", start=2,
         name=2, draw=None, qty=4, stock=4, price=7, total=8, note=10),
    dict(file="sklad_2.xls", loc="Склад 2", start=6,
         name=1, draw=2, qty=3, stock=4, price=6, total=7, note=None),
    dict(file="sklad_3.xls", loc="Склад 3", start=5,
         name=1, draw=2, qty=3, stock=3, price=5, total=6, note=None),
    dict(file="menedzherskaya.xls", loc="Менеджерская", start=2,
         name=1, draw=2, qty=3, stock=3, price=5, total=6, note=None),
    dict(file="foye.xls", loc="Фойе", start=4,
         name=1, draw=None, qty=2, stock=3, price=5, total=6, note=8),
    # склад электрики: цен в файле нет, только наименование и количество
    dict(file="sklad_elektriki.xls", loc="Склад электрики", start=4,
         name=1, draw=None, qty=2, stock=2, price=None, total=None, note=None),
    # перечень ЗИП и насосов: покрывает в том числе улицу, цен тоже нет
    dict(file="zip_nasosy.csv", loc="ЗИП и насосы", start=1,
         name=0, draw=None, qty=1, stock=1, price=None, total=None, note=None),
]

# В блоке контакторов менеджерской наименование сдвинуто на строку вниз
# относительно количества и цены, а количество лежит в колонке наличия.
CONTACTORS = {
    "menedzherskaya.xls": (112, 120),
}

COLUMNS = [
    "Локация", "Товарная группа", "Подгруппа", "Группа в файле",
    "Оборудование", "Наименование",
    "Чертеж", "Кол-во по учету", "Наличие",
    "Цена за ед., руб", "Сумма, руб", "Состояние", "Комплектность",
    "Происхождение", "Примечание",
]

# В разделе ЗИП подзаголовками идут модели оборудования: под ними лежат
# запчасти к этой модели. Без них запчасть теряет привязку — «Фильтр»
# без указания, что он к двигателю ЭТФ-3, продать невозможно.
EQUIPMENT_MODELS = (
    "цвс 10/40", "цвс 4/40", "нцкг", "1эцну", "нцв 100/30", "нцв 40/80",
    "нцв160/80", "нцв 160/80", "нцв 40/20", "нцв 63/20", "нцвс 63/20",
    "нцвс 40/20", "8nvd48a2u", "3д12", "6чн25/34", "экп 70/25",
    "двигатель этф-3", "двигатель д1.м",
    # склад электрики
    "пилстик 6чн40/46", "vd26/20", "3d6", "6 nvd26", "4ч 8,5/11",
    "4ч10,5/13", "zd 72/48", "траловая лебедка", "сепараторы зип",
    "насосы, сепаратор",
)

# строки-заголовки разделов внутри файлов
GROUP_MARKERS = (
    "арматура", "электрика", "зип", "прочее", "насосы", "захлопки", "кингстон",
    "клапан", "клапана", "кран", "коробка", "фланец", "фильтр", "задвижк",
    "гайка", "компенсатор", "крышка", "барабан", "эжектор", "стакан",
    "колонка", "контакторы", "регулятор", "судовое оборудование",
)


# Товарные группы для КП. Арматура — это клапаны и клинкеты (задвижки),
# плюс краны, захлопки, клапанные коробки и кингстоны. Порядок важен:
# позиция попадает в первую подошедшую подгруппу.
CATEGORIES = [
    ("Арматура", "Клинкеты и задвижки", r"клинкет|задвижк"),
    ("Арматура", "Захлопки", r"захлопк"),
    ("Арматура", "Кингстоны", r"кингстон"),
    ("Арматура", "Клапанные коробки", r"коробка \d|коробка 2-х|коробка 3-х"),
    ("Арматура", "Краны", r"^кран |^кран$|кран штуцерн|кран двухходов|кран трехходов"),
    ("Арматура", "Клапаны", r"клапан|кланан"),
    ("Трубопроводная обвязка", "Фильтры забортной воды", r"фильтр забортн|фильтр фланц|фильтр заборн|сетка на фильтр"),
    ("Трубопроводная обвязка", "Фланцы, стаканы, гайки", r"фланец|стакан переборочн|гайка|гайки"),
    ("Трубопроводная обвязка", "Головки, грибки, компенсаторы", r"головка воздушно|грибк|компенсатор|эжектор|колонка указательн|фонарь смотров"),
    ("Насосное и механическое", "Насосы", r"насос|^нцв|^нцвс|цвс |эвн |эсн-"),
    ("Насосное и механическое", "Сепараторы и барабаны", r"сепаратор|барабан сц"),
    ("Насосное и механическое", "Компрессоры и приводы", r"компрессор|электропривод|подрулив|брашпил|лебедк|кран-балк|шпил"),
    ("Насосное и механическое", "Двигатели и ЗИП", r"коленвал|поршень|плунжер|блок цилиндр|крышка цилиндра|тнвд|ротор|колесо раб|колесо вихрев|уплотнен|сальников|манжета|втулка|кольцо|шайба|седло клапана|направляющая|кулачков|фонарь выхлоп|корпус выхлоп|поплавок|двигатель"),
    ("Корпусная часть", "Двери", r"дверь|двери"),
    ("Корпусная часть", "Крышки и люки", r"крышка|крышки"),
    ("Корпусная часть", "Иллюминаторы", r"иллюминатор"),
    ("Электрика", "Светильники и комплектующие", r"светильник|стекло к светил|решетка сс|решетка|резинки для|патрон|прожектор|плафон"),
    ("Электрика", "Лампы и предохранители", r"лампоч|лампа|предохранит"),
    ("Электрика", "Автоматы и контакторы", r"автомат|контактор|км \d|кпм-|пускател|выключател|выключ"),
    ("Электрика", "Реле, датчики, приборы", r"реле|датчик|манометр|термометр|сельсин|тумблер|переключ|звонок|розетк|вилка|рш|ку-|кнопка|трт|мку|симметр|контроллер|пульт|ящик|коробка соедин|ся10|гпв|гпп|вк-|вс-|в-1|рс-|рдк|зип"),
    ("Электрика", "Трансформаторы и питание", r"трансформатор|водонагрев|воздухораспредел"),
    ("Прочее", "Камбузное оборудование", r"плита камбузн|плита пкэ|плита"),
    ("Прочее", "Блокформы и балласт", r"блокформ|балласт"),
    ("Насосное и механическое", "Регуляторы и топливная аппаратура", r"регулятор ртнд|ртнд"),
    ("Арматура", "Клапаны", r"судовая запорная арматура"),
    ("Электрика", "Автоматы и контакторы", r"^км ?\d|^кпм|рнп-|дпс-|змр"),
    ("Насосное и механическое", "Двигатели и ЗИП", r"^фильтр$|данфас"),
    ("Электрика", "Реле, датчики, приборы", r"извещатель"),
    ("Прочее", "Камбузное оборудование", r"посуда к плите"),
    ("Трубопроводная обвязка", "Пожарное оборудование", r"ствол пожарн"),
    ("Прочее", "Такелаж и корпусное", r"противовес|скоб|гак|весы|вентилятор|кранец|трал|диски тормоз|стекло с обогрев"),
]


def categorize(name):
    """Товарная группа и подгруппа по наименованию."""
    low = name.lower()
    for group, sub, pattern in CATEGORIES:
        if re.search(pattern, low):
            return group, sub
    return "Не разнесено", ""


def norm(v):
    if v is None:
        return ""
    if isinstance(v, float) and v == int(v):
        v = int(v)
    return re.sub(r"\s+", " ", str(v)).strip()


def num(v):
    v = norm(v)
    if not v:
        return None
    v = v.replace(" ", "").replace(" ", "").replace(",", ".")
    try:
        return float(v)
    except ValueError:
        return None


def read_rows(path):
    if path.suffix == ".csv":
        with path.open(encoding="utf-8") as fh:
            return [[norm(v) for v in row] for row in csv.reader(fh)]
    if path.suffix == ".xls":
        wb = xlrd.open_workbook(path)
        # в части файлов первые листы пустые, данные лежат дальше
        sh = next((s for s in wb.sheets() if s.nrows), wb.sheet_by_index(0))
        return [[norm(sh.cell_value(r, c)) for c in range(sh.ncols)]
                for r in range(sh.nrows)]
    ws = openpyxl.load_workbook(path, data_only=True).worksheets[0]
    return [[norm(v) for v in row] for row in ws.iter_rows(values_only=True)]


def cell(row, idx):
    if idx is None or idx >= len(row):
        return ""
    return row[idx]


def is_group_header(row, spec):
    """Строка-заголовок раздела: есть текст в наименовании, но нет ни цены,
    ни количества, ни суммы."""
    if not cell(row, spec["name"]):
        return False
    for key in ("qty", "stock", "price", "total"):
        if num(cell(row, spec[key])):
            return False
    low = cell(row, spec["name"]).lower()
    if any(low.startswith(m) for m in EQUIPMENT_MODELS):
        return True
    return any(low.startswith(m) or m in low for m in GROUP_MARKERS)


def condition(name):
    low = name.lower()
    if "б/у" in low or "б-у" in low:
        return "б/у"
    if "нов" in low:
        return "новый"
    return "не указано"


def completeness(name):
    low = name.lower()
    if "брак" in low:
        return "брак"
    if "не комплект" in low or "в разборе" in low or "без " in low:
        return "не комплект"
    return "комплект"


def origin(name):
    low = name.lower()
    for word, label in (("китай", "Китай"), ("корея", "Корея"),
                        ("росси", "Россия"), ("немец", "Германия"),
                        ("украин", "Украина")):
        if word in low:
            return label
    return ""


def parse(spec):
    path = SRC / spec["file"]
    rows = read_rows(path)
    lo, hi = CONTACTORS.get(spec["file"], (-1, -1))
    out, group, model = [], "", ""

    for i in range(spec["start"], len(rows)):
        row = rows[i]
        name = cell(row, spec["name"])
        if "итого" in " ".join(row).lower():
            continue

        # блок контакторов: имя берем со следующей строки
        if lo <= i < hi:
            name = cell(rows[i + 1], spec["name"]) if i + 1 < len(rows) else name
            qty = stock = num(cell(row, spec["stock"]))
        else:
            if not name:
                continue
            if is_group_header(row, spec):
                if any(name.lower().startswith(m) for m in EQUIPMENT_MODELS):
                    model = name
                else:
                    group, model = name, ""
                continue
            qty = num(cell(row, spec["qty"]))
            stock = num(cell(row, spec["stock"]))

        price = num(cell(row, spec["price"]))
        total = num(cell(row, spec["total"]))
        if not name or (qty is None and stock is None and price is None):
            continue

        cat, sub = categorize(name)
        if cat == "Не разнесено" and model:
            cat, sub = "Насосное и механическое", "Двигатели и ЗИП"
        out.append({
            "Локация": spec["loc"],
            "Товарная группа": cat,
            "Подгруппа": sub,
            "Группа в файле": group,
            "Оборудование": model,
            "Наименование": name,
            "Чертеж": cell(row, spec["draw"]),
            "Кол-во по учету": qty,
            "Наличие": stock,
            "Цена за ед., руб": price,
            "Сумма, руб": total,
            "Состояние": condition(name),
            "Комплектность": completeness(name),
            "Происхождение": origin(name),
            "Примечание": cell(row, spec["note"]),
        })
    return out


def fmt(v):
    if v is None:
        return ""
    if isinstance(v, float):
        return str(int(v)) if v == int(v) else f"{v:.2f}"
    return str(v)


def write_tsv(path, rows, columns=COLUMNS):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as fh:
        fh.write("\t".join(columns) + "\n")
        for r in rows:
            fh.write("\t".join(fmt(r.get(c)) for c in columns) + "\n")


def main():
    stamp = dt.date.today().isoformat()
    rows = []
    for spec in SOURCES:
        rows.extend(parse(spec))

    for r in rows:
        stock, price = r["Наличие"], r["Цена за ед., руб"]
        r["Сумма по наличию"] = stock * price if stock and price else 0

    in_stock = [r for r in rows if (r["Наличие"] or 0) > 0]
    priced = [r for r in in_stock if r["Цена за ед., руб"]]
    unpriced = [r for r in in_stock if not r["Цена за ед., руб"]]

    cols = COLUMNS + ["Сумма по наличию"]
    write_tsv(OUT / f"sudovoe_registry_{stamp}.tsv", rows, cols)
    write_tsv(OUT / f"sudovoe_price_list_{stamp}.tsv",
              sorted(priced, key=lambda r: -r["Сумма по наличию"]), cols)
    write_tsv(OUT / f"sudovoe_no_price_{stamp}.tsv",
              sorted(unpriced, key=lambda r: -(r["Наличие"] or 0)), cols)

    total = sum(r["Сумма по наличию"] for r in priced)
    print(f"позиций всего: {len(rows)}")
    print(f"в наличии: {len(in_stock)}, из них с ценой: {len(priced)}, "
          f"без цены: {len(unpriced)}")
    print(f"стоимость склада по наличию и цене: {total:,.0f} руб"
          .replace(",", " "))

    print("\nпо локациям (наличие > 0):")
    for spec in SOURCES:
        part = [r for r in in_stock if r["Локация"] == spec["loc"]]
        s = sum(r["Сумма по наличию"] for r in part)
        no_p = sum(1 for r in part if not r["Цена за ед., руб"])
        print(f"  {spec['loc']:<14} позиций {len(part):>4}  "
              f"единиц {sum(r['Наличие'] for r in part):>6.0f}  "
              f"сумма {s:>12,.0f}  без цены {no_p}".replace(",", " "))

    print("\nтоп-20 по стоимости:")
    for r in sorted(priced, key=lambda r: -r["Сумма по наличию"])[:20]:
        print(f"  {r['Сумма по наличию']:>11,.0f}  {r['Наличие']:>5.0f} x "
              f"{r['Цена за ед., руб']:>9,.0f}  {r['Локация']:<13} "
              f"{r['Наименование'][:60]}".replace(",", " "))

    print("\nтоп-15 позиций без цены (по количеству):")
    for r in sorted(unpriced, key=lambda r: -(r["Наличие"] or 0))[:15]:
        print(f"  {r['Наличие']:>6.0f}  {r['Локация']:<13} "
              f"{r['Наименование'][:65]}")


if __name__ == "__main__":
    main()
