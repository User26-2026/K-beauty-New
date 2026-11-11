#!/usr/bin/env python3
"""Упаковка фактической отгрузки 26.10.2026: что в каком месте лежит.

В Петербурге груз принимают по местам, и по пункту 2.2 Договора на
претензии по количеству есть 3 рабочих дня с даты получения. Значит
нужен список «место — что внутри», а не только общий перечень.

Основной источник — колонка «упаковка» общего перечня: она заполнена
по всем позициям и согласована с итогами файла. Отдельные упаковочные
листы в том же файле идут ПРОВЕРКОЙ: где лист читается, количество по
нему сверяется с перечнем.

Листы набирали в разное время и разными руками, поэтому раскладка
колонок у них гуляет: у листа № 2 шапка сдвинута на колонку
относительно данных, у листов с девятого по тринадцатый шапки нет
вовсе. Поэтому колонки определяются по тексту заголовка, а сдвиг
ловится по первой строке данных; где шапки нет, строка читается по
смыслу ячеек — номер места в конце строки, количество перед ним.
Лист, который так не читается, честно помечается неразобранным, а не
считается пустым.
"""

import pathlib
import re
import sys

from openpyxl import load_workbook

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
import container_list  # noqa: E402

SRC = container_list.SRC

# Как заголовок ложится на наше поле. Порядок важен: «маховик» надо
# проверить раньше «кол-во», иначе колонка «Маховик , кол-во» уедет
# в количество и перечень недосчитается штук.
HEADERS = [
    ("маховик", "Маховиков"),
    ("наимен", "Наименование"),
    ("чертеж", "Чертеж"),
    ("состоя", "Состояние"),
    ("сомтоя", "Состояние"),   # опечатка в листе № 8
    ("кол-во", "Штук"),
    ("вес 1 шт", "вес_ед"),
    ("вес всего", "вес"),
    ("итого вес", "вес"),
    ("всего вес", "вес"),
    ("вес", "вес_ед"),
    ("ду", "Ду"),
    ("ру", "Ру"),
]

PLACE_IN_TITLE = re.compile(r"№\s*(\d+)")
# Чертеж: цифры, точки и дефисы, обычно с буквенным префиксом серии.
DRAWING = re.compile(r"^[A-Za-zА-Яа-я]*\s*\d{3}[-.\d A-Za-zА-Яа-я]{3,}$")
COND = re.compile(r"^(нов\w*|б/?у|new)\s*$", re.I)
# «ящик №1», «коробка  № 3», «Коробка №13», «16 мешков», «2 мешка :»
PLACE_IN_CELL = re.compile(r"(ящик|коробк\w*|мешк\w*|мешок)[^\d]*(\d*)", re.I)


def place_of_title(title):
    m = PLACE_IN_TITLE.search(title)
    return int(m.group(1)) if m else 0


def place_of_cell(text):
    """Место из колонки «упаковка» перечня: «коробка №7» дает «№ 7»."""
    t = container_list.text(text)
    if not t:
        return ""
    m = PLACE_IN_CELL.search(t)
    if not m:
        return re.sub(r"\s+", " ", t)
    kind = m.group(1).lower()
    kind = ("ящик" if kind.startswith("ящик")
            else "коробка" if kind.startswith("короб") else "мешки")
    return f"{kind} № {m.group(2)}" if m.group(2) else re.sub(r"\s+", " ", t)


def field(header):
    h = str(header or "").strip().lower()
    if not h:
        return None
    for needle, name in HEADERS:
        if needle in h:
            return name
    return None


def header_row(ws):
    """Строка шапки и карта колонок. Первая колонка на поле — главная:
    если заголовок повторяется, правая его уже не перебьет."""
    for i, row in enumerate(ws.iter_rows(max_row=8, values_only=True), 1):
        cols = {}
        for j, c in enumerate(row):
            name = field(c)
            if name and name not in cols.values():
                cols[j] = name
        if "Наименование" in cols.values() and "Чертеж" in cols.values():
            return i, cols
    return None, {}


def shift_of(ws, start, cols):
    """На сколько колонок данные сдвинуты относительно шапки.

    У листа № 2 в шапке есть лишняя пустая колонка, и без поправки
    количество читается из пустой ячейки, а лист выходит пустым.
    Проверяем по первой строке данных: там, где по шапке стоит
    наименование, должен быть текст, а в количестве — число.
    """
    name_col = next(j for j, n in cols.items() if n == "Наименование")
    qty_col = next((j for j, n in cols.items() if n == "Штук"), None)
    for row in ws.iter_rows(min_row=start + 1, max_row=start + 1,
                            values_only=True):
        for d in (0, -1, 1):
            nj, qj = name_col + d, (qty_col + d) if qty_col is not None else None
            if not (0 <= nj < len(row)):
                continue
            if not container_list.text(row[nj]):
                continue
            if qj is None or not (0 <= qj < len(row)):
                return d
            if container_list.num(row[qj]):
                return d
    return 0


def sheets(src=SRC):
    """Разобранные упаковочные листы: {номер места: [строки]} и список
    неразобранных номеров."""
    wb = load_workbook(src, data_only=True)
    parsed, failed = {}, []
    for ws in wb.worksheets:
        if not ws.title.lower().startswith("упак"):
            continue
        place = place_of_title(ws.title)
        start, cols = header_row(ws)
        if not start:
            failed.append(place)
            continue
        d = shift_of(ws, start, cols)
        rows = []
        for row in ws.iter_rows(min_row=start + 1, values_only=True):
            rec = {n: row[j + d] for j, n in cols.items()
                   if 0 <= j + d < len(row)}
            name = container_list.text(rec.get("Наименование"))
            if not name or name.lower().startswith("итого"):
                continue
            qty = container_list.num(rec.get("Штук"))
            if not qty:
                continue
            draw = container_list.text(rec.get("Чертеж"))
            rows.append({
                "Наименование": name,
                "Чертеж": "" if container_list.NO_DRAW.search(draw) else draw,
                "Состояние": container_list.text(rec.get("Состояние")),
                "Штук": int(qty),
            })
        if rows:
            parsed[place] = rows
        else:
            failed.append(place)
    for ws in wb.worksheets:
        if not ws.title.lower().startswith("упак"):
            continue
        place = place_of_title(ws.title)
        if place not in failed:
            continue
        rows = headerless(ws)
        if rows:
            parsed[place] = rows
            failed.remove(place)
    return parsed, sorted(failed)


def headerless(ws):
    """Лист без шапки: читаем строку по смыслу ячеек.

    Опорная точка — ячейка с номером места в конце строки («коробка
    №9»). Количество — последнее число перед ней, чертеж и состояние
    узнаются по виду, наименование — первый длинный текст.
    """
    rows = []
    for row in ws.iter_rows(values_only=True):
        cells = list(row)
        anchor = None
        for j, c in enumerate(cells):
            t = container_list.text(c)
            if t and PLACE_IN_CELL.match(t):
                anchor = j
        if anchor is None:
            continue
        qty = None
        for c in reversed(cells[:anchor]):
            n = container_list.num(c)
            if n:
                qty = int(n)
                break
        name = draw = cond = ""
        for c in cells[:anchor]:
            t = container_list.text(c)
            if not t or container_list.num(t) is not None:
                continue
            if COND.match(t):
                cond = t
            elif DRAWING.match(t) and not draw:
                draw = t
            elif not name and len(t) > 6:
                name = t
        if not name or not qty:
            continue
        rows.append({
            "Наименование": name,
            "Чертеж": "" if container_list.NO_DRAW.search(draw) else draw,
            "Состояние": cond,
            "Штук": qty,
        })
    return rows


def by_place(src=SRC):
    """Позиции перечня, разложенные по местам упаковки."""
    out = {}
    for r in container_list.read(src):
        out.setdefault(place_of_cell(r["Упаковка"]) or "без упаковки",
                       []).append(r)
    return out


def check(src=SRC):
    """Сверка перечня с упаковочными листами по числу штук в месте."""
    parsed, failed = sheets(src)
    places = by_place(src)
    out = []
    for num, rows in sorted(parsed.items()):
        listed = sum(r["Штук"] for r in rows)
        keys = [k for k in places
                if re.search(rf"№\s*{num}$", k)]
        in_list = sum(r["Штук"] for k in keys for r in places[k])
        out.append({
            "Место": num,
            "По упаковочному листу, шт": listed,
            "По перечню, шт": in_list,
            "Расхождение": in_list - listed,
        })
    return out, failed


if __name__ == "__main__":
    places = by_place()
    print("ПО ПЕРЕЧНЮ, места упаковки:")
    for k in sorted(places, key=lambda x: (x == "без упаковки", x)):
        rs = places[k]
        print(f"  {k:<22}{len(rs):>3} поз, {sum(r['Штук'] for r in rs):>5} шт, "
              f"{sum(r['вес'] for r in rs):>8.1f} кг")
    rows, failed = check()
    print("\nСВЕРКА С УПАКОВОЧНЫМИ ЛИСТАМИ:")
    for r in rows:
        mark = "" if r["Расхождение"] == 0 else "  <-- расходится"
        print(f"  место № {r['Место']:<3} лист {r['По упаковочному листу, шт']:>4} шт, "
              f"перечень {r['По перечню, шт']:>4} шт{mark}")
    if failed:
        print(f"\nне разобрались листы: {', '.join('№' + str(f) for f in failed)}")
