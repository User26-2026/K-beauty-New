#!/usr/bin/env python3
"""Варианты отгрузки в Петербург: что дешевле увезти всю арматуру.

Считаем не ставку за контейнер, а полные расходы: контейнер плюс
автодоставка во Владивостоке плюс автодоставка в Петербурге. Плюс вес
необрезной доски, которая едет вместе с грузом и съедает лимит.

Сырая необрезная доска весит около 700 кг на кубометр, а на перекладку
уходит примерно 0,035 куб. м на тонну груза. То есть каждая тонна
арматуры тянет за собой 25 кг доски, и в двадцатифутовом это полтонны.

Остаток, который не влез, считаем по тарифу транспортной компании —
81 руб/кг по реальным распискам от 24.09.2026.

Погрузка тоже стоит денег и зависит от числа контейнеров: подача
погрузчика 12 000 руб плюс 2 000 руб за час, четверо грузчиков по
750 руб в час, на контейнер уходит около 4 часов. Если оба
двадцатифутовых грузим в один день, подача погрузчика платится один
раз — это 12 тысяч разницы.

Результат: outputs/marine_equipment/Логистика_варианты.xlsx
"""

import pathlib
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
from build_container_spb import collect  # noqa: E402

OUT = pathlib.Path(__file__).resolve().parents[2] / "outputs" / "marine_equipment"

FONT = "Arial"
HEAD_FILL = PatternFill("solid", fgColor="1F3864")
HEAD_FONT = Font(name=FONT, size=11, bold=True, color="FFFFFF")
TITLE = Font(name=FONT, size=14, bold=True, color="1F3864")
SUB = Font(name=FONT, size=11, bold=True, color="1F3864")
BOLD = Font(name=FONT, size=10, bold=True)
BASE = Font(name=FONT, size=10)
NOTE = Font(name=FONT, size=9, italic=True, color="666666")
WIN = PatternFill("solid", fgColor="EAF3EA")
LOSE = PatternFill("solid", fgColor="FDF2E3")
INPUT_FILL = PatternFill("solid", fgColor="FFFF00")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# сырая необрезная доска 25 мм
BOARD_RHO = 700          # кг в кубометре
BOARD_PER_T = 0.035      # кубометров доски на тонну груза
TK_RATE = 81             # руб за килограмм транспортной компанией

# погрузка во Владивостоке
FORKLIFT_CALL = 12_000   # подача погрузчика, за выезд
FORKLIFT_HOUR = 2_000    # час работы погрузчика
LOADERS = 4              # грузчиков в бригаде
LOADER_HOUR = 750        # руб за час с человека
HOURS_BOX = 4            # часов на контейнер
HOURS_REST = 2           # часов на остаток под транспортную компанию
SAME_DAY = True          # оба двадцатифутовых грузим в один день

# ставки на руках
PRICE = {"20 обычный": 236_000, "20 усиленный": 315_000, "40": 435_000}
AUTO_VL = {20: 27_000, 40: 37_000}
# вывоз контейнера в Петербурге: по двадцатифутовому ставка
# подтверждена, такая же как во Владивостоке. По сорокафутовому пока
# принята по аналогии
AUTO_SPB = {20: 27_000, 40: 37_000}

BOXES = {
    "20 обычный": dict(ft=20, kg=22_000, m3=33.0),
    "20 усиленный": dict(ft=20, kg=27_000, m3=33.0),
    "40": dict(ft=40, kg=26_000, m3=67.2),
}

SCENARIOS = [
    ("Один сорокафутовый", ["40"]),
    ("Два двадцатифутовых по 22 т", ["20 обычный", "20 обычный"]),
    ("Один двадцатифутовый 22 т", ["20 обычный"]),
    ("Один усиленный двадцатифутовый", ["20 усиленный"]),
    ("Сорокафутовый и двадцатифутовый", ["40", "20 обычный"]),
]


def loading(hours, calls=1):
    """Погрузка: подача погрузчика, его часы и бригада грузчиков."""
    return (FORKLIFT_CALL * calls + FORKLIFT_HOUR * hours
            + LOADERS * LOADER_HOUR * hours)


def run_box(kind):
    """Сколько влезет в контейнер с учетом веса и объема доски."""
    box = BOXES[kind]
    cargo_lim = box["kg"] / (1 + BOARD_PER_T * BOARD_RHO / 1000)
    board_m3 = cargo_lim / 1000 * BOARD_PER_T
    a, b, rest, kg, m3 = collect(cargo_lim, box["m3"] - board_m3)
    return dict(kind=kind, ft=box["ft"], kg=kg, m3=m3,
                board_kg=kg / 1000 * BOARD_PER_T * BOARD_RHO,
                board_m3=kg / 1000 * BOARD_PER_T,
                positions=len(a) + len(b),
                cost=PRICE[kind] + AUTO_VL[box["ft"]] + AUTO_SPB[box["ft"]])


def head(ws, row, ncols, height=34):
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


def main():
    _, _, rest_all, kg_all, _ = collect(10 ** 9, 10 ** 9)
    total = kg_all + sum(r["вес"] for r in rest_all)

    cache = {k: run_box(k) for k in BOXES}
    results = []
    for name, kinds in SCENARIOS:
        runs = [cache[k] for k in kinds]
        capacity = sum(r["kg"] for r in runs)
        moved = min(capacity, total)
        rest = total - moved
        cost = sum(r["cost"] for r in runs)
        # контейнеры: в один день можно погрузить только одинаковые
        # двадцатифутовые, подача погрузчика тогда одна
        calls = 1 if (SAME_DAY and len(set(kinds)) == 1) else len(kinds)
        load = loading(HOURS_BOX * len(kinds), calls)
        # остаток грузим отдельным днем: своя подача и свои часы
        load_rest = loading(HOURS_REST) if rest else 0
        full = cost + load + rest * TK_RATE + load_rest
        results.append(dict(
            name=name, kinds=kinds, runs=runs, capacity=capacity,
            moved=moved, rest=rest, cost=cost, load=load,
            load_rest=load_rest, calls=calls,
            tk=rest * TK_RATE, full=full,
            per_ton=(cost + load) / (moved / 1000) if moved else 0,
            board_kg=sum(r["board_kg"] for r in runs),
        ))

    wb = Workbook()
    ws = wb.active
    ws.title = "Варианты"
    ws["A1"] = "Чем везти арматуру в Петербург"
    ws["A1"].font = TITLE
    ws["A2"] = (f"Вся арматура — {total / 1000:.2f} т. Считаем полные "
                "расходы: контейнер, автодоставка во Владивостоке и "
                "автодоставка в Петербурге. Вес сырой необрезной доски "
                "входит в лимит контейнера.")
    ws["A2"].font = NOTE

    row = 4
    ws.cell(row=row, column=1, value="Ставки, с которыми считаем").font = SUB
    row += 1
    for label, value, note in [
        ("Двадцатифутовый обычный, 22 т", PRICE["20 обычный"], ""),
        ("Двадцатифутовый усиленный, 27 т", PRICE["20 усиленный"],
         "грузоподъемность не работает: объем кончается раньше"),
        ("Сорокафутовый, 26 т", PRICE["40"], ""),
        ("Автодоставка 20-футового", AUTO_VL[20],
         "во Владивостоке и столько же в Петербурге — подтверждено"),
        ("Автодоставка 40-футового", AUTO_VL[40],
         "во Владивостоке, в Петербурге принята по аналогии — уточнить"),
        ("Транспортная компания, руб за кг", TK_RATE,
         "по распискам от 24.09.2026"),
        ("Подача погрузчика, за выезд", FORKLIFT_CALL,
         "платится один раз за день погрузки"),
        ("Погрузчик, руб за час", FORKLIFT_HOUR, ""),
        ("Грузчики, человек", LOADERS, ""),
        ("Грузчик, руб за час", LOADER_HOUR,
         f"бригада обходится в {LOADERS * LOADER_HOUR:,.0f} руб в час"
         .replace(",", " ")),
        ("Часов на контейнер", HOURS_BOX,
         f"погрузка контейнера — {loading(HOURS_BOX):,.0f} руб"
         .replace(",", " ")),
        ("Доска сырая, кг в кубометре", BOARD_RHO, ""),
        ("Доски на тонну груза, куб. м", BOARD_PER_T,
         "43,4 кв. м настилов на 27,3 т арматуры"),
    ]:
        ws.cell(row=row, column=1, value=label).font = BASE
        cell = ws.cell(row=row, column=2, value=value)
        cell.font = BOLD
        cell.border = BORDER
        cell.fill = INPUT_FILL
        cell.number_format = "#,##0.000" if value < 1 else "#,##0"
        ws.cell(row=row, column=3, value=note).font = NOTE
        row += 1
    cell = ws.cell(row=row, column=1,
                   value="По двадцатифутовому ставка в Петербурге "
                         "подтверждена: 27 тысяч, как и во Владивостоке. "
                         "По сорокафутовому принята по аналогии, 37 тысяч "
                         "— уточнить. Все ставки в желтых ячейках.")
    cell.font = NOTE
    cell.alignment = Alignment(wrap_text=True)

    row += 3
    ws.cell(row=row, column=1, value="Варианты").font = SUB
    row += 1
    cols = ["Вариант", "Контейнеры", "Увезет, т", "Останется, т",
            "Доска, кг", "Контейнеры и авто, руб", "Погрузка, руб",
            "Остаток через ТК, руб", "Погрузка остатка, руб",
            "ВСЕГО за всю арматуру, руб", "Руб за увезенную тонну"]
    for i, h in enumerate(cols, start=1):
        ws.cell(row=row, column=i, value=h)
    head(ws, row, len(cols))

    best = min(results, key=lambda r: r["full"])
    first = row + 1
    for r in results:
        row += 1
        values = [r["name"], len(r["kinds"]), r["moved"] / 1000,
                  r["rest"] / 1000, round(r["board_kg"]), r["cost"],
                  r["load"], r["tk"], r["load_rest"] or None, r["full"],
                  r["per_ton"]]
        for i, v in enumerate(values, start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BOLD if r is best else BASE
            cell.border = BORDER
            if i in (3, 4):
                cell.number_format = "#,##0.00"
            if i >= 5:
                cell.number_format = "#,##0"
            if r is best:
                cell.fill = WIN
            elif i == 10:
                cell.fill = LOSE

    row += 2
    ws.cell(row=row, column=1, value="Что в каждом контейнере").font = SUB
    row += 1
    cols2 = ["Вариант", "Контейнер", "Груз, т", "Доска, кг", "Объем, куб. м",
             "Позиций", "Контейнер, руб", "Авто туда и обратно, руб"]
    for i, h in enumerate(cols2, start=1):
        ws.cell(row=row, column=i, value=h)
    head(ws, row, len(cols2), height=28)
    for r in results:
        for i, run in enumerate(r["runs"], start=1):
            row += 1
            auto = AUTO_VL[run["ft"]] + AUTO_SPB[run["ft"]]
            values = [r["name"] if i == 1 else "", f"{run['ft']} футов",
                      run["kg"] / 1000, round(run["board_kg"]),
                      round(run["m3"], 1), run["positions"],
                      PRICE[run["kind"]], auto]
            for j, v in enumerate(values, start=1):
                cell = ws.cell(row=row, column=j, value=v)
                cell.font = BASE
                cell.border = BORDER
                if j == 3:
                    cell.number_format = "#,##0.00"
                if j in (4, 5, 7, 8):
                    cell.number_format = "#,##0"

    row += 2
    ws.cell(row=row, column=1, value="Выводы").font = SUB
    row += 1
    forty = next(r for r in results if r["name"] == "Один сорокафутовый")
    two20 = next(r for r in results
                 if r["name"] == "Два двадцатифутовых по 22 т")
    one20 = next(r for r in results
                 if r["name"] == "Один двадцатифутовый 22 т")
    diff = forty["cost"] - two20["cost"]
    extra_t = (two20["moved"] - forty["moved"]) / 1000
    for text in [
        f"Чтобы увезти всю арматуру, дешевле всего «{best['name']}»: "
        f"{best['full']:,.0f} руб против "
        f"{max(r['full'] for r in results):,.0f} у худшего варианта."
        .replace(",", " "),
        f"Два двадцатифутовых дороже сорокафутового по расходам на "
        f"{-diff:,.0f} руб, но увозят на {extra_t:.2f} т больше и не "
        "оставляют ничего. Через транспортную компанию эти тонны стоили "
        f"бы {(two20['moved'] - forty['moved']) * TK_RATE:,.0f} руб."
        .replace(",", " "),
        f"Один двадцатифутовый самый дешевый по ставке — "
        f"{one20['cost']:,.0f} руб, и самая низкая цена тонны "
        f"{one20['per_ton']:,.0f}. Но он оставляет "
        f"{one20['rest'] / 1000:.2f} т, и если их везти транспортной "
        f"компанией, вариант становится худшим: {one20['full']:,.0f} руб."
        .replace(",", " "),
        "Усиленный двадцатифутовый бессмыслен при любом раскладе: он "
        "того же размера, что обычный, объем кончается на тех же тоннах, "
        "а стоит на 79 тысяч дороже.",
        f"Доска съедает {forty['board_kg']:.0f} кг в сорокафутовом и по "
        f"{two20['runs'][0]['board_kg']:.0f} кг в каждом "
        "двадцатифутовом. Это не мелочь: в двадцатифутовом полтонны "
        "доски это полтонны неувезенной арматуры.",
        "Автодоставка добавляет 54 тысячи на двадцатифутовый и 74 на "
        "сорокафутовый, считая оба конца. На двух контейнерах это "
        "108 тысяч — больше, чем разница в ставке за сами контейнеры.",
        f"Погрузка на двух двадцатифутовых стоит {two20['load']:,.0f} руб "
        f"против {forty['load']:,.0f} на сорокафутовом. Разница всего "
        f"{two20['load'] - forty['load']:,.0f} руб, и только потому, что "
        "оба контейнера грузим в один день и подачу погрузчика платим "
        f"один раз. Разными днями это {loading(HOURS_BOX, 1) * 2:,.0f} "
        "руб.".replace(",", " "),
        f"У сорокафутового погрузка не заканчивается контейнером: "
        f"{forty['rest'] / 1000:.2f} т остатка надо грузить отдельным "
        f"днем, это еще одна подача погрузчика и бригада — "
        f"{forty['load_rest']:,.0f} руб сверх тарифа ТК."
        .replace(",", " "),
        "Выгрузка в Петербурге в расчет не входит: ставок на тамошний "
        "погрузчик и грузчиков у нас нет. По владивостокским это "
        f"около {loading(HOURS_BOX):,.0f} руб на контейнер, и считать "
        "надо отдельно.".replace(",", " "),
        "Сохранность груза тоже за два двадцатифутовых: при равной "
        "развеске в каждом едет около 13,7 т вместо 25, пирог ниже, "
        "на нижние задвижки давит вдвое меньше, и мелочь не оказывается "
        "под тяжелым. Разбивку считает build_two_containers.py.",
    ]:
        cell = ws.cell(row=row, column=1, value="— " + text)
        cell.font = BASE
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[row].height = 32
        row += 1

    widths(ws, [38, 12, 11, 12, 11, 17, 13, 16, 15, 20, 15])

    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / "Логистика_варианты.xlsx"
    wb.save(path)

    print(f"готово: {path}\n")
    print(f"вся арматура: {total / 1000:.2f} т\n")
    print(f"{'вариант':<34}{'увезет':>9}{'останется':>11}"
          f"{'контейнеры':>12}{'погрузка':>10}{'остаток ТК':>12}"
          f"{'ВСЕГО':>11}{'руб/т':>9}")
    for r in results:
        mark = " <-- дешевле всего" if r is best else ""
        print(f"{r['name']:<34}{r['moved'] / 1000:>8.2f}т"
              f"{r['rest'] / 1000:>10.2f}т{r['cost']:>12,.0f}"
              f"{r['load'] + r['load_rest']:>10,.0f}{r['tk']:>12,.0f}"
              f"{r['full']:>11,.0f}{r['per_ton']:>9,.0f}{mark}"
              .replace(",", " "))


if __name__ == "__main__":
    main()
