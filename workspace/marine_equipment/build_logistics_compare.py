#!/usr/bin/env python3
"""Варианты отгрузки в Петербург: что дешевле увезти всю арматуру.

Считаем не ставку за контейнер, а полные расходы: контейнер, автодоставка
во Владивостоке, автодоставка в Петербурге и погрузка. Плюс сырая
необрезная доска, которая едет вместе с грузом и съедает и лимит по весу,
и место по объему.

Доска весит около 700 кг на кубометр, на перекладку уходит примерно
0,035 куб. м на тонну груза. Каждая тонна арматуры тянет за собой 25 кг
доски, и доска считается от фактически набранного веса, а не от
паспортной грузоподъемности.

Погрузка: подача погрузчика 12 000 руб, 2 000 руб за час работы, четверо
грузчиков по 750 руб в час, около 4 часов на контейнер. Подача платится за
выезд, поэтому одинаковые контейнеры в один день дешевле, чем в разные.

Остаток, который не влез, считаем по тарифу транспортной компании —
81 руб/кг по реальным распискам от 24.09.2026.

Главное, что показывает расчет: груз упирается в объем раньше, чем в
грузоподъемность. Поэтому каждый вариант считается в двух очередностях
погрузки: как просил покупатель (сначала крупные клинкеты) и
теоретический максимум (самое плотное вперед). Это разные тонны в одном
и том же корпусе, и для усиленного двадцатифутового разница решающая.

Результат: outputs/marine_equipment/Логистика_варианты.xlsx
"""

import pathlib
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
from build_container_spb import BLOCK_A, collect  # noqa: E402

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
BOARD_RHO = 700                    # кг в кубометре
BOARD_PER_T = 0.035                # кубометров доски на тонну груза
BOARD_PIECE = 4.0 * 0.20 * 0.025   # одна доска 4 м, куб. м
TK_RATE = 81                       # руб за килограмм транспортной компанией

# погрузка во Владивостоке
FORKLIFT_CALL = 12_000   # подача погрузчика, за выезд
FORKLIFT_HOUR = 2_000    # час работы погрузчика
LOADERS = 4              # грузчиков в бригаде
LOADER_HOUR = 750        # руб за час с человека
HOURS_BOX = 4            # часов на контейнер
HOURS_REST = 2           # часов на остаток под транспортную компанию
SAME_DAY = True          # одинаковые контейнеры грузим в один день

# ставки на руках
PRICE = {"20 обычный": 236_000, "20 усиленный": 315_000, "40": 435_000}
AUTO_VL = {20: 27_000, 40: 37_000}
# вывоз контейнера в Петербурге: по двадцатифутовому ставка подтверждена,
# такая же как во Владивостоке. По сорокафутовому принята по аналогии
AUTO_SPB = {20: 27_000, 40: 37_000}

BOXES = {
    "20 обычный": dict(ft=20, kg=22_000, m3=33.0, name="20 футов, 22 т"),
    "20 усиленный": dict(ft=20, kg=27_000, m3=33.0,
                         name="20 футов усиленный, 27 т"),
    "40": dict(ft=40, kg=26_000, m3=67.2, name="40 футов, 26 т"),
}

SCENARIOS = [
    ("Два двадцатифутовых по 22 т", ["20 обычный", "20 обычный"]),
    ("Один усиленный двадцатифутовый", ["20 усиленный"]),
    ("Один сорокафутовый", ["40"]),
    ("Один двадцатифутовый 22 т", ["20 обычный"]),
    ("Усиленный и обычный двадцатифутовые", ["20 усиленный", "20 обычный"]),
    ("Сорокафутовый и двадцатифутовый", ["40", "20 обычный"]),
]


def board(kg):
    """Доска под фактическую загрузку: вес и объем."""
    m3 = kg / 1000 * BOARD_PER_T
    return m3 * BOARD_RHO, m3


def by_queue(r):
    """Очередность покупателя: сначала согласованный список менеджера,
    потом добор со склада по цене за килограмм."""
    per_kg = r["Сумма"] / r["вес"] if r["вес"] and r["Сумма"] else 0
    return (0 if r["Блок"] == BLOCK_A else 1, -per_kg)


def by_density(r):
    """Теоретический максимум тонн: самое плотное вперед."""
    return (-(r["вес"] / r["объем"] if r["объем"] else 0),)


ORDERS = [
    ("По очередности покупателя", by_queue,
     "Сначала согласованный список менеджера целиком, потом добор со "
     "склада по цене за килограмм. Так и будем грузить."),
    ("Самое плотное вперед", by_density,
     "Теоретический максимум тонн в корпус. Так грузить нельзя: плотное "
     "— это мелочь в ящиках, а крупные клинкеты, которые покупатель "
     "просил первыми, самые объемные и уехали бы последними."),
]


def rows_all():
    a, b, rest, _, _ = collect(10 ** 9, 10 ** 9)
    out = []
    for r in a + b + rest:
        qty = int(r["Штук"] or 0)
        if qty <= 0:
            continue
        out.append(dict(r, **{"ост": qty, "объем_ед": r["объем"] / qty}))
    return out


def pack(pool, gross_kg, vol_m3, key):
    """Сколько влезет в один корпус. Позиции дробятся по штукам, доска
    пересчитывается на каждом шаге."""
    kg = m3 = 0.0
    taken = []
    for r in sorted(pool, key=key):
        n = int(r["ост"])
        if n <= 0:
            continue
        dw, dv = r["вес_ед"], r["объем_ед"]
        got = 0
        while got < n:
            bw, bv = board(kg + dw)
            if kg + dw + bw <= gross_kg and m3 + dv + bv <= vol_m3:
                kg += dw
                m3 += dv
                got += 1
            else:
                break
        if got:
            taken.append((r, got))
            r["ост"] -= got
    return taken, kg, m3


def loading(hours, calls=1):
    """Погрузка: подача погрузчика, его часы и бригада грузчиков."""
    return (FORKLIFT_CALL * calls + FORKLIFT_HOUR * hours
            + LOADERS * LOADER_HOUR * hours)


def run(name, kinds, key):
    """Загрузить контейнеры по очереди и посчитать остаток и расходы."""
    pool = rows_all()
    boxes = []
    for kind in kinds:
        spec = BOXES[kind]
        taken, kg, m3 = pack(pool, spec["kg"], spec["m3"], key)
        bw, bm3 = board(kg)
        boxes.append(dict(
            kind=kind, ft=spec["ft"], kg=kg, m3=m3, board_kg=bw,
            board_m3=bm3, positions=len(taken),
            units=sum(q for _, q in taken),
            gross=kg + bw, lim_kg=spec["kg"], lim_m3=spec["m3"],
            limit=("вес" if spec["kg"] - (kg + bw) < 150 else
                   "объем" if spec["m3"] - (m3 + bm3) < 0.5 else
                   "груз кончился"),
            cost=PRICE[kind] + AUTO_VL[spec["ft"]] + AUTO_SPB[spec["ft"]],
        ))
    rest_kg = sum(r["ост"] * r["вес_ед"] for r in pool)
    rest_pos = sum(1 for r in pool if r["ост"] > 0)

    calls = 1 if (SAME_DAY and len(set(kinds)) == 1) else len(kinds)
    load = loading(HOURS_BOX * len(kinds), calls)
    load_rest = loading(HOURS_REST) if rest_kg > 1 else 0
    cost = sum(b["cost"] for b in boxes)
    moved = sum(b["kg"] for b in boxes)
    tk = rest_kg * TK_RATE
    return dict(
        name=name, kinds=kinds, boxes=boxes, moved=moved, rest=rest_kg,
        rest_pos=rest_pos, cost=cost, load=load, load_rest=load_rest,
        tk=tk, full=cost + load + tk + load_rest,
        per_ton=(cost + load) / (moved / 1000) if moved else 0,
        board_kg=sum(b["board_kg"] for b in boxes),
        board_m3=sum(b["board_m3"] for b in boxes),
    )


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


def rates_block(ws, row):
    ws.cell(row=row, column=1, value="Ставки, с которыми считаем").font = SUB
    row += 1
    for label, value, note in [
        ("Двадцатифутовый обычный, 22 т", PRICE["20 обычный"], ""),
        ("Двадцатифутовый усиленный, 27 т", PRICE["20 усиленный"],
         "27 тонн паспортных, но внутри те же 33 куб. м"),
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
         "25 кг доски на каждую тонну арматуры"),
    ]:
        ws.cell(row=row, column=1, value=label).font = BASE
        cell = ws.cell(row=row, column=2, value=value)
        cell.font = BOLD
        cell.border = BORDER
        cell.fill = INPUT_FILL
        cell.number_format = "#,##0.000" if value < 1 else "#,##0"
        ws.cell(row=row, column=3, value=note).font = NOTE
        row += 1
    return row + 1


def table(ws, row, title, note, results):
    ws.cell(row=row, column=1, value=title).font = SUB
    row += 1
    cell = ws.cell(row=row, column=1, value=note)
    cell.font = NOTE
    cell.alignment = Alignment(wrap_text=True)
    row += 2

    cols = ["Вариант", "Корпусов", "Увезет груза, т", "Доска, кг",
            "Брутто, т", "Во что уперся", "Останется, т",
            "Контейнеры и авто, руб", "Погрузка, руб", "Остаток через ТК, руб",
            "Погрузка остатка, руб", "ВСЕГО, руб", "Руб за увезенную тонну"]
    for i, h in enumerate(cols, start=1):
        ws.cell(row=row, column=i, value=h)
    head(ws, row, len(cols), height=44)

    best = min(results, key=lambda r: r["full"])
    for r in results:
        row += 1
        limits = ", ".join(sorted({b["limit"] for b in r["boxes"]}))
        values = [r["name"], len(r["kinds"]), r["moved"] / 1000,
                  round(r["board_kg"]),
                  (r["moved"] + r["board_kg"]) / 1000, limits,
                  r["rest"] / 1000, r["cost"], r["load"], r["tk"],
                  r["load_rest"] or None, r["full"], r["per_ton"]]
        for i, v in enumerate(values, start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BOLD if r is best else BASE
            cell.border = BORDER
            if i in (3, 5, 7):
                cell.number_format = "#,##0.00"
            if i in (4, 8, 9, 10, 11, 12, 13):
                cell.number_format = "#,##0"
            if r is best:
                cell.fill = WIN
            elif i == 12:
                cell.fill = LOSE
    return row + 2, best


def boxes_block(ws, row, results, label):
    ws.cell(row=row, column=1,
            value=f"Что в каждом корпусе — {label.lower()}").font = SUB
    row += 1
    cols = ["Вариант", "Корпус", "Груз, т", "Доска, кг", "Брутто, т",
            "Из лимита", "Объем груза, куб. м", "Доска, куб. м",
            "Объем занят", "Во что уперся", "Позиций", "Штук"]
    for i, h in enumerate(cols, start=1):
        ws.cell(row=row, column=i, value=h)
    head(ws, row, len(cols), height=32)
    for r in results:
        for i, b in enumerate(r["boxes"], start=1):
            row += 1
            values = [r["name"] if i == 1 else "", BOXES[b["kind"]]["name"],
                      b["kg"] / 1000, round(b["board_kg"]),
                      b["gross"] / 1000, b["gross"] / b["lim_kg"],
                      round(b["m3"], 1), round(b["board_m3"], 2),
                      (b["m3"] + b["board_m3"]) / b["lim_m3"], b["limit"],
                      b["positions"], b["units"]]
            for j, v in enumerate(values, start=1):
                cell = ws.cell(row=row, column=j, value=v)
                cell.font = BASE
                cell.border = BORDER
                if j in (3, 5):
                    cell.number_format = "#,##0.00"
                if j in (4, 7, 11, 12):
                    cell.number_format = "#,##0"
                if j == 8:
                    cell.number_format = "#,##0.00"
                if j in (6, 9):
                    cell.number_format = "0%"
                    if v > 0.97:
                        cell.fill = LOSE
    return row + 2


def main():
    rows = rows_all()
    total = sum(r["вес"] for r in rows)
    total_m3 = sum(r["объем"] for r in rows)

    runs = {}
    for label, key, _ in ORDERS:
        runs[label] = [run(name, kinds, key) for name, kinds in SCENARIOS]

    wb = Workbook()
    ws = wb.active
    ws.title = "Варианты"
    ws["A1"] = "Чем везти арматуру в Петербург"
    ws["A1"].font = TITLE
    ws["A2"] = (f"Весь перечень — {total / 1000:.2f} т и {total_m3:.1f} куб. м, "
                f"это {total / total_m3:.0f} кг на кубометр. Расходы считаем "
                "полностью: контейнер, авто на обоих концах, погрузка, "
                "доска и остаток через транспортную компанию.")
    ws["A2"].font = NOTE

    row = rates_block(ws, 4)
    bests = {}
    for label, _, note in ORDERS:
        row, best = table(ws, row, label, note, runs[label])
        bests[label] = best
    main_label = ORDERS[0][0]
    row = boxes_block(ws, row, runs[main_label], main_label)

    row += 1
    ws.cell(row=row, column=1, value="Доска: сколько покупать").font = SUB
    row += 1
    for r in runs[main_label][:3]:
        cell = ws.cell(
            row=row, column=1,
            value=f"— {r['name']}: {r['board_m3']:.2f} куб. м, это "
                  f"{r['board_m3'] / BOARD_PIECE:.0f} досок 4 м х 200 х 25 "
                  f"и {r['board_kg']:.0f} кг веса, которые занимают место "
                  "груза.")
        cell.font = BASE
        row += 1

    row += 1
    ws.cell(row=row, column=1, value="Выводы").font = SUB
    row += 1
    q = {r["name"]: r for r in runs["По очередности покупателя"]}
    d = {r["name"]: r for r in runs["Самое плотное вперед"]}
    hard = d["Один усиленный двадцатифутовый"]["boxes"][0]
    soft = q["Один усиленный двадцатифутовый"]["boxes"][0]
    two20 = q["Два двадцатифутовых по 22 т"]
    forty = q["Один сорокафутовый"]
    best = bests["По очередности покупателя"]
    for text in [
        f"Усиленный двадцатифутовый паспортно берет 27 т, но до них не "
        f"догружается: внутри те же 33 куб. м. Даже если грузить самое "
        f"плотное вперед, в него входит {hard['kg'] / 1000:.2f} т груза "
        f"плюс {hard['board_kg']:.0f} кг доски — "
        f"{hard['gross'] / 1000:.2f} т брутто, и объем кончается. "
        f"Паспортной грузоподъемности не хватает "
        f"{(27_000 - hard['gross']) / 1000:.2f} т до предела.",
        f"В нашей очередности, с крупными клинкетами первыми, в усиленный "
        f"входит еще меньше — {soft['kg'] / 1000:.2f} т. Клинкеты самые "
        f"объемные на тонну: 530 кг в кубометре против 967 у средней "
        "арматуры. Чем раньше они грузятся, тем меньше тонн в корпус.",
        f"Доска не бесплатная ни по весу, ни по месту: "
        f"{hard['board_kg']:.0f} кг и {hard['board_m3']:.2f} куб. м в "
        "усиленном. Это те же полтонны неувезенной арматуры, и считать "
        "ее надо от фактического веса, а не от паспортного.",
        f"По деньгам в нашей очередности дешевле всего «{best['name']}»: "
        f"{best['full']:,.0f} руб против "
        f"{max(r['full'] for r in runs['По очередности покупателя']):,.0f} "
        "у худшего варианта.".replace(",", " "),
        f"Два обычных двадцатифутовых увозят {two20['moved'] / 1000:.2f} т "
        f"и не оставляют ничего. Сорокафутовый — {forty['moved'] / 1000:.2f} "
        f"т, остаток {forty['rest'] / 1000:.2f} т через транспортную "
        f"компанию стоит {forty['tk']:,.0f} руб плюс отдельный день "
        "погрузки.".replace(",", " "),
        "Сорокафутовый при этом загружен по объему примерно наполовину: "
        "у него упор в вес, а не в место. У любого двадцатифутового "
        "наоборот.",
        "Выгрузка в Петербурге в расчет не входит: ставок на тамошний "
        f"погрузчик и грузчиков нет. По владивостокским это около "
        f"{loading(HOURS_BOX):,.0f} руб на контейнер."
        .replace(",", " "),
    ]:
        cell = ws.cell(row=row, column=1, value="— " + text)
        cell.font = BASE
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[row].height = 32
        row += 1

    widths(ws, [38, 11, 14, 11, 11, 14, 12, 17, 13, 16, 15, 16, 15])

    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / "Логистика_варианты.xlsx"
    wb.save(path)

    print(f"готово: {path}\n")
    print(f"весь перечень: {total / 1000:.2f} т, {total_m3:.1f} куб. м, "
          f"{total / total_m3:.0f} кг/куб.м\n")
    for label, _, _ in ORDERS:
        print(f"=== {label} ===")
        print(f"{'вариант':<38}{'увезет':>9}{'брутто':>9}{'упор':>9}"
              f"{'останется':>11}{'ВСЕГО':>11}")
        best = min(runs[label], key=lambda r: r["full"])
        for r in runs[label]:
            mark = "  <-- дешевле" if r is best else ""
            limits = "/".join(sorted({b["limit"] for b in r["boxes"]}))
            print(f"{r['name']:<38}{r['moved'] / 1000:>8.2f}т"
                  f"{(r['moved'] + r['board_kg']) / 1000:>8.2f}т{limits:>9}"
                  f"{r['rest'] / 1000:>10.2f}т{r['full']:>11,.0f}{mark}"
                  .replace(",", " "))
        print()


if __name__ == "__main__":
    main()
