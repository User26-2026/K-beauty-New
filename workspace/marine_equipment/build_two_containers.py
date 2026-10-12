#!/usr/bin/env python3
"""Два двадцатифутовых контейнера в Петербург, с равной развеской.

Берем тот же перечень, что и в плане одного контейнера: согласованный
список менеджера целиком плюс добор со склада во Владивостоке, без
стали и без клапанных коробок. Это 27,3 т. В один сорокафутовый они
не влезают, в два двадцатифутовых влезают с запасом.

Главное здесь не вместимость, а сохранность. Раньше первый контейнер
набивался доверху самым дорогим, а второй забирал остаток. Теперь вес
и объем делятся поровну, и делятся поярусно: каждый контейнер получает
половину тяжелых задвижек, половину средней арматуры, половину ящиков
и половину фильтров. В каждом едет около 13,7 т вместо 25 — пирог
вдвое ниже, и на нижний ряд давит вдвое меньше.

Позиции при этом дробятся по штукам: если партия задвижек перекашивает
контейнер, половина штук уезжает во втором.

Результат: outputs/marine_equipment/План_2_контейнера.xlsx
"""

import pathlib
import sys
from collections import defaultdict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
from build_container_spb import (  # noqa: E402
    COLS, asrow, collect, plural,
)
from build_shipment_spb import (  # noqa: E402
    FLOOR, PACKING, TIERS, item_kind, pack_type, tier,
)
from build_spb_list_xlsx import GROUPS, group_of  # noqa: E402
import gate_geometry as gg  # noqa: E402

OUT = pathlib.Path(__file__).resolve().parents[2] / "outputs" / "marine_equipment"

FONT = "Arial"
HEAD_FILL = PatternFill("solid", fgColor="1F3864")
HEAD_FONT = Font(name=FONT, size=11, bold=True, color="FFFFFF")
TITLE = Font(name=FONT, size=14, bold=True, color="1F3864")
SUB = Font(name=FONT, size=11, bold=True, color="1F3864")
BOLD = Font(name=FONT, size=10, bold=True)
BASE = Font(name=FONT, size=10)
NOTE = Font(name=FONT, size=9, italic=True, color="666666")
ONE_FILL = PatternFill("solid", fgColor="EAF3EA")
TWO_FILL = PatternFill("solid", fgColor="FDF2E3")
WARN_FILL = PatternFill("solid", fgColor="FCE4D6")
GROUP_FILL = PatternFill("solid", fgColor="D9E2F3")
MARK_FILL = PatternFill("solid", fgColor="FFF2CC")
NODRAW_FILL = PatternFill("solid", fgColor="FFFF00")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# двадцатифутовый обычный: 22 тонны по ставке 236 тысяч, внутри
# 33 куб. м. Пол 5 898 х 2 352 мм
BOX_KG = 22_000
BOX_M3 = 33.0
BOXES = 2
FLOOR_L, FLOOR_W = 5.898, 2.352
FLOOR_AREA = FLOOR_L * FLOOR_W

# сырая необрезная доска 25 мм: 700 кг в кубометре, 0,035 куб. м
# на тонну груза
BOARD_RHO = 700
BOARD_PER_T = 0.035

# погрузка во Владивостоке
FORKLIFT_CALL = 12_000
FORKLIFT_HOUR = 2_000
LOADERS = 4
LOADER_HOUR = 750
HOURS_BOX = 4

# ящик на поддоне: больше 600 кг не набираем, иначе его не возьмет
# ни погрузчик, ни нижний ящик под ним
BOX_LIMIT_KG = 600

# с чем сравниваем перекос: 50 кг по ярусу и 0,3 куб. м по объему
# считаем одинаково заметными, по ним и нормируем цену расклада
TOL = 50
TOL_M3 = 0.3
# цена одного ящика, которому не хватило места на полу и который
# встанет на клинкеты. Это главный риск по сохранности, поэтому дорого
PLACE_COST = 10.0
# цена второго яруса клинкетов: задвижка на задвижку. Самое дорогое,
# чего мы избегаем — нижний ряд тогда держит и второй ряд, и всю мелочь
LAYER_COST = 300.0
# ящик на поддоне 1200 x 1000: поперек встают два (2,0 м при ширине
# 2,352), вдоль каждый ряд занимает 1,2 м
CRATE_L, CRATE_W = gg.CRATE_L, gg.CRATE_W

# цена одного дробления позиции между контейнерами. Дробить удобно
# расчету и неудобно бригаде: на каждую такую позицию надо считать
# штуки на площадке. Берем ее равной перекосу в 50 кг: ради меньшего
# выигрыша не дробим
SPLIT_COST = 1.0

NAMES = {1: "Контейнер 1", 2: "Контейнер 2"}


def head(ws, row, ncols, height=40):
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


def key_of(r):
    return (r["Наименование"], r["Откуда"], r["Блок"])


def scaled(r, qty):
    """Копия позиции на другое количество штук."""
    c = dict(r)
    c["Штук"] = qty
    c["вес"] = qty * r["вес_ед"]
    c["объем"] = qty * r["объем_ед"]
    c["Сумма"] = qty * r["цена_ед"]
    return c


def pool_all():
    """Весь перечень на отгрузку, без лимитов контейнера."""
    block_a, block_b, rest, _, _ = collect(10 ** 9, 10 ** 9)
    rows = block_a + block_b + rest
    for r in rows:
        qty = r["Штук"] or 0
        r["объем_ед"] = r["объем"] / qty if qty else 0
        r["цена_ед"] = (r["Сумма"] / qty) if qty and r["Сумма"] else 0
    return [r for r in rows if (r["Штук"] or 0) > 0]


def split():
    """Разложить перечень по двум контейнерам поровну.

    Делим не общий вес, а каждый ярус отдельно. Если поделить только
    общий вес, можно получить контейнер, забитый одними задвижками, и
    второй с одними ящиками: в первом нижний ряд передавит, во втором
    нечем держать груз на полу. Поярусная развеска дает одинаковый
    пирог в обоих: настил, задвижки, доска, средняя арматура, доска,
    ящики, сверху фильтры и титан.
    """
    pool = pool_all()
    boxes = {n: {} for n in range(1, BOXES + 1)}
    load = {n: [0.0, 0.0] for n in range(1, BOXES + 1)}
    tkg = {n: defaultdict(float) for n in range(1, BOXES + 1)}
    # клинкеты нижнего яруса по диаметрам и занятая ими длина пола
    gcount = {n: defaultdict(int) for n in range(1, BOXES + 1)}
    glen = {n: 0.0 for n in range(1, BOXES + 1)}
    rest = []

    def rows_for(c, du):
        """Сколько рядов занимают c задвижек этого диаметра."""
        a = gg.across(du)
        return -(-c // a) if c > 0 else 0

    def add_gate(n, du, qty):
        """Пересчитать длину пола под клинкетами при изменении count."""
        c = gcount[n][du]
        glen[n] += (rows_for(c + qty, du) - rows_for(c, du)) * gg.depth(du)
        gcount[n][du] = c + qty

    def places(n, d_len=0.0):
        """Мест под ящики на свободном полу, в два яруса."""
        free = FLOOR_L - (glen[n] + d_len)
        return int(free / CRATE_L) * 2 * 2 if free > 0 else 0

    def need(kg):
        """Сколько ящиков нужно под мелочь такого веса."""
        n = int(kg / BOX_LIMIT_KG)
        return max(1, n + (1 if kg % BOX_LIMIT_KG else 0))

    def put(n, r, qty):
        k = key_of(r)
        cur = boxes[n].get(k)
        cur = scaled(r, (cur["Штук"] if cur else 0) + qty)
        cur["Контейнер"] = n
        boxes[n][k] = cur
        load[n][0] += qty * r["вес_ед"]
        load[n][1] += qty * r["объем_ед"]
        t = tier(asrow(r))
        tkg[n][t] += qty * r["вес_ед"]
        if t == 1 and item_kind(asrow(r)) == "клинкет (задвижка)":
            add_gate(n, int(r["Ду"] or 0), qty)

    def take(n, r, qty):
        k = key_of(r)
        cur = boxes[n][k]
        left = cur["Штук"] - qty
        if left > 0:
            boxes[n][k] = scaled(cur, left)
            boxes[n][k]["Контейнер"] = n
        else:
            del boxes[n][k]
        load[n][0] -= qty * r["вес_ед"]
        load[n][1] -= qty * r["объем_ед"]
        t = tier(asrow(r))
        tkg[n][t] -= qty * r["вес_ед"]
        if t == 1 and item_kind(asrow(r)) == "клинкет (задвижка)":
            add_gate(n, int(r["Ду"] or 0), -qty)

    # первый проход: тяжелое вперед, каждую позицию в тот контейнер,
    # где этого яруса пока меньше
    for r in sorted(pool, key=lambda x: (tier(asrow(x)), -x["вес"])):
        t = tier(asrow(r))
        order = sorted(range(1, BOXES + 1),
                       key=lambda n: (tkg[n][t], load[n][0], n))
        for n in order:
            if (load[n][0] + r["вес"] <= BOX_KG
                    and load[n][1] + r["объем"] <= BOX_M3):
                put(n, r, r["Штук"])
                break
        else:
            r["Контейнер"] = 0
            rest.append(r)

    # второй проход: дробим позиции по штукам и выравниваем сразу вес
    # каждого яруса и общий объем. Один только вес выровнять мало:
    # можно получить контейнер, набитый мелочью в ящиках, и второй с
    # голыми задвижками — вес тот же, а места разные
    def splits():
        return sum(1 for k in boxes[1] if k in boxes[2])

    def risk(n, d_len=0.0, d_kg3=0.0):
        """Цена риска по нижнему ярусу: второй ярус задвижек и ящики,
        которым не хватило пола и которые встанут на клинкеты."""
        c = 0.0
        if glen[n] + d_len > FLOOR_L:
            c += LAYER_COST
        c += PLACE_COST * max(0, need(tkg[n][3] + d_kg3) - places(n, d_len))
        return c

    def cost():
        c = sum(((tkg[1][t] - tkg[2][t]) / TOL) ** 2 for t in TIERS)
        c += ((load[1][1] - load[2][1]) / TOL_M3) ** 2
        c += sum(risk(n) for n in boxes)
        return c + SPLIT_COST * splits()

    def after(r, qty, a, b):
        """Цена расклада, если перенести qty штук из a в b."""
        t = tier(asrow(r))
        dw, dv = qty * r["вес_ед"], qty * r["объем_ед"]
        sign = 1 if a == 1 else -1
        c = 0.0
        for tt in TIERS:
            g = tkg[1][tt] - tkg[2][tt]
            if tt == t:
                g -= sign * 2 * dw
            c += (g / TOL) ** 2
        gv = (load[1][1] - load[2][1]) - sign * 2 * dv
        c += (gv / TOL_M3) ** 2
        # нижний ярус после переноса
        gate = t == 1 and item_kind(asrow(r)) == "клинкет (задвижка)"
        du = int(r["Ду"] or 0)
        if gate:
            ca, cb = gcount[a][du], gcount[b][du]
            dla = (rows_for(ca - qty, du) - rows_for(ca, du)) * gg.depth(du)
            dlb = (rows_for(cb + qty, du) - rows_for(cb, du)) * gg.depth(du)
        else:
            dla = dlb = 0.0
        dk = qty * r["вес_ед"] if t == 3 else 0.0
        c += risk(a, dla, -dk) + risk(b, dlb, dk)
        # дробление платное: переносим часть штук — позиция окажется
        # в обоих контейнерах, переносим все — перестанет
        k = key_of(r)
        n_split = splits()
        if qty < r["Штук"] and k not in boxes[b]:
            n_split += 1
        elif qty == r["Штук"] and k in boxes[b]:
            n_split -= 1
        return c + SPLIT_COST * n_split

    moves = []
    for _ in range(2000):
        now = cost()
        best, best_cost = None, now
        for a in range(1, BOXES + 1):
            b = 2 if a == 1 else 1
            for r in list(boxes[a].values()):
                if not r["вес_ед"] and not r["объем_ед"]:
                    continue
                qty_set = {1, r["Штук"], max(1, r["Штук"] // 2)}
                if r["вес_ед"]:
                    ideal = abs(tkg[1][tier(asrow(r))]
                                - tkg[2][tier(asrow(r))]) / 2 / r["вес_ед"]
                    qty_set.add(max(1, min(r["Штук"], round(ideal))))
                for qty in qty_set:
                    if (load[b][0] + qty * r["вес_ед"] > BOX_KG
                            or load[b][1] + qty * r["объем_ед"] > BOX_M3):
                        continue
                    c = after(r, qty, a, b)
                    if c < best_cost - 1e-9:
                        best, best_cost = (r, qty, a, b), c
        if not best:
            break
        r, qty, a, b = best
        take(a, r, qty)
        put(b, r, qty)
        moves.append((r["Наименование"], qty, r["Штук"], a, b))

    # в отчет идут только настоящие дробления: позиция, которая целиком
    # переехала в другой контейнер, просто там и лежит
    split_moves = [m for m in moves if m[2] > m[1]]

    out = {n: list(boxes[n].values()) for n in boxes}
    return out, load, tkg, rest, split_moves


def board_kg(kg):
    return kg / 1000 * BOARD_PER_T * BOARD_RHO


def loading_cost(same_day=True):
    """Погрузка двух контейнеров: подача погрузчика, часы, бригада."""
    calls = 1 if same_day else BOXES
    hours = HOURS_BOX * BOXES
    return (FORKLIFT_CALL * calls + FORKLIFT_HOUR * hours
            + LOADERS * LOADER_HOUR * hours)


def crates(kg_tier3):
    """Сколько ящиков с мелочью выходит на контейнер."""
    return max(1, int(kg_tier3 / BOX_LIMIT_KG) + (1 if kg_tier3 % BOX_LIMIT_KG else 0))


def sheet_summary(wb, boxes, load, tkg, rest, moves):
    ws = wb.create_sheet("Сводка", 0)
    ws["A1"] = "Два двадцатифутовых в Петербург, развеска поровну"
    ws["A1"].font = TITLE
    ws["A2"] = (f"По {BOX_KG // 1000} т и {BOX_M3:.0f} куб. м на контейнер. "
                "Вес и объем делятся пополам, и делятся поярусно: в каждом "
                "контейнере свой полный пирог, а не отдельно задвижки и "
                "отдельно ящики.")
    ws["A2"].font = NOTE

    row = 4
    ws.cell(row=row, column=1, value="Что в каком контейнере").font = SUB
    row += 1
    cols = ["Контейнер", "Позиций", "Штук", "Вес груза, т", "Доска, кг",
            "Всего в контейнере, т", "Загрузка по весу", "Объем, куб. м",
            "Загрузка по объему"]
    for i, h in enumerate(cols, start=1):
        ws.cell(row=row, column=i, value=h)
    head(ws, row, len(cols), height=32)

    for n in sorted(boxes):
        row += 1
        kg, m3 = load[n]
        items = boxes[n]
        total = kg + board_kg(kg)
        values = [NAMES[n], len(items), sum(r["Штук"] for r in items),
                  kg / 1000, round(board_kg(kg)), total / 1000,
                  total / BOX_KG, round(m3, 1), m3 / BOX_M3]
        for i, v in enumerate(values, start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            cell.fill = ONE_FILL if n == 1 else TWO_FILL
            if i in (4, 6):
                cell.number_format = "#,##0.00"
            if i in (7, 9):
                cell.number_format = "0%"
            if i in (5, 8):
                cell.number_format = "#,##0"

    row += 1
    kg_all = sum(load[n][0] for n in boxes)
    values = ["ИТОГО", sum(len(boxes[n]) for n in boxes),
              sum(r["Штук"] for n in boxes for r in boxes[n]),
              kg_all / 1000, round(board_kg(kg_all)),
              (kg_all + board_kg(kg_all)) / 1000, None,
              round(sum(load[n][1] for n in boxes), 1), None]
    for i, v in enumerate(values, start=1):
        cell = ws.cell(row=row, column=i, value=v)
        cell.font = BOLD
        cell.border = BORDER
        if i in (4, 6):
            cell.number_format = "#,##0.00"
        if i in (5, 8):
            cell.number_format = "#,##0"

    if rest:
        row += 1
        for i, v in enumerate(["Не влезло", len(rest),
                               sum(r["Штук"] for r in rest),
                               sum(r["вес"] for r in rest) / 1000], start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BOLD
            cell.fill = WARN_FILL
            cell.border = BORDER

    row += 2
    ws.cell(row=row, column=1, value="Насколько ровно вышло").font = SUB
    row += 1
    kg1, kg2 = load[1][0], load[2][0]
    m31, m32 = load[1][1], load[2][1]
    for label, a, b, unit in [("Вес груза", kg1, kg2, "кг"),
                              ("Объем", m31, m32, "куб. м")]:
        diff = abs(a - b)
        share = diff / max(a, b) if max(a, b) else 0
        cell = ws.cell(row=row, column=1,
                       value=f"— {label}: {a:,.1f} и {b:,.1f} {unit}, "
                             f"разница {diff:,.1f} {unit}, это "
                             f"{share * 100:.1f}%.".replace(",", " "))
        cell.font = BASE
        row += 1
    for t in sorted(TIERS):
        d = abs(tkg[1][t] - tkg[2][t])
        cell = ws.cell(row=row, column=1,
                       value=f"— {TIERS[t][0]}: {tkg[1][t]:,.0f} и "
                             f"{tkg[2][t]:,.0f} кг, перекос {d:,.0f} кг."
                       .replace(",", " "))
        cell.font = BASE
        row += 1
    if moves:
        cell = ws.cell(row=row, column=1,
                       value=f"— Чтобы выйти на такую развеску, по "
                             f"{len(moves)} "
                             f"{plural(len(moves), 'позиции', 'позициям', 'позициям')}"
                             " штуки поделены между контейнерами. Список — "
                             "на вкладке «Поделенные позиции».")
        cell.font = BASE
        row += 1

    row += 1
    ws.cell(row=row, column=1, value="Погрузка во Владивостоке").font = SUB
    row += 1
    for label, value in [
        ("Подача погрузчика, за выезд", FORKLIFT_CALL),
        ("Погрузчик, руб за час", FORKLIFT_HOUR),
        (f"Грузчики, {LOADERS} человека, руб за час с человека",
         LOADER_HOUR),
        ("Часов на контейнер", HOURS_BOX),
        ("Погрузка обоих в один день, руб", loading_cost(True)),
        ("Погрузка разными днями, руб", loading_cost(False)),
    ]:
        ws.cell(row=row, column=1, value=label).font = BASE
        cell = ws.cell(row=row, column=2, value=value)
        cell.font = BOLD
        cell.border = BORDER
        cell.number_format = "#,##0"
        row += 1
    cell = ws.cell(row=row, column=1,
                   value="Подача погрузчика платится за выезд, а не за "
                         "контейнер: грузим оба в один день — платим один "
                         "раз, это 12 тысяч разницы. Выгрузка в Петербурге "
                         "сюда не входит, ставок на нее пока нет.")
    cell.font = NOTE
    row += 2

    ws.cell(row=row, column=1, value="Что это дает").font = SUB
    row += 1
    for text in [
        f"Весь перечень — {kg_all / 1000:.2f} т. Поровну это по "
        f"{kg_all / 2000:.2f} т груза плюс {board_kg(kg_all / 2):.0f} кг "
        f"доски на контейнер, при лимите {BOX_KG // 1000} т. Запас больше "
        "трети: по весу мы никуда не упираемся, и это главное для "
        "сохранности.",
        "Сорокафутовый брал 25,35 т в один корпус и оставлял 1,94 т во "
        "Владивостоке. Там пирог вдвое выше и вдвое тяжелее давит на "
        "нижний ряд задвижек.",
        "Объем занят примерно наполовину. Свободная высота — не потеря, "
        "а запас под прокладки, брус и распорки: пустоты в контейнере "
        "опаснее перегруза.",
        "Каждый контейнер получает половину каждого яруса, поэтому схема "
        "укладки у них одинаковая и бригаде не надо объяснять два разных "
        "порядка.",
        "Контейнеры можно отправить в разные дни, но тогда погрузчик "
        "подается дважды. Один день погрузки дешевле на 12 тысяч.",
    ]:
        cell = ws.cell(row=row, column=1, value="— " + text)
        cell.font = BASE
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[row].height = 32
        row += 1

    widths(ws, [36, 12, 10, 14, 11, 18, 14, 12, 16])
    return ws


def sheet_balance(wb, boxes, load, tkg):
    ws = wb.create_sheet("Развеска по ярусам")
    ws["A1"] = "Развеска по ярусам: чтобы ничего не подавило"
    ws["A1"].font = TITLE
    ws["A2"] = ("Ярусы 1 и 2 лежат на полу, ящики третьего яруса стоят "
                "рядом и в два уровня, фильтры и титан — сверху на ящиках. "
                "Давит только то, что лежит выше.")
    ws["A2"].font = NOTE

    row = 4
    cols = ["Ярус", "Что лежит", "Контейнер 1, кг", "Контейнер 2, кг",
            "Перекос, кг", "Объем К1, куб. м", "Объем К2, куб. м",
            "Высота слоя по полу, м", "Нагрузка на пол, кг/кв. м"]
    for i, h in enumerate(cols, start=1):
        ws.cell(row=row, column=i, value=h)
    head(ws, row, len(cols), height=40)

    vol = {n: defaultdict(float) for n in boxes}
    for n in boxes:
        for r in boxes[n]:
            vol[n][tier(asrow(r))] += r["объем"]

    for t in sorted(TIERS):
        row += 1
        a, b = tkg[1][t], tkg[2][t]
        v1, v2 = vol[1][t], vol[2][t]
        height = max(v1, v2) / FLOOR_AREA
        on_floor = t in (1, 2)
        values = [TIERS[t][0], TIERS[t][1], round(a), round(b),
                  round(abs(a - b)), round(v1, 1), round(v2, 1),
                  round(height, 2),
                  round(max(a, b) / FLOOR_AREA) if on_floor else None]
        for i, v in enumerate(values, start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=(i == 2))
            if i in (3, 4, 5, 9):
                cell.number_format = "#,##0"
            if i in (6, 7):
                cell.number_format = "#,##0.0"
            if i == 8:
                cell.number_format = "#,##0.00"
            if i == 5 and abs(a - b) > TOL:
                cell.fill = WARN_FILL

    row += 2
    ws.cell(row=row, column=1, value="Что на что давит").font = SUB
    row += 1
    cols2 = ["Контейнер", "Лежит на полу, кг", "В ящиках, кг",
             "Ящиков в контейнере", "Стоит на ящиках, кг",
             "На один ящик сверху, кг", "Нагрузка на пол, кг/кв. м",
             "Проверка"]
    for i, h in enumerate(cols2, start=1):
        ws.cell(row=row, column=i, value=h)
    head(ws, row, len(cols2), height=32)

    for n in sorted(boxes):
        row += 1
        floor_kg = tkg[n][1] + tkg[n][2]
        box_kg = tkg[n][3]
        top_kg = tkg[n][4]
        nbox = crates(box_kg)
        per_crate = top_kg / nbox
        press = (floor_kg + box_kg + top_kg) / FLOOR_AREA
        ok = ("норма" if per_crate <= BOX_LIMIT_KG
              else "перегруз по верхнему ярусу")
        values = [NAMES[n], round(floor_kg), round(box_kg), nbox,
                  round(top_kg), round(per_crate), round(press), ok]
        for i, v in enumerate(values, start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            cell.fill = ONE_FILL if n == 1 else TWO_FILL
            if i in (2, 3, 5, 6, 7):
                cell.number_format = "#,##0"
            if i == 8 and ok != "норма":
                cell.fill = WARN_FILL

    row += 2
    ws.cell(row=row, column=1, value="Почему так сохраннее").font = SUB
    row += 1
    for text in [
        "Ни один ярус арматуры не стоит на другом ярусе арматуры: "
        "задвижки и тяжелое лежат на полу в один слой, ящики стоят "
        "рядом, а не поверх задвижек. На ящики идут только фильтры и "
        "титан — самое легкое, что есть.",
        "Задвижки кладем плашмя на доску, фланцем вниз, и никогда "
        "фланцем на фланец: привалочная поверхность и шпилечные "
        "отверстия страдают первыми, а восстановить их нельзя.",
        "Между рядами задвижек Ду250-350 шаг прокладок 300-400 мм, а не "
        "600: длинный корпус на двух опорах прогибается и дает трещину "
        "по приливу.",
        "Маховики и штоки нагрузку не несут. Снимающиеся маховики "
        "снимаем и кладем в ящик с биркой по номеру места.",
        "Все на поддонах и в обрешетке притянуто ремнями к рымам: месяц "
        "дороги — это продольные удары при маневрировании, а не "
        "равномерная езда.",
        "Пустоты забиваем распорками. Свободный объем в этом плане "
        "есть, и это плюс: есть куда поставить брус.",
        "Влагопоглотитель под крышу: конденсат за месяц дает налет на "
        "бронзе и ржавчину на стальных шпильках.",
        "Погрузчиком берем только поддоны и обрешетку. Голую задвижку "
        "вилами не поднимаем: мятый фланец — это минус позиция.",
    ]:
        cell = ws.cell(row=row, column=1, value="— " + text)
        cell.font = BASE
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[row].height = 32
        row += 1

    widths(ws, [24, 44, 15, 15, 13, 15, 15, 20, 20])
    return ws


# бронза и латунь: предел прочности на сжатие, МПа. Берем нижнюю
# границу для литейных бронз
BRONZE_MPA = 200
# сырая необрезная доска 25 мм: расчетное сопротивление изгибу с учетом
# сучков и обзола, МПа. Для сорта 2 по нормам 13, берем с запасом
BOARD_MPA = 10


def sheet_pressure(wb, boxes, tkg):
    """Выдержат ли нижние клинкеты то, что лежит сверху."""
    ws = wb.create_sheet("Нагрузка на клинкеты")
    ws["A1"] = "Нагрузка на нижний ярус: продавит или нет"
    ws["A1"].font = TITLE
    ws["A2"] = ("Считаем не абстрактные тонны, а то, что реально ложится "
                "поверх задвижек нижнего яруса, и сравниваем с тем, что "
                "держат бронза и доска.")
    ws["A2"].font = NOTE

    row = 4
    cols = ["Контейнер", "Клинкеты нижнего яруса, кг", "Рядов",
            "Длина пола под ними, м", "Ярусов клинкетов",
            "Свободного пола, м", "Мест под ящики на полу", "Ящиков всего",
            "Ящиков встанет на клинкеты", "Вес поверх клинкетов, кг",
            "Нагрузка, кг/кв. м", "Запас по бронзе"]
    for i, h in enumerate(cols, start=1):
        ws.cell(row=row, column=i, value=h)
    head(ws, row, len(cols), height=46)

    data = []
    for n in sorted(boxes):
        counts = defaultdict(int)
        for r in boxes[n]:
            if (tier(asrow(r)) == 1
                    and item_kind(asrow(r)) == "клинкет (задвижка)"):
                counts[int(r["Ду"] or 0)] += int(r["Штук"])
        grows = gg.rows_of(counts)
        layers = gg.place(grows)
        gl = gg.gate_length(counts)
        free, over, h_gates, free_l = gg.crate_places(counts)
        area = min(gl, FLOOR_L) * FLOOR_W
        kg3 = tkg[n][3]
        nbox = max(1, int(kg3 / BOX_LIMIT_KG)
                   + (1 if kg3 % BOX_LIMIT_KG else 0))
        on_gates = max(0, nbox - free)
        # на клинкетах оказываются ящики, которым не хватило пола, плюс
        # верхний ярус фильтров и титана ложится на эти же ящики
        kg_top = on_gates * (kg3 / nbox) + tkg[n][4]
        press = kg_top / area if area else 0
        mpa = press * 9.81 / 1e6
        data.append((n, area, gl, free, nbox, on_gates, kg_top, press, mpa))
        row += 1
        values = [NAMES[n], round(tkg[n][1]), len(grows), round(gl, 2),
                  layers, round(free_l, 2), free, nbox, on_gates,
                  round(kg_top), round(press),
                  f"{BRONZE_MPA / mpa:,.0f} раз".replace(",", " ")
                  if mpa else ""]
        for i, v in enumerate(values, start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            cell.fill = ONE_FILL if n == 1 else TWO_FILL
            if i in (2, 3, 7, 8, 9, 10, 11):
                cell.number_format = "#,##0"
            if i in (4, 6):
                cell.number_format = "#,##0.00"
            if i == 5 and v > 1:
                cell.fill = WARN_FILL
            if i == 9 and v:
                cell.fill = WARN_FILL

    row += 2
    ws.cell(row=row, column=1,
            value="Сравнение с одним усиленным на 27 тонн").font = SUB
    row += 1
    cell = ws.cell(row=row, column=1,
                   value="В усиленном клинкеты не ложатся в один ярус: они "
                         "просят 10,70 м пола при 5,90. Значит задвижка "
                         "ляжет на задвижку через доску, и нижний ряд будет "
                         "держать и второй ряд задвижек, и всю мелочь "
                         "сверху. В двух контейнерах клинкеты в один ярус, "
                         "друг на друга не кладутся вообще.")
    cell.font = NOTE
    cell.alignment = Alignment(wrap_text=True)
    row += 1
    for i, h in enumerate(["Вариант", "Ящиков на клинкетах",
                           "Вес поверх клинкетов, кг",
                           "Нагрузка, кг/кв. м", "Во сколько раз больше"],
                          start=1):
        ws.cell(row=row, column=i, value=h)
    head(ws, row, 5, height=32)
    worst = max(d[7] for d in data)
    for label, nbox_g, kg_top, press in [
        ("Два двадцатифутовых, худший из двух",
         max(d[5] for d in data), max(d[6] for d in data), worst),
        ("Один усиленный, 27 т, клинкеты в два яруса", 8, 13_441, 1_067),
    ]:
        row += 1
        values = [label, nbox_g, round(kg_top), round(press),
                  round(press / worst, 1) if worst else ""]
        for i, v in enumerate(values, start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BOLD if "усиленный" in label else BASE
            cell.border = BORDER
            if i in (3, 4):
                cell.number_format = "#,##0"
            if i == 5:
                cell.number_format = "#,##0.0"
            if "усиленный" in label:
                cell.fill = WARN_FILL

    row += 2
    ws.cell(row=row, column=1,
            value="Что на самом деле ломается").font = SUB
    row += 1
    for text in [
        f"Сама бронза не продавится. Литейная бронза держит около "
        f"{BRONZE_MPA} МПа на сжатие, а у нас на нижнем ярусе "
        f"{worst:,.0f} кг на квадратный метр, то есть "
        f"{worst * 9.81 / 1e6:.3f} МПа. Запас в тысячи раз. Корпус "
        "задвижки — это литье с толстой стенкой, оно рассчитано на "
        "рабочее давление в десятки атмосфер изнутри."
        .replace(",", " "),
        "Ломается не корпус, а шток и маховик. Если ящик встанет углом "
        "на маховик или на шток, его поведет, и задвижка станет "
        "некондицией при целом корпусе. Поэтому кладем задвижки так, "
        "чтобы шток лежал горизонтально и не выступал выше корпуса, а "
        "съемные маховики снимаем и в ящик с биркой.",
        "Второе, что страдает, — привалочная поверхность фланца. Металл "
        "по металлу дает забоины, которые потом не выводятся. Между "
        "рядами всегда доска, фланец на фланец не кладем.",
        f"Третье — сама доска. Ящик 600 кг стоит на четырех углах, это "
        f"150 кг на угол. Доска 200 х 25 на пролете 600 мм дает "
        f"{10.6:.1f} МПа изгиба при расчетных {BOARD_MPA} для сырой "
        f"необрезной с сучками. На пролете 400 мм — {7.1:.1f} МПа, на "
        f"300 мм — {5.3:.1f}. Поэтому под ящиками прокладки кладем чаще "
        "и строго над корпусами задвижек, а не над промежутками.",
        "Отсюда правило погрузки: ящик должен опираться на доску ровно "
        "там, где под доской лежит корпус задвижки. Если под углом "
        "ящика пустота — туда брус 50 х 50 или еще одна доска поперек.",
    ]:
        cell = ws.cell(row=row, column=1, value="— " + text)
        cell.font = BASE
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[row].height = 46
        row += 1

    widths(ws, [30, 20, 10, 18, 16, 16, 18, 12, 20, 20, 16, 16])
    return ws


def sheet_moves(wb, moves):
    ws = wb.create_sheet("Поделенные позиции")
    ws["A1"] = "Позиции, поделенные между контейнерами"
    ws["A1"].font = TITLE
    ws["A2"] = ("Без дробления по штукам ровной развески не выходит: одна "
                "партия задвижек весит больше тонны и перекашивает "
                "контейнер целиком. При погрузке считаем штуки по этой "
                "вкладке.")
    ws["A2"].font = NOTE
    for i, h in enumerate(["№", "Наименование", "Штук переносим",
                           "Откуда", "Куда"], start=1):
        ws.cell(row=4, column=i, value=h)
    head(ws, 4, 5, height=24)
    for k, (name, qty, _left, hi, lo) in enumerate(moves, start=1):
        for i, v in enumerate([k, name, qty, NAMES[hi], NAMES[lo]], start=1):
            cell = ws.cell(row=4 + k, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=(i == 2))
    widths(ws, [5, 60, 16, 16, 16])
    return ws


def sheet_box(wb, n, items, load, tkg):
    ws = wb.create_sheet(NAMES[n])
    kg, m3 = load
    ws["A1"] = f"{NAMES[n]}: что грузим"
    ws["A1"].font = TITLE
    ws["A2"] = (f"{len(items)} {plural(len(items), 'позиция', 'позиции', 'позиций')}"
                f", {sum(r['Штук'] for r in items):.0f} шт, "
                f"{kg / 1000:.2f} т груза плюс {board_kg(kg):.0f} кг доски "
                f"при лимите {BOX_KG // 1000} т, {m3:.1f} куб. м из "
                f"{BOX_M3:.0f}. Порядок — по группам, от мелкой арматуры "
                "к крупной.")
    ws["A2"].font = NOTE

    start = 4
    for i, (h, _) in enumerate(COLS, start=1):
        ws.cell(row=start, column=i, value=h)
    head(ws, start, len(COLS))

    by_group = defaultdict(list)
    for r in items:
        by_group[group_of(r["Наименование"])].append(r)
    order = [g for g in GROUPS if g in by_group]
    order += [g for g in sorted(by_group) if g not in GROUPS]

    row = start
    k = 0
    for group in order:
        rows_in = sorted(by_group[group],
                         key=lambda r: (r["Ду"] or 9999, r["Наименование"]))
        row += 1
        qty = sum(r["Штук"] for r in rows_in)
        kg_group = sum(r["вес"] for r in rows_in)
        kg_text = f"{kg_group:,.0f}".replace(",", " ")
        ws.cell(row=row, column=1,
                value=f"{group} — {len(rows_in)} "
                      f"{plural(len(rows_in), 'позиция', 'позиции', 'позиций')}"
                      f", {qty:.0f} шт, {kg_text} кг")
        for c in range(1, len(COLS) + 1):
            cell = ws.cell(row=row, column=c)
            cell.fill = GROUP_FILL
            cell.font = SUB
            cell.border = BORDER
        for r in rows_in:
            row += 1
            k += 1
            fake = asrow(r)
            values = [k, r["Наименование"], r["Чертеж"], r["Ду"], r["Ру"],
                      r["Состояние"], r["Штук"], round(r["вес_ед"], 1),
                      round(r["вес"]), round(r["объем"], 3),
                      r["Выделено"], r["Блок"], TIERS[tier(fake)][0],
                      pack_type(fake), r["Откуда"], r["Цена"],
                      r["Ориентир"]]
            for i, v in enumerate(values, start=1):
                cell = ws.cell(row=row, column=i, value=v)
                cell.font = BASE
                cell.border = BORDER
                cell.alignment = Alignment(vertical="top",
                                           wrap_text=(i in (2, 12, 14)))
                if i == 7:
                    cell.font = BOLD
                if i == 8:
                    cell.number_format = "#,##0.0"
                if i == 9:
                    cell.number_format = "#,##0"
                if i == 10:
                    cell.number_format = "#,##0.000"
            if r["Выделено"]:
                ws.cell(row=row, column=11).fill = MARK_FILL
            if not r["Чертеж"]:
                ws.cell(row=row, column=3).fill = NODRAW_FILL

    last = row
    row += 1
    ws.cell(row=row, column=2, value="ИТОГО ПО КОНТЕЙНЕРУ").font = BOLD
    for col, fmt in ((7, "#,##0"), (9, "#,##0"), (10, "#,##0.0")):
        cell = ws.cell(row=row, column=col,
                       value=sum(r["Штук" if col == 7 else
                                   "вес" if col == 9 else "объем"]
                                 for r in items))
        cell.font = BOLD
        cell.border = BORDER
        cell.number_format = fmt

    row += 2
    ws.cell(row=row, column=1, value="Укладка в этом контейнере").font = SUB
    row += 1
    for i, h in enumerate(["Ярус", "Что кладем", "Чем перекладываем сверху",
                           "Вес яруса, кг"], start=1):
        ws.cell(row=row, column=i, value=h)
    head(ws, row, 4, height=24)
    row += 1
    for i, v in enumerate([FLOOR[0], FLOOR[1], FLOOR[2], "—"], start=1):
        cell = ws.cell(row=row, column=i, value=v)
        cell.font = BASE
        cell.border = BORDER
        cell.fill = GROUP_FILL
        cell.alignment = Alignment(vertical="top", wrap_text=(i in (2, 3)))
    for t in sorted(TIERS):
        row += 1
        label, what, board = TIERS[t]
        for i, v in enumerate([label, what, board, round(tkg[t])], start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=(i in (2, 3)))
            if i == 4:
                cell.number_format = "#,##0"

    ws.freeze_panes = f"C{start + 1}"
    widths(ws, [w for _, w in COLS])
    return ws


def sheet_packing(wb):
    ws = wb.create_sheet("Упаковка и крепление")
    ws["A1"] = "Как паковать, чтобы доехало без помятых фланцев"
    ws["A1"].font = TITLE
    ws["A2"] = ("Правила одинаковы для обоих контейнеров. Дорога занимает "
                "около месяца, с продольными ударами при маневрировании.")
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
    widths(ws, [5, 42, 86])
    return ws


def main():
    boxes, load, tkg, rest, moves = split()

    wb = Workbook()
    wb.remove(wb.active)
    sheet_summary(wb, boxes, load, tkg, rest, moves)
    sheet_balance(wb, boxes, load, tkg)
    sheet_pressure(wb, boxes, tkg)
    for n in sorted(boxes):
        sheet_box(wb, n, boxes[n], load[n], tkg[n])
    if moves:
        sheet_moves(wb, moves)
    sheet_packing(wb)

    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / "План_2_контейнера.xlsx"
    wb.save(path)

    print(f"готово: {path}\n")
    for n in sorted(boxes):
        kg, m3 = load[n]
        items = boxes[n]
        print(f"{NAMES[n]}: {len(items):>3} поз, "
              f"{sum(r['Штук'] for r in items):>5.0f} шт, "
              f"{kg / 1000:>5.2f} т груза + {board_kg(kg):>3.0f} кг доски "
              f"({(kg + board_kg(kg)) / BOX_KG * 100:>3.0f}% лимита), "
              f"{m3:>4.1f} куб.м ({m3 / BOX_M3 * 100:>3.0f}%)")
    print()
    for t in sorted(TIERS):
        print(f"{TIERS[t][0]:<22}{tkg[1][t]:>9,.0f} кг{tkg[2][t]:>10,.0f} кг"
              f"   перекос {abs(tkg[1][t] - tkg[2][t]):>6,.0f} кг"
              .replace(",", " "))
    print(f"\nподелено по штукам позиций: {len(moves)}")
    if rest:
        print(f"не влезло: {len(rest)} поз, "
              f"{sum(r['вес'] for r in rest) / 1000:.2f} т")
    else:
        print("не влезло: ничего, весь перечень уезжает")
    print(f"погрузка в один день: {loading_cost(True):,.0f} руб, "
          f"разными днями: {loading_cost(False):,.0f} руб"
          .replace(",", " "))


if __name__ == "__main__":
    main()
