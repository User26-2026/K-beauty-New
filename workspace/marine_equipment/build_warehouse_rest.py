#!/usr/bin/env python3
"""Что осталось на складе после отгрузок: по группам товара.

Менеджер прислала остатки за сентябрь 2026 по четырем локациям, но
отгруженное из них не сняла. Видно это сразу: «Клинкет Ду 150,
1 фланец» — 45 штук, а ровно 45 таких уехало в Керчь 24.09. Так же
стоят на месте головки, грибки, задвижки Ду80 и фильтры, которые уже
в контейнере на Петербург.

Поэтому файл делает две вещи: вычитает отгруженное и раскладывает
остаток по товарным группам, чтобы было видно, чем склад набит.

Вычитаем ВСЕ отгрузки, а не только контейнер:
  24.09.2026 — задвижки Ду150 в Керчь, 45 шт;
  24.09.2026 — фильтры Ду300 в Петербург, 2 шт;
  29.09.2026 — запчасти в Петербург, 182 шт;
  26.10.2026 — контейнер в Петербург, 222 позиции, 1 810 шт.

Как сводится отгруженное с остатком. Сначала по номеру чертежа: он
надежнее всего, и в остатках он местами проставлен. Потом по паре
«группа плюс диаметр»: в остатках одна строка «Клапан фланцевый
проходной Ду50 Ру 10/25» на 47 штук, а в отгрузке те же клапаны
разбиты на семь строк по чертежам. Складывать их приходится по
диаметру, иначе не сойдется.

Где отгружено больше, чем числится, остаток уходит в ноль, а разница
попадает на вкладку «Не сошлось». Это не ошибка расчета: это значит,
что учет и отгрузка расходятся, и цифру надо проверять глазами.

Результат: outputs/marine_equipment/Остаток_склада_по_группам.xlsx
"""

import pathlib
import sys
from collections import defaultdict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
import container_list  # noqa: E402
import stock_groups as sg  # noqa: E402
import stock_sheets  # noqa: E402
import upd_shipments  # noqa: E402
from build_container_spb import plural  # noqa: E402
from build_contract_annexes import norm_draw  # noqa: E402

OUT = pathlib.Path(__file__).resolve().parents[2] / "outputs" / "marine_equipment"

# Все локации склада по нашему учету и то, прислала ли менеджер по ним
# остатки за сентябрь 2026. Это главное ограничение расчета: из
# одиннадцати файлов пришло четыре, и вся штуцерная мелочь, которой
# набивали контейнер, лежит в тех, что не пришли. Поэтому отгруженное
# по ним вычесть не из чего.
LOCATIONS = [
    ("Склад 1А, 1Б", "sklad_1a_1b.xlsx", True),
    ("Фойе", "foye.xls", True),
    ("Цех", "в файле фойе", True),
    ("Гараж 2", "garazh_2.xlsx", True),
    ("Контейнер 2", "konteiner_2.xls", True),
    ("Гараж 1", "garazh_1.xls", False),
    ("Контейнер 1", "konteiner_1.xls", False),
    ("Менеджерская", "menedzherskaya.xls", False),
    ("Склад 2", "sklad_2.xls", False),
    ("Склад 3", "sklad_3.xls", False),
    ("Склад электрики", "sklad_elektriki.xls", False),
    ("ЗИП насосы", "zip_nasosy.csv", False),
]

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
        cell.alignment = Alignment(vertical="top", wrap_text=(i in wrap),
                                   horizontal="center" if i in center else None)
        if i in money:
            cell.number_format = "#,##0"
        if fill:
            cell.fill = fill
    return row + 1


def widths(ws, values):
    for i, w in enumerate(values, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def shipped():
    """Все отгруженное: контейнер плюс три ранние партии."""
    out = []
    for r in container_list.read():
        out.append({
            "Откуда": "Контейнер 26.10.2026",
            "Наименование": r["Наименование"],
            "Чертеж": r["Чертеж"],
            "Ду": r["Ду"],
            "Штук": r["Штук"],
        })
    for s in upd_shipments.SHIPMENTS:
        for it in s["items"]:
            out.append({
                "Откуда": f"{s['title']}, {s['shipped']}",
                "Наименование": it["name"],
                "Чертеж": it.get("draw", ""),
                "Ду": sg.du_of(it["name"]),
                "Штук": it["qty"],
            })
    for r in out:
        r["Группа"], r["Подгруппа"] = sg.classify(r["Наименование"])
    return out


def match(stock, ship):
    """Вычитает отгруженное из остатка. Возвращает строки остатка с
    полями «Снято» и «Остаток», и то, что снять не удалось."""
    by_draw = defaultdict(int)
    by_type = defaultdict(int)
    for s in ship:
        for d in [x.strip() for x in str(s["Чертеж"]).split(",") if x.strip()]:
            by_draw[norm_draw(d)] += s["Штук"]
        by_type[(s["Подгруппа"], s["Ду"])] += s["Штук"]

    rows = []
    for r in stock:
        g, sub = sg.classify(r["Наименование"], r["Подгруппа"])
        du = sg.du_of(r["Наименование"]) if g == sg.ARMATURE else 0
        left = r["Штук"]
        took, how = 0, []

        if sg.broken(r["Наименование"]):
            # брак не отгружали: в перечнях отгрузок такого слова нет
            how.append("брак, не отгружался")
            price = r["Цена"] or (r["Сумма"] / r["Штук"] if r["Штук"] else 0)
            rows.append(dict(r, Группа=g, Подгруппа2=sub, Ду=du, Снято=0,
                             Остаток=left, Как=how[0], Цена1=price,
                             СуммаОстаток=round(price * left), Брак=True))
            continue

        for d in sg.draw_of(r["Наименование"]):
            key = norm_draw(d)
            if by_draw.get(key):
                n = min(left, by_draw[key])
                if n:
                    by_draw[key] -= n
                    by_type[(sub, du)] = max(0, by_type[(sub, du)] - n)
                    left -= n
                    took += n
                    how.append(f"по чертежу {d}: {n}")

        if left and by_type.get((sub, du)):
            n = min(left, by_type[(sub, du)])
            by_type[(sub, du)] -= n
            left -= n
            took += n
            how.append(f"по группе и диаметру ({sub}, Ду{du}): {n}")

        price = r["Цена"] or (r["Сумма"] / r["Штук"] if r["Штук"] else 0)
        rows.append(dict(r, Группа=g, Подгруппа2=sub, Ду=du, Снято=took,
                         Остаток=left, Как=("; ".join(how) or "не отгружалось"),
                         Цена1=price, СуммаОстаток=round(price * left),
                         Брак=sg.flawed(r["Наименование"])))

    unmatched = []
    for (sub, du), n in sorted(by_type.items(), key=lambda kv: -kv[1]):
        if n > 0:
            unmatched.append({"Подгруппа": sub, "Ду": du, "Штук": n})
    return rows, unmatched


def sheet_groups(wb, rows):
    ws = wb.active
    ws.title = "По группам"
    # Снятое считаем по ВСЕМ строкам группы, а позиции и сумму — только
    # по тем, где что-то осталось. Иначе в колонке «снято» будут нули
    # там, где позицию увезли целиком, а это как раз самое интересное.
    agg = defaultdict(lambda: defaultdict(lambda: [0, 0, 0, 0.0]))
    for r in rows:
        a = agg[r["Группа"]][r["Подгруппа2"]]
        a[2] += r["Снято"]
        if r["Остаток"] <= 0:
            continue
        a[0] += 1
        a[1] += r["Остаток"]
        a[3] += r["СуммаОстаток"]

    total_sum = sum(v[3] for g in agg.values() for v in g.values())
    gone = sum(r["Снято"] for r in rows)

    ws["A1"] = "Что осталось на складе: по группам товара"
    ws["A1"].font = TITLE
    have = sum(1 for _, _, g in LOCATIONS if g)
    ws["A2"] = (f"Остатки менеджера за сентябрь 2026, из них вычтено "
                f"отгруженное — {rub(gone)} шт по четырем отгрузкам. "
                f"Присланы {have} локаций из {len(LOCATIONS)}, то есть "
                f"это НЕ весь склад: чего не хватает, видно на вкладке "
                f"«Источники». Сумма по ценам из тех же файлов.")
    ws["A2"].font = NOTE

    cols = [("Группа товара", 30), ("Подгруппа", 36), ("Позиций", 9),
            ("Осталось, шт", 12), ("Снято как отгруженное, шт", 13),
            ("Сумма остатка, руб", 16), ("Доля суммы", 10)]
    start = 4
    for i, (h, _) in enumerate(cols, start=1):
        ws.cell(row=start, column=i, value=h)
    head(ws, start, len(cols))
    row = start + 1
    for g in sorted(agg, key=lambda x: -sum(v[3] for v in agg[x].values())):
        subs = agg[g]
        tot = [sum(v[i] for v in subs.values()) for i in range(3)]
        tots = sum(v[3] for v in subs.values())
        row = put(ws, row, [
            g, "", tot[0], tot[1], tot[2], round(tots),
            f"{tots / total_sum:.1%}" if total_sum else "",
        ], center=(3, 4, 5, 7), money=(6,), font=BOLD, fill=GROUP_FILL)
        for sub, v in sorted(subs.items(), key=lambda kv: -kv[1][3]):
            row = put(ws, row, ["", sub, v[0], v[1], v[2], round(v[3]), ""],
                      wrap=(2,), center=(3, 4, 5), money=(6,))
    row = put(ws, row, [
        "ИТОГО", "", sum(v[0] for g in agg.values() for v in g.values()),
        sum(v[1] for g in agg.values() for v in g.values()),
        sum(v[2] for g in agg.values() for v in g.values()),
        round(total_sum), "100%",
    ], center=(3, 4, 5, 7), money=(6,), font=BOLD, fill=OK_FILL)

    row += 1
    for t in [
        "«Снято как отгруженное» — сколько штук из этой строки уже "
        "уехало. Если снято все, позиция в остаток не попала и здесь "
        "не показана.",
        "Сумма считается по цене из файлов менеджера. Это ее оценка, "
        "сделками она не проверена, и по части позиций цена стоит "
        "нулем.",
        "Брак из остатка НЕ вычитался: в перечнях отгрузок нет ни "
        "одной строки со словом «брак», состояние там только «новый» "
        "или «б/у». Некомплект и «в разборе» вычитались — такие "
        "позиции в контейнере есть.",
        "Брак, некомплект и разобранное в сумму входят по той же цене, "
        "хотя стоят дешевле. Они помечены на вкладке «Остаток».",
    ]:
        cell = ws.cell(row=row, column=1, value="— " + t)
        cell.font = BASE
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=row, start_column=1,
                       end_row=row, end_column=len(cols))
        ws.row_dimensions[row].height = 28
        row += 1

    ws.freeze_panes = f"A{start + 1}"
    widths(ws, [w for _, w in cols])
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1


COLS_REST = [
    ("№", 5), ("Группа", 24), ("Подгруппа", 28), ("Наименование", 56),
    ("Локация", 16), ("Ду", 6), ("Было, шт", 9), ("Снято, шт", 9),
    ("ОСТАЛОСЬ, шт", 11), ("Цена, руб", 11), ("Сумма остатка, руб", 13),
    ("Что сняли", 40),
]


def sheet_rest(wb, rows):
    ws = wb.create_sheet("Остаток")
    ws["A1"] = "Остаток построчно"
    ws["A1"].font = TITLE
    ws["A2"] = ("Оранжевым — строки, где отгрузили столько же или "
                "больше, чем числилось. Желтым — брак, некомплект и "
                "разобранное.")
    ws["A2"].font = NOTE
    start = 4
    for i, (h, _) in enumerate(COLS_REST, start=1):
        ws.cell(row=start, column=i, value=h)
    head(ws, start, len(COLS_REST))
    row = start + 1
    n = 0
    for r in sorted(rows, key=lambda x: (x["Группа"], x["Подгруппа2"],
                                         -x["СуммаОстаток"])):
        n += 1
        fill = (WARN_FILL if r["Остаток"] == 0 and r["Снято"]
                else MARK_FILL if r["Брак"] else None)
        row = put(ws, row, [
            n, r["Группа"], r["Подгруппа2"], r["Наименование"],
            r["Локация"], r["Ду"] or "", r["Штук"], r["Снято"] or "",
            r["Остаток"], round(r["Цена1"]), r["СуммаОстаток"], r["Как"],
        ], wrap=(4, 12), center=(1, 6, 7, 8, 9), money=(10, 11), fill=fill)
    ws.freeze_panes = f"A{start + 1}"
    ws.auto_filter.ref = f"A{start}:{get_column_letter(len(COLS_REST))}{row - 1}"
    widths(ws, [w for _, w in COLS_REST])
    ws.print_title_rows = f"{start}:{start}"
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1


def sheet_unmatched(wb, unmatched, ship):
    ws = wb.create_sheet("Не сошлось")
    ws["A1"] = "Отгружено, но в остатках не нашлось"
    ws["A1"].font = TITLE
    ws["A2"] = ("Эти штуки уехали, а строки под них в остатках менеджера "
                "нет или в ней меньше количество. Значит учет и отгрузка "
                "расходятся — проверять глазами, а не считать.")
    ws["A2"].font = NOTE
    cols = [("Подгруппа", 34), ("Ду", 8), ("Не сошлось, шт", 14),
            ("Что это было в отгрузке", 76)]
    start = 4
    for i, (h, _) in enumerate(cols, start=1):
        ws.cell(row=start, column=i, value=h)
    head(ws, start, len(cols))
    row = start + 1
    for u in unmatched:
        names = sorted({s["Наименование"] for s in ship
                        if s["Подгруппа"] == u["Подгруппа"]
                        and s["Ду"] == u["Ду"]})
        row = put(ws, row, [
            u["Подгруппа"], u["Ду"] or "", u["Штук"],
            "; ".join(x[:46] for x in names[:4])
            + (f"; и еще {len(names) - 4}" if len(names) > 4 else ""),
        ], wrap=(1, 4), center=(2, 3), fill=WARN_FILL)
    row = put(ws, row, ["ИТОГО", "", sum(u["Штук"] for u in unmatched), ""],
              center=(2, 3), font=BOLD, fill=GROUP_FILL)
    widths(ws, [w for _, w in cols])


def sheet_sources(wb, rows, ship):
    ws = wb.create_sheet("Источники")
    ws["A1"] = "Откуда цифры"
    ws["A1"].font = TITLE
    row = 3
    ws.cell(row=row, column=1, value="Остатки менеджера, сентябрь 2026").font = SUB
    row += 1
    cols = ["Файл", "Лист", "Позиций", "Штук", "Сумма, руб"]
    for i, h in enumerate(cols, start=1):
        ws.cell(row=row, column=i, value=h)
    head(ws, row, len(cols), height=20)
    row += 1
    agg = defaultdict(lambda: [0, 0, 0.0])
    for r in rows:
        a = agg[(r["Файл"], r["Лист"])]
        a[0] += 1
        a[1] += r["Штук"]
        a[2] += r["Сумма"]
    for (f, s), v in agg.items():
        row = put(ws, row, [f, s, v[0], v[1], round(v[2])],
                  center=(3, 4), money=(5,))
    row = put(ws, row, ["ИТОГО", "", sum(v[0] for v in agg.values()),
                        sum(v[1] for v in agg.values()),
                        round(sum(v[2] for v in agg.values()))],
              center=(3, 4), money=(5,), font=BOLD, fill=GROUP_FILL)

    row += 2
    ws.cell(row=row, column=1, value="Что вычтено как отгруженное").font = SUB
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
            value="Локации склада: что вошло в расчет").font = SUB
    row += 1
    for i, h in enumerate(["Локация", "Учетный файл", "Остатки за сентябрь"],
                          start=1):
        ws.cell(row=row, column=i, value=h)
    head(ws, row, 3, height=20)
    row += 1
    for name, fname, got in LOCATIONS:
        row = put(ws, row, [name, fname,
                            "прислана" if got else "НЕ ПРИСЛАНА"],
                  center=(3,), fill=None if got else WARN_FILL)
    have = sum(1 for _, _, g in LOCATIONS if g)
    row += 1
    for t in [
        f"Прислано {have} локаций из {len(LOCATIONS)}. Расчет сделан "
        f"только по ним, и это не весь склад.",
        "Поэтому часть отгруженного вычесть не из чего: штуцерная "
        "мелочь, которой набивали контейнер, лежит в локациях, по "
        "которым остатков нет. Эти штуки — на вкладке «Не сошлось».",
        "Чтобы увидеть склад целиком, нужны остатки за сентябрь по "
        "остальным локациям в той же форме.",
    ]:
        cell = ws.cell(row=row, column=1, value="— " + t)
        cell.font = BASE
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=row, start_column=1,
                       end_row=row, end_column=5)
        ws.row_dimensions[row].height = 28
        row += 1

    widths(ws, [40, 24, 18, 12, 16])


def main():
    stock = stock_sheets.read()
    ship = shipped()
    rows, unmatched = match(stock, ship)

    wb = Workbook()
    sheet_groups(wb, rows)
    sheet_rest(wb, rows)
    sheet_unmatched(wb, unmatched, ship)
    sheet_sources(wb, rows, ship)

    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / "Остаток_склада_по_группам.xlsx"
    wb.save(path)

    was = sum(r["Штук"] for r in rows)
    took = sum(r["Снято"] for r in rows)
    left = sum(r["Остаток"] for r in rows)
    print(f"готово: {path}\n")
    print(f"в файлах менеджера {len(rows)} позиций, {rub(was)} шт")
    print(f"снято как отгруженное {rub(took)} шт")
    print(f"ОСТАЛОСЬ {rub(left)} шт, "
          f"{len([r for r in rows if r['Остаток'] > 0])} позиций, "
          f"{rub(sum(r['СуммаОстаток'] for r in rows))} руб\n")
    agg = defaultdict(lambda: [0, 0, 0.0])
    for r in rows:
        if r["Остаток"] > 0:
            a = agg[r["Группа"]]
            a[0] += 1
            a[1] += r["Остаток"]
            a[2] += r["СуммаОстаток"]
    for g, v in sorted(agg.items(), key=lambda kv: -kv[1][2]):
        print(f"  {g:<26}{v[0]:>3} поз, {v[1]:>5} шт, {rub(v[2]):>12} руб")
    have = sum(1 for _, _, g in LOCATIONS if g)
    print(f"\nприслано {have} локаций из {len(LOCATIONS)} — это не весь склад")
    if unmatched:
        print(f"не сошлось {rub(sum(u['Штук'] for u in unmatched))} шт: "
              f"отгружено из локаций, по которым остатков не прислали")


if __name__ == "__main__":
    main()
