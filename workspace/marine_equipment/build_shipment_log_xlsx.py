#!/usr/bin/env python3
"""Учет отгрузки: состояние товара и доставка, разнесенная на единицу.

При отправке контейнера надо знать по каждой позиции две вещи: новая
она или б/у, и сколько на нее упало доставки. Без второй цифры нельзя
назвать покупателю цену: доставка Владивосток — Петербург на тяжелой
арматуре съедает заметную часть цены, а на мелком клапане почти
ничего.

Расходы вводятся один раз на листе «Расходы и доставка» и
разносятся по позициям. Способ разнесения выбирается там же: по весу,
по объему, поровну на штуку или пропорционально стоимости. Для
арматуры по умолчанию по весу — за контейнер платят за тонны.

Состояние в учетных файлах проставлено только у 54 позиций из 810,
поэтому колонка «Состояние» заполняется руками при погрузке: тогда же,
когда позицию берут в руки и кладут в ящик.

Результат: outputs/marine_equipment/Отгрузка_Питер_учет.xlsx
"""

import pathlib
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
from build_shipment_spb import (  # noqa: E402
    QUEUE, TIERS, item_kind, material, queue, select, tier,
)
from build_weight_xlsx import du  # noqa: E402

OUT = pathlib.Path(__file__).resolve().parents[2] / "outputs" / "marine_equipment"

FONT = "Arial"
HEAD_FILL = PatternFill("solid", fgColor="1F3864")
HEAD_FONT = Font(name=FONT, size=11, bold=True, color="FFFFFF")
TITLE = Font(name=FONT, size=14, bold=True, color="1F3864")
BOLD = Font(name=FONT, size=10, bold=True)
BASE = Font(name=FONT, size=10)
NOTE = Font(name=FONT, size=9, italic=True, color="666666")
INPUT_FILL = PatternFill("solid", fgColor="FFFF00")
CALC_FILL = PatternFill("solid", fgColor="EAF3EA")
SUB_FILL = PatternFill("solid", fgColor="D9E2F3")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

COSTS = "Расходы и доставка"
ITEMS = "Отгрузка"

# что вводим руками: строка листа расходов -> подпись
COST_LINES = [
    "Контейнер: фрахт Владивосток — Петербург",
    "Автодоставка до контейнера во Владивостоке",
    "Погрузка и выгрузка, такелаж",
    "Тара: поддоны, доска, ящики",
    "Упаковка: стрейч, лента, фанера на фланцы",
    "Экспедирование и документы",
    "Страхование груза",
    "Прочие расходы",
]

METHODS = ["по весу", "по объему", "поровну на штуку", "по стоимости"]
CONDITIONS = ["Новое", "Б/у", "Восстановленное"]
COMPLETE = ["Комплект", "Не комплект", "Брак"]

COLUMNS = [
    ("№", 5),
    ("Очередь", 26),
    ("Ярус", 15),
    ("Тип", 22),
    ("Диаметр", 10),
    ("Наименование", 46),
    ("Чертеж", 18),
    ("Состояние", 15),
    ("Комплектность", 15),
    ("Материал", 20),
    ("Числится, шт", 12),
    ("Отгружено, шт", 13),
    ("Вес ед., кг", 12),
    ("Вес всего, кг", 13),
    ("Объем всего, м³", 14),
    ("Доля в отгрузке", 14),
    ("Доставка на позицию, руб", 16),
    ("Доставка на ед., руб", 15),
    ("Цена за ед., руб", 14),
    ("Сумма позиции, руб", 16),
    ("Цена с доставкой за ед., руб", 17),
    ("Доставка в цене, %", 13),
    ("№ места", 10),
    ("Локация", 15),
    ("Примечание", 30),
]

# буквы колонок, чтобы формулы читались
C = {name: get_column_letter(i) for i, (name, _) in enumerate(COLUMNS, start=1)}

START = 5  # первая строка данных


def head(ws, row, ncols, height=40):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEAD_FILL
        cell.font = HEAD_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[row].height = height


def sheet_costs(wb, last):
    ws = wb.create_sheet(COSTS, 0)
    ws["A1"] = "Расходы на отгрузку и как они ложатся на товар"
    ws["A1"].font = TITLE
    ws["A2"] = ("Желтые ячейки заполняем сами. Все остальное считается. "
                "Пока расходы не введены, доставка на единицу будет нулевой.")
    ws["A2"].font = NOTE

    row = 4
    ws.cell(row=row, column=1, value="Расход").font = BOLD
    ws.cell(row=row, column=2, value="Сумма, руб").font = BOLD
    ws.cell(row=row, column=3, value="Комментарий").font = BOLD
    head(ws, row, 3, height=22)

    first = row + 1
    for name in COST_LINES:
        row += 1
        ws.cell(row=row, column=1, value=name).font = BASE
        cell = ws.cell(row=row, column=2)
        cell.fill = INPUT_FILL
        cell.number_format = "#,##0"
        for c in (1, 2, 3):
            ws.cell(row=row, column=c).border = BORDER

    row += 1
    total_row = row
    ws.cell(row=row, column=1, value="ИТОГО расходов на отгрузку").font = BOLD
    cell = ws.cell(row=row, column=2, value=f"=SUM(B{first}:B{row - 1})")
    cell.font = BOLD
    cell.number_format = "#,##0"
    cell.border = BORDER
    cell.fill = CALC_FILL

    row += 2
    ws.cell(row=row, column=1, value="Как разносим на позиции").font = BOLD
    method_row = row
    cell = ws.cell(row=row, column=2, value="по весу")
    cell.fill = INPUT_FILL
    cell.border = BORDER
    cell.font = BASE
    ws.cell(row=row, column=3,
            value="За контейнер платят за тонны, поэтому по умолчанию "
                  "по весу").font = NOTE
    dv = DataValidation(type="list", formula1='"' + ",".join(METHODS) + '"',
                        allow_blank=False)
    ws.add_data_validation(dv)
    dv.add(cell)

    row += 2
    ws.cell(row=row, column=1, value="Что набралось в отгрузке").font = TITLE
    row += 1
    base = [
        ("Вес всего, кг", f"=SUM('{ITEMS}'!{C['Вес всего, кг']}{START}:"
                          f"{C['Вес всего, кг']}{last})", "#,##0"),
        ("Объем всего, м³", f"=SUM('{ITEMS}'!{C['Объем всего, м³']}{START}:"
                            f"{C['Объем всего, м³']}{last})", "#,##0.0"),
        ("Штук всего", f"=SUM('{ITEMS}'!{C['Отгружено, шт']}{START}:"
                       f"{C['Отгружено, шт']}{last})", "#,##0"),
        ("Стоимость товара, руб",
         f"=SUM('{ITEMS}'!{C['Сумма позиции, руб']}{START}:"
         f"{C['Сумма позиции, руб']}{last})", "#,##0"),
    ]
    refs = {}
    for name, formula, fmt in base:
        row += 1
        ws.cell(row=row, column=1, value=name).font = BASE
        cell = ws.cell(row=row, column=2, value=formula)
        cell.font = BOLD
        cell.number_format = fmt
        cell.border = BORDER
        cell.fill = CALC_FILL
        refs[name] = row

    row += 2
    ws.cell(row=row, column=1, value="Сколько стоит перевозка").font = TITLE
    kg_row, sum_row = refs["Вес всего, кг"], refs["Стоимость товара, руб"]
    for name, formula, fmt in [
        ("Доставка на килограмм, руб",
         f"=IFERROR(B{total_row}/B{kg_row},0)", "#,##0.00"),
        ("Доставка на штуку в среднем, руб",
         f"=IFERROR(B{total_row}/B{refs['Штук всего']},0)", "#,##0.00"),
        ("Доставка в процентах от стоимости товара",
         f"=IFERROR(B{total_row}/B{sum_row},0)", "0.0%"),
        ("Товар с доставкой, руб",
         f"=B{sum_row}+B{total_row}", "#,##0"),
    ]:
        row += 1
        ws.cell(row=row, column=1, value=name).font = BASE
        cell = ws.cell(row=row, column=2, value=formula)
        cell.font = BOLD
        cell.number_format = fmt
        cell.border = BORDER
        cell.fill = CALC_FILL

    ws.column_dimensions["A"].width = 46
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 52
    return total_row, method_row, refs


def share_formula(row, total_row, method_row, refs):
    """Доля позиции в отгрузке по выбранному способу разнесения."""
    m = f"'{COSTS}'!$B${method_row}"

    def ref(name):
        return f"'{COSTS}'!$B${refs[name]}"

    weight = f"{C['Вес всего, кг']}{row}/{ref('Вес всего, кг')}"
    volume = f"{C['Объем всего, м³']}{row}/{ref('Объем всего, м³')}"
    pieces = f"{C['Отгружено, шт']}{row}/{ref('Штук всего')}"
    money = f"{C['Сумма позиции, руб']}{row}/{ref('Стоимость товара, руб')}"
    return (f'=IFERROR(IF({m}="по весу",{weight},'
            f'IF({m}="по объему",{volume},'
            f'IF({m}="поровну на штуку",{pieces},{money}))),0)')


def sheet_items(wb, go, total_row, method_row, refs):
    ws = wb.create_sheet(ITEMS)
    ws["A1"] = "Отгрузка в Санкт-Петербург: что грузим и почем выходит"
    ws["A1"].font = TITLE
    ws["A2"] = ("Желтые колонки заполняем при погрузке. Состояние ставим "
                "на месте: в учете оно проставлено только у части позиций. "
                "Зеленые колонки считаются сами.")
    ws["A2"].font = NOTE

    for i, (name, _) in enumerate(COLUMNS, start=1):
        ws.cell(row=4, column=i, value=name)
    head(ws, 4, len(COLUMNS))

    cond_map = {"новый": "Новое", "б/у": "Б/у"}
    comp_map = {"комплект": "Комплект", "не комплект": "Не комплект",
                "брак": "Брак"}

    for n, r in enumerate(go, start=1):
        row = START + n - 1
        d = du(r["Наименование"])
        values = [
            n,
            QUEUE[queue(r)],
            TIERS[tier(r)][0],
            item_kind(r),
            f"Ду{d}" if d else "",
            r["Наименование"],
            r["Чертеж"],
            cond_map.get(r["Состояние"], ""),
            comp_map.get(r["Комплектность"], "Комплект"),
            material(r),
            r["Наличие"],
            r["Наличие"],
            r["вес_ед"],
            f"={C['Отгружено, шт']}{row}*{C['Вес ед., кг']}{row}",
            round(r["объем"], 3),
            share_formula(row, total_row, method_row, refs),
            f"={C['Доля в отгрузке']}{row}*'{COSTS}'!$B${total_row}",
            f"=IFERROR({C['Доставка на позицию, руб']}{row}"
            f"/{C['Отгружено, шт']}{row},0)",
            r["Цена за ед., руб"] or None,
            f"={C['Отгружено, шт']}{row}*{C['Цена за ед., руб']}{row}",
            f"={C['Цена за ед., руб']}{row}"
            f"+{C['Доставка на ед., руб']}{row}",
            f"=IFERROR({C['Доставка на ед., руб']}{row}"
            f"/{C['Цена за ед., руб']}{row},\"\")",
            "",
            r["Локация"],
            "",
        ]
        for i, v in enumerate(values, start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top",
                                       wrap_text=(i in (2, 6, 25)))
        for name in ("Вес всего, кг", "Доставка на позицию, руб",
                     "Доставка на ед., руб", "Сумма позиции, руб",
                     "Цена с доставкой за ед., руб"):
            ws[f"{C[name]}{row}"].number_format = "#,##0"
        ws[f"{C['Цена за ед., руб']}{row}"].number_format = "#,##0"
        ws[f"{C['Объем всего, м³']}{row}"].number_format = "#,##0.000"
        ws[f"{C['Доля в отгрузке']}{row}"].number_format = "0.00%"
        ws[f"{C['Доставка в цене, %']}{row}"].number_format = "0.0%"
        ws[f"{C['Доставка на ед., руб']}{row}"].number_format = "#,##0.00"

        # руками заполняем состояние, факт отгрузки, вес, цену и место
        for name in ("Состояние", "Комплектность", "Отгружено, шт",
                     "Вес ед., кг", "Цена за ед., руб", "№ места",
                     "Примечание"):
            ws[f"{C[name]}{row}"].fill = INPUT_FILL
        for name in ("Вес всего, кг", "Доля в отгрузке",
                     "Доставка на позицию, руб", "Доставка на ед., руб",
                     "Сумма позиции, руб", "Цена с доставкой за ед., руб",
                     "Доставка в цене, %"):
            ws[f"{C[name]}{row}"].fill = CALC_FILL

    last = START + len(go) - 1

    dv_cond = DataValidation(type="list",
                             formula1='"' + ",".join(CONDITIONS) + '"',
                             allow_blank=True)
    dv_comp = DataValidation(type="list",
                             formula1='"' + ",".join(COMPLETE) + '"',
                             allow_blank=True)
    ws.add_data_validation(dv_cond)
    ws.add_data_validation(dv_comp)
    dv_cond.add(f"{C['Состояние']}{START}:{C['Состояние']}{last}")
    dv_comp.add(f"{C['Комплектность']}{START}:{C['Комплектность']}{last}")

    total = last + 1
    ws.cell(row=total, column=6, value="ИТОГО").font = BOLD
    for name in ("Числится, шт", "Отгружено, шт", "Вес всего, кг",
                 "Объем всего, м³", "Доставка на позицию, руб",
                 "Сумма позиции, руб"):
        letter = C[name]
        cell = ws.cell(row=total, column=COLUMNS.index(
            next(c for c in COLUMNS if c[0] == name)) + 1,
            value=f"=SUM({letter}{START}:{letter}{last})")
        cell.font = BOLD
        cell.border = BORDER
        cell.number_format = ("#,##0.0" if name == "Объем всего, м³"
                              else "#,##0")

    ws.auto_filter.ref = f"A4:{get_column_letter(len(COLUMNS))}{last}"
    ws.freeze_panes = f"{C['Чертеж']}{START}"
    for i, (_, width) in enumerate(COLUMNS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width
    return last


def sheet_totals(wb, last, total_row):
    ws = wb.create_sheet("Итоги по состоянию")
    ws["A1"] = "Новое и б/у: сколько чего уехало"
    ws["A1"].font = TITLE
    ws["A2"] = ("Считается по колонке «Состояние» листа «Отгрузка». "
                "Пока она не заполнена, все попадет в строку «не указано».")
    ws["A2"].font = NOTE

    cond = f"'{ITEMS}'!{C['Состояние']}${START}:{C['Состояние']}${last}"
    pcs = f"'{ITEMS}'!{C['Отгружено, шт']}${START}:{C['Отгружено, шт']}${last}"
    kg = f"'{ITEMS}'!{C['Вес всего, кг']}${START}:{C['Вес всего, кг']}${last}"
    money = (f"'{ITEMS}'!{C['Сумма позиции, руб']}${START}:"
             f"{C['Сумма позиции, руб']}${last}")
    ship = (f"'{ITEMS}'!{C['Доставка на позицию, руб']}${START}:"
            f"{C['Доставка на позицию, руб']}${last}")

    row = 4
    for i, h in enumerate(["Состояние", "Позиций", "Штук", "Вес, кг",
                           "Стоимость, руб", "Доставка, руб",
                           "Всего с доставкой, руб"], start=1):
        ws.cell(row=row, column=i, value=h)
    head(ws, row, 7, height=30)

    for name in CONDITIONS + ["не указано"]:
        row += 1
        crit = f'"{name}"' if name != "не указано" else '""'
        # пустую ячейку COUNTIF считает ненадежно, для нее COUNTBLANK
        count = (f"=COUNTBLANK({cond})" if name == "не указано"
                 else f"=COUNTIF({cond},{crit})")
        values = [
            name,
            count,
            f"=SUMIF({cond},{crit},{pcs})",
            f"=SUMIF({cond},{crit},{kg})",
            f"=SUMIF({cond},{crit},{money})",
            f"=SUMIF({cond},{crit},{ship})",
            f"=E{row}+F{row}",
        ]
        for i, v in enumerate(values, start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            if i >= 4:
                cell.number_format = "#,##0"

    row += 1
    ws.cell(row=row, column=1, value="ИТОГО").font = BOLD
    for col in range(2, 8):
        letter = get_column_letter(col)
        cell = ws.cell(row=row, column=col,
                       value=f"=SUM({letter}{row - 4}:{letter}{row - 1})")
        cell.font = BOLD
        cell.border = BORDER
        cell.number_format = "#,##0"

    row += 2
    ws.cell(row=row, column=1, value="Проверка").font = TITLE
    for name, formula in [
        ("Разнесено доставки, руб",
         f"=SUM({ship})"),
        ("Введено расходов, руб",
         f"='{COSTS}'!$B${total_row}"),
        ("Расхождение, руб",
         f"=SUM({ship})-'{COSTS}'!$B${total_row}"),
    ]:
        row += 1
        ws.cell(row=row, column=1, value=name).font = BASE
        cell = ws.cell(row=row, column=2, value=formula)
        cell.font = BOLD
        cell.number_format = "#,##0"
        cell.border = BORDER
    row += 1
    cell = ws.cell(row=row, column=1,
                   value="Расхождение должно быть нулевым. Если не ноль — "
                         "в отгрузке есть строка без веса или без цены, "
                         "и доля по ней не посчиталась.")
    cell.font = NOTE
    cell.alignment = Alignment(wrap_text=True)

    for i, w in enumerate([34, 12, 12, 14, 18, 16, 20], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    return ws


HOWTO = [
    ("Состояние ставим при погрузке, не после",
     "Позицию все равно берут в руки и кладут в ящик — тогда и видно, "
     "новая она или б/у. В учете состояние проставлено только у 54 "
     "позиций из 810, по остальным его просто нет. Заполнить потом по "
     "памяти не получится."),
    ("Новое и б/у не смешиваем в одном месте",
     "Покупатель принимает груз по упаковочному листу. Если в ящике "
     "лежит и новое, и б/у, спор о состоянии придет на весь ящик."),
    ("Сначала расходы, потом цены",
     "Доставка разносится от суммы на листе расходов. Пока там нули, "
     "колонка «Доставка на ед.» будет нулевой, и цена с доставкой "
     "совпадет с ценой товара."),
    ("Разносим по весу",
     "За контейнер платят за тонны, поэтому доставка ложится на вес. "
     "На задвижке Ду250 это будут тысячи рублей на штуку, на штуцерном "
     "клапане — рубли. Разнести поровну на штуку значит завысить цену "
     "мелочи и занизить цену тяжелой арматуры."),
    ("Вес можно заменить на фактический",
     "Колонка «Вес ед.» желтая: там стоит справочная масса с "
     "погрешностью около трети. Взвесили место — впишите факт, "
     "разнесение пересчитается."),
    ("Отгружено, а не числится",
     "Колонка «Числится» — то, что стоит в учете на 16.06.2025 и не "
     "подтверждено инвентаризацией. Считаем по колонке «Отгружено»: "
     "сколько реально положили в контейнер."),
    ("Цена с доставкой — это не цена продажи",
     "Это то, во что позиция обошлась на складе в Петербурге. Цену "
     "покупателю ставим от нее, иначе тяжелая арматура уедет в минус."),
    ("Каждое место — свой номер",
     "Номер места в отгрузке должен совпадать с биркой на ящике и "
     "строкой упаковочного листа. Иначе принимать груз придется "
     "вскрытием всего подряд."),
]


def sheet_howto(wb):
    ws = wb.create_sheet("Как заполнять", 0)
    ws["A1"] = "Как вести этот файл"
    ws["A1"].font = TITLE
    ws["A2"] = ("Желтое — вводим руками. Зеленое — считается. "
                "Файл ведется во время погрузки, а не после отправки.")
    ws["A2"].font = NOTE
    for i, h in enumerate(["№", "Правило", "Почему так"], start=1):
        ws.cell(row=4, column=i, value=h)
    head(ws, 4, 3, height=22)
    for n, (rule, why) in enumerate(HOWTO, start=1):
        row = 4 + n
        for i, v in enumerate([n, rule, why], start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=(i > 1))
        ws.row_dimensions[row].height = 56
    for i, w in enumerate([5, 40, 88], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    return ws


def main():
    go, _, _, kg, _ = select()
    last = START + len(go) - 1

    wb = Workbook()
    wb.remove(wb.active)
    total_row, method_row, refs = sheet_costs(wb, last)
    sheet_items(wb, go, total_row, method_row, refs)
    sheet_totals(wb, last, total_row)
    sheet_howto(wb)
    wb.move_sheet("Как заполнять", offset=-wb.index(wb["Как заполнять"]))

    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / "Отгрузка_Питер_учет.xlsx"
    wb.save(path)

    known = sum(1 for r in go if r["Состояние"] in ("новый", "б/у"))
    print(f"готово: {path}")
    print(f"позиций в отгрузке: {len(go)}, "
          f"{sum(r['Наличие'] for r in go):.0f} шт, {kg / 1000:.1f} т")
    print(f"состояние известно по {known} позициям, "
          f"по остальным {len(go) - known} заполняется при погрузке")


if __name__ == "__main__":
    main()
