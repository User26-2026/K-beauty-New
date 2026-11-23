#!/usr/bin/env python3
"""Бланк отгрузки: что отправили, новое или б/у, почем и сколько
доставки легло на штуку.

Форма на 300 строк. Одна строка — одна позиция в одной отгрузке, так
что в файле живут и контейнер, и посылки транспортной компанией: канал,
город и номер расписки стоят в самой строке.

Вписываем наименование, единицу измерения, состояние, количество, вес,
цену за штуку и доставку за штуку — остальное считается.

Верхний блок повторяет экспедиторскую расписку: перевозка это только
часть счета, организация перевозки, упаковка и страхование дают больше.
Разделите итог на вес — получите цену килограмма.

Две первые строки заполнены реальными отгрузками от 24.09.2026 как
пример, их стираем перед работой.

Результат: outputs/marine_equipment/Отгрузка_образец.xlsx
"""

import pathlib
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
from shipments_done import SHIPMENTS, cost, per_kg  # noqa: E402

OUT = pathlib.Path(__file__).resolve().parents[2] / "outputs" / "marine_equipment"

FONT = "Arial"
HEAD_FILL = PatternFill("solid", fgColor="1F3864")
HEAD_FONT = Font(name=FONT, size=11, bold=True, color="FFFFFF")
TITLE = Font(name=FONT, size=14, bold=True, color="1F3864")
SUB = Font(name=FONT, size=11, bold=True, color="1F3864")
BOLD = Font(name=FONT, size=10, bold=True)
BASE = Font(name=FONT, size=10)
NOTE = Font(name=FONT, size=9, italic=True, color="666666")
EXAMPLE = Font(name=FONT, size=10, italic=True, color="808080")
INPUT_FILL = PatternFill("solid", fgColor="FFFF00")
CALC_FILL = PatternFill("solid", fgColor="EAF3EA")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

SHEET = "Отгрузка"
ROWS = 300

UNITS = ["шт", "компл", "кг", "м", "пара", "набор"]
CONDITIONS = ["Новое", "Б/у"]
CHANNELS = ["Контейнер", "ТК Энергия", "Другая ТК", "Самовывоз"]

# строки счета — как они названы в экспедиторской расписке
COST_LINES = [
    "Организация перевозки",
    "Перевозка",
    "Упаковка, доп. сервис отправителю",
    "Забор от отправителя",
    "Доставка до получателя",
    "Страхование",
    "Оповещения",
]

# колонка: (заголовок, ширина, вводим руками)
COLUMNS = [
    ("№", 5, False),
    ("Дата отгрузки", 13, True),
    ("Куда", 16, True),
    ("Покупатель", 24, True),
    ("Чем отправлено", 15, True),
    ("Документ, расписка", 16, True),
    ("Наименование товара", 46, True),
    ("Ед. изм.", 9, True),
    ("Новое или б/у", 13, True),
    ("Количество", 11, True),
    ("Мест", 8, True),
    ("Вес всего, кг", 12, True),
    ("Вес за шт, кг", 12, False),
    ("Цена за шт, руб", 13, True),
    ("Стоимость доставки за шт, руб", 15, True),
    ("Цена с доставкой за шт, руб", 15, False),
    ("Сумма, руб", 13, False),
    ("Доставка всего, руб", 14, False),
    ("Итого по строке, руб", 14, False),
    ("Объявленная ценность, руб", 15, True),
    ("№ места", 9, True),
    ("Примечание", 28, True),
]

L = {name: get_column_letter(i) for i, (name, _, _) in
     enumerate(COLUMNS, start=1)}

HOWTO = [
    ("Одна строка — одна позиция одной отгрузки",
     "Канал, город и номер расписки стоят в самой строке, поэтому в "
     "файле спокойно живут и контейнер, и посылки транспортной "
     "компанией. Итоги считаются по каналам отдельно."),
    ("Состояние ставим при погрузке",
     "Новое или б/у видно только тогда, когда позицию берут в руки и "
     "кладут в ящик. Потом по памяти это не восстановить. В учетных "
     "файлах состояние стоит только у 54 позиций из 810."),
    ("Вес пишем по расписке, брутто",
     "Экспедитор считает деньги от веса вместе с тарой. По этому же "
     "весу разносим доставку. Заодно накапливаем настоящие веса: "
     "справочные цифры по задвижкам оказались завышены вдвое."),
    ("Доставка за штуку считается от веса",
     "Разделите итог по расписке на общий вес груза — получится цена "
     "килограмма, умножьте на вес штуки. Блок для этого счета стоит "
     "вверху листа."),
    ("Поровну на штуку разносить нельзя",
     "Задвижка Ду250 весит 220 кг, штуцерный клапан — килограмм. Если "
     "разделить доставку поровну, мелочь подорожает в разы, а тяжелая "
     "арматура уедет дешевле себестоимости."),
    ("Перевозка — это не весь счет",
     "По расписке в Керчь перевозка 27 000 из 135 912, остальное "
     "организация перевозки. В Питер перевозка 7 852 из 47 225, плюс "
     "упаковка 6 150 и страховка 1 800. Считать надо по итогу."),
    ("Цена с доставкой — это не цена продажи",
     "Это то, во что позиция обошлась на складе у покупателя. Цену "
     "ставим от нее, иначе тяжелое уедет в минус."),
    ("Объявленная ценность — от нее страховка",
     "Занизите — страховка дешевле, но при бое получите гроши. "
     "Завысите — переплатите за страхование."),
    ("Новое и б/у не смешиваем в одном месте",
     "Покупатель принимает груз по упаковочному листу. Если в ящике "
     "вперемешку, спор о состоянии придет на весь ящик."),
]


def head(ws, row, ncols, height=44):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEAD_FILL
        cell.font = HEAD_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[row].height = height


def examples():
    """Первые строки бланка — реальные отгрузки, чтобы было видно,
    как он заполняется."""
    out = []
    for s in SHIPMENTS:
        for i in s["items"]:
            qty = i["qty"]
            out.append([
                s["date"], s["city"], s["buyer"], "ТК Энергия", s["doc"],
                i["name"], i["unit"], i["condition"] or None, qty,
                s["places"], s["kg"], None,
                None, round(cost(s) / qty, 2), None, None, None, None,
                s["declared"], None,
                f"доставка вышла {per_kg(s):.0f} руб за кг",
            ])
    return out


def cost_block(ws):
    """Верхний блок: расписка и цена килограмма."""
    ws["A4"] = "Расходы по расписке"
    ws["A4"].font = SUB
    row = 5
    first = row
    for name in COST_LINES:
        ws.cell(row=row, column=1, value=name).font = BASE
        cell = ws.cell(row=row, column=2)
        cell.fill = INPUT_FILL
        cell.border = BORDER
        cell.number_format = "#,##0"
        ws.cell(row=row, column=1).border = BORDER
        row += 1

    ws.cell(row=row, column=1, value="Все расходы на доставку").font = BOLD
    total = ws.cell(row=row, column=2, value=f"=SUM(B{first}:B{row - 1})")
    total.font = BOLD
    total.fill = CALC_FILL
    total.border = BORDER
    total.number_format = "#,##0"
    total_row = row

    row += 1
    ws.cell(row=row, column=1, value="Общий вес груза, кг").font = BASE
    cell = ws.cell(row=row, column=2)
    cell.fill = INPUT_FILL
    cell.border = BORDER
    cell.number_format = "#,##0"
    weight_row = row

    row += 1
    ws.cell(row=row, column=1, value="Доставка за килограмм, руб").font = BOLD
    cell = ws.cell(row=row, column=2,
                   value=f"=IFERROR(B{total_row}/B{weight_row},0)")
    cell.font = BOLD
    cell.fill = CALC_FILL
    cell.border = BORDER
    cell.number_format = "#,##0.00"

    rate = (sum(cost(s) for s in SHIPMENTS)
            / sum(s["kg"] for s in SHIPMENTS))
    parts = ", ".join(f"{s['city']} {per_kg(s):.0f}" for s in SHIPMENTS)
    cell = ws.cell(row=row, column=3,
                   value=f"по факту 24.09.2026 вышло {rate:.0f} руб за кг "
                         f"({parts}). Умножьте на вес штуки — получится "
                         f"доставка за штуку")
    cell.font = NOTE
    return row


def sheet_form(wb):
    ws = wb.active
    ws.title = SHEET
    ws["A1"] = "Отгрузка: что отправили и почем это обошлось"
    ws["A1"].font = TITLE
    ws["A2"] = ("Желтое вписываем сами, зеленое считается. Две первые "
                "строки — реальные отгрузки от 24.09.2026 как пример, "
                "сотрите их перед работой.")
    ws["A2"].font = NOTE

    last_block = cost_block(ws)
    header = last_block + 2
    start = header + 1

    for i, (name, _, _) in enumerate(COLUMNS, start=1):
        ws.cell(row=header, column=i, value=name)
    head(ws, header, len(COLUMNS))

    ex = examples()
    last = start + ROWS - 1
    for n in range(1, ROWS + 1):
        row = start + n - 1
        sample = ex[n - 1] if n <= len(ex) else None
        values = [n] + (sample if sample else [None] * (len(COLUMNS) - 1))
        # расчетные колонки ставим поверх примера
        values[COLUMNS.index(("Вес за шт, кг", 12, False)) ] = (
            f'=IF({L["Количество"]}{row}="","",'
            f'{L["Вес всего, кг"]}{row}/{L["Количество"]}{row})')
        values[15] = (f'=IF({L["Цена за шт, руб"]}{row}="","",'
                      f'{L["Цена за шт, руб"]}{row}'
                      f'+{L["Стоимость доставки за шт, руб"]}{row})')
        values[16] = (f'=IF({L["Цена за шт, руб"]}{row}="","",'
                      f'{L["Количество"]}{row}*{L["Цена за шт, руб"]}{row})')
        values[17] = (f'=IF({L["Стоимость доставки за шт, руб"]}{row}="","",'
                      f'{L["Количество"]}{row}'
                      f'*{L["Стоимость доставки за шт, руб"]}{row})')
        values[18] = (f'=IF({L["Цена за шт, руб"]}{row}="","",'
                      f'{L["Сумма, руб"]}{row}'
                      f'+{L["Доставка всего, руб"]}{row})')

        for i, v in enumerate(values, start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = EXAMPLE if sample else BASE
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top",
                                       wrap_text=(i in (7, 22)))
        for name in ("Вес всего, кг", "Вес за шт, кг", "Цена за шт, руб",
                     "Стоимость доставки за шт, руб",
                     "Цена с доставкой за шт, руб", "Сумма, руб",
                     "Доставка всего, руб", "Итого по строке, руб",
                     "Объявленная ценность, руб"):
            ws[f"{L[name]}{row}"].number_format = "#,##0"
        ws[f'{L["Вес за шт, кг"]}{row}'].number_format = "#,##0.0"
        ws[f'{L["Стоимость доставки за шт, руб"]}{row}'
           ].number_format = "#,##0.00"
        for i, (name, _, manual) in enumerate(COLUMNS, start=1):
            if i == 1:
                continue
            ws[f"{get_column_letter(i)}{row}"].fill = (
                INPUT_FILL if manual else CALC_FILL)

    dv = {
        "Ед. изм.": UNITS,
        "Новое или б/у": CONDITIONS,
        "Чем отправлено": CHANNELS,
    }
    for name, options in dv.items():
        v = DataValidation(type="list",
                           formula1='"' + ",".join(options) + '"',
                           allow_blank=True)
        ws.add_data_validation(v)
        v.add(f"{L[name]}{start}:{L[name]}{last}")

    total = last + 1
    ws.cell(row=total, column=7, value="ИТОГО").font = BOLD
    for name in ("Количество", "Мест", "Вес всего, кг", "Сумма, руб",
                 "Доставка всего, руб", "Итого по строке, руб",
                 "Объявленная ценность, руб"):
        cell = ws[f"{L[name]}{total}"]
        cell.value = f"=SUM({L[name]}{start}:{L[name]}{last})"
        cell.font = BOLD
        cell.border = BORDER
        cell.number_format = "#,##0"
    cell = ws.cell(row=total, column=len(COLUMNS))
    cell.value = (f'="строк заполнено: "&COUNTA({L["Наименование товара"]}'
                  f'{start}:{L["Наименование товара"]}{last})')
    cell.font = NOTE

    ws.auto_filter.ref = (f"A{header}:"
                          f"{get_column_letter(len(COLUMNS))}{last}")
    ws.freeze_panes = f"G{start}"
    for i, (_, width, _) in enumerate(COLUMNS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width
    return start, last


def sheet_totals(wb, start, last):
    ws = wb.create_sheet("Итоги")
    ws["A1"] = "Итоги по состоянию и по каналу отгрузки"
    ws["A1"].font = TITLE
    ws["A2"] = ("Считается по листу «Отгрузка». Строка «не указано» "
                "показывает, что еще не заполнили.")
    ws["A2"].font = NOTE

    def rng(name):
        return f"'{SHEET}'!${L[name]}${start}:${L[name]}${last}"

    def block(row, title, key, values, label):
        ws.cell(row=row, column=1, value=title).font = SUB
        row += 1
        for i, h in enumerate([label, "Строк",
                               "Количество", "Вес, кг", "Сумма, руб",
                               "Доставка, руб", "Итого с доставкой, руб"],
                              start=1):
            ws.cell(row=row, column=i, value=h)
        head(ws, row, 7, height=28)
        first = row + 1
        for name in values + ["не указано"]:
            row += 1
            if name == "не указано":
                count = (f'=COUNTIFS({rng("Наименование товара")},"<>",'
                         f'{rng(key)},"")')
                crit = '""'
            else:
                count = f'=COUNTIF({rng(key)},"{name}")'
                crit = f'"{name}"'
            cells = [
                name, count,
                f'=SUMIF({rng(key)},{crit},{rng("Количество")})',
                f'=SUMIF({rng(key)},{crit},{rng("Вес всего, кг")})',
                f'=SUMIF({rng(key)},{crit},{rng("Сумма, руб")})',
                f'=SUMIF({rng(key)},{crit},{rng("Доставка всего, руб")})',
                f"=E{row}+F{row}",
            ]
            for i, v in enumerate(cells, start=1):
                cell = ws.cell(row=row, column=i, value=v)
                cell.font = BASE
                cell.border = BORDER
                if i >= 3:
                    cell.number_format = "#,##0"
        row += 1
        ws.cell(row=row, column=1, value="ИТОГО").font = BOLD
        for col in range(2, 8):
            letter = get_column_letter(col)
            cell = ws.cell(row=row, column=col,
                           value=f"=SUM({letter}{first}:{letter}{row - 1})")
            cell.font = BOLD
            cell.border = BORDER
            cell.number_format = "#,##0"
        return row + 2

    row = block(4, "Новое и б/у", "Новое или б/у", CONDITIONS, "Состояние")
    row = block(row, "По каналу отгрузки", "Чем отправлено", CHANNELS,
                "Чем отправлено")

    ws.cell(row=row, column=1, value="Доля доставки в цене").font = SUB
    row += 1
    for name, formula, fmt in [
        ("Товар без доставки, руб", f'=SUM({rng("Сумма, руб")})', "#,##0"),
        ("Доставка, руб", f'=SUM({rng("Доставка всего, руб")})', "#,##0"),
        ("Доставка в процентах от товара",
         f'=IFERROR(SUM({rng("Доставка всего, руб")})'
         f'/SUM({rng("Сумма, руб")}),0)', "0.0%"),
        ("Доставка за килограмм в среднем, руб",
         f'=IFERROR(SUM({rng("Доставка всего, руб")})'
         f'/SUM({rng("Вес всего, кг")}),0)', "#,##0.00"),
        ("Всего с доставкой, руб",
         f'=SUM({rng("Сумма, руб")})+SUM({rng("Доставка всего, руб")})',
         "#,##0"),
    ]:
        ws.cell(row=row, column=1, value=name).font = BASE
        cell = ws.cell(row=row, column=2, value=formula)
        cell.font = BOLD
        cell.border = BORDER
        cell.number_format = fmt
        row += 1

    for i, w in enumerate([36, 14, 14, 14, 18, 16, 20], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    return ws


def sheet_howto(wb):
    ws = wb.create_sheet("Как заполнять")
    ws["A1"] = "Как вести бланк"
    ws["A1"].font = TITLE
    for i, h in enumerate(["№", "Правило", "Почему так"], start=1):
        ws.cell(row=3, column=i, value=h)
    head(ws, 3, 3, height=22)
    for n, (rule, why) in enumerate(HOWTO, start=1):
        row = 3 + n
        for i, v in enumerate([n, rule, why], start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = BASE
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=(i > 1))
        ws.row_dimensions[row].height = 50
    for i, w in enumerate([5, 42, 86], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    return ws


def main():
    wb = Workbook()
    start, last = sheet_form(wb)
    sheet_totals(wb, start, last)
    sheet_howto(wb)

    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / "Отгрузка_образец.xlsx"
    wb.save(path)
    print(f"готово: {path}")
    print(f"строк под заполнение: {ROWS}, колонок: {len(COLUMNS)}, "
          f"данные с {start} строки")
    print(f"пример: {len(examples())} строки из реальных отгрузок")


if __name__ == "__main__":
    main()
