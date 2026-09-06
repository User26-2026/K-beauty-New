"""Заказ покупателю: берем только позиции с остатком от 3 тысяч штук."""

import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import name_match

STOCK = "data/stock_costs/Остатки_31.08.2026.xlsx"
PLAN = "data/shipments/2026-09_plan_raspredeleniya.xlsx"
OUT = "outputs/Заказ покупателю.xlsx"
MIN_STOCK = 3000
GIVE = 3000
# Отдельные позиции даем меньше стандартных 3000: по ним свои причины.
OVERRIDES = {
    "FARMSTAY - Black Garlic Nourishing Shampoo": 1000,
    "FARMSTAY - Collagen Water Full Shampoo": 1000,
    # Сыворотку придерживаем: нужен запас, чтобы торговать самим.
    "CELIMAX THE VITA-A RETINOL SHOT TIGHTENING SERUM": 1000,
}
# Позиции, которых на складе меньше порога, но в заказ их берем отдельно.
ADDITIONS = {
    "FARMSTAY - Argan Oil Complete Volume Up Shampoo": 1000,
    # Линейка DEOPROCE уходит целиком, остатки там маленькие.
    "DEOPROCE SHAMPOO - BLACK GARLIC INTENSME ENERGY [200ml]": 240,
    "DEOPROCE SHAMPOO - BLACK GARLIC INTENSME ENERGY [1000ml]": 400,
    "DEOPROCE RINSE - BLACK GARLIC INTENSME ENERGY [1000ml]": 400,
    "ELIZAVECCA - CER-100 Collagen Ceramide Coating Protein Treatment": 300,
}

HEADER_FILL = PatternFill("solid", fgColor="DDEBF7")
TOTAL_FILL = PatternFill("solid", fgColor="FFF2CC")


def read_stock(path):
    sheet = openpyxl.load_workbook(path, data_only=True)["Лист1"]
    rows = []
    for line in range(3, sheet.max_row + 1):
        name = sheet.cell(line, 1).value
        if not name:
            continue
        rows.append({"Товар": str(name).strip(),
                     "Остаток, шт": sheet.cell(line, 3).value or 0,
                     "Себестоимость, руб": sheet.cell(line, 4).value,
                     "Срок годности": sheet.cell(line, 5).value,
                     "Просил покупатель, шт": sheet.cell(line, 6).value or 0})
    table = pd.DataFrame(rows)
    for column in ("Остаток, шт", "Себестоимость, руб", "Просил покупатель, шт"):
        table[column] = pd.to_numeric(table[column], errors="coerce")
    return table


stock = read_stock(STOCK)
plan = pd.read_excel(PLAN, header=0)
plan = plan[plan["Наименование"].notna()].reset_index(drop=True)
plan["Приход, шт"] = (pd.to_numeric(plan["Машина"], errors="coerce").fillna(0)
                      + pd.to_numeric(plan["Контейнер"], errors="coerce").fillna(0))

def added(name):
    for key in ADDITIONS:
        if key.lower() in str(name).lower():
            return True
    return False


picked = (stock["Остаток, шт"] >= MIN_STOCK) | stock["Товар"].map(added)
order = stock[picked].copy()
pairs = name_match.match(order["Товар"], plan["Наименование"])
order["Приход, шт"] = [plan.loc[pairs[index], "Приход, шт"] if index in pairs else 0
                       for index in order.index]

def give(name):
    for source in (OVERRIDES, ADDITIONS):
        for key, quantity in source.items():
            if key.lower() in str(name).lower():
                return quantity
    return GIVE


order["Отгружаем, шт"] = order["Товар"].map(give)
order["Сумма, руб"] = (order["Отгружаем, шт"] * order["Себестоимость, руб"]).round(0)
order["Останется на складе, шт"] = order["Остаток, шт"] - order["Отгружаем, шт"]
order["С учетом прихода, шт"] = order["Останется на складе, шт"] + order["Приход, шт"]
order = order.sort_values("Сумма, руб", ascending=False)

columns = ["Товар", "Остаток, шт", "Отгружаем, шт", "Себестоимость, руб", "Сумма, руб",
           "Останется на складе, шт", "Приход, шт", "С учетом прихода, шт",
           "Просил покупатель, шт", "Срок годности"]
order = order[columns]

footer = {column: None for column in columns}
footer.update({"Товар": "ИТОГО",
               "Остаток, шт": order["Остаток, шт"].sum(),
               "Отгружаем, шт": order["Отгружаем, шт"].sum(),
               "Сумма, руб": order["Сумма, руб"].sum(),
               "Останется на складе, шт": order["Останется на складе, шт"].sum(),
               "Приход, шт": order["Приход, шт"].sum(),
               "С учетом прихода, шт": order["С учетом прихода, шт"].sum(),
               "Просил покупатель, шт": order["Просил покупатель, шт"].sum()})
order = pd.concat([order, pd.DataFrame([footer])], ignore_index=True)

# Что было в его заявке и не попало в новый заказ.
dropped = stock[(stock["Просил покупатель, шт"] > 0) & (stock["Остаток, шт"] < MIN_STOCK)].copy()
dropped["Сумма по его заявке, руб"] = (dropped["Просил покупатель, шт"]
                                       * dropped["Себестоимость, руб"]).round(0)
dropped = dropped[["Товар", "Остаток, шт", "Просил покупатель, шт",
                   "Себестоимость, руб", "Сумма по его заявке, руб"]]
dropped = dropped.sort_values("Сумма по его заявке, руб", ascending=False)
tail = {column: None for column in dropped.columns}
tail.update({"Товар": "ИТОГО",
             "Просил покупатель, шт": dropped["Просил покупатель, шт"].sum(),
             "Сумма по его заявке, руб": dropped["Сумма по его заявке, руб"].sum()})
dropped = pd.concat([dropped, pd.DataFrame([tail])], ignore_index=True)

asked_total = (stock["Просил покупатель, шт"] * stock["Себестоимость, руб"]).sum()
summary = pd.DataFrame([
    {"Показатель": "СУММА ЗАКАЗА, РУБ", "Значение": int(order["Сумма, руб"].iloc[-1])},
    {"Показатель": "Позиций в заказе", "Значение": len(order) - 1},
    {"Показатель": "Отгружаем, шт", "Значение": int(order["Отгружаем, шт"].iloc[-1])},
    {"Показатель": "Отбор: остаток на складе не меньше, шт", "Значение": MIN_STOCK},
    {"Показатель": "По каждой позиции даем, шт", "Значение": GIVE},
    {"Показатель": "Он просил, руб", "Значение": int(asked_total)},
    {"Показатель": "Разница с его заявкой, руб",
     "Значение": int(order["Сумма, руб"].iloc[-1] - asked_total)},
])

with pd.ExcelWriter(OUT) as writer:
    summary.to_excel(writer, sheet_name="ИТОГО", index=False)
    order.to_excel(writer, sheet_name="ЗАКАЗ", index=False)
    dropped.to_excel(writer, sheet_name="НЕ ВОШЛО", index=False)

    for name in writer.book.sheetnames:
        sheet = writer.book[name]
        sheet.freeze_panes = "B2"
        titles = [str(c.value or "") for c in sheet[1]]
        for index, title in enumerate(titles, start=1):
            cell = sheet.cell(row=1, column=index)
            cell.font = Font(bold=True)
            cell.fill = HEADER_FILL
            cell.alignment = Alignment(wrap_text=True, vertical="center")
            sheet.column_dimensions[get_column_letter(index)].width = (
                70 if title == "Товар" else 38 if title == "Показатель" else 14)
            if "руб" in title or "шт" in title:
                for row in sheet.iter_rows(min_row=2, min_col=index, max_col=index):
                    row[0].number_format = "# ##0"
        last = sheet.cell(row=sheet.max_row, column=1).value
        if isinstance(last, str) and last.startswith("ИТОГО"):
            for cell in sheet[sheet.max_row]:
                cell.font = Font(bold=True)
                cell.fill = TOTAL_FILL
        if name == "ИТОГО":
            for cell in sheet[2]:
                cell.font = Font(bold=True, size=14, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="C00000")
            sheet.cell(row=2, column=2).number_format = "# ##0 ₽"
            sheet.row_dimensions[2].height = 24

print(summary.to_string(index=False))
print()
print(order[["Товар", "Остаток, шт", "Отгружаем, шт", "Сумма, руб",
             "Останется на складе, шт", "Приход, шт"]].to_string(index=False))
print(f"\nФайл: {OUT}")
