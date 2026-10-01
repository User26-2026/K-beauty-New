"""Заказ покупателю: отдаем излишек, себе оставляем запас на продажи.

Логика простая. По каждой позиции считаем, сколько нужно держать на
складе, чтобы торговать самим: спрос в месяц умножаем на срок запаса. Что
сверх этого — предлагаем оптовому покупателю. Товар в контейнере и машине
идет в наш запас: если приход закрывает потребность, со склада можно
отдать больше.

Позиции без спроса отдаем целиком — держать их незачем. Там, где данных о
спросе нет вовсе, страхуемся и оставляем себе половину остатка. Мелочь
ниже порога в опт не берем: возиться с сотней штук покупателю неинтересно.

Заказ набираем на нужную сумму, начиная с того, чего у нас больше всего
относительно собственных продаж: покупатель забирает излишки, а не то,
чем мы торгуем сами.

Спрос берем по лучшему из двух периодов воронки WB: провал продаж после
пожара на складах не должен завышать наш запас.

Запуск:
    python3 tools/build_customer_order.py --months 4
"""

import argparse

import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import name_match

STOCK = "data/stock_costs/Остатки_31.08.2026.xlsx"
PLAN = "data/shipments/2026-09_plan_raspredeleniya.xlsx"
DEMAND = "outputs/illiquid_check.xlsx"
OUT = "outputs/Заказ покупателю.xlsx"

NO_DEMAND = ("неликвид: смотрят и не берут", "продаж не было")
# Ручные правки поверх расчета: по этим позициям решение уже принято.
OVERRIDES = {
    "FARMSTAY - Black Garlic Nourishing Shampoo": 1000,
    "FARMSTAY - Collagen Water Full Shampoo": 1000,
    "FARMSTAY - Argan Oil Complete Volume Up Shampoo": 1000,
    "CELIMAX THE VITA-A RETINOL SHOT TIGHTENING SERUM": 1000,
    # Линейка DEOPROCE и маска ELIZAVECCA уходят целиком: решено отдельно.
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


def manual(name):
    for key, quantity in OVERRIDES.items():
        if key.lower() in str(name).lower():
            return quantity
    return None


def with_total(table, sums, label="ИТОГО"):
    footer = {column: None for column in table.columns}
    footer[table.columns[0]] = label
    for column in sums:
        footer[column] = table[column].sum()
    return pd.concat([table, pd.DataFrame([footer])], ignore_index=True)


def main(months, min_lot, step, target):
    stock = read_stock(STOCK)

    plan = pd.read_excel(PLAN, header=0)
    plan = plan[plan["Наименование"].notna()].reset_index(drop=True)
    plan["Приход, шт"] = (pd.to_numeric(plan["Машина"], errors="coerce").fillna(0)
                          + pd.to_numeric(plan["Контейнер"], errors="coerce").fillna(0))
    demand = pd.read_excel(DEMAND, sheet_name="ВСЕ ПОЗИЦИИ")
    plan = plan.merge(demand[["Наименование", "Спрос, шт/мес", "Вердикт"]],
                      on="Наименование", how="left")

    pairs = name_match.match(stock["Товар"], plan["Наименование"])
    for column, default in (("Приход, шт", 0), ("Спрос, шт/мес", 0), ("Вердикт", "нет данных")):
        stock[column] = [plan.loc[pairs[index], column] if index in pairs else default
                         for index in stock.index]
    for column in ("Спрос, шт/мес", "Приход, шт"):
        stock[column] = pd.to_numeric(stock[column], errors="coerce").fillna(0)

    # Сколько держим под свои продажи. Товар в пути — тоже наш запас.
    stock["Нужно себе, шт"] = (stock["Спрос, шт/мес"] * months).round(0)
    stock.loc[stock["Вердикт"].isin(NO_DEMAND), "Нужно себе, шт"] = 0
    # Спроса не знаем — не значит, что его нет: половину придерживаем.
    unknown = (stock["Спрос, шт/мес"] == 0) & ~stock["Вердикт"].isin(NO_DEMAND)
    stock.loc[unknown, "Нужно себе, шт"] = (stock.loc[unknown, "Остаток, шт"] / 2).round(0)
    stock["Держим на складе, шт"] = (stock["Нужно себе, шт"] - stock["Приход, шт"]).clip(lower=0)

    free = (stock["Остаток, шт"] - stock["Держим на складе, шт"]).clip(lower=0)
    stock["Отгружаем, шт"] = (free // step * step).astype(int)

    fixed = stock["Товар"].map(manual)
    stock.loc[fixed.notna(), "Отгружаем, шт"] = fixed[fixed.notna()]
    stock["Отгружаем, шт"] = stock[["Отгружаем, шт", "Остаток, шт"]].min(axis=1)
    # Мелкие партии в опт не отдаем, кроме назначенных вручную.
    stock.loc[(stock["Отгружаем, шт"] < min_lot) & fixed.isna(), "Отгружаем, шт"] = 0

    order = stock[stock["Отгружаем, шт"] > 0].copy()
    order["Сумма, руб"] = (order["Отгружаем, шт"] * order["Себестоимость, руб"]).round(0)

    # Набираем заказ на нужную сумму: сначала то, чего у нас больше всего
    # относительно собственных продаж, и то, что нам не нужно совсем.
    order["Излишек, мес"] = (order["Остаток, шт"] /
                             order["Спрос, шт/мес"].replace(0, float("nan")))
    order["Излишек, мес"] = order["Излишек, мес"].fillna(999)
    picked, running = [], 0.0
    forced = order["Товар"].map(manual).notna()
    for index, row in order.assign(_forced=forced).sort_values(
            ["_forced", "Излишек, мес"], ascending=[False, False]).iterrows():
        if row["_forced"]:
            picked.append(index)
            running += row["Сумма, руб"]
            continue
        if running >= target:
            continue
        # Последнюю позицию подрезаем, чтобы заказ не перевалил за сумму.
        if running + row["Сумма, руб"] > target:
            room = (target - running) / row["Себестоимость, руб"]
            quantity = int(room // step * step)
            if quantity < min_lot:
                continue
            order.loc[index, "Отгружаем, шт"] = quantity
            order.loc[index, "Сумма, руб"] = round(quantity * row["Себестоимость, руб"])
        picked.append(index)
        running += order.loc[index, "Сумма, руб"]
    order = order.loc[picked]
    order["Останется на складе, шт"] = order["Остаток, шт"] - order["Отгружаем, шт"]
    order["Останется на складе, шт"] = order["Остаток, шт"] - order["Отгружаем, шт"]
    order["С учетом прихода, шт"] = order["Останется на складе, шт"] + order["Приход, шт"]
    order["Хватит себе на, мес"] = (order["С учетом прихода, шт"] /
                                    order["Спрос, шт/мес"].replace(0, float("nan"))).round(1)
    order = order.sort_values("Сумма, руб", ascending=False)

    columns = ["Товар", "Остаток, шт", "Спрос, шт/мес", "Приход, шт", "Отгружаем, шт",
               "Себестоимость, руб", "Сумма, руб", "Останется на складе, шт",
               "С учетом прихода, шт", "Хватит себе на, мес", "Просил покупатель, шт",
               "Вердикт", "Срок годности"]
    order = with_total(order[columns],
                       ["Остаток, шт", "Приход, шт", "Отгружаем, шт", "Сумма, руб",
                        "Останется на складе, шт", "С учетом прихода, шт",
                        "Просил покупатель, шт"])

    # Что он просил, но мы не отдаем: держим под свои продажи.
    held = stock[(stock["Просил покупатель, шт"] > 0) & (stock["Отгружаем, шт"] == 0)].copy()
    held["Сумма по его заявке, руб"] = (held["Просил покупатель, шт"]
                                        * held["Себестоимость, руб"]).round(0)
    held = with_total(held[["Товар", "Остаток, шт", "Спрос, шт/мес", "Нужно себе, шт",
                            "Просил покупатель, шт", "Сумма по его заявке, руб", "Вердикт"]],
                      ["Просил покупатель, шт", "Сумма по его заявке, руб"])

    asked = (stock["Просил покупатель, шт"] * stock["Себестоимость, руб"]).sum()
    total = int(order["Сумма, руб"].iloc[-1])
    summary = pd.DataFrame([
        {"Показатель": "СУММА ЗАКАЗА, РУБ", "Значение": total},
        {"Показатель": "Позиций в заказе", "Значение": len(order) - 1},
        {"Показатель": "Отгружаем, шт", "Значение": int(order["Отгружаем, шт"].iloc[-1])},
        {"Показатель": "Запас себе, месяцев продаж", "Значение": months},
        {"Показатель": "Минимальная партия в опт, шт", "Значение": min_lot},
        {"Показатель": "Он просил, руб", "Значение": int(asked)},
        {"Показатель": "Разница с его заявкой, руб", "Значение": int(total - asked)},
    ])

    with pd.ExcelWriter(OUT) as writer:
        summary.to_excel(writer, sheet_name="ИТОГО", index=False)
        order.to_excel(writer, sheet_name="ЗАКАЗ", index=False)
        held.to_excel(writer, sheet_name="ДЕРЖИМ СЕБЕ", index=False)

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
                    70 if title == "Товар" else 38 if title == "Показатель" else
                    30 if title == "Вердикт" else 14)
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
    print(order[["Товар", "Остаток, шт", "Спрос, шт/мес", "Отгружаем, шт", "Сумма, руб",
                 "С учетом прихода, шт", "Хватит себе на, мес"]].to_string(index=False))
    print(f"\nФайл: {OUT}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Заказ покупателю из излишков склада")
    parser.add_argument("--months", type=float, default=4, help="запас себе, месяцев продаж")
    parser.add_argument("--target", type=float, default=13300000,
                        help="на какую сумму собираем заказ, руб")
    parser.add_argument("--min-lot", type=int, default=300, help="минимальная партия в опт")
    parser.add_argument("--step", type=int, default=100, help="округление партии")
    args = parser.parse_args()
    main(args.months, args.min_lot, args.step, args.target)
