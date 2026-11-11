"""Прайс SAFIYA по брендам и товарам.

Лист на каждый бренд плюс своды. Основа — персональное предложение нам от
25.08: по нему мы и покупаем, и оно шире прайса. Рядом ставим действующий
B2B-прайс, предложение знакомому от 12.08 и публичную цену: видно, сколько
SAFIYA уступила нам с прайса, лучше ли наши условия, чем у знакомого, и
сколько мы выигрываем к публичной цене.

Публичная цена идет с НДС 22%, прайсы в долларах — без НДС, поэтому к
публичной считаем две разницы: как есть и к цене, очищенной от НДС.

Запуск:
    python3 tools/build_safiya_offer.py
"""

import csv
import os
import re
import sys

import pandas as pd
from openpyxl import Workbook
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from rates import SAFIYA_USD_RUB

ROOT = "data/price_lists/safiya"
B2B = f"{ROOT}/safiya_2026-09-01_b2b.xlsx"
PERSONAL = f"{ROOT}/archive/safiya_2026-08-25_personal.xlsx"
FRIEND = f"{ROOT}/archive/safiya_2026-08-12_znakomomu.xlsx"
PUBLIC = f"{ROOT}/official/safiya_2026-08-04_official.csv"
TARGET = "outputs/SAFIYA_прайс_по_брендам.xlsx"
VAT = 1.22

HEAD = PatternFill("solid", fgColor="1F3864")
TOTAL = PatternFill("solid", fgColor="E2EFDA")
WHITE = Font(color="FFFFFF", bold=True)

COLUMNS = ["Наименование", "Объем", "Штрихкод", "Наша цена, $", "Наша цена, руб",
           "Прайс B2B, $", "Скидка к прайсу, %", "Знакомому 12.08, $",
           "Мы дешевле знакомого, %", "Публичная цена A, руб",
           "Дешевле публичной, %"]


def read_b2b():
    """B2B-прайс: бренд стоит заголовком секции, товары — строками ниже."""
    table = pd.read_excel(B2B, header=0)
    table.columns = ["Код", "Наименование", "Объем", "Цена, $", "Кол-во", "Сумма"]
    goods, brand = [], ""
    for _, row in table.iterrows():
        code = str(row["Код"]).strip()
        if code.isdigit() and len(code) >= 12:
            goods.append({"Бренд": brand, "Штрихкод": code,
                          "Наименование": str(row["Наименование"]).strip(),
                          "Объем": row["Объем"], "Цена B2B, $": float(row["Цена, $"])})
        elif code and code != "nan" and pd.isna(row["Цена, $"]):
            brand = code
    goods = pd.DataFrame(goods)
    return goods.set_index("Штрихкод")[["Цена B2B, $"]]


def read_personal():
    """Предложение нам: бренд стоит в своей колонке у каждой строки."""
    table = pd.read_excel(PERSONAL, skiprows=4, header=None)
    table.columns = ["Бренд", "Наименование", "Штрихкод", "Объем", "Наша цена, $"]
    table = table.dropna(subset=["Наша цена, $", "Наименование"])
    table["Штрихкод"] = table["Штрихкод"].astype(str).str.replace(r"\D", "", regex=True)
    table["Наименование"] = table["Наименование"].astype(str).str.strip()
    # Цены SAFIYA переводим по их курсу, а не по рыночному: разница курсов —
    # это их скрытая наценка, и платим мы именно по их счету.
    table["Наша цена, руб"] = (table["Наша цена, $"] * SAFIYA_USD_RUB).round()
    return table.reset_index(drop=True)


def read_friend():
    """Предложение, которое та же компания дала знакомому в августе."""
    table = pd.read_excel(FRIEND, skiprows=4, header=None)
    table.columns = ["Бренд", "Категория", "Наименование", "Штрихкод",
                     "Объем", "Остаток на 12.08, шт", "Знакомому 12.08, $"]
    table = table.dropna(subset=["Знакомому 12.08, $"])
    table["Штрихкод"] = table["Штрихкод"].astype(str).str.replace(r"\D", "", regex=True)
    return table.set_index("Штрихкод")[["Знакомому 12.08, $"]]


def read_public():
    with open(PUBLIC, encoding="utf-8") as handle:
        return {row["Штрихкод"]: int(row["Цена A, руб"])
                for row in csv.DictReader(handle, delimiter=";")}


def build():
    goods = read_personal()
    public = read_public()
    goods = goods.join(read_b2b(), on="Штрихкод").join(read_friend(), on="Штрихкод")

    # Скидка положительным числом: сколько процентов цены мы не платим.
    goods["Скидка к прайсу, %"] = [
        round(100 - our / base * 100, 1) if pd.notna(base) else ""
        for our, base in zip(goods["Наша цена, $"], goods["Цена B2B, $"])
    ]
    goods["Мы дешевле знакомого, %"] = [
        round(100 - our / base * 100, 1) if pd.notna(base) else ""
        for our, base in zip(goods["Наша цена, $"], goods["Знакомому 12.08, $"])
    ]
    goods["Публичная цена A, руб"] = [public.get(code, "") for code in goods["Штрихкод"]]
    goods["Дешевле публичной, %"] = [
        round(100 - price / base * 100, 1) if base else ""
        for price, base in zip(goods["Наша цена, руб"], goods["Публичная цена A, руб"])
    ]
    goods["Без НДС, %"] = [
        round(100 - price / (base / VAT) * 100, 1) if base else ""
        for price, base in zip(goods["Наша цена, руб"], goods["Публичная цена A, руб"])
    ]
    goods = goods.rename(columns={"Цена B2B, $": "Прайс B2B, $"})
    return goods.fillna({"Прайс B2B, $": "", "Знакомому 12.08, $": ""})


def sheet_name(brand):
    """Имя листа: Excel не берет длиннее 31 знака и служебные символы."""
    return re.sub(r"[\\/*?:\[\]]", "-", brand)[:31]


def write(ws, columns, rows, widths, total=None):
    ws.append(columns)
    for cell in ws[1]:
        cell.fill, cell.font = HEAD, WHITE
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    for row in rows:
        ws.append(row)
    if total:
        ws.append(total)
        for cell in ws[ws.max_row]:
            cell.fill, cell.font = TOTAL, Font(bold=True)
    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width
    ws.freeze_panes = "A2"
    return ws


def money(ws, letters, fmt="# ##0"):
    for letter in letters:
        for cell in ws[letter][1:]:
            cell.number_format = fmt


def scale(ws, letter, rows, low, high, reverse=False):
    if rows < 2:
        return
    colors = ("F8696B", "FFEB84", "63BE7B")
    if reverse:
        colors = colors[::-1]
    ws.conditional_formatting.add(
        f"{letter}2:{letter}{rows}",
        ColorScaleRule(start_type="num", start_value=low, start_color=colors[0],
                       mid_type="num", mid_value=(low + high) / 2, mid_color=colors[1],
                       end_type="num", end_value=high, end_color=colors[2]))


def median(values):
    values = [value for value in values if value != ""]
    return round(pd.Series(values).median(), 1) if values else ""


def main():
    goods = build()
    book = Workbook()
    book.remove(book.active)

    summary = []
    for brand, part in goods.groupby("Бренд", sort=True):
        summary.append([
            brand, len(part),
            int(part["Наша цена, руб"].mean()),
            int(part["Наша цена, руб"].min()),
            int(part["Наша цена, руб"].max()),
            int((part["Скидка к прайсу, %"] != "").sum()),
            median(part["Скидка к прайсу, %"]),
            median(part["Мы дешевле знакомого, %"]),
            int((part["Дешевле публичной, %"] != "").sum()),
            median(part["Дешевле публичной, %"]),
        ])
    итого = write(
        book.create_sheet("ИТОГО"),
        ["Бренд", "Позиций", "Средняя цена, руб", "Мин, руб", "Макс, руб",
         "Есть в прайсе B2B, поз.", "Скидка к прайсу, %",
         "Мы дешевле знакомого, %", "Есть в публичном, поз.",
         "Дешевле публичной, %"],
        summary,
        [24, 9, 15, 11, 11, 14, 13, 14, 14, 14],
        ["ВСЕГО", len(goods), int(goods["Наша цена, руб"].mean()),
         int(goods["Наша цена, руб"].min()), int(goods["Наша цена, руб"].max()),
         int((goods["Скидка к прайсу, %"] != "").sum()),
         median(goods["Скидка к прайсу, %"]),
         median(goods["Мы дешевле знакомого, %"]),
         int((goods["Дешевле публичной, %"] != "").sum()),
         median(goods["Дешевле публичной, %"])],
    )
    money(итого, "CDE")
    scale(итого, "G", итого.max_row - 1, 0, 20)
    scale(итого, "J", итого.max_row - 1, 0, 45)

    скидка = goods[goods["Скидка к прайсу, %"] != ""].sort_values(
        "Скидка к прайсу, %", ascending=False)
    лист = write(
        book.create_sheet("СКИДКА К ПРАЙСУ"),
        ["Бренд", "Наименование", "Объем", "Штрихкод", "Прайс B2B, $",
         "Наша цена, $", "Скидка к прайсу, %"],
        скидка[["Бренд", "Наименование", "Объем", "Штрихкод",
                "Прайс B2B, $", "Наша цена, $", "Скидка к прайсу, %"]].values.tolist(),
        [18, 58, 12, 16, 12, 13, 12],
    )
    money(лист, "EF", "0.00")
    scale(лист, "G", лист.max_row, 0, 20)

    знакомый = goods[goods["Мы дешевле знакомого, %"] != ""].sort_values(
        "Мы дешевле знакомого, %")
    лист = write(
        book.create_sheet("ПРОТИВ ЦЕНЫ ЗНАКОМОГО"),
        ["Бренд", "Наименование", "Объем", "Штрихкод", "Знакомому 12.08, $",
         "Наша цена, $", "Мы дешевле знакомого, %"],
        знакомый[["Бренд", "Наименование", "Объем", "Штрихкод",
                  "Знакомому 12.08, $", "Наша цена, $",
                  "Мы дешевле знакомого, %"]].values.tolist(),
        [18, 58, 12, 16, 13, 13, 14],
    )
    money(лист, "EF", "0.00")
    scale(лист, "G", лист.max_row, -5, 15)

    # Персональное предложение шире прайса: эти позиции есть только в нем.
    новые = goods[goods["Прайс B2B, $"] == ""].sort_values(["Бренд", "Наименование"])
    лист = write(
        book.create_sheet("НЕТ В ПРАЙСЕ B2B"),
        ["Бренд", "Наименование", "Объем", "Штрихкод", "Наша цена, $",
         "Наша цена, руб"],
        новые[["Бренд", "Наименование", "Объем", "Штрихкод",
               "Наша цена, $", "Наша цена, руб"]].values.tolist(),
        [18, 58, 12, 16, 13, 14],
    )
    money(лист, "E", "0.00")
    money(лист, "F")

    сверка = goods[goods["Дешевле публичной, %"] != ""].sort_values(
        "Дешевле публичной, %", ascending=False)
    лист = write(
        book.create_sheet("СКИДКА К ПУБЛИЧНОМУ"),
        ["Бренд", "Наименование", "Объем", "Штрихкод", "Наша цена, руб",
         "Публичная цена A, руб", "Дешевле публичной, %", "Без НДС, %"],
        сверка[["Бренд", "Наименование", "Объем", "Штрихкод", "Наша цена, руб",
                "Публичная цена A, руб", "Дешевле публичной, %",
                "Без НДС, %"]].values.tolist(),
        [18, 58, 12, 16, 13, 15, 14, 12],
    )
    money(лист, "EF")
    scale(лист, "G", лист.max_row, 0, 45)

    for brand, part in goods.groupby("Бренд", sort=True):
        part = part.sort_values("Наименование")
        ws = write(
            book.create_sheet(sheet_name(brand)),
            COLUMNS,
            part[COLUMNS].values.tolist(),
            [58, 12, 16, 12, 13, 12, 13, 13, 14, 15, 14],
            ["ИТОГО", f"позиций: {len(part)}", "",
             round(part["Наша цена, $"].mean(), 2),
             int(part["Наша цена, руб"].mean()), "",
             median(part["Скидка к прайсу, %"]), "",
             median(part["Мы дешевле знакомого, %"]), "",
             median(part["Дешевле публичной, %"])],
        )
        money(ws, "EJ")
        money(ws, "DF", "0.00")
        scale(ws, "G", ws.max_row - 1, 0, 20)
        scale(ws, "K", ws.max_row - 1, 0, 45)

    os.makedirs("outputs", exist_ok=True)
    book.save(TARGET)
    print(f"Брендов: {goods['Бренд'].nunique()}   позиций: {len(goods)}")
    print(f"Скидка к прайсу B2B: {median(goods['Скидка к прайсу, %'])}% "
          f"по {(goods['Скидка к прайсу, %'] != '').sum()} позициям")
    print(f"Против цены знакомого: {median(goods['Мы дешевле знакомого, %'])}% "
          f"по {(goods['Мы дешевле знакомого, %'] != '').sum()} позициям")
    print(f"Нет в прайсе B2B: {(goods['Прайс B2B, $'] == '').sum()} позиций")
    print(f"Дешевле публичного прайса: {median(goods['Дешевле публичной, %'])}% "
          f"по {(goods['Дешевле публичной, %'] != '').sum()} позициям")
    print(f"Сохранено: {TARGET}")


if __name__ == "__main__":
    main()
