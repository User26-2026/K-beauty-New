"""Заявка в Корею: запрос цен по всему нашему ассортименту.

Задача не закупить, а узнать текущие цены. Поэтому берем все позиции со
склада и просим количество пропорционально остатку — так заявка выглядит
как настоящий заказ, а не как список на прайс.

Количество: доля от остатка, округленная до полусотни, с потолком, чтобы
по крупным позициям не выходили десятки тысяч штук.

Штрихкод и бренд подтягиваем из сводной таблицы прайсов: без штрихкода
кореец не поймет, о какой именно фасовке речь.

Запуск:
    python3 tools/build_rfq.py --share 10 --cap 1000
"""

import argparse
import re

import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import name_match

# Форма заказа: колонки Наименование, Остаток, себестоимость, срок, заказ.
STOCK = "data/stock_costs/Остатки_31.08.2026_форма_заказа.xlsx"
PRICES = "outputs/prices_normalized.xlsx"
OUT = "outputs/Заявка в Корею.xlsx"

# По этим брендам спрашиваем весь остаток, а не долю: цены по ним нужны
# точные, под реальный объем закупки.
FULL_BRANDS = ("PETITFEE", "MANYO", "MA:NYO")

HEADER_FILL = PatternFill("solid", fgColor="DDEBF7")
ASK_FILL = PatternFill("solid", fgColor="FFF2CC")
TOTAL_FILL = PatternFill("solid", fgColor="C6EFCE")


def read_stock(path):
    sheet = openpyxl.load_workbook(path, data_only=True)["Лист1"]
    rows = []
    for line in range(3, sheet.max_row + 1):
        name = sheet.cell(line, 1).value
        if not name:
            continue
        quantity = sheet.cell(line, 2).value
        if isinstance(quantity, str):
            quantity = quantity.replace("\xa0", "").replace(" ", "")
        rows.append({"Товар": str(name).strip(),
                     "Остаток, шт": pd.to_numeric(quantity, errors="coerce"),
                     "Себестоимость, руб": sheet.cell(line, 3).value})
    return pd.DataFrame(rows)


def korean_prices(path):
    """Корейские прайсы: нужны штрихкод, бренд и последняя известная цена."""
    df = pd.read_excel(path, dtype={"Штрихкод": str})
    df = df[(df["Страна"] == "KR") & df["Закупка, KRW"].notna()]
    df = df[df["Штрихкод"].fillna("").str.len() >= 8]
    best = df.sort_values("Закупка, KRW").drop_duplicates("Название EN")
    return best[["Название EN", "Название RU", "Бренд", "Штрихкод",
                 "Закупка, KRW", "Поставщик"]].reset_index(drop=True)


def brand_from_name(name):
    """Бренд из начала названия — для позиций, которых нет в корейских прайсах."""
    head = re.split(r"[\[/,]| - ", str(name))[0].strip()
    words = head.split()
    return words[0].upper() if words else "БЕЗ БРЕНДА"


def latin(name):
    """Английская часть названия: русский хвост корейцу не нужен."""
    head = re.split(r"[/|]", str(name))[0]
    head = re.sub(r"[а-яё]", "", head, flags=re.IGNORECASE)
    return re.sub(r"\s+", " ", head).strip(" -")


def main(share, cap, step, floor):
    stock = read_stock(STOCK)
    prices = korean_prices(PRICES)

    # Пару подтверждаем брендом: без этого крем ENOUGH сводится с кремом
    # PYUNKANG YUL просто потому, что слова в названиях похожи.
    def same_brand(left, right):
        brand = str(prices.loc[right, "Бренд"] or "").upper()
        return bool(brand) and brand.replace(" ", "") in \
            str(stock.loc[left, "Товар"]).upper().replace(" ", "")

    pairs = name_match.match(stock["Товар"], prices["Название EN"],
                             loose=True, guard=same_brand)
    # Проверяем каждую пару брендом: он есть в обоих названиях или пары нет.
    pairs = {left: right for left, right in pairs.items() if same_brand(left, right)}
    for column in ("Бренд", "Штрихкод", "Закупка, KRW", "Поставщик"):
        stock[column] = [prices.loc[pairs[index], column] if index in pairs else None
                         for index in stock.index]

    quantity = (stock["Остаток, шт"] * share / 100).clip(upper=cap)
    quantity = (quantity / step).round() * step
    quantity = quantity.clip(lower=floor).fillna(floor)
    full = stock["Товар"].str.upper().str.startswith(FULL_BRANDS)
    quantity[full] = stock.loc[full, "Остаток, шт"]
    stock["Запрашиваем, шт"] = quantity.fillna(floor).astype(int)
    stock["Бренд"] = stock["Бренд"].fillna(stock["Товар"].map(brand_from_name))
    stock["Название для заявки"] = stock["Товар"].map(latin)
    # Штрихкод — текст: иначе Excel покажет его как 8,8096E+12.
    stock["Штрихкод"] = stock["Штрихкод"].fillna("").astype(str).str.replace(r"\.0$", "", regex=True)
    stock["Цена, KRW"] = None
    stock["Срок поставки"] = None
    stock["Комментарий поставщика"] = None

    columns = ["Бренд", "Название для заявки", "Штрихкод", "Запрашиваем, шт",
               "Цена, KRW", "Срок поставки", "Комментарий поставщика",
               "Остаток, шт", "Товар", "Последняя известная цена, KRW", "Где видели"]
    stock = stock.rename(columns={"Закупка, KRW": "Последняя известная цена, KRW",
                                  "Поставщик": "Где видели"})
    order = stock.sort_values(["Бренд", "Название для заявки"], na_position="last")[columns]
    order.insert(0, "№", range(1, len(order) + 1))

    footer = {column: None for column in order.columns}
    footer.update({"Бренд": "ИТОГО",
                   "Запрашиваем, шт": order["Запрашиваем, шт"].sum(),
                   "Остаток, шт": order["Остаток, шт"].sum()})
    order = pd.concat([order, pd.DataFrame([footer])], ignore_index=True)

    with pd.ExcelWriter(OUT) as writer:
        order.to_excel(writer, sheet_name="ЗАЯВКА", index=False)
        sheet = writer.book["ЗАЯВКА"]
        sheet.freeze_panes = "B2"
        titles = [str(c.value or "") for c in sheet[1]]
        for index, title in enumerate(titles, start=1):
            cell = sheet.cell(row=1, column=index)
            cell.font = Font(bold=True)
            cell.fill = ASK_FILL if title in ("Цена, KRW", "Срок поставки",
                                              "Комментарий поставщика") else HEADER_FILL
            cell.alignment = Alignment(wrap_text=True, vertical="center")
            sheet.column_dimensions[get_column_letter(index)].width = (
                58 if title in ("Название для заявки", "Товар") else
                22 if title in ("Бренд", "Комментарий поставщика") else
                16 if title == "Штрихкод" else 13)
            if "шт" in title or "KRW" in title:
                for row in sheet.iter_rows(min_row=2, min_col=index, max_col=index):
                    row[0].number_format = "# ##0"
        for cell in sheet[sheet.max_row]:
            cell.font = Font(bold=True)
            cell.fill = TOTAL_FILL

    known = int((order["Штрихкод"].fillna("").astype(str).str.len() >= 8).sum())
    print(f"Позиций в заявке: {len(order) - 1}   со штрихкодом: {known}")
    print(f"Запрашиваем всего: {int(order['Запрашиваем, шт'].iloc[-1])} шт "
          f"при остатке {int(order['Остаток, шт'].iloc[-1])} шт")
    print(f"Доля от остатка: {share}%, потолок {cap} шт, минимум {floor} шт")
    print(f"\nФайл: {OUT}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Заявка на запрос цен в Корею")
    parser.add_argument("--share", type=float, default=10, help="доля от остатка, %%")
    parser.add_argument("--cap", type=int, default=1000, help="потолок по позиции, шт")
    parser.add_argument("--step", type=int, default=50, help="округление, шт")
    parser.add_argument("--floor", type=int, default=100, help="минимум по позиции, шт")
    args = parser.parse_args()
    main(args.share, args.cap, args.step, args.floor)
