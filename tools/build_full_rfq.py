"""Заявка на весь ассортимент: склад плюс товар в пути.

Собираем один список из трех источников — остатки, контейнер и машина, —
и по нему запрашиваем цены у поставщиков. Штрихкод берем из инвойса, где
он есть, остальное ищем в сводной таблице прайсов по бренду и названию.

Количество — годовая потребность: продажи месяца, разложенные по сезону
на двенадцать месяцев. Где продаж нет, ставим то, что лежит на складе:
это единственная известная величина спроса.

Запуск:
    python3 tools/build_full_rfq.py <остатки.xlsx> --container <инвойс.xlsx>
"""

import argparse
import os
import re
import sys

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import add_barcodes
import brand_names
import name_match
from add_sales_price import read_stock, sales_by_stock
from offer_to_customer import need_for, read_net
from stock_with_incoming import read_container, truck

TARGET = "outputs/Заявка на весь ассортимент.xlsx"
PRICES = "outputs/prices_normalized.xlsx"
KEEP = 12                # на сколько месяцев считаем потребность
STEP = 100               # округление количества
LOW, HIGH = 200, 500     # рамки количества в заявке

HEAD = PatternFill("solid", fgColor="1F3864")
ASK = PatternFill("solid", fgColor="FFF2CC")
TOTAL = PatternFill("solid", fgColor="E2EFDA")
BRAND = PatternFill("solid", fgColor="DDEBF7")
WHITE = Font(color="FFFFFF", bold=True)

SHOWN = ["№", "Бренд", "Штрихкод", "Товар", "Количество, шт", "Цена",
         "Валюта", "Срок поставки", "Комментарий поставщика"]
WIDTHS = [5, 18, 16, 74, 15, 12, 10, 14, 26]
# Бренд в названии стоит первым словом, но у части позиций он из двух слов.
SERVICE = re.compile(r"container|loading|freight|costs?\b", re.IGNORECASE)
# Позиции, которых нет ни в остатках, ни в приходе, но они нужны в заявке.
# Пары «штрихкод — название» сверены по прайсам поставщиков.
ADD_POSITIONS = [
    ("ETUDE", "8809820683118",
     "ETUDE - Baking Powder Crunch Pore Scrub [200ml] / Очищающий скраб для лица с содой"),
    ("ETUDE", "8809668028089",
     "ETUDE - Baking Powder Pore Cleansing Foam [160ml] / Пенка для умывания с содой"),
    ("ETUDE", "8809668028041",
     "ETUDE - Baking Powder BB Deep Cleansing Foam [160ml] / "
     "Пенка с содой для глубокого очищения и снятия макияжа"),
    ("ETUDE", "8809820692707",
     "ETUDE - Baking Powder Crunch Pore Scrub [24*7ml] / Очищающий скраб для лица с содой"),
    ("ETUDE", "8809668028058",
     "ETUDE - Baking Powder BB Deep Cleansing Foam [30ml] / "
     "Пенка с содой для глубокого очищения"),
]
TWO_WORDS = {"ROUND", "SOME", "THE", "I'M", "DR."}
# Один бренд пишут по-разному, в заявке он должен быть один.
SAME_BRAND = {
    "ROUND": "ROUND LAB", "ROUND LAB": "ROUND LAB", "1025": "ROUND LAB",
    "VT": "VT COSMETICS", "VT COSMETICS": "VT COSMETICS",
    "SOME BY": "SOME BY MI", "SOME BY MI": "SOME BY MI",
    "MA:NYO": "MANYO", "MANYO": "MANYO",
    "DR. ALTHEA": "DR.ALTHEA", "DR.ALTHEA": "DR.ALTHEA",
}


def brand_of(name):
    parts = re.split(r"[\s\-_,]+", str(name).strip())
    if not parts:
        return ""
    first = parts[0].upper()
    if first in TWO_WORDS and len(parts) > 1:
        first = f"{first} {parts[1].upper()}"
    return SAME_BRAND.get(first, first)


def same(brand):
    return SAME_BRAND.get(str(brand).strip().upper(), str(brand).strip().upper())


def read_invoice(path):
    """Инвойс контейнера: бренд, название, штрихкод и количество одной
    выборкой — иначе колонки разъезжаются при сбросе индекса."""
    table = pd.read_excel(path, sheet_name=0, header=None).iloc[3:]
    table = table[[1, 2, 3, 4]]
    table.columns = ["Бренд", "Товар", "Штрихкод", "Кол-во"]
    table = table[table["Товар"].notna()]
    table["Кол-во"] = pd.to_numeric(table["Кол-во"], errors="coerce")
    table = table.dropna(subset=["Кол-во"])
    table["Штрихкод"] = (pd.to_numeric(table["Штрихкод"], errors="coerce")
                         .astype("Int64").astype(str)
                         .str.replace("<NA>", "", regex=False))
    table["Бренд"] = table["Бренд"].fillna("").astype(str).str.strip()
    # В инвойсе есть служебные строки — погрузка контейнера и фрахт.
    table = table[~table["Товар"].astype(str).str.contains(SERVICE)]
    return table.reset_index(drop=True)


def whole_brand(brand):
    """Весь каталог бренда из прайсов поставщиков.

    Берем по одной строке на штрихкод и самое подробное название из
    имеющихся: у одного поставщика оно урезано до сорока знаков, у
    другого записано целиком.
    """
    table = pd.read_excel(PRICES, dtype={"Штрихкод": str})
    table = table[table["Страна"] == "KR"].copy()
    table["Бренд в прайсе"] = table["Бренд"]
    table["Бренд"] = brand_names.resolve(table).fillna("")
    mark = re.sub(r"[^A-Z0-9]", "", brand.upper())
    table = table[table["Бренд"].astype(str).str.upper()
                  .str.replace(r"[^A-Z0-9]", "", regex=True) == mark]
    table = table[table["Штрихкод"].notna() & (table["Штрихкод"].str.len() >= 8)]
    if table.empty:
        return pd.DataFrame(columns=["Товар", "Остаток, шт", "Штрихкод",
                                     "Бренд", "Едет, шт"])
    table["Длина"] = table["Название EN"].astype(str).str.len()
    best = table.sort_values("Длина", ascending=False).drop_duplicates("Штрихкод")
    names = []
    for name, volume in zip(best["Название EN"], best["Объем"]):
        # Сокращение бренда в начале названия разворачиваем полностью.
        clean = re.sub(r"^ET\.\s*", "", str(name)).strip()
        if pd.notna(volume) and str(volume) not in clean:
            clean = f"{clean} [{volume}]"
        names.append(f"{brand} {clean}")
    return pd.DataFrame({"Товар": names, "Остаток, шт": 0.0,
                         "Штрихкод": best["Штрихкод"].values,
                         "Бренд": brand, "Едет, шт": 0.0})


def collect(stock_path, container_path):
    """Один список позиций из остатков, контейнера и машины."""
    stock = read_stock(stock_path)[["Наименование", "Остаток, шт"]]
    stock = stock.rename(columns={"Наименование": "Товар"})
    stock["Штрихкод"] = ""
    stock["Бренд"] = stock["Товар"].map(brand_of)
    stock["Едет, шт"] = 0.0

    coming = read_invoice(container_path)
    coming = pd.concat([coming, truck().assign(Штрихкод="", Бренд="")], ignore_index=True)

    pairs = name_match.to_one(coming["Товар"], stock["Товар"])
    extra = []
    for index, item in coming.iterrows():
        if index in pairs:
            target = pairs[index]
            stock.at[target, "Едет, шт"] += item["Кол-во"]
            if not stock.at[target, "Штрихкод"] and item["Штрихкод"]:
                stock.at[target, "Штрихкод"] = item["Штрихкод"]
            if item["Бренд"]:
                stock.at[target, "Бренд"] = item["Бренд"]
            continue
        extra.append({"Товар": item["Товар"], "Остаток, шт": 0.0,
                      "Штрихкод": item["Штрихкод"],
                      "Бренд": item["Бренд"] or brand_of(item["Товар"]),
                      "Едет, шт": item["Кол-во"]})
    return pd.concat([stock, pd.DataFrame(extra)], ignore_index=True)


# Одна и та же тональная пишется как "#13" и как "13 тон", а название
# идет то с приставкой Premium, то без нее и с русским хвостом.
NOISE = {"PREMIUM", "THE", "AND", "FOR", "NEW", "ML", "G", "GR", "EA", "PCS"}
TONE = re.compile(r"#\s*(\d{1,3})|\b(\d{1,3})\s*тон", re.IGNORECASE)


def tone_of(name):
    found = TONE.search(str(name))
    return next((part for part in found.groups() if part), "") if found else ""


def key_words(name):
    """Значимые латинские слова названия без тона, фасовки и приставок."""
    text = re.sub(r"\((\d)\)?\s*in\s*(\d)", r"\1in\2", str(name), flags=re.I)
    text = re.sub(r"(\d)\s*in\s*(\d)", r"\1in\2", text, flags=re.I)
    text = re.sub(r"\[[^\]]*\]", " ", text)          # фасовка в скобках
    text = text.split("/")[0]                       # русский хвост
    parts = re.split(r"[^A-Za-z0-9]+", text.upper())
    skip = {tone_of(name)}
    return {part for part in parts
            if part and part not in NOISE and part not in skip
            and not re.fullmatch(r"\d+(ML|G|GR|EA|PCS)?", part)}


def same_item(one, two):
    """Один и тот же товар: совпал тон и совпал набор слов.

    Вложенности набора мало: "collagen moisture foundation" входит в
    "collagen whitening moisture foundation", а это разные тональные.
    """
    if tone_of(one) != tone_of(two):
        return False
    left, right = key_words(one), key_words(two)
    return bool(left) and left == right


def read_needs(path):
    """Заявка покупателя: название, потребность в месяц, проходная цена.

    Строки без количества — заголовки разделов («Пудра:»), их пропускаем.
    """
    table = pd.read_excel(path, sheet_name=0, header=None).iloc[2:]
    table = table[[0, 1]]
    table.columns = ["Товар", "Нужно в месяц, шт"]
    table = table[table["Товар"].notna()]
    table["Нужно в месяц, шт"] = pd.to_numeric(table["Нужно в месяц, шт"],
                                               errors="coerce")
    table = table.dropna(subset=["Нужно в месяц, шт"])
    table["Товар"] = table["Товар"].astype(str).str.strip()
    # В заявке покупателя бренд местами набран с опечаткой.
    table["Товар"] = table["Товар"].str.replace(r"^NOUGH\b", "ENOUGH", regex=True)
    table["Бренд"] = table["Товар"].map(brand_of)
    return table.reset_index(drop=True)


def add_needs(table, needs):
    """Позиции из заявки покупателя, которых у нас еще нет."""
    rows = []
    for _, item in needs.iterrows():
        brand = item["Бренд"]
        same = table[table["Бренд"] == brand]["Товар"]
        if any(same_item(item["Товар"], other) for other in same):
            continue
        rows.append({"Товар": item["Товар"], "Остаток, шт": 0.0, "Штрихкод": "",
                     "Бренд": brand, "Едет, шт": 0.0})
    print(f"Из заявки покупателя добавлено: {len(rows)} из {len(needs)} позиций")
    return pd.concat([table, pd.DataFrame(rows)], ignore_index=True) if rows else table


def add_positions(table):
    """Дописываем позиции из списка.

    Если такая позиция уже есть по названию, новую строку не заводим, а
    проставляем ей штрихкод: в остатках он не хранится.
    """
    known = set(table["Штрихкод"]) - {""}
    by_name = {name_match.normal(name): index
               for index, name in table["Товар"].items()}
    rows = []
    for brand, code, name in ADD_POSITIONS:
        if code in known:
            continue
        found = by_name.get(name_match.normal(name))
        if found is not None:
            table.at[found, "Штрихкод"] = code
            continue
        rows.append({"Товар": name, "Остаток, шт": 0.0, "Штрихкод": code,
                     "Бренд": brand, "Едет, шт": 0.0})
    return pd.concat([table, pd.DataFrame(rows)], ignore_index=True) if rows else table


def add_brands(table, brands):
    """Добавляем весь каталог названных брендов, без своих дублей."""
    for brand in brands:
        catalog = whole_brand(brand)
        if catalog.empty:
            print(f"В прайсах нет бренда {brand}")
            continue
        known = set(table["Штрихкод"]) - {""}
        catalog = catalog[~catalog["Штрихкод"].isin(known)]
        table = pd.concat([table, catalog], ignore_index=True)
    return table


def quantities(table, sales_path):
    """Количество для запроса цены: годовая потребность в рамках 200-500.

    Поставщику нужен объем, а не точная потребность: по нему он назовет
    цену. Поэтому потребность округляем до сотен и зажимаем в рамки —
    иначе по одной позиции запрос был бы на 30 тысяч штук, а по другой на
    десяток, и цены вышли бы несравнимыми.
    """
    sales = read_net(sales_path).rename(columns={"К нам, руб": "Выручка, руб"})
    found = sales_by_stock(table.rename(columns={"Товар": "Наименование"}), sales)
    need = []
    for monthly, rest, coming in zip(found["Продано, шт"], table["Остаток, шт"],
                                     table["Едет, шт"]):
        value = need_for(monthly, KEEP) if monthly else rest + coming
        value = int(round(value / STEP)) * STEP
        need.append(min(HIGH, max(LOW, value)))
    return need


def barcodes(table):
    """Чего не было в инвойсе, ищем в прайсах по бренду и названию.

    Код из инвойса — правда, найденный по названию — предположение.
    Поэтому один и тот же найденный код второй раз не ставим: у пэдов и
    патчей названия почти совпадают, а фасовка разная, и код уехал бы не
    на ту позицию. Пустую ячейку заполнит поставщик.
    """
    prices = add_barcodes.korean_prices(PRICES)
    used = {code for code in table["Штрихкод"] if code}
    filled = []
    for code, name, brand in zip(table["Штрихкод"], table["Товар"], table["Бренд"]):
        if code:
            filled.append(code)
            continue
        found = add_barcodes.candidates(name, brand, prices)
        best = max(found)[1] if found else None
        guess = prices.at[best, "Штрихкод"] if best is not None else ""
        if guess and guess in used:
            guess = ""
        if guess:
            used.add(guess)
        filled.append(guess)
    return filled


def main(stock_path, container_path, sales_path, target, brands, needs_path):
    table = add_positions(collect(stock_path, container_path))
    if needs_path:
        table = add_needs(table, read_needs(needs_path))
    if brands:
        table = add_brands(table, brands)
    table["Количество, шт"] = quantities(table, sales_path)
    table["Штрихкод"] = barcodes(table)
    table["Бренд"] = table["Бренд"].map(same)
    table = table.sort_values(["Бренд", "Товар"]).reset_index(drop=True)
    table["№"] = range(1, len(table) + 1)
    for column in ("Цена", "Валюта", "Срок поставки", "Комментарий поставщика"):
        table[column] = None

    book = Workbook()
    ws = book.active
    ws.title = "ЗАЯВКА"
    ws.append(SHOWN)
    for index, cell in enumerate(ws[1], start=1):
        cell.fill, cell.font = HEAD, WHITE
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    brand = None
    for row in table[SHOWN].where(table.notna(), None).values.tolist():
        if row[1] != brand:
            # Заголовок бренда: список общий, но читается по брендам.
            brand = row[1]
            ws.append([None, brand])
            for cell in ws[ws.max_row]:
                cell.fill, cell.font = BRAND, Font(bold=True)
        ws.append(row)
        # Желтым — то, что заполняет поставщик.
        for cell in ws[ws.max_row][5:]:
            cell.fill = ASK
    ws.append([None, "ИТОГО", None, f"позиций: {len(table)}",
               int(table["Количество, шт"].sum()), None, None, None, None])
    for cell in ws[ws.max_row]:
        cell.fill, cell.font = TOTAL, Font(bold=True)
    for index, width in enumerate(WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width
    for cell in ws["E"][1:]:
        cell.number_format = "# ##0"
    ws.freeze_panes = "B2"

    свод = book.create_sheet("ПО БРЕНДАМ")
    свод.append(["Бренд", "Позиций", "Количество, шт"])
    for cell in свод[1]:
        cell.fill, cell.font = HEAD, WHITE
    for name, part in table.groupby("Бренд", sort=True):
        свод.append([name, len(part), int(part["Количество, шт"].sum())])
    свод.append(["ВСЕГО", len(table), int(table["Количество, шт"].sum())])
    for cell in свод[свод.max_row]:
        cell.fill, cell.font = TOTAL, Font(bold=True)
    for index, width in enumerate([24, 10, 16], start=1):
        свод.column_dimensions[get_column_letter(index)].width = width
    for cell in свод["C"][1:]:
        cell.number_format = "# ##0"
    свод.freeze_panes = "A2"

    os.makedirs("outputs", exist_ok=True)
    book.save(target)
    with_code = (table["Штрихкод"].astype(str).str.len() >= 8).sum()
    print(f"Позиций: {len(table)}   штук: {int(table['Количество, шт'].sum()):,}"
          .replace(",", " "))
    print(f"Со штрихкодом: {with_code}   без кода: {len(table) - with_code}")
    print(f"Брендов: {table['Бренд'].nunique()}")
    print(f"Сохранено: {target}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Заявка на весь ассортимент")
    parser.add_argument("source")
    parser.add_argument("--container", required=True)
    parser.add_argument("--sales", default="data/sales/wb_sales_2026-08.xls")
    parser.add_argument("--out", default=TARGET)
    parser.add_argument("--brand", action="append", default=[],
                        help="добавить весь каталог бренда из прайсов")
    parser.add_argument("--needs", help="заявка покупателя с его позициями")
    args = parser.parse_args()
    main(args.source, args.container, args.sales, args.out, args.brand, args.needs)
