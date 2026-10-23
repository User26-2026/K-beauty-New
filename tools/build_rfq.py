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
# Заявка покупателя с его месячной потребностью: по этим позициям
# количество известно точно, и спрашивать долю от остатка незачем.
MONTHLY = "data/pricing/Ежемесячная потребность.xlsx"
MATCH_SCORE = 0.65
PRICES = "outputs/prices_normalized.xlsx"
OUT = "outputs/Заявка в Корею.xlsx"

# По этим брендам спрашиваем весь остаток, а не долю: цены по ним нужны
# точные, под реальный объем закупки.
FULL_BRANDS = ("PETITFEE", "MANYO", "MA:NYO")
# Добавляем в заявку позиции, которых у нас на складе нет: по этим линейкам
# нужны цены на все тона, а не только на те, что уже стоят на полке.
EXTRA = {"ENOUGH": r"FOUNDATION"}

HEADER_FILL = PatternFill("solid", fgColor="DDEBF7")
ASK_FILL = PatternFill("solid", fgColor="FFF2CC")
TOTAL_FILL = PatternFill("solid", fgColor="C6EFCE")


TONE = re.compile(r"(?:#|\\#)?\b(\d{2}[A-Z]?|N\d{2})\b(?:\s*тон)?", re.IGNORECASE)
SKIP_WORDS = {"ТОН", "SPF", "PA", "ML", "ГР", "КРЕМ", "ENOUGH", "NOUGH"}


def line_key(name):
    """Ключ линейки и тон отдельно: «Ultra X10 Cover Up ... 13 тон» и
    «ENOUGH Ultra X10 Cover up Collagen Foundation #13» — один товар."""
    text = str(name).upper().replace("\\", "")
    tone = ""
    for match in TONE.finditer(text):
        candidate = match.group(1)
        if candidate not in {"15", "50", "10", "3X", "X10"}:
            tone = candidate
    words = {word.strip("#+,.") for word in re.split(r"[^A-ZА-Я0-9#+]+", text)}
    words = {word for word in words
             if word and word not in SKIP_WORDS and not word.isdigit()}
    return frozenset(words), tone


def brand_of(name):
    brand = re.split(r"[ -]", str(name).strip())[0].upper()
    return "ENOUGH" if brand == "NOUGH" else brand


def match_line(name, prices):
    """Ищем позицию в корейских прайсах по бренду, словам линейки и тону."""
    candidates = prices[prices["Марка"] == brand_of(name).replace(" ", "")]
    words, tone = line_key(name)
    best, score = None, 0.0
    for index, row in candidates.iterrows():
        other, other_tone = line_key(row["Чистое"])
        if tone != other_tone or not (words | other):
            continue
        value = len(words & other) / len(words | other)
        if value > score:
            best, score = index, value
    return best if score >= MATCH_SCORE else None


def read_monthly(path):
    """Список месячной потребности: название, количество, проходная цена."""
    table = pd.read_excel(path, header=None, names=["Товар", "Нужно, шт",
                                                   "Проходная цена, руб", "Сумма"])
    table = table[pd.to_numeric(table["Нужно, шт"], errors="coerce").notna()]
    table["Нужно, шт"] = table["Нужно, шт"].astype(int)
    table["Товар"] = table["Товар"].astype(str).str.strip()
    return table.reset_index(drop=True)


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


def latin(name, brand=None):
    """Английская часть названия: русский хвост корейцу не нужен.

    В прайсе Классика тон вынесен и в начало строки, и в конец, а после
    вычистки кириллицы остаются обрывки вроде «3x» и «X10». Поэтому режем
    строку от бренда и убираем повторы слов.
    """
    head = re.split(r"[/|]", str(name))[0].replace("\\", "")
    if brand and brand.upper() in head.upper():
        head = head[head.upper().index(brand.upper()):]
    head = re.sub(r"[а-яё]", "", head, flags=re.IGNORECASE)
    words, seen = [], set()
    for word in head.split():
        key = word.upper().strip(",")
        if key in seen or word == ",":
            continue
        seen.add(key)
        words.append(word)
    return re.sub(r"\s+", " ", " ".join(words)).strip(" -,")


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

    # Дотягиваем недостающие тона по линейкам из EXTRA.
    known = set(stock["Штрихкод"].dropna())
    additions = []
    for brand, pattern in EXTRA.items():
        rows = prices[(prices["Бренд"].fillna("").str.upper() == brand)
                      & prices["Название EN"].str.upper().str.contains(pattern)]
        for _, row in rows.iterrows():
            if row["Штрихкод"] in known:
                continue
            additions.append({"Товар": row["Название EN"], "Остаток, шт": 0,
                              "Бренд": brand, "Штрихкод": row["Штрихкод"],
                              "Запрашиваем, шт": floor,
                              "Закупка, KRW": row["Закупка, KRW"],
                              "Поставщик": row["Поставщик"]})
    if additions:
        stock = pd.concat([stock, pd.DataFrame(additions)], ignore_index=True)
        print(f"Добавлено позиций, которых нет на складе: {len(additions)}")
    # Заявка покупателя: там количество названо прямо, и оно главнее доли.
    monthly = read_monthly(MONTHLY)
    prices = prices.assign(
        Чистое=[latin(name, brand) for name, brand
                in zip(prices["Название EN"], prices["Бренд"])],
        Марка=prices["Бренд"].fillna("").str.upper().str.replace(" ", ""))
    wanted = []
    for _, row in monthly.iterrows():
        match = match_line(row["Товар"], prices)
        wanted.append({
            "Товар": prices.loc[match, "Чистое"] if match is not None else row["Товар"],
            "Бренд": prices.loc[match, "Бренд"] if match is not None else brand_of(row["Товар"]),
            "Штрихкод": prices.loc[match, "Штрихкод"] if match is not None else None,
            "Остаток, шт": 0,
            "Запрашиваем, шт": int(row["Нужно, шт"]),
            "Закупка, KRW": prices.loc[match, "Закупка, KRW"] if match is not None else None,
            "Поставщик": prices.loc[match, "Поставщик"] if match is not None else None,
        })
    wanted = pd.DataFrame(wanted)
    # Две строки заявки могут указывать на один товар: берем большее число.
    with_code = wanted[wanted["Штрихкод"].notna()]
    if not with_code.empty:
        biggest = with_code.groupby("Штрихкод")["Запрашиваем, шт"].transform("max")
        wanted = pd.concat([wanted[wanted["Штрихкод"].isna()],
                            with_code.assign(**{"Запрашиваем, шт": biggest})
                            .drop_duplicates("Штрихкод")], ignore_index=True)

    # Если позиция уже есть в заявке, количество берем из заявки покупателя.
    by_code = {code: index for index, code in stock["Штрихкод"].items()
               if isinstance(code, str) and len(code) >= 8}
    keep = []
    for index, row in wanted.iterrows():
        target = by_code.get(row["Штрихкод"])
        if target is None:
            keep.append(index)
            continue
        # Ставим ровно то количество, которое назвал покупатель.
        stock.loc[target, "Запрашиваем, шт"] = row["Запрашиваем, шт"]
    if keep:
        stock = pd.concat([stock, wanted.loc[keep]], ignore_index=True)
        print(f"Из заявки покупателя добавлено позиций: {len(keep)}")

    stock["Название для заявки"] = [latin(name, brand) for name, brand
                                    in zip(stock["Товар"], stock["Бренд"])]
    # Круглые числа: меньше сотни округляем до сотни, дальше до сотен.
    # 614 штук в заявке выглядят как выгрузка из учета, 600 — как заказ.
    quantity = pd.to_numeric(stock["Запрашиваем, шт"], errors="coerce").fillna(0)
    rounded = ((quantity + 50) // 100 * 100).clip(lower=100)
    stock["Запрашиваем, шт"] = rounded.astype(int)

    # Штрихкод — текст: иначе Excel покажет его как 8,8096E+12.
    stock["Штрихкод"] = stock["Штрихкод"].fillna("").astype(str).str.replace(r"\.0$", "", regex=True)
    stock["Цена, KRW"] = None
    stock["Срок поставки"] = None
    stock["Комментарий поставщика"] = None

    # В файл, который уходит поставщику, лишнего не кладем: только то, что
    # ему нужно, чтобы понять товар и проставить цену.
    columns = ["Бренд", "Название для заявки", "Штрихкод", "Запрашиваем, шт",
               "Цена, KRW", "Срок поставки", "Комментарий поставщика"]
    order = stock.sort_values(["Бренд", "Название для заявки"], na_position="last")[columns]
    order.insert(0, "№", range(1, len(order) + 1))

    footer = {column: None for column in order.columns}
    footer.update({"Бренд": "ИТОГО",
                   "Запрашиваем, шт": order["Запрашиваем, шт"].sum()})
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
    print(f"Запрашиваем всего: {int(order['Запрашиваем, шт'].iloc[-1])} шт")
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
