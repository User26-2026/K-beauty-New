"""Проставляет штрихкоды в заявку по названию и бренду.

Заявка приходит от поставщика без штрихкодов, а без них не сверить ни
цену, ни фасовку. Ищем товар в сводной таблице прайсов: сначала внутри
того же бренда по совпадению слов, потом по похожести, и обязательно
подтверждаем бренд — иначе крем одного бренда сводится с кремом другого.

Запуск:
    python3 tools/add_barcodes.py "outputs/Заявка №2 в Корею.xlsx"
"""

import argparse
import re

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

PRICES = "outputs/prices_normalized.xlsx"
SCORE = 0.5

HEADER_FILL = PatternFill("solid", fgColor="DDEBF7")
ASK_FILL = PatternFill("solid", fgColor="FFF2CC")
TOTAL_FILL = PatternFill("solid", fgColor="C6EFCE")

NOISE = {"THE", "AND", "FOR", "NEW", "ML", "GR", "EA", "PCS", "G", "SET"}
TONE = re.compile(r"(?:#|No\.?)\s*(\d{1,3}[A-Z]?)", re.IGNORECASE)


def words(name):
    parts = re.split(r"[^A-Za-z0-9]+", str(name).upper())
    return {part for part in parts if part and part not in NOISE and len(part) > 1}


def tone(name):
    found = TONE.findall(str(name))
    return found[-1].upper() if found else ""


def korean_prices(path):
    table = pd.read_excel(path, dtype={"Штрихкод": str})
    table = table[(table["Страна"] == "KR") & table["Штрихкод"].notna()]
    table = table[table["Штрихкод"].str.len() >= 8].copy()
    table["Марка"] = table["Бренд"].fillna("").astype(str).str.upper().str.replace(r"[^A-Z0-9]", "", regex=True)
    table["Слова"] = table["Название EN"].map(words)
    table["Тон"] = table["Название EN"].map(tone)
    return table.reset_index(drop=True)


def candidates(name, brand, prices):
    """Все подходящие строки прайса с оценкой похожести."""
    mark = re.sub(r"[^A-Z0-9]", "", str(brand).upper())
    pool = prices[prices["Марка"] == mark]
    if pool.empty and len(mark) >= 5:
        pool = prices[prices["Марка"].str.startswith(mark[:5])]
    if pool.empty:
        return []
    mine, my_tone = words(name), tone(name)
    found = []
    for index, row in pool.iterrows():
        if my_tone and row["Тон"] and my_tone != row["Тон"]:
            continue
        union = mine | row["Слова"]
        if not union:
            continue
        value = len(mine & row["Слова"]) / len(union)
        if value >= SCORE:
            found.append((value, index))
    return found


def main(path, sheet):
    order = pd.read_excel(path, sheet_name=sheet)
    order = order[order["Товар"].notna() & (order["Бренд"] != "ИТОГО")].copy()
    prices = korean_prices(PRICES)

    # Один штрихкод — один товар: раздаем совпадения по убыванию похожести,
    # иначе пять разных патчей получают код одного и того же.
    scored = []
    for position, row in order.iterrows():
        for value, index in candidates(row["Товар"], row["Бренд"], prices):
            scored.append((value, position, index))
    scored.sort(reverse=True)

    codes, taken = {}, set()
    for _, position, index in scored:
        code = prices.loc[index, "Штрихкод"]
        if position in codes or code in taken:
            continue
        codes[position] = code
        taken.add(code)
    order["Штрихкод"] = [codes.get(position, "") for position in order.index]

    # Прайсы иногда дают разные штрихкоды на один и тот же товар. Молча
    # выбрать один нельзя: пометим, чтобы уточнили у поставщика.
    conflicts = []
    for position, row in order.iterrows():
        near = {prices.loc[index, "Штрихкод"]
                for value, index in candidates(row["Товар"], row["Бренд"], prices)
                if value >= 0.9}
        conflicts.append(" / ".join(sorted(near)) if len(near) > 1 else "")
    order["Коды расходятся"] = conflicts

    columns = ["№", "Бренд", "Штрихкод", "Товар", "Количество, шт",
               "Цена, KRW", "Срок поставки", "Комментарий поставщика",
               "Коды расходятся"]
    order = order[[column for column in columns if column in order.columns]]

    footer = {column: None for column in order.columns}
    footer.update({"Бренд": "ИТОГО", "Количество, шт": order["Количество, шт"].sum()})
    order = pd.concat([order, pd.DataFrame([footer])], ignore_index=True)

    with pd.ExcelWriter(path) as writer:
        order.to_excel(writer, sheet_name=sheet, index=False)
        page = writer.book[sheet]
        page.freeze_panes = "B2"
        titles = [str(c.value or "") for c in page[1]]
        for index, title in enumerate(titles, start=1):
            cell = page.cell(row=1, column=index)
            cell.font = Font(bold=True)
            cell.fill = ASK_FILL if title in ("Цена, KRW", "Срок поставки",
                                              "Комментарий поставщика") else HEADER_FILL
            cell.alignment = Alignment(wrap_text=True, vertical="center")
            page.column_dimensions[get_column_letter(index)].width = (
                62 if title == "Товар" else 22 if title in ("Бренд", "Комментарий поставщика")
                else 16 if title == "Штрихкод" else 14)
            if "шт" in title or "KRW" in title:
                for row in page.iter_rows(min_row=2, min_col=index, max_col=index):
                    row[0].number_format = "# ##0"
        for cell in page[page.max_row]:
            cell.font = Font(bold=True)
            cell.fill = TOTAL_FILL

    found = int((order["Штрихкод"].fillna("").astype(str).str.len() >= 8).sum())
    disputed = int((order["Коды расходятся"].fillna("") != "").sum())
    print(f"Позиций: {len(order) - 1}   штрихкод нашелся у {found}   "
          f"спорных кодов: {disputed}")
    print(f"Файл: {path}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Штрихкоды в заявку")
    parser.add_argument("path")
    parser.add_argument("--sheet", default="ЗАЯВКА 2")
    args = parser.parse_args()
    main(args.path, args.sheet)
