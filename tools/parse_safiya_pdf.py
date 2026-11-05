"""Официальный прайс SAFIYA из PDF в CSV.

Прайс печатается таблицей с картинками. Название товара разбито по
вертикали: русская часть стоит строкой выше штрихкода, английская —
строкой ниже, а иногда часть слов попадает прямо в строку со
штрихкодом. Поэтому слова названия собираем не по строкам, а по
близости к строке товара.

Колонки различаем по горизонтали, границы берем из шапки
(«МРЦ*», «РРЦ*», «Цена A», «Цена B», «Цена C»). Числа в PDF идут с
хвостом «-» и отдельным значком рубля, значок отбрасываем.

Цены три: A — заказ от 20 тысяч, B — от 500 тысяч, C — от 2 миллионов.
Все с НДС 22%.

Запуск:
    python3 tools/parse_safiya_pdf.py <файл.pdf> <куда.csv>
"""

import argparse
import csv
import re
from collections import defaultdict

import pdfplumber

BARCODE = re.compile(r"^\d{12,14}$")
VALUE = re.compile(r"^[\d,]+-$")
NAME_LEFT, NAME_RIGHT = 60, 240
VOLUME_LEFT, VOLUME_RIGHT = 230, 282
PRICE_COLUMNS = ("МРЦ", "РРЦ", "Цена A, руб", "Цена B, руб", "Цена C, руб")
# Шапка печатается в две-три строки, привязываемся к первому слову колонки.
HEADERS = (("МРЦ", "МРЦ*"), ("РРЦ", "РРЦ*"), ("Цена A, руб", "A"),
           ("Цена B, руб", "B"), ("Цена C, руб", "C"))


def lines_of(page):
    """Слова страницы, сгруппированные в строки по вертикали."""
    rows = defaultdict(list)
    for word in page.extract_words():
        rows[round(word["top"] / 3)].append(word)
    return [sorted(words, key=lambda word: word["x0"]) for _, words in sorted(rows.items())]


def price_bounds(pages):
    """Левые границы ценовых колонок по шапке прайса."""
    starts = {}
    for page in pages:
        for words in lines_of(page):
            for word in words:
                for column, label in HEADERS:
                    if word["text"] == label and word["x0"] > 300 and column not in starts:
                        starts[column] = word["x0"]
        if len(starts) == len(HEADERS):
            break
    if len(starts) < len(HEADERS):
        raise SystemExit("не нашли шапку с колонками цен")
    order = sorted(starts.items(), key=lambda item: item[1])
    bounds = []
    for index, (column, left) in enumerate(order):
        right = order[index + 1][1] if index + 1 < len(order) else 10_000
        # Число печатается правее подписи колонки, но левее следующей.
        bounds.append((column, left - 16, right - 16))
    return bounds


def value(text):
    return int(re.sub(r"[^\d]", "", text)) if re.sub(r"[^\d]", "", text) else ""


def is_brand(words):
    """Заголовок бренда: одна строка по центру, латиница, без цифр цен."""
    if not words or any(word["x0"] < NAME_RIGHT for word in words):
        return False
    joined = " ".join(word["text"] for word in words)
    return bool(re.fullmatch(r"[A-Z][A-Z0-9&.:\- ]{2,30}", joined))


def parse(path):
    goods = []
    with pdfplumber.open(path) as pdf:
        bounds = price_bounds(pdf.pages)
        brand = ""
        for page in pdf.pages:
            rows = lines_of(page)
            head = next((words[0]["top"] for words in rows
                         if any(word["text"] == "Наименование" for word in words)), 0)
            products = [words for words in rows
                        if any(BARCODE.match(word["text"]) for word in words)]
            if not products:
                continue
            anchors = [min(word["top"] for word in words) for words in products]
            names = defaultdict(list)
            headers = []

            for words in rows:
                if words[0]["top"] <= head or any(BARCODE.match(w["text"]) for w in words):
                    continue
                if is_brand(words):
                    headers.append((words[0]["top"], " ".join(w["text"] for w in words)))
                    continue
                for word in words:
                    if not NAME_LEFT < word["x0"] < NAME_RIGHT:
                        continue
                    nearest = min(range(len(anchors)),
                                  key=lambda index: abs(anchors[index] - word["top"]))
                    if abs(anchors[nearest] - word["top"]) < 40:
                        names[nearest].append(word)

            for index, words in enumerate(products):
                top = anchors[index]
                # Бренд печатается заголовком над своей группой товаров,
                # на следующую страницу заголовок не повторяется.
                above = [name for line, name in headers if line < top]
                brand = above[-1] if above else brand
                own = [word for word in words if NAME_LEFT < word["x0"] < NAME_RIGHT]
                parts = sorted(names.get(index, []) + own,
                               key=lambda word: (round(word["top"] / 3), word["x0"]))
                volume = " ".join(word["text"] for word in words
                                  if VOLUME_LEFT < word["x0"] < VOLUME_RIGHT)
                barcode = next(word["text"] for word in words if BARCODE.match(word["text"]))
                row = {
                    "Бренд": brand,
                    "Наименование": re.sub(r"\s+", " ", " ".join(w["text"] for w in parts)).strip(),
                    "Объем": volume.replace(" / ", "/"),
                    "Штрихкод": barcode,
                }
                for column, left, right in bounds:
                    found = [word["text"] for word in words
                             if VALUE.match(word["text"]) and left <= word["x0"] < right]
                    row[column] = value(found[0]) if found else ""
                goods.append(row)
            brand = headers[-1][1] if headers else brand
    return goods


def main(source, target):
    goods = parse(source)
    columns = ["Бренд", "Наименование", "Объем", "Штрихкод", *PRICE_COLUMNS]
    with open(target, "w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=columns, delimiter=";")
        writer.writeheader()
        writer.writerows(goods)
    priced = sum(1 for row in goods if row["Цена A, руб"])
    print(f"Позиций: {len(goods)}   с ценой A: {priced}   "
          f"брендов: {len({row['Бренд'] for row in goods})}")
    broken = [row for row in goods
              if not row["Наименование"] or not all(row[column] for column in PRICE_COLUMNS)]
    if broken:
        print(f"Разобрано не полностью: {len(broken)}")
        for row in broken[:5]:
            print("   ", row["Штрихкод"], row["Наименование"][:50])
    print(f"Сохранено: {target}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Прайс SAFIYA из PDF")
    parser.add_argument("source")
    parser.add_argument("target")
    args = parser.parse_args()
    main(args.source, args.target)
