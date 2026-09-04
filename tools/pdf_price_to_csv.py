"""Прайс из PDF в CSV — для поставщиков, которые шлют выгрузку из 1С.

Korea Global (Бишкек) присылает прайс печатной формой: колонки бренда,
названия, штрихкода и двух цен — в сомах и в долларах. Таблица в PDF
размечена линиями, поэтому строки читаются как есть, без разбора текста
по координатам.

Название бренда стоит отдельной строкой в колонке с картинкой и
относится ко всем товарам ниже, до следующего такого заголовка.

Запуск:
    python3 tools/pdf_price_to_csv.py <файл.pdf> <куда.csv>
"""

import argparse
import csv
import re

import pdfplumber

HEADER_WORDS = ("наименование", "штрихкод", "изображение", "картинка")

# Раскладка колонок отличается у поставщиков, но обе таблицы размечены
# линиями, поэтому достаточно знать, где что стоит.
LAYOUTS = {
    # Korea Global: картинка, название, штрихкод, цена в сомах, цена в долларах.
    "koreaglobal": {"brand": 1, "name": 2, "barcode": 3, "kgs": 4, "usd": 5},
    # Aibeauty: картинка, штрихкод, название, пусто, цена в долларах.
    "aibeauty": {"brand": 1, "barcode": 2, "name": 3, "kgs": None, "usd": 5},
}


def detect_layout(pages):
    """Раскладку узнаем по данным, а не по шапке.

    Продолжение прайса приходит отдельным файлом и начинается сразу со
    строк товара, без заголовков. Поэтому проверяем обе раскладки и берем
    ту, где в колонке штрихкода действительно стоят штрихкоды.
    """
    scores = dict.fromkeys(LAYOUTS, 0)
    for page in pages[:2]:
        for row in page.extract_table() or []:
            cells = [clean(cell) for cell in row]
            if len(cells) < 6:
                continue
            for name, columns in LAYOUTS.items():
                digits = re.sub(r"\D", "", cells[columns["barcode"]])
                if len(digits) >= 12 and to_number(cells[columns["usd"]]):
                    scores[name] += 1
    return max(scores, key=scores.get)


BRAND_IN_NAME = re.compile(r"^\(([^)]{2,40})\)")


def clean(value):
    return re.sub(r"\s+", " ", str(value or "")).strip()


def to_number(value):
    """1 465 и 16,65 — обе записи чисел встречаются в одном прайсе."""
    text = clean(value).replace(" ", "").replace(" ", "").replace(",", ".")
    return text if re.fullmatch(r"-?\d+(\.\d+)?", text) else ""


def rows_from_pdf(path):
    brand = ""
    with pdfplumber.open(path) as pdf:
        columns = LAYOUTS[detect_layout(pdf.pages)]
        for page in pdf.pages:
            for row in page.extract_table() or []:
                cells = [clean(cell) for cell in row]
                if len(cells) < 6:
                    continue
                picture = cells[columns["brand"]]
                name = cells[columns["name"]].replace("\n", " ")
                if any(word in picture.lower() for word in HEADER_WORDS):
                    continue
                price_usd = to_number(cells[columns["usd"]])
                # Строка с одним заполненным полем в колонке картинки — бренд.
                if picture and not name and not price_usd:
                    brand = picture
                    continue
                if not name:
                    continue
                # Aibeauty пишет бренд прямо в названии: «(Anua) Крем ...».
                inside = BRAND_IN_NAME.match(name)
                yield {
                    "Бренд": inside.group(1) if inside else brand,
                    "Наименование": name,
                    "Штрихкод": re.sub(r"\D", "", cells[columns["barcode"]]),
                    "Цена, сом": (to_number(cells[columns["kgs"]])
                                  if columns["kgs"] is not None else ""),
                    "Цена, $": price_usd,
                }


def main(source, target):
    rows = [row for row in rows_from_pdf(source) if row["Цена, $"]]
    with open(target, "w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(rows[0]), delimiter=";")
        writer.writeheader()
        writer.writerows(rows)

    brands = {row["Бренд"] for row in rows if row["Бренд"]}
    without_barcode = sum(1 for row in rows if not row["Штрихкод"])
    print(f"Позиций: {len(rows)}   брендов: {len(brands)}   "
          f"без штрихкода: {without_barcode}")
    print(f"Сохранено: {target}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Прайс из PDF в CSV")
    parser.add_argument("source")
    parser.add_argument("target")
    args = parser.parse_args()
    main(args.source, args.target)
