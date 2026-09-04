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

HEADER_WORDS = ("наименование", "штрихкод", "изображение")


def clean(value):
    return re.sub(r"\s+", " ", str(value or "")).strip()


def to_number(value):
    """1 465 и 16,65 — обе записи чисел встречаются в одном прайсе."""
    text = clean(value).replace(" ", "").replace(" ", "").replace(",", ".")
    return text if re.fullmatch(r"-?\d+(\.\d+)?", text) else ""


def rows_from_pdf(path):
    brand = ""
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            for row in page.extract_table() or []:
                cells = [clean(cell) for cell in row]
                if len(cells) < 6:
                    continue
                _, picture, name, barcode, price_kgs, price_usd = cells[:6]
                if any(word in picture.lower() for word in HEADER_WORDS):
                    continue
                # Строка с одним заполненным полем в колонке картинки — бренд.
                if picture and not name and not price_usd:
                    brand = picture
                    continue
                if not name:
                    continue
                yield {
                    "Бренд": brand,
                    "Наименование": name.replace("\n", " "),
                    "Штрихкод": re.sub(r"\D", "", barcode),
                    "Цена, сом": to_number(price_kgs),
                    "Цена, $": to_number(price_usd),
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
