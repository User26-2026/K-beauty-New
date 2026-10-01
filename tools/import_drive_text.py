"""Сохранение прайса, вычитанного из Google Диска, в папку поставщика.

Коннектор отдает таблицу одной длинной строкой CSV: ячейки разделены
запятыми, переносов строк нет, а ширина строк в исходной таблице гуляет —
восстановить разбиение по ширине нельзя.

Опорой берем штрихкод: он встречается в каждой товарной строке на одном и
том же месте относительно остальных полей. Смещения полей от штрихкода
задаются на конкретный прайс.

Запуск:
    python3 tools/import_drive_text.py <json> <поставщик> <имя_файла> <раскладка>
"""

import csv
import json
import os
import re
import sys

PRICE_ROOT = "data/price_lists"
BARCODE = re.compile(r"\d{13}")

# Смещения полей относительно ячейки со штрихкодом.
LAYOUTS = {
    "enough": {"Product name": -7, "Volume": -6, "Qty per box": -5,
               "MSRP": -4, "Supply price": -3},
    "masil": {"Product name": -8, "Volume": -6, "Qty per box": -3,
              "Supply price": 1},
}


def records(cells, layout):
    """Собираем строки товаров, отталкиваясь от ячейки со штрихкодом."""
    rows = []
    for index, cell in enumerate(cells):
        match = BARCODE.fullmatch(cell.strip())
        if not match:
            continue
        row = {"Barcode": match.group(0)}
        for field, shift in layout.items():
            spot = index + shift
            row[field] = cells[spot].strip() if 0 <= spot < len(cells) else ""
        rows.append(row)
    return rows


def main(source, supplier, filename, layout_name):
    layout = LAYOUTS.get(layout_name)
    if layout is None:
        raise SystemExit(f"Нет раскладки {layout_name}. Есть: {', '.join(LAYOUTS)}")

    with open(source, encoding="utf-8") as handle:
        text = json.load(handle)["fileContent"]
    if not text.strip():
        raise SystemExit("Ответ пустой — коннектор не смог прочитать этот файл")

    cells = next(csv.reader([text]))
    rows = records(cells, layout)
    if not rows:
        raise SystemExit("Штрихкодов не найдено")

    dest_dir = os.path.join(PRICE_ROOT, supplier)
    os.makedirs(dest_dir, exist_ok=True)
    dest = os.path.join(dest_dir, filename)
    with open(dest, "w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(rows[0]))
        writer.writeheader()
        writer.writerows(rows)

    priced = sum(1 for r in rows if re.search(r"\d", r.get("Supply price", "")))
    print(f"Сохранено: {dest}")
    print(f"  строк: {len(rows)}, с ценой: {priced}")


if __name__ == "__main__":
    if len(sys.argv) != 5:
        sys.exit("Запуск: python3 tools/import_drive_text.py <json> <поставщик> <файл> <раскладка>")
    main(*sys.argv[1:])
