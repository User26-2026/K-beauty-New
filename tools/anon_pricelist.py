"""Прайс без упоминания поставщика.

Корейские прайсы приходят бланком заказа, и в первой ячейке стоит имя
компании: «PAPACOSMETIC CELIMAX PRICE LIST». Такой файл нельзя переслать
дальше, не раскрыв, у кого мы берем товар. Скрипт собирает новый файл с
теми же позициями и ценами, но с нейтральной шапкой, без свойств
документа и без имени поставщика в названии листа и файла.

Берем значения, а не копию листа: в копии остаются скрытые свойства,
колонтитулы и автор файла.

Запуск:
    python3 tools/anon_pricelist.py прайс.xlsx --sheet Sheet2 --brand "ROUND LAB"
"""

import argparse
import os
import re

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

OUT_DIR = "outputs"

HEAD = PatternFill("solid", fgColor="DDEBF7")
TOTAL = PatternFill("solid", fgColor="E2EFDA")
THIN = Side(style="thin", color="B4C6E7")
GRID = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

COLUMNS = ["NO", "Bar Code", "Product Name (KR)", "Product Name (EN)",
           "RETAIL PRICE (KRW)", "SUPPLY PRICE (KRW)", "PCS / OUTBOX",
           "ORDER BOXES", "TOTAL QUANTITIES", "TOTAL SUPPLY PRICE"]
WIDTHS = [5, 17, 34, 56, 15, 15, 12, 12, 14, 18]


def find_header(rows):
    """Строка шапки — та, где стоит Bar Code."""
    for index, row in enumerate(rows):
        if any("bar code" in str(cell).lower() for cell in row):
            return index
    raise SystemExit("в файле не нашлась строка с Bar Code")


def read_goods(path, sheet):
    table = pd.read_excel(path, sheet_name=sheet, header=None)
    rows = table.values.tolist()
    start = find_header(rows)
    titles = [re.sub(r"\s+", " ", str(cell)).strip().lower() for cell in rows[start]]

    def column(*words):
        for index, title in enumerate(titles):
            if any(word in title for word in words):
                return index
        return None

    spots = {
        "code": column("bar code"),
        "name_kr": column("product name(kr)", "product name (kr)"),
        "name_en": column("product name(en)", "product name (en)"),
        "retail": column("소비자가", "retail"),
        "supply": column("공급가", "supply price"),
        "box": column("outbox", "pcs /"),
    }
    goods, unpriced = [], []
    for row in rows[start + 1:]:
        code = re.sub(r"\D", "", str(row[spots["code"]]))
        supply = pd.to_numeric(row[spots["supply"]], errors="coerce")
        if len(code) < 8:
            continue
        # Позицию без цены в пересылаемый прайс не ставим: пустая или нулевая
        # цена читается как ошибка, а не как «цена по запросу».
        if pd.isna(supply) or supply <= 0:
            unpriced.append(str(row[spots["name_en"]]).strip())
            continue
        goods.append({
            "code": code,
            "name_kr": str(row[spots["name_kr"]]).strip()
            if spots["name_kr"] is not None else "",
            "name_en": str(row[spots["name_en"]]).strip(),
            "retail": pd.to_numeric(row[spots["retail"]], errors="coerce"),
            "supply": supply,
            "box": pd.to_numeric(row[spots["box"]], errors="coerce"),
        })
    return goods, unpriced


def write(goods, brand, date, note, target):
    book = Workbook()
    # Свойства документа заводим заново: в исходном файле стоят автор и
    # организация, а их видно в свойствах даже после правки ячеек.
    book.properties.creator = brand
    book.properties.lastModifiedBy = brand
    book.properties.title = f"{brand} PRICE LIST"
    book.properties.description = None
    book.properties.company = None

    ws = book.active
    ws.title = brand[:31]
    ws.append([f"{brand} PRICE LIST"])
    ws["A1"].font = Font(bold=True, size=14)
    ws.append([f"Date: {date}"] + ([note] if note else []))
    ws.append([])

    ws.append(COLUMNS)
    for cell in ws[4]:
        cell.fill, cell.font = HEAD, Font(bold=True)
        cell.alignment = Alignment(wrap_text=True, vertical="center",
                                   horizontal="center")
        cell.border = GRID

    for number, item in enumerate(goods, start=1):
        line = ws.max_row + 1
        ws.append([
            number, item["code"], item["name_kr"], item["name_en"],
            item["retail"] if pd.notna(item["retail"]) else None,
            item["supply"],
            item["box"] if pd.notna(item["box"]) else None,
            None,
            f"=IF(H{line}=\"\",0,H{line}*G{line})",
            f"=IF(H{line}=\"\",0,I{line}*F{line})",
        ])
        for cell in ws[line]:
            cell.border = GRID
        # Штрихкод текстом: иначе Excel покажет его как 8,80985E+12.
        ws[f"B{line}"].number_format = "@"
        ws[f"B{line}"].alignment = Alignment(horizontal="left")

    first, last = 5, ws.max_row
    ws.append([None, "TOTAL", f"позиций: {len(goods)}", None, None, None, None,
               f"=SUM(H{first}:H{last})", f"=SUM(I{first}:I{last})",
               f"=SUM(J{first}:J{last})"])
    for cell in ws[ws.max_row]:
        cell.fill, cell.font, cell.border = TOTAL, Font(bold=True), GRID

    for index, width in enumerate(WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width
    for letter in "EFGHIJ":
        for cell in ws[letter][3:]:
            cell.number_format = "# ##0"
    ws.freeze_panes = "A5"

    os.makedirs(os.path.dirname(target) or ".", exist_ok=True)
    book.save(target)
    return ws.max_row


def main(source, sheet, brand, date, note, target):
    goods, unpriced = read_goods(source, sheet)
    if not goods:
        raise SystemExit("в файле не нашлось позиций с ценой")
    target = target or os.path.join(OUT_DIR, f"Прайс {brand} {date}.xlsx")
    write(goods, brand, date, note, target)
    print(f"{brand}: {len(goods)} позиций, цены "
          f"{int(min(item['supply'] for item in goods)):,}-"
          f"{int(max(item['supply'] for item in goods)):,} KRW".replace(",", " "))
    if unpriced:
        print(f"Без цены у поставщика, в файл не вошли: {len(unpriced)} — "
              + "; ".join(unpriced))
    print(f"Сохранено: {target}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Прайс без имени поставщика")
    parser.add_argument("source")
    parser.add_argument("--sheet", default=0)
    parser.add_argument("--brand", required=True)
    parser.add_argument("--date", required=True)
    parser.add_argument("--note", default="", help="строка условий, например MOA")
    parser.add_argument("--out", default="")
    args = parser.parse_args()
    main(args.source, args.sheet, args.brand, args.date, args.note, args.out)
