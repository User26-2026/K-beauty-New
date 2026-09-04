"""На сколько Бишкек дороже Кореи.

Обычно цены сравнивают только внутри одной страны: у поставщиков разный
маршрут и разные расходы после отгрузки. Здесь задача другая — не выбрать
поставщика, а измерить разрыв. Поэтому считаем в рублях и показываем две
базы сразу:

- корейская цена как есть, EXW со склада поставщика;
- она же с множителем импорта — сколько товар стоит уже у нас.

Первая цифра показывает наценку посредника вместе с логистикой, вторая
отвечает на вопрос, выгоднее ли везти самим.

Запуск:
    python3 tools/compare_regions.py --base KR --against KG
"""

import argparse
import os

import pandas as pd
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import brand_names
from price_unit import unify_packs
from rates import IMPORT_MULTIPLIER, rub_per_unit
from supplier_brand_matrix import split_conflicts

SRC = "outputs/prices_normalized.xlsx"
OUT_DIR = "outputs"
COLUMN = "Цена за штуку (сводно)"

HEADER_FILL = PatternFill("solid", fgColor="DDEBF7")


def load(country):
    df = pd.read_excel(SRC, dtype={"Штрихкод": str})
    df = df[df["Страна"] == country].copy()
    df["Штрихкод"] = df["Штрихкод"].fillna("").astype(str).str.replace(r"\D", "", regex=True)
    df = df[df["Штрихкод"].str.len() >= 8]
    df = df[df["Закупка, KRW"].notna() & (df["Закупка, KRW"] > 0)]
    df["Бренд в прайсе"] = df["Бренд"]
    df["Бренд"] = brand_names.resolve(df).fillna("БЕЗ БРЕНДА")
    df, _ = split_conflicts(df)
    df, _ = unify_packs(df)
    return df


def cheapest(df, imported):
    """По каждому штрихкоду — самое дешевое предложение страны, в рублях."""
    money = sorted(df["Валюта"].dropna().unique())[0]
    rate = rub_per_unit(money, imported=imported)
    df = df.copy()
    df["Цена, руб"] = df[COLUMN] * rate
    best = df.sort_values("Цена, руб").drop_duplicates("Штрихкод")
    return best.set_index("Штрихкод"), money


def main(base_country, against_country):
    base = load(base_country)
    against = load(against_country)

    # Корею берем дважды: как есть и с логистикой, пошлиной и приемкой.
    base_raw, base_money = cheapest(base, imported=False)
    base_landed, _ = cheapest(base, imported=True)

    rows = []
    for supplier, group in against.groupby("Поставщик"):
        money = sorted(group["Валюта"].dropna().unique())[0]
        rate = rub_per_unit(money, imported=False)
        group = group.copy()
        group["Цена, руб"] = group[COLUMN] * rate
        group = group.sort_values("Цена, руб").drop_duplicates("Штрихкод").set_index("Штрихкод")

        shared = group.index.intersection(base_raw.index)
        if shared.empty:
            continue
        table = pd.DataFrame({
            "Поставщик": supplier,
            "Бренд": group.loc[shared, "Бренд"],
            "Товар": group.loc[shared, "Название EN"].fillna(group.loc[shared, "Название RU"]),
            f"Цена {against_country}, руб": group.loc[shared, "Цена, руб"].round(2),
            f"Цена {base_country} EXW, руб": base_raw.loc[shared, "Цена, руб"].round(2),
            f"Цена {base_country} с импортом, руб": base_landed.loc[shared, "Цена, руб"].round(2),
            f"Кто дешевле в {base_country}": base_raw.loc[shared, "Поставщик"],
        })
        table["Дороже EXW, %"] = ((table[f"Цена {against_country}, руб"] /
                                   table[f"Цена {base_country} EXW, руб"] - 1) * 100).round(1)
        table["Дороже с импортом, %"] = (
            (table[f"Цена {against_country}, руб"] /
             table[f"Цена {base_country} с импортом, руб"] - 1) * 100).round(1)
        rows.append(table.reset_index())

    if not rows:
        raise SystemExit("Общих штрихкодов между странами нет")
    items = pd.concat(rows, ignore_index=True).sort_values("Дороже с импортом, %")

    by_supplier = (items.groupby("Поставщик")
                   .agg(**{"Общих позиций": ("Товар", "size"),
                           "Медиана дороже EXW, %": ("Дороже EXW, %", "median"),
                           "Медиана дороже с импортом, %": ("Дороже с импортом, %", "median"),
                           "Дешевле нашего импорта, позиций":
                               ("Дороже с импортом, %", lambda values: int((values < 0).sum()))})
                   .round(1).reset_index()
                   .sort_values("Медиана дороже с импортом, %"))

    by_brand = (items.groupby("Бренд")
                .agg(**{"Позиций": ("Товар", "size"),
                        "Медиана дороже EXW, %": ("Дороже EXW, %", "median"),
                        "Медиана дороже с импортом, %": ("Дороже с импортом, %", "median")})
                .round(1).reset_index()
                .sort_values("Позиций", ascending=False))

    total = pd.DataFrame([
        {"Показатель": f"Позиций {against_country}, совпавших с {base_country}", "Значение": len(items)},
        {"Показатель": "Множитель импорта для Кореи", "Значение": IMPORT_MULTIPLIER},
        {"Показатель": "Медиана: дороже корейской цены EXW, %",
         "Значение": round(items["Дороже EXW, %"].median(), 1)},
        {"Показатель": "Медиана: дороже нашей себестоимости с импортом, %",
         "Значение": round(items["Дороже с импортом, %"].median(), 1)},
        {"Показатель": "Позиций, где местная цена ниже нашего импорта",
         "Значение": int((items["Дороже с импортом, %"] < 0).sum())},
    ])

    out_path = os.path.join(OUT_DIR, f"regions_{against_country.lower()}_vs_{base_country.lower()}.xlsx")
    with pd.ExcelWriter(out_path) as writer:
        total.to_excel(writer, sheet_name="ИТОГО", index=False)
        by_supplier.to_excel(writer, sheet_name="ПО ПОСТАВЩИКАМ", index=False)
        by_brand.to_excel(writer, sheet_name="ПО БРЕНДАМ", index=False)
        items.to_excel(writer, sheet_name="ПОЗИЦИИ", index=False)

        for name in writer.book.sheetnames:
            sheet = writer.book[name]
            sheet.freeze_panes = "B2"
            titles = [str(c.value or "") for c in sheet[1]]
            for index, title in enumerate(titles, start=1):
                cell = sheet.cell(row=1, column=index)
                cell.font = Font(bold=True)
                cell.fill = HEADER_FILL
                cell.alignment = Alignment(wrap_text=True, vertical="center")
                sheet.column_dimensions[get_column_letter(index)].width = (
                    52 if title == "Товар" else 44 if title == "Показатель" else
                    20 if "Поставщик" in title or title == "Бренд" else 14)
                if "%" in title or "руб" in title:
                    for row in sheet.iter_rows(min_row=2, min_col=index, max_col=index):
                        row[0].number_format = "0.0" if "%" in title else "# ##0.00"
                if "%" in title and name != "ИТОГО":
                    letter = get_column_letter(index)
                    sheet.conditional_formatting.add(
                        f"{letter}2:{letter}{sheet.max_row}",
                        ColorScaleRule(start_type="num", start_value=-20, start_color="63BE7B",
                                       mid_type="num", mid_value=20, mid_color="FFEB84",
                                       end_type="num", end_value=100, end_color="F8696B"))

    print(total.to_string(index=False))
    print("\nПо поставщикам:")
    print(by_supplier.to_string(index=False))
    print(f"\nФайл: {out_path}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Разрыв цен между странами")
    parser.add_argument("--base", default="KR")
    parser.add_argument("--against", default="KG")
    args = parser.parse_args()
    main(args.base, args.against)
