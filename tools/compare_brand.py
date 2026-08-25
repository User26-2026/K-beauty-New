"""Сравнение цен на один бренд у всех поставщиков.

Товары сопоставляем по штрихкоду: названия у поставщиков расходятся
(MEDI-PEEL против MEDIPEEL, разные хвосты и приписки), а штрихкод один.

Запуск:
    python3 tools/compare_brand.py MEDIPEEL
    python3 tools/compare_brand.py "ROUND LAB" --min-diff 5
"""

import argparse
import os
import re

import pandas as pd

SRC = "outputs/prices_normalized.xlsx"
OUT_DIR = "outputs"


def load_brand(brand):
    df = pd.read_excel(SRC)
    mask = (df["Бренд"].fillna("").astype(str).str.upper().str.contains(brand.upper())
            | df["Название EN"].fillna("").astype(str).str.upper().str.contains(brand.upper()))
    rows = df[mask].copy()
    rows["Штрихкод"] = rows["Штрихкод"].astype(str).str.replace(r"\D", "", regex=True)
    return rows[rows["Штрихкод"].str.len() >= 8]


def main(brand, min_diff):
    if not os.path.exists(SRC):
        raise SystemExit(f"Нет файла {SRC} — сначала запустите parse_price_lists.py")

    rows = load_brand(brand)
    if rows.empty:
        raise SystemExit(f"Бренд {brand} не найден")

    # У одного поставщика товар может встретиться дважды — берем дешевле.
    best = (rows.sort_values("Закупка, KRW")
                .drop_duplicates(["Штрихкод", "Поставщик"]))
    suppliers = sorted(best["Поставщик"].unique())
    print(f"{brand}: SKU со штрихкодом по поставщикам")
    print(best.groupby("Поставщик").size().to_string())

    prices = best.pivot(index="Штрихкод", columns="Поставщик", values="Закупка, KRW")
    info = (best.sort_values("Название EN", key=lambda s: s.str.len(), ascending=False)
                .drop_duplicates("Штрихкод")
                .set_index("Штрихкод")[["Название EN", "Объем", "MSRP, KRW"]])

    table = info.join(prices)
    shared = table[(prices.notna().sum(axis=1) > 1).reindex(table.index, fill_value=False)].copy()
    print(f"\nЕсть более чем у одного поставщика: {len(shared)} из {len(table)}")
    if shared.empty:
        return

    price_cols = [c for c in suppliers if c in shared.columns]
    shared["Дешевле у"] = shared[price_cols].idxmin(axis=1)
    shared["Мин, KRW"] = shared[price_cols].min(axis=1)
    shared["Макс, KRW"] = shared[price_cols].max(axis=1)
    shared["Разница, %"] = ((1 - shared["Мин, KRW"] / shared["Макс, KRW"]) * 100).round(1)
    shared["Экономия, руб"] = ((shared["Макс, KRW"] - shared["Мин, KRW"]) * 0.058 * 1.4).round(0)

    print("\nКто дешевле, по числу позиций:")
    print(shared["Дешевле у"].value_counts().to_string())
    print(f"\nМедиана расхождения: {shared['Разница, %'].median():.1f}%")

    shared = shared.sort_values("Разница, %", ascending=False)
    out_path = os.path.join(OUT_DIR, f"brand_{re.sub(r'[^a-z0-9]+', '_', brand.lower())}_by_supplier.xlsx")
    shared.to_excel(out_path)

    top = shared[shared["Разница, %"] >= min_diff]
    print(f"\nТоп расхождений (от {min_diff}%):")
    print(top.head(15).to_string())
    print(f"\nСохранено: {out_path}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("brand", help="название бренда, часть тоже подойдет")
    parser.add_argument("--min-diff", type=float, default=0.0, help="порог расхождения, %%")
    ns = parser.parse_args()
    main(ns.brand, ns.min_diff)
