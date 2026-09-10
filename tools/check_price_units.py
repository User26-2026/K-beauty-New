"""Поиск позиций, где единица цены у поставщиков расходится.

Два признака расхождения:

1. Прямой — по одному штрихкоду один поставщик пишет цену за штуку, другой
   за набор, или количество в упаковке у них разное.
2. Косвенный — фасовка записана одинаково, но цены отличаются примерно в
   целое число раз. Это тоже похоже на разную единицу цены: цена за десять
   масок против цены за одну.

Запуск:
    python3 tools/check_price_units.py
"""

import os

import pandas as pd

SRC = "outputs/prices_normalized.xlsx"
OUT = "outputs/price_unit_check.xlsx"
# Во сколько раз цены должны отличаться, чтобы заподозрить разную единицу.
RATIOS = [2, 3, 4, 5, 6, 10, 12, 20, 30, 50, 60, 70, 100]
TOLERANCE = 0.12


def close_to_pack_ratio(ratio):
    """Отношение цен похоже на кратность упаковки."""
    for candidate in RATIOS:
        if abs(ratio - candidate) / candidate <= TOLERANCE:
            return candidate
    return None


def main():
    if not os.path.exists(SRC):
        raise SystemExit(f"Нет файла {SRC} — сначала запустите parse_price_lists.py")

    df = pd.read_excel(SRC)
    df["EAN"] = df["Штрихкод"].astype(str).str.replace(r"\D", "", regex=True)
    df = df[(df["EAN"].str.len() >= 8) & df["Закупка, KRW"].notna()]
    best = df.sort_values("Закупка, KRW").drop_duplicates(["EAN", "Поставщик"])

    rows = []
    for ean, group in best.groupby("EAN"):
        if group["Поставщик"].nunique() < 2:
            continue
        cheap, dear = group.iloc[0], group.iloc[-1]
        units = set(group["Единица цены"])
        packs = set(group["Штук в упаковке"].dropna())
        ratio = dear["Закупка, KRW"] / cheap["Закупка, KRW"] if cheap["Закупка, KRW"] else 0
        multiple = close_to_pack_ratio(ratio)

        if len(units) > 1:
            reason = "разная единица цены в прайсах"
        elif len(packs) > 1:
            reason = f"разное число штук в упаковке: {sorted(int(p) for p in packs)}"
        elif multiple:
            reason = f"цены отличаются примерно в {multiple} раз"
        else:
            continue

        rows.append({
            "Штрихкод": ean,
            "Бренд": cheap["Бренд"],
            "Товар": dear["Название EN"],
            "Причина": reason,
            "Дешевле у": cheap["Поставщик"],
            "Цена мин, KRW": cheap["Закупка, KRW"],
            "Единица мин": cheap["Единица цены"],
            "Фасовка мин": cheap["Объем"],
            "Дороже у": dear["Поставщик"],
            "Цена макс, KRW": dear["Закупка, KRW"],
            "Единица макс": dear["Единица цены"],
            "Фасовка макс": dear["Объем"],
            "Отношение цен": round(ratio, 2),
        })

    result = pd.DataFrame(rows)
    total = best.groupby("EAN")["Поставщик"].nunique()
    shared = int((total > 1).sum())
    print(f"Товаров более чем у одного поставщика: {shared}")
    if result.empty:
        print("Расхождений по единице цены не найдено.")
        return

    result = result.sort_values("Отношение цен", ascending=False)
    result.to_excel(OUT, index=False)
    print(f"Позиций под вопросом: {len(result)}\n")
    print(result["Причина"].value_counts().to_string())
    print("\nТоп-15:")
    cols = ["Бренд", "Товар", "Причина", "Дешевле у", "Цена мин, KRW",
            "Фасовка мин", "Дороже у", "Цена макс, KRW", "Фасовка макс"]
    print(result[cols].head(15).to_string(index=False))
    print(f"\nСохранено: {OUT}")


if __name__ == "__main__":
    main()
