"""Сравнение цен поставщиков: где один и тот же товар дешевле.

Работает по outputs/prices_normalized.xlsx, который делает parse_price_lists.py.

Один товар у разных поставщиков сопоставляем по штрихкоду — это единственный
надежный ключ, названия у всех пишутся по-своему. Там, где штрихкода нет,
подстраховываемся нормализованным названием: латиница в нижнем регистре без
служебных символов, плюс объем.

Запуск:
    python3 tools/compare_suppliers.py
    python3 tools/compare_suppliers.py --min-diff 10   # только разница от 10%
"""

import argparse
import os
import re

import pandas as pd

SRC = "outputs/prices_normalized.xlsx"
OUT_DIR = "outputs"


def norm_name(text):
    """Название для сопоставления: только буквы и цифры в нижнем регистре."""
    text = re.sub(r"[^0-9a-zA-Zа-яА-Я]+", " ", str(text)).lower()
    return re.sub(r"\s+", " ", text).strip()


def norm_volume(text):
    """Объем к единому виду: '150 mL' и '150ml' — одно и то же."""
    match = re.search(r"(\d+[.,]?\d*)\s*(ml|g|мл|г|ea|шт)", str(text), re.I)
    if not match:
        return ""
    return f"{match.group(1).replace(',', '.')}{match.group(2).lower()}"


def name_similarity(left, right):
    """Доля общих слов в названиях — от 0 до 1.

    Штрихкоды в прайсах иногда проставлены с ошибкой, и один код стоит на
    разных товарах у разных поставщиков. Совпадение названий это ловит.
    """
    a = set(norm_name(left).split())
    b = set(norm_name(right).split())
    if not a or not b:
        return None
    return len(a & b) / min(len(a), len(b))


def build_key(row):
    """Штрихкод — основной ключ, название с объемом — запасной."""
    barcode = re.sub(r"\D", "", str(row["Штрихкод"]))
    if len(barcode) >= 8:
        return "ean:" + barcode
    name = norm_name(row["Название EN"]) or norm_name(row["Название KR"])
    if not name:
        return None
    return "name:" + name + "|" + norm_volume(row["Объем"])


def main(min_diff):
    if not os.path.exists(SRC):
        raise SystemExit(f"Нет файла {SRC} — сначала запустите parse_price_lists.py")

    df = pd.read_excel(SRC)
    suppliers = sorted(df["Поставщик"].dropna().unique())
    print("Поставщиков в данных:", ", ".join(suppliers))
    if len(suppliers) < 2:
        print("Сравнивать не с чем: нужен минимум второй поставщик.")
        return

    df["Ключ"] = df.apply(build_key, axis=1)
    df = df[df["Ключ"].notna()]

    # У одного поставщика товар может встретиться дважды — берем дешевле.
    best = (df.sort_values("Закупка, KRW")
              .groupby(["Ключ", "Поставщик"], as_index=False)
              .first())

    shared = best.groupby("Ключ")["Поставщик"].nunique()
    shared = shared[shared > 1].index
    overlap = best[best["Ключ"].isin(shared)]
    print(f"Товаров есть более чем у одного поставщика: {overlap['Ключ'].nunique()}")
    if overlap.empty:
        return

    rows = []
    for key, group in overlap.groupby("Ключ"):
        group = group.sort_values("Закупка, KRW")
        cheapest, priciest = group.iloc[0], group.iloc[-1]
        diff = (priciest["Закупка, KRW"] - cheapest["Закупка, KRW"]) / priciest["Закупка, KRW"] * 100
        similarity = name_similarity(cheapest["Название EN"], priciest["Название EN"])
        row = {
            "Ключ": key,
            "Бренд": cheapest["Бренд"],
            "Товар": cheapest["Название EN"],
            "Объем": cheapest["Объем"],
            "Дешевле у": cheapest["Поставщик"],
            "Цена, KRW": cheapest["Закупка, KRW"],
            "Себестоимость, руб": cheapest["Себестоимость, руб"],
            "Дороже у": priciest["Поставщик"],
            "Цена дороже, KRW": priciest["Закупка, KRW"],
            "Выгода, %": round(diff, 1),
            "Выгода, руб": round(priciest["Себестоимость, руб"] - cheapest["Себестоимость, руб"], 2),
            "Товар у второго": priciest["Название EN"],
            "Схожесть названий": None if similarity is None else round(similarity, 2),
        }
        for supplier in suppliers:
            match = group[group["Поставщик"] == supplier]
            row[f"{supplier}, KRW"] = match.iloc[0]["Закупка, KRW"] if len(match) else None
        rows.append(row)

    result = pd.DataFrame(rows).sort_values("Выгода, %", ascending=False)
    result = result[result["Выгода, %"] >= min_diff]

    # Один штрихкод на разных товарах — ошибка в прайсе, а не выгодная цена.
    suspect = result["Схожесть названий"].notna() & (result["Схожесть названий"] < 0.2)
    result["Проверить"] = suspect.map({True: "штрихкод у разных товаров", False: ""})
    clean = result[~suspect]

    out_path = os.path.join(OUT_DIR, "supplier_price_comparison.xlsx")
    result.to_excel(out_path, index=False)

    print(f"\nПозиций с разницей от {min_diff}%: {len(result)}"
          f"   из них сомнительных по штрихкоду: {int(suspect.sum())}")
    if len(clean):
        print("\nКто чаще дешевле:")
        print(clean["Дешевле у"].value_counts().to_string())
        print("\nТоп-15 по выгоде:")
        cols = ["Бренд", "Товар", "Объем", "Дешевле у", "Цена, KRW", "Дороже у",
                "Цена дороже, KRW", "Выгода, %"]
        print(clean[cols].head(15).to_string(index=False))
    print(f"\nСохранено: {out_path}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--min-diff", type=float, default=0.0,
                        help="показывать только позиции с разницей не меньше, %%")
    main(parser.parse_args().min_diff)
