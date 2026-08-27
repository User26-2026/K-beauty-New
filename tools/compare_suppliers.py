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

from price_unit import unify_packs

SRC = "outputs/prices_normalized.xlsx"
OUT_DIR = "outputs"


def norm_name(text):
    """Название для сопоставления: только буквы и цифры в нижнем регистре."""
    text = re.sub(r"[^0-9a-zA-Zа-яА-Я]+", " ", str(text)).lower()
    return re.sub(r"\s+", " ", text).strip()


def norm_volume(text):
    """Объем к единому виду: '150 mL' и '150ml' — одно и то же.

    Фасовку в наборе учитываем отдельно: '38g*4ea' и '38g' — разные позиции,
    и сравнивать их цены нельзя.
    """
    text = str(text)
    match = re.search(r"(\d+[.,]?\d*)\s*(ml|g|мл|г|ea|шт)", text, re.I)
    if not match:
        return ""
    volume = f"{match.group(1).replace(',', '.')}{match.group(2).lower()}"
    pack = re.search(r"[*xх]\s*(\d+)\s*(ea|шт|pcs)?", text[match.end():], re.I)
    return volume + (f"*{pack.group(1)}" if pack else "")


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

    # Фасовку сводим по штрихкоду: часть поставщиков объем не пишет вовсе,
    # и тогда набор из десяти считался бы за штуку.
    df, pack_conflicts = unify_packs(df)
    if pack_conflicts:
        print(f"Разное число штук в упаковке: {len(pack_conflicts)} позиций")

    df["Ключ"] = df.apply(build_key, axis=1)
    df = df[df["Ключ"].notna()]

    # У одного поставщика товар может встретиться дважды — берем дешевле.
    best = (df.sort_values("Цена за штуку (сводно)")
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
        # Сравниваем цену за штуку: у одного поставщика это может быть набор.
        group = group.sort_values("Цена за штуку (сводно)")
        cheapest, priciest = group.iloc[0], group.iloc[-1]
        diff = ((priciest["Цена за штуку (сводно)"] - cheapest["Цена за штуку (сводно)"])
                / priciest["Цена за штуку (сводно)"] * 100)
        similarity = name_similarity(cheapest["Название EN"], priciest["Название EN"])
        # Фасовка сопоставима, только если совпали и единица цены, и число
        # штук в упаковке. Иначе делить цену на штуки нельзя: неизвестно,
        # относится ли цена второго поставщика к той же упаковке.
        packs = {cheapest["Штук в упаковке (сводно)"], priciest["Штук в упаковке (сводно)"]}
        same_pack = len({0 if pd.isna(n) else n for n in packs}) == 1
        row = {
            "Ключ": key,
            "Бренд": cheapest["Бренд"],
            "Товар": cheapest["Название EN"],
            "Объем": cheapest["Объем"],
            "Дешевле у": cheapest["Поставщик"],
            "Цена за шт, KRW": cheapest["Цена за штуку (сводно)"],
            "Цена в прайсе, KRW": cheapest["Закупка, KRW"],
            "Единица": cheapest["Единица цены"],
            "Себестоимость штуки, руб": round(cheapest["Цена за штуку (сводно)"] * 0.058 * 1.4, 2),
            "Дороже у": priciest["Поставщик"],
            "Цена за шт дороже, KRW": priciest["Цена за штуку (сводно)"],
            "Единица у второго": priciest["Единица цены"],
            "Выгода, %": round(diff, 1),
            "Выгода, руб": round((priciest["Цена за штуку (сводно)"]
                                  - cheapest["Цена за штуку (сводно)"]) * 0.058 * 1.4, 2),
            "Товар у второго": priciest["Название EN"],
            "Схожесть названий": None if similarity is None else round(similarity, 2),
            "Объем у второго": priciest["Объем"],
            "Штук в упаковке": cheapest["Штук в упаковке (сводно)"],
            "Штук в упаковке у второго": priciest["Штук в упаковке (сводно)"],
            "Разная фасовка": not same_pack,
        }
        for supplier in suppliers:
            match = group[group["Поставщик"] == supplier]
            row[f"{supplier}, KRW"] = match.iloc[0]["Закупка, KRW"] if len(match) else None
        rows.append(row)

    result = pd.DataFrame(rows).sort_values("Выгода, %", ascending=False)
    result = result[result["Выгода, %"] >= min_diff]

    # Один штрихкод на разных товарах — ошибка в прайсе, а не выгодная цена.
    # Один штрихкод может стоять на разных товарах — это не выгодная цена, а
    # несопоставимая пара. Разную фасовку сравнивать уже можно: цены
    # приведены к штуке, но помечаем ее, чтобы можно было перепроверить.
    bad_name = result["Схожесть названий"].notna() & (result["Схожесть названий"] < 0.2)
    result["Проверить"] = ""
    result.loc[result["Разная фасовка"], "Проверить"] = "разная фасовка, уточнить у поставщика"
    result.loc[bad_name, "Проверить"] = "штрихкод у разных товаров"
    # Крупная разница бывает и настоящей: по MEDI-PEEL RED LACTO SUN SCREEN
    # расхождение в 57% подтвердилось. Поэтому такие позиции остаются в
    # рейтинге, но помечаются на сверку с поставщиком.
    huge = result["Выгода, %"] > 50
    result.loc[huge & (result["Проверить"] == ""), "Проверить"] = "крупная разница, сверить"
    suspect = bad_name | result["Разная фасовка"]
    # Приведение к штуке имеет смысл только при совпавшей фасовке, поэтому
    # позиции с разной фасовкой в рейтинг не идут — их надо уточнить.
    clean = result[~suspect]

    out_path = os.path.join(OUT_DIR, "supplier_price_comparison.xlsx")
    result.to_excel(out_path, index=False)

    print(f"\nПозиций с разницей от {min_diff}%: {len(result)}"
          f"   несопоставимых пар: {int(suspect.sum())}"
          f" (чужой штрихкод {int(bad_name.sum())}, разная фасовка"
          f" {int(result['Разная фасовка'].sum())});"
          f" помечено на сверку по величине разницы: {int(huge.sum())}")
    if len(clean):
        print("\nКто чаще дешевле:")
        print(clean["Дешевле у"].value_counts().to_string())
        print("\nТоп-15 по выгоде:")
        cols = ["Бренд", "Товар", "Объем", "Дешевле у", "Цена за шт, KRW", "Дороже у",
                "Цена за шт дороже, KRW", "Выгода, %"]
        print(clean[cols].head(15).to_string(index=False))
    print(f"\nСохранено: {out_path}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--min-diff", type=float, default=0.0,
                        help="показывать только позиции с разницей не меньше, %%")
    main(parser.parse_args().min_diff)
