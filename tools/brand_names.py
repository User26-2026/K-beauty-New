"""Приведение бренда к одному написанию.

Две беды в прайсах. Первая: один товар у разных поставщиков подписан
по-разному (MEDI-PEEL и MEDIPEEL). Вторая: в части прайсов колонка бренда
заполнена не брендом, а категорией (SERUM, CUSHION, TONER / TONER PAD),
из-за чего один бренд разваливается на несколько псевдобрендов.

Чиним в два шага: по штрихкоду берем то написание, которое чаще
встречается как бренд, а строки с категорией вместо бренда отдаем тому
бренду, который назван в имени файла. Подбренды (CHOGONGJIN у MISSHA,
려 у AMORE) не трогаем: они настоящие, поставщики их так и продают.
"""

import re

import pandas as pd

# Слова, из которых состоит категория, а не название бренда.
CATEGORY_WORDS = {
    "SERUM", "AMPOULE", "ESSENCE", "TONER", "PAD", "PADS", "CREAM", "MOISTURIZER",
    "CLEANSER", "CLEANSING", "FOAM", "MASK", "MASKS", "SHEET", "CUSHION", "POWDER",
    "LIP", "EYE", "SPF", "SUN", "SUNSCREEN", "MIST", "OIL", "GEL", "LOTION",
    "EMULSION", "PACK", "PATCH", "SET", "KIT", "BODY", "HAIR", "SHAMPOO", "BALM",
    "SCRUB", "PEELING", "SOAP", "MAKEUP", "BASE", "PRIMER", "STICK", "TINT",
    "RELIEF", "SKINCARE", "CARE", "AND", "OTHERS", "ETC", "NEW", "ITEM", "ITEMS",
}


def _clean(series):
    return (series.fillna("").astype(str).str.upper().str.strip()
            .str.replace(r"\s+", " ", regex=True).replace("", pd.NA))


def _token(text):
    return re.sub(r"[^A-Z0-9]", "", str(text).upper())


def is_category(value):
    """True, если в колонке бренда стоит категория товара, а не бренд."""
    if not isinstance(value, str):
        return False
    words = [word for word in re.split(r"[^A-Z0-9]+", value.upper()) if word]
    if not words:
        return False
    return all(word in CATEGORY_WORDS or word.isdigit() for word in words)


def resolve(df, barcode="Штрихкод", brand="Бренд", source="Файл"):
    """Возвращает колонку бренда, сведенную по штрихкоду и по имени файла."""
    brands = _clean(df[brand])
    counts = brands.value_counts()

    # Шаг 1: один штрихкод — один бренд, берем самое частое написание.
    keys = df[barcode].fillna("").astype(str)
    frame = pd.DataFrame({"ean": keys, "brand": brands}).dropna()
    frame = frame[frame["ean"].str.len() >= 8]
    frame["weight"] = frame["brand"].map(counts)
    by_ean = (frame.sort_values("weight", ascending=False)
                   .drop_duplicates("ean").set_index("ean")["brand"])
    resolved = keys.map(by_ean).fillna(brands)

    # Шаг 2: где вместо бренда категория — берем бренд, названный в файле.
    if source not in df.columns:
        return resolved
    category = resolved.apply(is_category)
    for file_name, group in resolved.groupby(df[source]):
        broken = category.loc[group.index]
        if not broken.any():
            continue
        named = [value for value in group[~broken].dropna().unique()
                 if _token(value) and _token(value) in _token(file_name)]
        if len(named) == 1:
            resolved.loc[group.index[broken]] = named[0]
    return resolved
