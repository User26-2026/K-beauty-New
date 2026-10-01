"""Определение единицы цены: за штуку или за набор.

Самая дорогая ошибка при сведении прайсов — сравнить цену за одну маску с
ценой за упаковку из десяти. Поставщики пишут фасовку по-разному, поэтому
разбираем ее из объема, а название используем только как подсказку.

Что считаем набором, а что штукой:

- `25ml*10ea`, `28gx12ea`, `(1.5g + 33g) *4ea` — несколько одинаковых единиц,
  цена за упаковку целиком: НАБОР.
- `20mlx1ea/15mlx2ea/15gx1ea` — комплект из разных средств: НАБОР.
- `SET`, `KIT`, `세트`: НАБОР.
- `160ml (70EA)`, `1.6g(60pcs)`, `50매`, `1 EA` — одна банка или одна пачка,
  а число в скобках это ее содержимое: ШТУКА.

Скобку из названия в расчет не берем: `MASK (4EA)` означает упаковку из
четырех масок, а `160ml (70EA)` — одну банку с семьюдесятью подушечками.
Из названия эти случаи не различить, поэтому решает только объем.
"""

import re

PIECE = "за шт"
PACK = "за набор"

VOLUME = r"\d+(?:[.,]\d+)?\s*(?:ml|g|kg|oz|мл|г)"
COUNT_WORD = r"(?:ea|pcs|шт|매|sheets?|pads?)"
# Между объемом и знаком умножения бывает закрывающая скобка: (1.5g + 33g ) *4ea
MULTIPLIER = rf"(?:{VOLUME}|\d+\s*{COUNT_WORD})\s*\)?\s*[x*×х]\s*(\d+)"


def _from_volume(volume):
    """Разбор фасовки из колонки объема. Возвращает пару или None."""
    if not volume:
        return None

    # Комплект из разных средств: несколько объемов с множителями.
    if len(re.findall(MULTIPLIER, volume, re.I)) > 1:
        return PACK, None

    multiplied = re.search(MULTIPLIER, volume, re.I)
    if multiplied and int(multiplied.group(1)) > 1:
        return PACK, int(multiplied.group(1))

    # Число в скобках после объема — содержимое одной упаковки.
    inside = re.search(rf"{VOLUME}\s*\(\s*(\d+)\s*{COUNT_WORD}", volume, re.I)
    if inside and int(inside.group(1)) > 1:
        return PIECE, int(inside.group(1))

    # Пачка подушечек или масок продается как одна единица: 50매, 12ea.
    single = re.search(rf"^\s*(\d+)\s*{COUNT_WORD}\b", volume, re.I)
    if single and int(single.group(1)) > 1:
        return PIECE, int(single.group(1))

    if re.search(r"\b\d*\s*(set|kit)\b|세트", volume, re.I):
        return PACK, None
    return None


def _from_name(name):
    """Подсказка из названия: только явное умножение и слова SET/KIT."""
    if not name:
        return None
    multiplied = re.search(MULTIPLIER, name, re.I)
    if multiplied and int(multiplied.group(1)) > 1:
        return PACK, int(multiplied.group(1))
    if re.search(r"\b\d*\s*(set|kit)\b|세트", name, re.I):
        return PACK, None
    return None


def detect(volume, name=""):
    """Единица цены и число штук в упаковке.

    Возвращает пару: ("за шт" | "за набор", количество или None).
    """
    volume = re.sub(r"\s+", " ", str(volume or "")).strip()
    name = re.sub(r"\s+", " ", str(name or "")).strip()
    return _from_volume(volume) or _from_name(name) or (PIECE, None)


def unify_packs(df):
    """Сводит фасовку по штрихкоду и пересчитывает цену за штуку.

    Фасовка — свойство товара, а не поставщика. Если один поставщик написал
    `27ml * 10EA`, а другой не указал объем вовсе, упаковка у них одна и та
    же: у обоих один штрихкод. Без этого цена набора у второго делится на
    единицу и он выглядит дороже в десять раз.

    Добавляет колонки «Штук в упаковке (сводно)» и «Цена за штуку (сводно)».
    Возвращает пару: таблица и список штрихкодов, где поставщики указали
    разное число штук — их надо уточнять, а не усреднять.
    """
    import pandas as pd

    result = df.copy()
    ean = result["Штрихкод"].fillna("").astype(str).str.replace(r"\D", "", regex=True)
    result["_EAN"] = ean.where(ean.str.len() >= 8)

    stated = result.dropna(subset=["_EAN", "Штук в упаковке"])
    counts = stated.groupby("_EAN")["Штук в упаковке"]
    conflicts = sorted(counts.nunique()[lambda s: s > 1].index)
    packs = counts.max()

    unified = result["_EAN"].map(packs)
    # Где никто фасовку не указал, товар считаем штучным.
    result["Штук в упаковке (сводно)"] = unified.fillna(result["Штук в упаковке"])
    divisor = result["Штук в упаковке (сводно)"].fillna(1).replace(0, 1)
    result["Цена за штуку (сводно)"] = (result["Закупка, KRW"] / divisor).round(2)
    return result.drop(columns="_EAN"), conflicts
