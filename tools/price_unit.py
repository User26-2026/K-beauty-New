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

    if re.search(r"\b(set|kit)\b|세트", volume, re.I):
        return PACK, None
    return None


def _from_name(name):
    """Подсказка из названия: только явное умножение и слова SET/KIT."""
    if not name:
        return None
    multiplied = re.search(MULTIPLIER, name, re.I)
    if multiplied and int(multiplied.group(1)) > 1:
        return PACK, int(multiplied.group(1))
    if re.search(r"\b(set|kit)\b|세트", name, re.I):
        return PACK, None
    return None


def detect(volume, name=""):
    """Единица цены и число штук в упаковке.

    Возвращает пару: ("за шт" | "за набор", количество или None).
    """
    volume = re.sub(r"\s+", " ", str(volume or "")).strip()
    name = re.sub(r"\s+", " ", str(name or "")).strip()
    return _from_volume(volume) or _from_name(name) or (PIECE, None)
