"""Сопоставление товаров по названию, когда общего кода нет.

В плане распределения есть артикул WB, в остатках его нет, в инвойсе —
только штрихкод. Общий у трех документов только текст названия, и он
везде разный: в плане обрезан ровно на 75 символах прямо посреди слова,
в инвойсе идет без бренда, в остатках — с русским хвостом.

Поэтому сначала сравниваем начало строки (обрезанное название — это
префикс полного), потом набор значимых слов. Пару берем, только если она
единственная: два похожих товара лучше оставить несопоставленными, чем
свести не тот с не тем.
"""

import re

NOISE = {"для", "лица", "или", "the", "and", "for", "мл", "гр", "шт"}


def normal(name):
    return re.sub(r"[^0-9a-zа-яё]+", "", str(name).lower())


VOLUME = re.compile(r"\d+(ml|g|kg|l|ea|pcs|мл|гр|г|шт)$")


def words(name, drop_volume=False):
    """Значимые слова. Короткие слова с цифрами оставляем: номер тона —
    единственное, чем отличаются #13 и #21."""
    text = re.sub(r"[^0-9a-zа-яё]+", " ", str(name).lower())
    parts = {w for w in text.split()
             if w not in NOISE and (len(w) > 2 or any(c.isdigit() for c in w))}
    if drop_volume:
        parts = {w for w in parts if not VOLUME.fullmatch(w)}
    return parts


def _unique(candidates):
    return candidates[0] if len(candidates) == 1 else None


def match(left_names, right_names):
    """Возвращает {индекс слева: индекс справа} для однозначных пар."""
    left_text = {index: normal(name) for index, name in left_names.items()}
    right_text = {index: normal(name) for index, name in right_names.items()}
    left_words = {index: words(name) for index, name in left_names.items()}
    right_words = {index: words(name) for index, name in right_names.items()}
    left_bare = {index: words(name, True) for index, name in left_names.items()}
    right_bare = {index: words(name, True) for index, name in right_names.items()}

    pairs, taken = {}, set()

    def run(rule):
        for index in left_names.index:
            if index in pairs:
                continue
            fits = [other for other in right_names.index
                    if other not in taken and rule(index, other)]
            found = _unique(fits)
            if found is not None:
                pairs[index] = found
                taken.add(found)

    # Полное совпадение строки, потом совпадение по началу (обрезанное
    # название), потом вложенность наборов слов.
    run(lambda a, b: left_text[a] and left_text[a] == right_text[b])
    run(lambda a, b: len(left_text[a]) > 20 and len(right_text[b]) > 20 and (
        left_text[a].startswith(right_text[b][:len(left_text[a])])
        or right_text[b].startswith(left_text[a][:len(right_text[b])])))
    run(lambda a, b: bool(left_words[a]) and bool(right_words[b]) and (
        left_words[a] <= right_words[b] or right_words[b] <= left_words[a]))
    # Последний заход — без объемов: один документ пишет 100ml в названии,
    # другой выносит фасовку в отдельную колонку.
    run(lambda a, b: bool(left_bare[a]) and bool(right_bare[b]) and (
        left_bare[a] <= right_bare[b] or right_bare[b] <= left_bare[a]))
    return pairs
