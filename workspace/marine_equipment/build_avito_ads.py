#!/usr/bin/env python3
"""Заготовка объявлений для Авито и Фарпоста из реестра склада.

Логика простая: дорогие и штучные позиции идут отдельными объявлениями,
мелочь и массовка — лотами по подгруппам. Позиции без цены не выгружаются:
объявление без цены на этом рынке не работает, покупатели первым делом
спрашивают цену.

Заголовок Авито — до 50 символов, поэтому он собирается отдельно
и обрезается по словам.

Результат: outputs/marine_equipment/avito_obyavleniya_<дата>.tsv
"""

import datetime as dt
import pathlib
import re
import sys
from collections import defaultdict

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
from build_marine_registry import SOURCES, parse  # noqa: E402

OUT = pathlib.Path(__file__).resolve().parents[2] / "outputs" / "marine_equipment"

# позиция дороже этого порога получает отдельное объявление
SOLO_THRESHOLD = 100_000
TITLE_LIMIT = 50

COLUMNS = [
    "Приоритет", "Тип", "Заголовок", "Цена, руб", "Категория Авито",
    "Описание", "Что входит", "Локация", "Нужны фото",
]

CATEGORY = {
    "Арматура": "Оборудование для бизнеса / Промышленное",
    "Насосное и механическое": "Оборудование для бизнеса / Промышленное",
    "Электрика": "Оборудование для бизнеса / Промышленное",
    "Корпусная часть": "Запчасти и аксессуары / Водный транспорт",
    "Трубопроводная обвязка": "Оборудование для бизнеса / Промышленное",
    "Прочее": "Оборудование для бизнеса / Промышленное",
}


def money(v):
    return f"{v:,.0f}".replace(",", " ")


def shorten(text, limit=TITLE_LIMIT):
    text = re.sub(r"\s+", " ", text).strip()
    if len(text) <= limit:
        return text
    out = []
    for word in text.split():
        if len(" ".join(out + [word])) > limit:
            break
        out.append(word)
    return " ".join(out) if out else text[:limit]


def solo_ad(r):
    """Отдельное объявление под дорогую или штучную позицию."""
    name = r["Наименование"]
    draw = r["Чертеж"]
    parts = [name]
    if draw and draw.lower() not in ("без чертежа", "бЕЗ ЧЕРТЕЖА".lower()):
        parts.append(f"чертеж {draw}")
    parts.append(f"В наличии {r['Наличие']:.0f} шт, Владивосток.")
    if r["Состояние"] != "не указано":
        parts.append(f"Состояние: {r['Состояние']}.")
    if r["Комплектность"] != "комплект":
        parts.append(f"Внимание: {r['Комплектность']}.")
    if r["Происхождение"]:
        parts.append(f"Производство: {r['Происхождение']}.")
    parts.append("Самовывоз или отправка транспортной компанией по России.")
    parts.append("Есть другое судовое оборудование, пришлем перечень.")

    return {
        "Приоритет": 1 if r["Сумма"] >= 500_000 else 2,
        "Тип": "отдельное",
        "Заголовок": shorten(f"{name} судовой"),
        "Цена, руб": money(r["Цена за ед., руб"]),
        "Категория Авито": CATEGORY.get(r["Товарная группа"], ""),
        "Описание": " ".join(parts),
        "Что входит": f"{name} — {r['Наличие']:.0f} шт",
        "Локация": r["Локация"],
        "Нужны фото": "да",
    }


def lot_ad(group, sub, items):
    """Лотовое объявление по подгруппе."""
    units = sum(r["Наличие"] for r in items)
    total = sum(r["Сумма"] for r in items)
    prices = [r["Цена за ед., руб"] for r in items]
    top = sorted(items, key=lambda r: -r["Сумма"])[:12]

    lines = [
        f"Судовое оборудование со склада во Владивостоке: {sub.lower()}.",
        f"Всего {len(items)} наименований, {units:.0f} единиц.",
        f"Цены от {money(min(prices))} до {money(max(prices))} руб за единицу.",
        "",
        "Основные позиции:",
    ]
    for r in top:
        draw = f", чертеж {r['Чертеж']}" if r["Чертеж"] else ""
        lines.append(f"- {r['Наименование']}{draw} — {r['Наличие']:.0f} шт, "
                     f"{money(r['Цена за ед., руб'])} руб")
    if len(items) > len(top):
        lines.append(f"- и еще {len(items) - len(top)} наименований")
    lines += [
        "",
        "Продаем поштучно и лотом, на лот скидка.",
        "Пришлем полный перечень с ценами и фото по запросу.",
        "Самовывоз во Владивостоке или отправка транспортной компанией.",
    ]

    return {
        "Приоритет": 2 if total >= 1_000_000 else 3,
        "Тип": "лот",
        "Заголовок": shorten(f"{sub} судовая, склад Владивосток"),
        "Цена, руб": money(min(prices)),
        "Категория Авито": CATEGORY.get(group, ""),
        "Описание": "\n".join(lines),
        "Что входит": f"{len(items)} наименований, {units:.0f} единиц, "
                      f"на {money(total)} руб",
        "Локация": "склад Владивосток",
        "Нужны фото": "да, общий план и крупные позиции",
    }


def main():
    rows = []
    for spec in SOURCES:
        rows.extend(parse(spec))
    for r in rows:
        r["Сумма"] = (r["Наличие"] or 0) * (r["Цена за ед., руб"] or 0)

    sellable = [r for r in rows
                if (r["Наличие"] or 0) > 0 and r["Цена за ед., руб"]]

    solo = [r for r in sellable if r["Цена за ед., руб"] >= SOLO_THRESHOLD]
    rest = [r for r in sellable if r not in solo]

    ads = [solo_ad(r) for r in sorted(solo, key=lambda r: -r["Сумма"])]

    buckets = defaultdict(list)
    for r in rest:
        buckets[(r["Товарная группа"], r["Подгруппа"] or "прочее")].append(r)
    for (group, sub), items in sorted(buckets.items(),
                                      key=lambda kv: -sum(r["Сумма"] for r in kv[1])):
        ads.append(lot_ad(group, sub, items))

    ads.sort(key=lambda a: (a["Приоритет"], a["Тип"]))

    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / f"avito_obyavleniya_{dt.date.today().isoformat()}.tsv"
    with path.open("w", encoding="utf-8", newline="") as fh:
        fh.write("\t".join(COLUMNS) + "\n")
        for a in ads:
            fh.write("\t".join(str(a[c]).replace("\n", " | ") for c in COLUMNS) + "\n")

    no_price = [r for r in rows
                if (r["Наличие"] or 0) > 0 and not r["Цена за ед., руб"]]
    print(f"объявлений всего: {len(ads)}")
    print(f"  отдельных (дороже {money(SOLO_THRESHOLD)} руб за единицу): "
          f"{sum(1 for a in ads if a['Тип'] == 'отдельное')}")
    print(f"  лотовых по подгруппам: {sum(1 for a in ads if a['Тип'] == 'лот')}")
    print(f"позиций не выгружено из-за отсутствия цены: {len(no_price)}")
    print(f"\nфайл: {path}")

    print("\nпорядок публикации:")
    for level, label in ((1, "первая волна, дороже 500 тыс"),
                         (2, "вторая волна"),
                         (3, "третья волна, мелкие лоты")):
        part = [a for a in ads if a["Приоритет"] == level]
        print(f"  приоритет {level} ({label}): {len(part)} объявлений")


if __name__ == "__main__":
    main()
