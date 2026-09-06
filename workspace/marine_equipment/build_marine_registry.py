#!/usr/bin/env python3
"""Сводный реестр склада судового оборудования и заготовка прайса.

На входе — CSV из data/marine_equipment/ (выгрузка перечней склада).
На выходе:
  outputs/marine_equipment/sudovoe_registry_<дата>.tsv       — единый реестр
  outputs/marine_equipment/sudovoe_price_template_<дата>.tsv — реестр + колонки под цены

Цен в исходных перечнях нет, поэтому колонки цены выводятся пустыми:
их заполняем по прайсам поставщиков и по факту сделок.
"""

import csv
import datetime as dt
import pathlib
import re

ROOT = pathlib.Path(__file__).resolve().parents[2]
SRC = ROOT / "data" / "marine_equipment"
OUT = ROOT / "outputs" / "marine_equipment"

# файл -> (категория, колонка склада, колонка группы)
SOURCES = [
    ("armatura_ulica_foye.csv", "Арматура (улица/фойе)", "Место хранения", None),
    ("sklad_1a_1b_2_3_armatura.csv", "Арматура (склады 1А, 1Б, 2, 3)", "Склад", "Группа"),
    ("elektrika.csv", "Электрика", "Место хранения", "Группа"),
    ("dveri_kryshki_illuminatory.csv", "Двери, крышки, иллюминаторы", "Место хранения", None),
    ("brashpil_nasos_lebedki.csv", "Брашпиль, насосы, лебедки", "Место хранения", None),
]

REGISTRY_COLS = [
    "Категория", "Группа", "Место хранения", "Наименование", "Чертеж",
    "Кол-во", "Состояние", "Комплектность", "Происхождение",
]
PRICE_COLS = REGISTRY_COLS + [
    "Цена за ед., руб", "Сумма, руб", "Источник цены", "Комментарий",
]


def condition(name: str) -> str:
    low = name.lower()
    if "б/у" in low:
        return "б/у"
    if "нов" in low:
        return "новый"
    return "не указано"


def completeness(name: str) -> str:
    return "не комплект" if "не комплект" in name.lower() else "комплект"


def origin(name: str) -> str:
    low = name.lower()
    if "китай" in low:
        return "Китай"
    if "росси" in low:
        return "Россия"
    if "немец" in low:
        return "Германия"
    return ""


def clean(name: str) -> str:
    return re.sub(r"\s+", " ", name).strip()


def load_rows():
    rows = []
    for fname, category, place_col, group_col in SOURCES:
        path = SRC / fname
        with path.open(encoding="utf-8") as fh:
            for raw in csv.DictReader(fh):
                name = clean(raw["Наименование"])
                if not name:
                    continue
                qty = (raw.get("Кол-во") or "").strip()
                rows.append({
                    "Категория": category,
                    "Группа": clean(raw.get(group_col) or "") if group_col else "",
                    "Место хранения": clean(raw.get(place_col) or ""),
                    "Наименование": name,
                    "Чертеж": clean(raw.get("Чертеж") or ""),
                    "Кол-во": qty,
                    "Состояние": condition(name),
                    "Комплектность": completeness(name),
                    "Происхождение": origin(name),
                })
    return rows


def write_tsv(path, columns, rows):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=columns, delimiter="\t",
                                extrasaction="ignore", restval="")
        writer.writeheader()
        writer.writerows(rows)


def main():
    stamp = dt.date.today().isoformat()
    rows = load_rows()

    write_tsv(OUT / f"sudovoe_registry_{stamp}.tsv", REGISTRY_COLS, rows)
    write_tsv(OUT / f"sudovoe_price_template_{stamp}.tsv", PRICE_COLS, rows)

    total_units = sum(int(r["Кол-во"]) for r in rows if r["Кол-во"].isdigit())
    no_qty = [r for r in rows if not r["Кол-во"].isdigit()]

    print(f"позиций: {len(rows)}")
    print(f"единиц хранения: {total_units}")
    print(f"позиций без количества: {len(no_qty)}")
    for r in no_qty:
        print(f"  - {r['Место хранения']}: {r['Наименование']}")

    print("\nпо категориям:")
    for _, category, _, _ in SOURCES:
        part = [r for r in rows if r["Категория"] == category]
        units = sum(int(r["Кол-во"]) for r in part if r["Кол-во"].isdigit())
        print(f"  {category}: позиций {len(part)}, единиц {units}")

    print("\nпо состоянию:")
    for state in ("новый", "б/у", "не указано"):
        part = [r for r in rows if r["Состояние"] == state]
        units = sum(int(r["Кол-во"]) for r in part if r["Кол-во"].isdigit())
        print(f"  {state}: позиций {len(part)}, единиц {units}")


if __name__ == "__main__":
    main()
