#!/usr/bin/env python3
"""Разбор официального файла тарифов Яндекс Маркета за услуги маркетплейса.

Файл Маркета содержит 4 листа (FBY, FBS, DBS, Express), в каждом дерево
категорий до 8 уровней и колонка с тарифом. Скрипт сводит все модели в одну
таблицу и выгружает срез по категории `Товары для красоты` в TSV.

Использование:
    python3 tools/parse_ym_rates.py data/yandex_market/marketplace_services_rates_01072026.xlsx \
        --category "Товары для красоты" \
        --out outputs/yandex_market/ym_rates_beauty_01072026.tsv
"""

import argparse
import statistics
from pathlib import Path

import openpyxl

MODELS = ["FBY", "FBS", "Express", "DBS"]


def read_rates(path):
    """Возвращает {модель: {кортеж категорий: тариф}}."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    rates = {}
    for ws in wb.worksheets:
        model = ws.title.split()[-1]
        table = {}
        for row in ws.iter_rows(min_row=2, values_only=True):
            path_cats = tuple(c for c in row[:8] if c not in (None, ""))
            if not path_cats:
                continue
            table[path_cats] = row[8]
        rates[model] = table
    return rates


def subtree(rates, category):
    """Ключи дерева одной категории уровня 1, в исходном порядке файла."""
    return [k for k in rates["FBY"] if k[0] == category]


def is_terminal(key, keys):
    return not any(k[: len(key)] == key and len(k) > len(key) for k in keys)


def write_tsv(rates, keys, out_path):
    out_path.parent.mkdir(parents=True, exist_ok=True)
    header = [
        "Уровень 1",
        "Уровень 2",
        "Уровень 3",
        "Уровень 4",
        "Уровень 5",
        "Путь",
        "Конечная категория",
        "FBY %",
        "FBS %",
        "Express %",
        "DBS %",
        "FBS минус FBY, п.п.",
    ]
    lines = ["\t".join(header)]
    for key in keys:
        levels = list(key[:5]) + [""] * (5 - len(key[:5]))
        row = levels + [
            " > ".join(key),
            "да" if is_terminal(key, keys) else "нет",
            f"{rates['FBY'][key] * 100:.1f}",
            f"{rates['FBS'][key] * 100:.1f}",
            f"{rates['Express'][key] * 100:.1f}",
            f"{rates['DBS'][key] * 100:.1f}",
            f"{(rates['FBS'][key] - rates['FBY'][key]) * 100:.1f}",
        ]
        lines.append("\t".join(str(c) for c in row))
    out_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return len(lines) - 1


def summary(rates, keys, category):
    fby = [rates["FBY"][k] for k in keys]
    term = [k for k in keys if is_terminal(k, keys)]
    term_fby = [rates["FBY"][k] for k in term]
    print(f"Категория: {category}")
    print(f"  строк всего: {len(keys)}, конечных категорий: {len(term)}")
    print(
        "  FBY: min %.1f%%  median %.1f%%  max %.1f%%"
        % (min(fby) * 100, statistics.median(fby) * 100, max(fby) * 100)
    )
    print(
        "  FBY по конечным: min %.1f%%  median %.1f%%  max %.1f%%"
        % (min(term_fby) * 100, statistics.median(term_fby) * 100, max(term_fby) * 100)
    )
    same_express = all(rates["Express"][k] == rates["FBY"][k] for k in keys)
    same_dbs = all(rates["DBS"][k] == rates["FBS"][k] for k in keys)
    diffs = {round(rates["FBS"][k] - rates["FBY"][k], 4) for k in keys}
    print(f"  Express == FBY: {same_express}; DBS == FBS: {same_dbs}")
    print(f"  разница FBS - FBY: {sorted(round(d * 100, 1) for d in diffs)} п.п.")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx", help="официальный файл тарифов Маркета")
    ap.add_argument("--category", default="Товары для красоты")
    ap.add_argument("--out", default="outputs/yandex_market/ym_rates_beauty.tsv")
    args = ap.parse_args()

    rates = read_rates(args.xlsx)
    keys = subtree(rates, args.category)
    if not keys:
        raise SystemExit(f"категория не найдена: {args.category}")
    summary(rates, keys, args.category)
    written = write_tsv(rates, keys, Path(args.out))
    print(f"  записано строк в {args.out}: {written}")


if __name__ == "__main__":
    main()
