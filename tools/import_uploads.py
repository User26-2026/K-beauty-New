"""Перенос загруженных прайсов в папку поставщика.

Из имени файла убираем служебный хеш и хвостовые подчеркивания, добавляем
префикс с именем поставщика. Уже перенесенные файлы пропускаем.

Кириллица в именах загруженных файлов часто превращается в подчеркивания —
их убираем.

Запуск:
    python3 tools/import_uploads.py <поставщик> <папка_или_файл> [...]
"""

import argparse
import glob
import os
import re
import shutil

PRICE_ROOT = "data/price_lists"


def clean_name(filename, supplier):
    name = re.sub(r"^[0-9a-f]{8}-", "", filename)   # служебный хеш загрузки
    name = re.sub(r"_+(\.xlsx|\.xls|\.csv)$", r"\1", name)
    name = re.sub(r"_+", "_", name).strip("_")
    return f"{supplier}_{name}"


def collect(paths):
    """Разворачиваем папки в список файлов прайсов, файлы берем как есть."""
    files = []
    for path in paths:
        if os.path.isdir(path):
            files += glob.glob(os.path.join(path, "*.xls*"))
            files += glob.glob(os.path.join(path, "*.csv"))
        else:
            files.append(path)
    return sorted(set(files))


def main(supplier, paths):
    dest = os.path.join(PRICE_ROOT, supplier)
    os.makedirs(dest, exist_ok=True)
    added = 0
    for path in collect(paths):
        target = os.path.join(dest, clean_name(os.path.basename(path), supplier))
        if os.path.exists(target):
            continue
        shutil.copy2(path, target)
        print("+", os.path.basename(target))
        added += 1
    total = len(glob.glob(os.path.join(dest, "*.xls*")) + glob.glob(os.path.join(dest, "*.csv")))
    print(f"Добавлено файлов: {added}. Всего у поставщика {supplier}: {total}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("supplier", help="имя папки поставщика в data/price_lists")
    parser.add_argument("paths", nargs="+", help="папка загрузок или конкретные файлы")
    ns = parser.parse_args()
    main(ns.supplier, ns.paths)
