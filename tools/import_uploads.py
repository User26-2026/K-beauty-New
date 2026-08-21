"""Перенос загруженных прайсов в папку поставщика.

Из имени файла убираем служебный хеш и хвостовые подчеркивания, добавляем
префикс с именем поставщика. Уже перенесенные файлы пропускаем.

Запуск:
    python3 tools/import_uploads.py <папка_загрузок> <поставщик>
"""

import glob
import os
import re
import shutil
import sys

PRICE_ROOT = "data/price_lists"


def clean_name(filename, supplier):
    name = re.sub(r"^[0-9a-f]{8}-", "", filename)   # служебный хеш загрузки
    name = re.sub(r"_+(\.xlsx|\.xls|\.csv)$", r"\1", name)
    return f"{supplier}_" + re.sub(r"_+", "_", name)


def main(src_dir, supplier):
    dest = os.path.join(PRICE_ROOT, supplier)
    os.makedirs(dest, exist_ok=True)
    added = 0
    sources = glob.glob(os.path.join(src_dir, "*.xls*")) + glob.glob(os.path.join(src_dir, "*.csv"))
    for path in sorted(sources):
        target = os.path.join(dest, clean_name(os.path.basename(path), supplier))
        if os.path.exists(target):
            continue
        shutil.copy2(path, target)
        print("+", os.path.basename(target))
        added += 1
    total = len(glob.glob(os.path.join(dest, "*.xls*")) + glob.glob(os.path.join(dest, "*.csv")))
    print(f"Добавлено файлов: {added}. Всего у поставщика {supplier}: {total}")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit("Запуск: python3 tools/import_uploads.py <папка_загрузок> <поставщик>")
    main(sys.argv[1], sys.argv[2])
