"""Перенос загруженных прайсов Классика в data/price_lists/classic.

Из имени файла убираем служебный хеш и хвостовые подчеркивания, добавляем
префикс classic_. Уже перенесенные файлы пропускаем.

Запуск:
    python3 tools/import_classic_uploads.py <папка_загрузок>
"""

import glob
import os
import re
import shutil
import sys

DEST = "data/price_lists/classic"


def clean_name(filename):
    name = re.sub(r"^[0-9a-f]{8}-", "", filename)   # служебный хеш загрузки
    name = re.sub(r"_+(\.xlsx|\.xls|\.csv)$", r"\1", name)
    return "classic_" + re.sub(r"_+", "_", name)


def main(src_dir):
    os.makedirs(DEST, exist_ok=True)
    added = 0
    for path in sorted(glob.glob(os.path.join(src_dir, "*.xls*")) + glob.glob(os.path.join(src_dir, "*.csv"))):
        dest = os.path.join(DEST, clean_name(os.path.basename(path)))
        if os.path.exists(dest):
            continue
        shutil.copy2(path, dest)
        print("+", os.path.basename(dest))
        added += 1
    print(f"Добавлено файлов: {added}. Всего в папке: {len(glob.glob(DEST + '/*.xls*'))}")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit("Укажите папку с загруженными прайсами")
    main(sys.argv[1])
