#!/usr/bin/env python3
"""Цены, проставленные покупателем по отгруженному контейнеру.

Покупатель вернул наш же файл Контейнер_отгружено_26.10.xlsx, дописав
на вкладке «Перечень» две колонки:

  P «Рыночная закупка» — его цена закупа за единицу. Заполнена у 190
    позиций из 222, в том числе у 73, где у нас стояло «по
    согласованию»;
  Q — пометки. Их две, обе про материал.

Кроме этого он поправил материал одной строки прямо в колонке O:
паровой клапан Ду65Ру10 (581-03.034) из бронзы стал сталью. Остальные
расхождения с нашим файлом — это округления с плавающей точкой от
пересохранения в Excel, товара они не касаются.

Читаем из файла только то, что он дописал: номер позиции, цену и
пометку. Все остальное — наименование, вес, материал, наш закуп —
берется из нашего расчета, иначе сверка потеряет смысл: сравнивать
надо его цену с нашей, а не его файл с его же файлом.
"""

import pathlib
import re

ROOT = pathlib.Path(__file__).resolve().parents[2]
SRC = (ROOT / "data" / "marine_equipment" / "quotes"
       / "Цены_Бурыкина_контейнер_26.10.xlsx")

SHEET = "Перечень"
HEAD_ROW = 4
COL_NUM = 0      # A, номер позиции
COL_NAME = 1     # B
COL_DRAW = 2     # C
COL_MAT = 14     # O, материал — он правил его руками
COL_PRICE = 15   # P, его цена закупа
COL_NOTE = 16    # Q, его пометка

# Дата, когда он вернул файл
DATED = "20.11.2026"
WHO = "Антон Бурыкин, ООО ПКФ «АРМА»"


def text(v):
    return re.sub(r"\s+", " ", str(v)).strip() if v is not None else ""


def num(v):
    if isinstance(v, (int, float)):
        return float(v)
    s = text(v).replace(" ", "").replace(" ", "").replace(",", ".")
    return float(s) if re.match(r"^-?\d+(\.\d+)?$", s) else None


def read(src=SRC):
    """Его цены по номеру позиции перечня."""
    from openpyxl import load_workbook
    ws = load_workbook(src, data_only=True)[SHEET]
    out = {}
    for row in ws.iter_rows(min_row=HEAD_ROW + 1, values_only=True):
        n = text(row[COL_NUM] if len(row) > COL_NUM else "")
        if not re.match(r"^\d+$", n):
            continue
        price = num(row[COL_PRICE] if len(row) > COL_PRICE else None)
        out[int(n)] = dict(
            Позиция=int(n),
            Наименование=text(row[COL_NAME] if len(row) > COL_NAME else ""),
            Чертеж=text(row[COL_DRAW] if len(row) > COL_DRAW else ""),
            Материал=text(row[COL_MAT] if len(row) > COL_MAT else ""),
            Цена=price if price else None,
            Пометка=text(row[COL_NOTE] if len(row) > COL_NOTE else ""),
        )
    return out


# Материал, который покупатель назвал сам. Владельцем НЕ подтвержден,
# поэтому в materials.py не кладется и в расчет цены не идет: сначала
# подтвердить, потом считать. Но знать об этом надо — сталь вместо
# бронзы меняет и лом в 39 раз, и цену продажи втрое.
CLAIMED_STEEL = [
    dict(draw="427-03.192-2", name="Фильтр забортной воды Ду 100",
         was="бронза", now="сталь", how="пометка «по чертежу сталь»"),
    dict(draw="581-03.034", name="Клапан паровой Ду65Ру10",
         was="бронза", now="сталь", how="материал исправлен в файле"),
]


if __name__ == "__main__":
    data = read()
    have = [v for v in data.values() if v["Цена"]]
    print(f"позиций в его файле: {len(data)}")
    print(f"цена проставлена:    {len(have)}")
    print(f"сумма за единицу:    {sum(v['Цена'] for v in have):,.0f} руб"
          .replace(",", " "))
    print(f"\nпометки:")
    for v in data.values():
        if v["Пометка"]:
            print(f"  поз. {v['Позиция']:>3}  {v['Наименование'][:46]:<48}"
                  f"{v['Пометка']}")
    print(f"\nматериал, названный покупателем (владельцем не подтвержден):")
    for c in CLAIMED_STEEL:
        print(f"  {c['draw']:<16}{c['name'][:40]:<42}{c['was']} -> "
              f"{c['now']}  ({c['how']})")
