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

# позиции с противоречиями в учете: публикуем только после проверки наличия
DISPUTED = ("ртнд-150", "иллюминатор круглый в сборе ф300мм створчатый",
            "сундук")
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

    low = name.lower()
    disputed = any(d in low for d in DISPUTED)
    if disputed:
        parts.insert(0, "НЕ ПУБЛИКОВАТЬ, ПОКА НЕ ПРОВЕРЕНО НАЛИЧИЕ.")

    return {
        "Приоритет": 9 if disputed else (1 if r["Сумма"] >= 500_000 else 2),
        "Тип": "проверить" if disputed else "отдельное",
        "Заголовок": shorten(f"{name} судовой"),
        "Цена, руб": money(r["Цена за ед., руб"]),
        "Цена число": float(r["Цена за ед., руб"]),
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
        "Цена число": float(min(prices)),
        "Категория Авито": CATEGORY.get(group, ""),
        "Описание": "\n".join(lines),
        "Что входит": f"{len(items)} наименований, {units:.0f} единиц, "
                      f"на {money(total)} руб",
        "Локация": "склад Владивосток",
        "Нужны фото": "да, общий план и крупные позиции",
    }


RULES = [
    ("Номер чертежа обязательно",
     "Снабженцы ищут не «арматуру», а «521-01.468-07» или «544-03.067». "
     "В описаниях чертежи уже проставлены из реестра. Это главное отличие "
     "от обычного объявления, по чертежу вас найдут те, кому нужна "
     "конкретная деталь."),
    ("Цена обязательна",
     "По журналу обзвона первый вопрос всегда про цену. Объявление без цены "
     "на этом рынке не работает. Поэтому 100 позиций без цены сюда не попали, "
     "их сначала надо оценить."),
    ("Не выкладывать спорные позиции",
     "Регулятор РТНД-150 числится с наличием 3 при учете 0, иллюминатор Ф300 "
     "створчатый не пересчитан. Продать то, чего нет, хуже, чем не продать."),
    ("Регион Владивосток, отправка по России",
     "Покупатели по журналу обзвона есть в Петербурге, Нижнем Новгороде, "
     "Москве, на Камчатке и в Магадане. В объявлении пишем про отправку "
     "транспортной компанией."),
    ("Фото обязательны",
     "На Google Drive есть папка «Фото» по судовому оборудованию. Объявление "
     "без фото на промышленном рынке не смотрят. Нужно разобрать архив "
     "по позициям."),
    ("Фарпост важнее Авито во Владивостоке",
     "По промышленному и судовому товару на Дальнем Востоке Фарпост сильнее. "
     "Выкладывать на обе площадки, отмечать в таблице обе колонки."),
    ("Один телефон и один ответственный",
     "Все звонки идут на Олесю. Иначе покупатель услышит разные цены "
     "от разных людей и уйдет."),
    ("Бизнес-профиль",
     "76 объявлений разовыми платными размещениями выйдут дороже, чем тариф "
     "для профессиональных продавцов. Сравнить стоимость до публикации."),
    ("Допродажа в каждом объявлении",
     "В описании есть строка о том, что склад большой и перечень вышлем "
     "по запросу. Человек пришел за клинкетом, а забрал еще и фильтры."),
]


def write_xlsx(path, ads):
    """Тот же список объявлений, но в виде рабочей таблицы."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation

    font = "Arial"
    head_fill = PatternFill("solid", fgColor="1F3864")
    head_font = Font(name=font, size=11, bold=True, color="FFFFFF")
    base = Font(name=font, size=10)
    input_fill = PatternFill("solid", fgColor="FFFF00")
    prio_fills = {1: PatternFill("solid", fgColor="FCE4E4"),
                  2: PatternFill("solid", fgColor="FFF6E0"),
                  3: PatternFill("solid", fgColor="EAF1EA"),
                  9: PatternFill("solid", fgColor="D9D9D9")}
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    wb = Workbook()
    ws = wb.active
    ws.title = "Объявления"

    ws["A1"] = "Объявления для Авито и Фарпоста"
    ws["A1"].font = Font(name=font, size=14, bold=True, color="1F3864")
    ws["A2"] = ("Публикуем волнами: сначала приоритет 1, затем 2, затем 3. "
                "Желтые колонки заполняете сами. Правила — на листе «Как публиковать».")
    ws["A2"].font = Font(name=font, size=9, italic=True, color="666666")

    headers = ["№", "Приоритет", "Тип", "Заголовок для объявления", "Цена, руб",
               "Категория", "Текст объявления", "Что входит", "Локация",
               "На Авито", "На Фарпост", "Дата публикации", "Отклики"]
    start = 4
    for i, h in enumerate(headers, start=1):
        cell = ws.cell(row=start, column=i, value=h)
        cell.fill = head_fill
        cell.font = head_font
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        cell.border = border
    ws.row_dimensions[start].height = 30

    for n, a in enumerate(ads, start=1):
        row = start + n
        values = [n, a["Приоритет"], a["Тип"], a["Заголовок"], a["Цена число"],
                  a["Категория Авито"], a["Описание"], a["Что входит"],
                  a["Локация"], "Нет", "Нет", "", ""]
        for i, v in enumerate(values, start=1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.font = base
            cell.border = border
            cell.alignment = Alignment(vertical="top",
                                       wrap_text=(i in (4, 6, 7, 8)))
            if i == 5:
                cell.number_format = "#,##0"
            if i <= 3:
                cell.fill = prio_fills.get(a["Приоритет"], prio_fills[9])
            if i in (10, 11, 12, 13):
                cell.fill = input_fill
        ws.row_dimensions[row].height = 58

    last = start + len(ads)
    dv = DataValidation(type="list", formula1='"Да,Нет"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"J{start + 1}:K{last}")

    s = last + 2
    ws.cell(row=s, column=3, value="Всего объявлений").font = base
    ws.cell(row=s, column=4, value=f"=COUNTA(D{start + 1}:D{last})").font = base
    ws.cell(row=s + 1, column=3, value="Опубликовано на Авито").font = base
    ws.cell(row=s + 1, column=4,
            value=f'=COUNTIF(J{start + 1}:J{last},"Да")').font = base
    ws.cell(row=s + 2, column=3, value="Опубликовано на Фарпост").font = base
    ws.cell(row=s + 2, column=4,
            value=f'=COUNTIF(K{start + 1}:K{last},"Да")').font = base
    ws.cell(row=s + 3, column=3, value="Осталось выложить на Авито").font = base
    ws.cell(row=s + 3, column=4,
            value=f"=D{s}-D{s + 1}").font = Font(name=font, bold=True, color="C00000")

    ws.auto_filter.ref = f"A{start}:M{last}"
    ws.freeze_panes = f"D{start + 1}"
    for i, w in enumerate([5, 10, 11, 40, 12, 30, 85, 34, 18, 10, 12, 15, 24], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # лист с правилами
    ws2 = wb.create_sheet("Как публиковать")
    ws2["A1"] = "Правила публикации"
    ws2["A1"].font = Font(name=font, size=14, bold=True, color="1F3864")
    for i, h in enumerate(["№", "Правило", "Почему"], start=1):
        cell = ws2.cell(row=3, column=i, value=h)
        cell.fill = head_fill
        cell.font = head_font
        cell.border = border
    for n, (rule, why) in enumerate(RULES, start=1):
        row = 3 + n
        for i, v in enumerate([n, rule, why], start=1):
            cell = ws2.cell(row=row, column=i, value=v)
            cell.font = base
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=(i in (2, 3)))
        ws2.row_dimensions[row].height = 52
    for i, w in enumerate([5, 38, 95], start=1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    wb.save(path)


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

    xlsx = OUT / "Объявления_Авито_и_Фарпост.xlsx"
    write_xlsx(xlsx, ads)

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
                         (3, "третья волна, мелкие лоты"),
                         (9, "не публиковать до проверки наличия")):
        part = [a for a in ads if a["Приоритет"] == level]
        print(f"  приоритет {level} ({label}): {len(part)} объявлений")


if __name__ == "__main__":
    main()
