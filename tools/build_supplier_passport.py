"""Паспорт поставщиков, журнал прайсов и бренды по поставщикам.

Три листа рабочей книги закупок, заполненные тем, что удалось вытащить из
самих прайсов. Все, чего в прайсах нет — город, условия оплаты, документы
для ввоза в РФ — остается со словом «уточнить»: это вопросы менеджерам, а
не то, что можно вывести из данных.

Запуск:
    python3 tools/build_supplier_passport.py
"""

import os

import pandas as pd

REGISTRY = "outputs/registry_prices.xlsx"
PRICES = "outputs/prices_normalized.xlsx"
OUT = "outputs/supplier_passport.xlsx"

# Что известно о компаниях помимо прайсов. Ограничения по рынкам взяты из
# шапок самих прайсов и требуют подтверждения у менеджера.
KNOWN = {
    "KR_classic": {
        "Название компании": "Классик",
        "Тип": "посредник",
        "Контакт": "уточнить",
        "Примечание": "TENZERO — эксклюзив на Россию у ULYBIKA, 304 SKU, продавать нельзя "
                      "без согласования. ARENCIA: запрет на Турцию, Польшу, Великобританию, "
                      "Тайвань, Корею и Amazon. DAENG GI MEO RI: запрет на Украину, Болгарию, "
                      "Венгрию, Грецию, Сербию, Польшу, Сингапур, Таиланд, Китай, Узбекистан, "
                      "США, Ирландию, Казахстан",
    },
    "KR_ge_global": {
        "Название компании": "G&E Global",
        "Тип": "посредник",
        "Контакт": "менеджер Инна",
        "Примечание": "файлы приходят с именами на «Инна»; часть прайсов в старом формате .xls",
    },
    "KR_annecy": {
        "Название компании": "Аннеси (Annecy)",
        "Тип": "посредник",
        "Контакт": "annecy.kr",
        "Примечание": "по CELIMAX в шапке прайса указано, что запрещенных к экспорту "
                      "стран нет; у LAMELIN две цены по объему заказа — уточнить пороги",
    },
    "KR_papacosmetic": {
        "Название компании": "Papacosmetic Co., Ltd",
        "Тип": "экспортер",
        "Контакт": "г. Кимпхо",
        "Примечание": "минимальный заказ 100 000 000 KRW; оплата 50% + 50%; прайс CELIMAX "
                      "содержит два листа с одинаковым набором товаров и ценами, "
                      "отличающимися на 2.5% — взят более дорогой лист",
    },
    "KR_glowbeauty": {
        "Название компании": "GlowBeauty",
        "Тип": "посредник",
        "Контакт": "менеджер Асель",
        "Примечание": "в именах файлов нет месяца, только год — просить присылать с датой",
    },
}


def found_values(series, total_files):
    """Сводка по значению, которое в прайсах встречается не везде.

    У посредника десятки прайсов от разных брендов, и базис или минималка
    указаны лишь в некоторых. Показывать самое частое значение как условие
    поставщика нельзя — это условие одного бренда, а не компании.
    """
    found = series[series.astype(str).str.strip().ne("") & series.ne("уточнить")].dropna()
    if found.empty:
        return "уточнить"
    counts = found.value_counts()
    parts = [f"{value} (прайсов: {count})" for value, count in counts.head(3).items()]
    tail = "" if len(counts) <= 3 else f"; и еще {len(counts) - 3} значений"
    prefix = "" if len(found) == total_files else "указано не во всех прайсах: "
    return prefix + "; ".join(parts) + tail


def suppliers_sheet(reg):
    rows = []
    for code, group in reg.groupby("Код поставщика"):
        known = KNOWN.get(code, {})
        brands = sorted(group["Бренд"].dropna().unique())
        # Считаем по прайсам, а не по позициям: условие относится к файлу.
        by_file = group.drop_duplicates("Файл-источник")
        files = len(by_file)
        note = known.get("Примечание", "")
        if code not in KNOWN:
            note = f"прямой прайс бренда; позиций: {len(group)}"
        rows.append({
            "Код поставщика": code,
            "Страна": "KR",
            "Город": "уточнить",
            "Название компании": known.get("Название компании", brands[0] if brands else code),
            "Тип": known.get("Тип", "бренд"),
            "Контакт": known.get("Контакт", "уточнить"),
            # Подтверждено по всем поставщикам.
            "Базис поставки": "EXW",
            "Валюта прайса": "KRW",
            "Мин. заказ": found_values(by_file["Мин. партия"], files),
            "Условия оплаты": "уточнить",
            "Документы для РФ": "уточнить",
            "Брендов в прайсах": len(brands),
            "Позиций": len(group),
            "Прайсов": group["Файл-источник"].nunique(),
            "Примечание": note,
        })
    return pd.DataFrame(rows).sort_values("Позиций", ascending=False)


def journal_sheet(reg):
    rows = []
    for filename, group in reg.groupby("Файл-источник"):
        note = group["Примечание"].iloc[0]
        rows.append({
            "Имя файла": filename,
            "Код поставщика": group["Код поставщика"].iloc[0],
            "Страна": "KR",
            "Дата прайса": group["Дата прайса"].iloc[0],
            "Валюта": "KRW",
            "Формат": os.path.splitext(filename)[1].lstrip("."),
            "Кол-во позиций": len(group),
            "Внесён в реестр": "да",
            "Кто внёс": "разбор скриптом",
            "Примечание": "дата прайса: только год" if "только год" in note else "",
        })
    return pd.DataFrame(rows).sort_values(["Код поставщика", "Дата прайса"])


def brands_sheet(prices):
    """Какой бренд у каких поставщиков и у кого он дешевле по общим позициям."""
    prices = prices.copy()
    prices["EAN"] = prices["Штрихкод"].astype(str).str.replace(r"\D", "", regex=True)
    best = (prices[prices["EAN"].str.len() >= 8]
            .sort_values("Цена за штуку, KRW")
            .drop_duplicates(["EAN", "Поставщик"]))

    rows = []
    for brand, group in prices.groupby("Бренд"):
        suppliers = sorted(group["Поставщик"].dropna().unique())
        shared = best[best["Бренд"] == brand]
        counts = (shared.groupby("EAN")
                        .filter(lambda g: g["Поставщик"].nunique() > 1)
                        .sort_values("Цена за штуку, KRW")
                        .groupby("EAN").first()["Поставщик"].value_counts())
        rows.append({
            "Бренд": brand,
            "Поставщиков": len(suppliers),
            "У кого есть": ", ".join(suppliers),
            "Общих позиций": int(counts.sum()),
            "Чаще дешевле у": counts.index[0] if len(counts) else "—",
            "Позиций в его пользу": int(counts.iloc[0]) if len(counts) else 0,
            "Комментарий": "выбора нет, поставщик один" if len(suppliers) == 1 else "",
        })
    return pd.DataFrame(rows).sort_values(["Поставщиков", "Общих позиций"], ascending=False)


def main():
    for path in (REGISTRY, PRICES):
        if not os.path.exists(path):
            raise SystemExit(f"Нет файла {path}")
    reg = pd.read_excel(REGISTRY, dtype={"Штрихкод": str})
    prices = pd.read_excel(PRICES, dtype={"Штрихкод": str})

    suppliers = suppliers_sheet(reg)
    journal = journal_sheet(reg)
    brands = brands_sheet(prices)

    with pd.ExcelWriter(OUT) as writer:
        suppliers.to_excel(writer, sheet_name="ПОСТАВЩИКИ", index=False)
        journal.to_excel(writer, sheet_name="ЖУРНАЛ ПРАЙСОВ", index=False)
        brands.to_excel(writer, sheet_name="БРЕНДЫ ПО ПОСТАВЩИКАМ", index=False)

    print(f"ПОСТАВЩИКИ: {len(suppliers)} строк")
    print(suppliers[["Код поставщика", "Название компании", "Тип", "Базис поставки",
                     "Мин. заказ", "Прайсов", "Позиций"]].to_string(index=False))
    print(f"\nЖУРНАЛ ПРАЙСОВ: {len(journal)} строк")
    print(f"БРЕНДЫ ПО ПОСТАВЩИКАМ: {len(brands)} строк")
    multi = brands[brands["Поставщиков"] > 1]
    print(f"\nБренды больше чем у одного поставщика: {len(multi)}")
    print(multi[["Бренд", "Поставщиков", "Общих позиций", "Чаще дешевле у",
                 "Позиций в его пользу"]].head(15).to_string(index=False))
    print(f"\nСохранено: {OUT}")


if __name__ == "__main__":
    main()
