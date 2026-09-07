"""Предложение покупателю по уже отправленным ценам.

Файл уходит покупателю, поэтому в нем нет ни себестоимости, ни нашей
прибыли, ни остатков — только позиции, количества и цены прайса.

Что можем отдать: складываем остаток с товаром в пути, вычитаем то, что
нужно нам самим на горизонт планирования с учетом сезона, и остаток
предлагаем. Отгрузить можем только то, что уже лежит на складе, поэтому
предложение ограничено остатком, а не будущим приходом.

Второй лист — по его заявке: его позиции и его количества, а рядом
сколько за тот же товар мы получаем на маркетплейсе после комиссии и
логистики. Разница между двумя суммами и есть то, что мы теряем на
сделке.

Запуск:
    python3 tools/offer_to_customer.py <заказ.xlsx> --sales <отчет WB.xls> \
        --container <инвойс.xlsx> --keep 12
"""

import argparse
import os
import re
import sys

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import add_barcodes
import brand_names
import name_match
from add_sales_price import attach, normal, read_prices, sales_by_stock
from rates import USD_RUB
from check_customer_order import read_order
from stock_forecast import SEASON, START, run_out
from brand_names import brand_of

PRICES_TABLE = "outputs/prices_normalized.xlsx"
from stock_with_incoming import read_container, truck
from wholesale_vs_wb import MARGIN, SKIP

TARGET = "outputs/Предложение покупателю.xlsx"

HEAD = PatternFill("solid", fgColor="1F3864")
TOTAL = PatternFill("solid", fgColor="E2EFDA")
ASKED = PatternFill("solid", fgColor="FFF2CC")
WHITE = Font(color="FFFFFF", bold=True)

OFFER = ["№", "Наименование", "Количество, шт", "Цена, руб", "Сумма, руб"]
OFFER_WIDTHS = [5, 88, 15, 12, 15]
ORDER = ["№", "Наименование", "Количество, шт", "Цена для вас, руб",
         "Сумма для вас, руб", "Выручка на маркетплейсе за штуку, руб",
         "Она же на это количество, руб", "Разница по выручке, руб"]
# Выручка маркетплейса — это уже за вычетом его комиссии и логистики, но
# до хранения, рекламы и налога. Пишем это в файле, чтобы разницу не
# читали как разницу в прибыли.
NOTE = ("Выручка на маркетплейсе указана после комиссии площадки и логистики, "
        "но до хранения, рекламы и налога.")
# Маски держим себе. Пэды и патчи сюда не относим: это отдельный товар,
# и под фильтр они попадали только из-за русского хвоста в названии.
MASK = re.compile(r"\bmask|маск", re.IGNORECASE)
ORDER_WIDTHS = [5, 76, 15, 14, 16, 18, 18, 15]
# Внутренний лист: он остается у нас, поэтому здесь видно и остаток, и
# сколько останется после отгрузки, и на сколько этого хватит.
INSIDE = ["№", "Наименование", "Остаток, шт", "В пути, шт", "Продажи в месяц, шт",
          "Можем отдать, шт", "Останется у нас, шт", "Нам хватит на, мес",
          "Цена, руб", "Сумма, руб", "Прибыль, руб"]
INSIDE_WIDTHS = [5, 76, 12, 12, 15, 14, 15, 14, 11, 14, 14]
# Что теряем, отдавая ходовое: на WB та же партия зарабатывает больше.
LOSS = ["№", "Наименование", "Просит, шт", "Продажи в месяц, шт",
        "Прибыль оптом, руб", "Прибыль на WB, руб", "Потеряем, руб",
        "Во сколько раз WB выгоднее"]
LOSS_WIDTHS = [5, 74, 12, 15, 15, 15, 14, 16]
# Киргизия как запасной канал: цена местного прайса плюс довоз до Москвы.
KG = ["№", "Наименование", "Наша себестоимость, руб", "Цена в Киргизии, руб",
      "Поставщик", "Доставка, руб", "Итого из Киргизии, руб", "Разница, руб",
      "Разница, %", "Замечание"]
KG_WIDTHS = [5, 62, 16, 16, 16, 12, 16, 13, 12, 24]
# Вид товара: маска и сыворотка с одинаковыми словами — разные вещи.
KIND = {"MASK", "PAD", "PADS", "EYE", "SET", "KIT", "SERUM", "CREAM", "TONER",
        "CLEANSER", "FOAM", "AMPOULE", "ESSENCE", "OIL", "BALM", "STICK",
        "PATCH", "MIST", "GEL", "SUNSCREEN", "SHAMPOO", "LOTION", "SCRUB"}
SCORE = 0.6


def read_report(path, column):
    """Колонка отчета WB рядом с продажами: выручка или прибыль."""
    columns = {1: "Товар", 15: "Продано, шт", column: "Выручка, руб"}
    table = pd.read_excel(path, sheet_name=0, header=None).iloc[SKIP:]
    table = table[list(columns)].rename(columns=columns)
    table = table[table["Товар"].notna()]
    for name in ("Продано, шт", "Выручка, руб"):
        table[name] = pd.to_numeric(table[name], errors="coerce").fillna(0)
    return table.reset_index(drop=True)


def read_net(path):
    """Сколько нам приходит с маркетплейса после комиссии и логистики."""
    columns = {1: "Товар", 15: "Продано, шт", 29: "К нам, руб"}
    table = pd.read_excel(path, sheet_name=0, header=None).iloc[SKIP:]
    table = table[list(columns)].rename(columns=columns)
    table = table[table["Товар"].notna()]
    for column in ("Продано, шт", "К нам, руб"):
        table[column] = pd.to_numeric(table[column], errors="coerce").fillna(0)
    return table.reset_index(drop=True)


def incoming(base):
    """Товар в пути к позициям остатков: общего кода нет, сводим по названию."""
    coming = pd.concat([read_container(PATHS["container"]), truck()], ignore_index=True)
    # Один товар приезжает и контейнером, и машиной — обе строки должны
    # лечь на одну позицию, поэтому сводим many-to-one.
    pairs = name_match.to_one(coming["Товар"], base["Наименование"])
    totals = pd.Series(0.0, index=base.index)
    for index, item in coming.iterrows():
        if index in pairs:
            totals[pairs[index]] += item["Кол-во"]
    return totals


def need_for(monthly, keep):
    """Сколько штук нужно нам самим на horizon месяцев с учетом сезона."""
    total, month = 0.0, START[1]
    for _ in range(keep):
        total += monthly * SEASON.get(month, 1.0)
        month = 1 if month == 12 else month + 1
    return total


def kind(name):
    return {word for word in re.findall(r"[A-Z0-9]+", str(name).upper())
            if word in KIND}


def kyrgyz_prices():
    """Прайсы киргизских поставщиков, приведенные к рублям."""
    table = pd.read_excel(PRICES_TABLE, dtype={"Штрихкод": str})
    table = table[(table["Страна"] == "KG") & table["Закупка, KRW"].notna()
                  & (table["Закупка, KRW"] > 0)].copy()
    table["Бренд в прайсе"] = table["Бренд"]
    table["Бренд"] = brand_names.resolve(table).fillna("")
    table["Марка"] = (table["Бренд"].astype(str).str.upper()
                      .str.replace(r"[^A-Z0-9]", "", regex=True))
    table["Слова"] = table["Название EN"].map(add_barcodes.words)
    table["Тон"] = table["Название EN"].map(add_barcodes.tone)
    table["Цена, руб"] = (table["Закупка, KRW"] * USD_RUB).round()
    return table.reset_index(drop=True)


def cheapest_kg(name, brand, prices):
    """Самое дешевое предложение Киргизии по позиции, если оно надежное."""
    fits = [(score, index) for score, index in
            add_barcodes.candidates(name, brand, prices)
            if score >= SCORE and kind(name) == kind(prices.at[index, "Название EN"])]
    if not fits:
        return None
    return min(fits, key=lambda pair: prices.at[pair[1], "Цена, руб"])[1]


def forced(name, rules):
    """Ручное количество по позиции: «часть названия=штук» или «=половина»."""
    for fragment, value in rules.items():
        if fragment.lower() in str(name).lower():
            return value
    return None


def drop(table, skip_brands, no_masks):
    """Убираем из предложения бренды и виды товара, которые не отдаем."""
    keep_rows = []
    for _, item in table.iterrows():
        name = item["Наименование"]
        if brand_of(name) in skip_brands:
            continue
        if no_masks and MASK.search(str(name)):
            continue
        keep_rows.append(item)
    return pd.DataFrame(keep_rows).reset_index(drop=True)


def build(order, report, keep, give_rules=None, markup=None):
    give_rules = give_rules or {}
    base = attach(order, read_prices())
    if markup is not None:
        # В складском файле в колонке «С/с» лежит уже цена с наценкой.
        # Значит, она и есть цена продажи, а себестоимость получаем
        # делением: ставить наценку второй раз нельзя.
        base["Продажная цена, руб"] = base["Себестоимость, руб"].round()
        base["Себестоимость, руб"] = base["Себестоимость, руб"] / (1 + markup / 100)
    found = sales_by_stock(base, report.rename(columns={"К нам, руб": "Выручка, руб"}))
    base["Продажи в месяц, шт"] = found["Продано, шт"]
    # Выручка к перечислению на штуку — цена сравнения, без себестоимости.
    base["Выручка на штуку, руб"] = [
        round(money / sold) if sold else None
        for money, sold in zip(found["Выручка, руб"], found["Продано, шт"])
    ]
    # Фактическая прибыль WB на штуку — из отчета за месяц.
    earned = sales_by_stock(base, read_report(PATHS["sales"], 123))
    base["Прибыль WB на штуку, руб"] = [
        round(money / sold, 1) if sold else None
        for money, sold in zip(earned["Выручка, руб"], earned["Продано, шт"])
    ]
    base["В пути, шт"] = incoming(base)

    offer, ask, inside, loss = [], [], [], []
    for _, item in base.iterrows():
        rest, coming = item["Остаток, шт"], item["В пути, шт"]
        monthly, price = item["Продажи в месяц, шт"], item["Продажная цена, руб"]
        if pd.isna(price):
            continue
        # Владелец может назначить количество сам — тогда расчет запаса
        # не спорим, но больше остатка не отдаем.
        hand = forced(item["Наименование"], give_rules)
        if hand is None:
            give = int(max(0, min((rest + coming) - need_for(monthly, keep), rest)))
        else:
            give = int(min(rest / 2 if hand == "половина" else float(hand), rest))
        if give:
            offer.append([len(offer) + 1, item["Наименование"], give, price,
                          round(price * give)])
            left = rest - give
            if monthly:
                _, months, _ = run_out(left + coming, monthly)
                months = round(months, 1)
            else:
                months = ""
            inside.append([len(inside) + 1, item["Наименование"], int(rest),
                           int(coming), int(monthly), give, int(left), months,
                           price, round(price * give),
                           round((price - item["Себестоимость, руб"]) * give)])
        want = int(item["Просит, шт"])
        # Ходовое отдавать больно: считаем разницу на его количестве.
        gain = item["Прибыль WB на штуку, руб"]
        if want and monthly and pd.notna(gain):
            opt = round((price - item["Себестоимость, руб"]) * want)
            wb = round(gain * want)
            if wb > opt:
                loss.append([len(loss) + 1, item["Наименование"], want,
                             int(monthly), opt, wb, wb - opt,
                             round(wb / opt, 1) if opt > 0 else ""])
        net = item["Выручка на штуку, руб"]
        if want and pd.notna(net):
            ask.append([len(ask) + 1, item["Наименование"], want, price,
                        round(price * want), net, round(net * want),
                        round((net - price) * want)])
    return (pd.DataFrame(offer, columns=OFFER),
            pd.DataFrame(ask, columns=ORDER),
            pd.DataFrame(inside, columns=INSIDE),
            pd.DataFrame(loss, columns=LOSS))


def write(book, title, columns, rows, widths, total, money, price, note=None):
    ws = book.create_sheet(title)
    ws.append(columns)
    for cell in ws[1]:
        cell.fill, cell.font = HEAD, WHITE
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    for row in rows:
        ws.append(row)
    ws.append(total)
    for cell in ws[ws.max_row]:
        cell.fill, cell.font = TOTAL, Font(bold=True)
    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width
    for letter in money:
        for cell in ws[letter][1:]:
            cell.number_format = "# ##0"
    for letter in price:
        for cell in ws[letter][1:]:
            cell.number_format = "# ##0.00"
    if note:
        ws.append([])
        ws.append(["", note])
    ws.freeze_panes = "B2"
    return ws


PATHS = {}


def main(source, sales_path, container_path, keep, target, list_only=False,
         everything=False, skip_brands=(), no_masks=False, give_rules=None,
         markup=None, kg_delivery=None):
    PATHS["container"] = container_path
    order = read_order(source, everything)
    if skip_brands or no_masks:
        before = len(order)
        order = drop(order, {brand.upper() for brand in skip_brands}, no_masks)
        print(f"Исключено позиций: {before - len(order)}")
    PATHS["sales"] = sales_path
    offer, ask, inside, loss = build(order, read_net(sales_path), keep,
                                     give_rules, markup)
    offer = offer.sort_values("Сумма, руб", ascending=False)
    ask = ask.sort_values("Разница по выручке, руб", ascending=False)
    offer["№"] = range(1, len(offer) + 1)
    ask["№"] = range(1, len(ask) + 1)

    book = Workbook()
    book.remove(book.active)
    write(book, "ЧТО МОЖЕМ ОТГРУЗИТЬ", OFFER, offer.values.tolist(), OFFER_WIDTHS,
          ["", "ИТОГО", int(offer["Количество, шт"].sum()), "",
           int(offer["Сумма, руб"].sum())], "CE", "D")
    if everything:
        inside = inside.sort_values("Сумма, руб", ascending=False)
        inside["№"] = range(1, len(inside) + 1)
        write(book, "ЧТО ОСТАНЕТСЯ У НАС", INSIDE, inside.values.tolist(),
              INSIDE_WIDTHS,
              ["", "ИТОГО", int(inside["Остаток, шт"].sum()),
               int(inside["В пути, шт"].sum()), int(inside["Продажи в месяц, шт"].sum()),
               int(inside["Можем отдать, шт"].sum()),
               int(inside["Останется у нас, шт"].sum()), "", "",
               int(inside["Сумма, руб"].sum()), int(inside["Прибыль, руб"].sum())],
              "CDEFGJK", "I")
    if kg_delivery is not None:
        prices = kyrgyz_prices()
        rows = []
        for _, item in read_order(source, True).iterrows():
            cost = item["Себестоимость, руб"] / 1.1
            found = cheapest_kg(item["Наименование"],
                                brand_of(item["Наименование"]), prices)
            if found is None:
                continue
            price = prices.at[found, "Цена, руб"]
            total = price + kg_delivery
            gap = round(total / cost * 100 - 100, 1)
            # Разрыв в разы — это не цена, а разная фасовка: пробник из
            # десяти пэдов против банки на шестьдесят.
            rows.append([len(rows) + 1, item["Наименование"], round(cost), price,
                         prices.at[found, "Поставщик"], kg_delivery, round(total),
                         round(total - cost), gap,
                         "сверить фасовку" if gap > 200 else ""])
        kg_table = pd.DataFrame(rows, columns=KG).sort_values("Разница, %")
        kg_table["№"] = range(1, len(kg_table) + 1)
        write(book, "ЦЕНЫ В КИРГИЗИИ", KG, kg_table.values.tolist(), KG_WIDTHS,
              ["", "ИТОГО", "", "", f"позиций: {len(kg_table)}", "", "", "", "", ""],
              "CDFGH", "")
        print(f"Сверено с Киргизией: {len(kg_table)} позиций, "
              f"дешевле нашей себестоимости: {(kg_table['Разница, %'] < 0).sum()}")
    if len(loss):
        loss = loss.sort_values("Потеряем, руб", ascending=False)
        loss["№"] = range(1, len(loss) + 1)
        write(book, "ЧТО ТЕРЯЕМ НА ХОДОВОМ", LOSS, loss.values.tolist(),
              LOSS_WIDTHS,
              ["", "ИТОГО", int(loss["Просит, шт"].sum()),
               int(loss["Продажи в месяц, шт"].sum()),
               int(loss["Прибыль оптом, руб"].sum()),
               int(loss["Прибыль на WB, руб"].sum()),
               int(loss["Потеряем, руб"].sum()), ""], "CDEFG", "")
    if not list_only:
        write(book, "ВАША ЗАЯВКА", ORDER, ask.values.tolist(), ORDER_WIDTHS,
               ["", "ИТОГО", int(ask["Количество, шт"].sum()), "",
                int(ask["Сумма для вас, руб"].sum()), "",
                int(ask["Она же на это количество, руб"].sum()),
                int(ask["Разница по выручке, руб"].sum())], "CEGH", "DF", NOTE)

    os.makedirs("outputs", exist_ok=True)
    book.save(target)
    print(f"Можем отгрузить: {len(offer)} позиций, "
          f"{int(offer['Количество, шт'].sum()):,} шт на "
          f"{int(offer['Сумма, руб'].sum()):,} руб".replace(",", " "))
    print(f"Сохранено: {target}")
    if list_only:
        return
    print(f"Его заявка: {len(ask)} позиций, {int(ask['Количество, шт'].sum()):,} шт")
    print(f"  по нашим ценам:   {int(ask['Сумма для вас, руб'].sum()):,} руб"
          .replace(",", " "))
    print(f"  на маркетплейсе:  {int(ask['Она же на это количество, руб'].sum()):,} руб"
          .replace(",", " "))
    print(f"  разница:          "
          f"{int(ask['Разница по выручке, руб'].sum()):,} руб".replace(",", " "))


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Предложение покупателю из излишка")
    parser.add_argument("source")
    parser.add_argument("--sales", default="data/sales/wb_sales_2026-08.xls")
    parser.add_argument("--container", required=True)
    parser.add_argument("--keep", type=int, default=12)
    parser.add_argument("--out", default=TARGET)
    parser.add_argument("--list-only", action="store_true",
                        help="только список товаров, без сравнения с маркетплейсом")
    parser.add_argument("--all", action="store_true", dest="everything",
                        help="весь склад, а не только позиции из заявки покупателя")
    parser.add_argument("--skip-brand", action="append", default=[],
                        help="бренд, который не отдаем")
    parser.add_argument("--no-masks", action="store_true",
                        help="не отдавать маски")
    parser.add_argument("--kg-delivery", type=float,
                        help="довоз из Киргизии до Москвы, руб на штуку")
    parser.add_argument("--markup", type=float,
                        help="ставить цену как себестоимость плюс наценку, %%")
    parser.add_argument("--give", action="append", default=[],
                        help="назначить количество: «часть названия=штук» "
                             "или «часть названия=половина»")
    args = parser.parse_args()
    rules = dict(rule.split("=", 1) for rule in args.give)
    main(args.source, args.sales, args.container, args.keep, args.out,
         args.list_only, args.everything, args.skip_brand, args.no_masks, rules,
         args.markup, args.kg_delivery)
