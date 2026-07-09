"""Монитор вложений в товар.

Считает сколько денег заморожено в товаре на всех стадиях и сравнивает с лимитом,
чтобы не допускать перевложения:

    Всего в товаре = Заказано за границей (в пути) + Склад + Товар на маркетах

Источники:
  - Заказано за границей   -> задаётся вручную в CONFIG (валюта + курсы).
  - Склад и маркеты        -> считаются из выгрузок остатков (data/stock_costs),
                              как sum(себестоимость * количество) по реальным SKU.
  - Свободный кэш и баланс WB -> задаются вручную в CONFIG.

Запуск:
    python tools/monitor_investment.py                # печать дашборда
    python tools/monitor_investment.py --report        # + запись отчёта в outputs/
    python tools/monitor_investment.py --korea-sheet 30,06,26 --china-sheet 09.07.26
"""
from __future__ import annotations

import argparse
import datetime as dt
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
STOCK_DIR = ROOT / "data/stock_costs"
OUT_DIR = ROOT / "outputs"

# --- Параметры, которые задаёт владелец (обновляй руками) -------------------

CONFIG = {
    # Заказано за границей: оплачено поставщику, товар ещё в пути.
    "abroad_orders": [
        {"label": "Корея (вон)", "amount": 790_000_000, "currency": "KRW",
         "krw_per_usd": 1500, "rub_per_usd": 82},
        {"label": "Китай (USD)", "amount": 67_000, "currency": "USD",
         "rub_per_usd": 82},
    ],
    "cash_bank": 14_000_000,          # свободный кэш в банках, руб
    "wb_balance": 15_000_000,         # остаток на WB, руб (заморожен частично)
    "wb_withdraw_pct": 0.30,          # можно выводить в понедельник, доля
    "limit": 100_000_000,            # лимит вложений в товар, руб
    "warn_ratio": 0.80,               # жёлтая зона: >=80% лимита

    # Выгрузки остатков. По каждой: колонки (0-индекс) и лист.
    # Если sheet=None -> берём последний лист без пометки "копия".
    "stock_files": [
        {
            "label": "Корея (косметика)",
            "path": STOCK_DIR / "ostatki_korea_full.xlsx",
            "sheet": None,
            "cost_col": 20,           # Себестоимость по последнему приходу
            "sklad_col": 21,          # Остаток на складе
            "wb_cols": [22, 23, 24],  # ОстатокВБ: Коротич / Скляров / Крона
            "ozon_cols": [26, 27],    # Остаток ОЗОН
            "art_cols": [1, 2, 3, 4], # артикулы -> признак реальной строки SKU
        },
        {
            "label": "Китай (товары)",
            "path": STOCK_DIR / "ostatki_china_full.xlsx",
            "sheet": None,
            "cost_col": 10,           # Себестоимость
            "sklad_col": 14,          # Остаток склад
            "wb_cols": [15, 16, 17],  # ОстатокВБ
            "ozon_cols": [19, 20],    # Остаток ОЗОН
            "art_cols": [1, 2, 3],
        },
    ],
}


def num(v) -> float:
    """Безопасный парсинг числа: терпит пробелы, %, запятую как разделитель."""
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace("\xa0", "").replace(" ", "").replace("%", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def pick_sheet(wb, requested):
    if requested:
        return wb[requested]
    for name in reversed(wb.sheetnames):
        if "копия" not in name.lower():
            return wb[name]
    return wb[wb.sheetnames[-1]]


def analyze_stock(cfg: dict, sheet_override: str | None):
    wb = openpyxl.load_workbook(cfg["path"], read_only=True, data_only=True)
    ws = pick_sheet(wb, sheet_override or cfg["sheet"])
    sklad_val = wb_val = ozon_val = 0.0
    n_sklad = n_wb = 0
    for row in ws.iter_rows(values_only=True):
        def g(c):
            return row[c] if c < len(row) else None
        cost = num(g(cfg["cost_col"]))
        # реальная строка SKU: есть артикул и себестоимость (не подытог/заголовок)
        if cost <= 0 or not any(num(g(c)) > 0 for c in cfg["art_cols"]):
            continue
        sk = num(g(cfg["sklad_col"]))
        wbq = sum(num(g(c)) for c in cfg["wb_cols"])
        oz = sum(num(g(c)) for c in cfg["ozon_cols"])
        sklad_val += cost * sk
        wb_val += cost * wbq
        ozon_val += cost * oz
        if sk > 0:
            n_sklad += 1
        if wbq > 0:
            n_wb += 1
    return {
        "label": cfg["label"], "sheet": ws.title,
        "sklad": sklad_val, "wb": wb_val, "ozon": ozon_val,
        "n_sklad": n_sklad, "n_wb": n_wb,
    }


def order_to_rub(o: dict) -> float:
    if o["currency"] == "USD":
        return o["amount"] * o["rub_per_usd"]
    if o["currency"] == "KRW":
        return o["amount"] / o["krw_per_usd"] * o["rub_per_usd"]
    raise ValueError(f"неизвестная валюта {o['currency']}")


def rub(x: float) -> str:
    return f"{x:,.0f}".replace(",", " ") + " ₽"


def build(cfg: dict, korea_sheet=None, china_sheet=None) -> dict:
    stocks = []
    for sf in cfg["stock_files"]:
        override = korea_sheet if "Корея" in sf["label"] else china_sheet
        stocks.append(analyze_stock(sf, override))

    abroad = sum(order_to_rub(o) for o in cfg["abroad_orders"])
    sklad = sum(s["sklad"] for s in stocks)
    wb = sum(s["wb"] + s["ozon"] for s in stocks)
    total = abroad + sklad + wb

    wb_liquid = cfg["wb_balance"] * cfg["wb_withdraw_pct"]
    cash_live = cfg["cash_bank"] + wb_liquid

    limit = cfg["limit"]
    free = limit - total
    if total > limit:
        flag = "🔴 СТОП — перевложение"
    elif total >= limit * cfg["warn_ratio"]:
        flag = "🟡 Осторожно — запас < 20%"
    else:
        flag = "🟢 ОК"

    return {
        "abroad": abroad, "sklad": sklad, "wb": wb, "total": total,
        "limit": limit, "free": free, "flag": flag,
        "cash_bank": cfg["cash_bank"], "wb_liquid": wb_liquid, "cash_live": cash_live,
        "coverage": cash_live / total if total else 0.0,
        "stocks": stocks, "orders": cfg["abroad_orders"],
    }


def render(r: dict) -> str:
    L = []
    a = L.append
    today = dt.date.today().isoformat()
    a(f"# Монитор вложений в товар — {today}\n")
    a(f"## {r['flag']}\n")
    a("## Всего в товаре\n")
    a("| Стадия | Сумма |")
    a("|---|---|")
    a(f"| 🌍 Заказано за границей (в пути) | {rub(r['abroad'])} |")
    a(f"| 📦 На складе | {rub(r['sklad'])} |")
    a(f"| 🛒 На маркетах (товар) | {rub(r['wb'])} |")
    a(f"| **💰 ВСЕГО В ТОВАРЕ** | **{rub(r['total'])}** |")
    a(f"| 🎯 Лимит | {rub(r['limit'])} |")
    label = "Свободный лимит" if r["free"] >= 0 else "Превышение лимита"
    a(f"| **{label}** | **{rub(r['free'])}** |\n")

    a("## Свободные деньги\n")
    a("| Источник | Сумма |")
    a("|---|---|")
    a(f"| Кэш в банках | {rub(r['cash_bank'])} |")
    a(f"| Вывод с WB в понедельник (30%) | {rub(r['wb_liquid'])} |")
    a(f"| **Живой кэш** | **{rub(r['cash_live'])}** |")
    a(f"| Покрытие товара живым кэшом | {r['coverage']*100:.0f}% |\n")

    a("## Разбивка «в товаре» по линейкам\n")
    a("| Линейка | Лист | Склад | Маркеты |")
    a("|---|---|---|---|")
    for s in r["stocks"]:
        a(f"| {s['label']} | {s['sheet']} | {rub(s['sklad'])} | {rub(s['wb']+s['ozon'])} |")
    a("")
    a("## Заказано за границей\n")
    a("| Заказ | Сумма (валюта) | В рублях |")
    a("|---|---|---|")
    for o in r["orders"]:
        amt = f"{o['amount']:,.0f}".replace(",", " ") + f" {o['currency']}"
        a(f"| {o['label']} | {amt} | {rub(order_to_rub(o))} |")
    a("")
    return "\n".join(L)


def main():
    ap = argparse.ArgumentParser(description="Монитор вложений в товар")
    ap.add_argument("--report", action="store_true", help="записать отчёт в outputs/")
    ap.add_argument("--korea-sheet", default=None, help="лист в файле Кореи")
    ap.add_argument("--china-sheet", default=None, help="лист в файле Китая")
    args = ap.parse_args()

    r = build(CONFIG, args.korea_sheet, args.china_sheet)
    text = render(r)
    print(text)

    if args.report:
        OUT_DIR.mkdir(exist_ok=True)
        path = OUT_DIR / f"investment_monitor_{dt.date.today().isoformat()}.md"
        path.write_text(text, encoding="utf-8")
        print(f"\n[отчёт записан] {path}")


if __name__ == "__main__":
    main()
