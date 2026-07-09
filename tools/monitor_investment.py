"""Монитор вложений в товар.

Считает сколько денег заморожено в товаре на всех стадиях и сравнивает с лимитом,
чтобы не допускать перевложения:

    Всего в товаре = Заказано за границей (в пути) + Склад + Товар на маркетах

Источники:
  - Заказано за границей   -> задаётся вручную в CONFIG (валюта + курсы).
  - Склад                  -> из ручных выгрузок остатков (data/stock_costs),
                              sum(себестоимость * остаток на складе) по SKU.
  - Товар на маркетах (ВБ) -> для косметики из ЖИВОГО отчёта ВБ по API
                              (колонка "Сумма остатка в себестоимости API").
                              Где в отчёте себестоимость пустая (кабинеты
                              КРОНА/Скляров), подставляем её по Артикулу ВБ из
                              ручной таблицы. Для Китая — из ручной выгрузки.
  - Свободный кэш и баланс WB -> задаются вручную в CONFIG.

Запуск:
    python tools/monitor_investment.py                 # печать дашборда
    python tools/monitor_investment.py --report         # + запись отчёта в outputs/

Зависимости: openpyxl (xlsx), xlrd (для .xls отчёта ВБ).
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

    # Ручные выгрузки остатков: берём отсюда СКЛАД (по всем линейкам) и
    # МАРКЕТЫ по тем линейкам, где нет живого отчёта (Китай).
    "stock_files": [
        {
            "label": "Корея (косметика)",
            "path": STOCK_DIR / "ostatki_korea_full.xlsx",
            "sheet": None,
            "cost_col": 20, "sklad_col": 21,
            "wb_cols": [22, 23, 24], "ozon_cols": [26, 27],
            "art_cols": [1, 2, 3, 4],
            "wb_art_cols": [2, 3, 4],   # Артикул WB (Коротич/Скляров/Крона)
            "use_marketplace": False,   # маркеты Кореи берём из отчёта ВБ по API
        },
        {
            "label": "Китай (товары)",
            "path": STOCK_DIR / "ostatki_china_full.xlsx",
            "sheet": None,
            "cost_col": 10, "sklad_col": 14,
            "wb_cols": [15, 16, 17], "ozon_cols": [19, 20],
            "art_cols": [1, 2, 3],
            "wb_art_cols": [1, 2, 3],
            "use_marketplace": True,    # у Китая живого отчёта нет
        },
    ],

    # Живой отчёт ВБ по API (косметика). col27 = "Сумма остатка в
    # себестоимости API". Пропуски себестоимости заполняем по Артикулу ВБ из
    # соответствующей ручной таблицы (cost_source).
    "wb_api": {
        "label": "ВБ косметика (API)",
        "path": STOCK_DIR / "wb_api_report.xls",
        "art_col": 0, "stock_col": 7, "unit_cost_col": 18, "cost_sum_col": 27,
        "cabinet_names": ("Коротич", "КРОНА", "Скляров"),
        "cost_source": STOCK_DIR / "ostatki_korea_full.xlsx",  # для пропусков
    },
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


def is_num(v) -> bool:
    try:
        float(v)
        return True
    except (TypeError, ValueError):
        return False


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
    sklad_val = mkt_val = 0.0
    n_sklad = n_mkt = 0
    for row in ws.iter_rows(values_only=True):
        def g(c):
            return row[c] if c < len(row) else None
        cost = num(g(cfg["cost_col"]))
        if cost <= 0 or not any(num(g(c)) > 0 for c in cfg["art_cols"]):
            continue
        sk = num(g(cfg["sklad_col"]))
        mkt = sum(num(g(c)) for c in cfg["wb_cols"] + cfg["ozon_cols"])
        sklad_val += cost * sk
        mkt_val += cost * mkt
        if sk > 0:
            n_sklad += 1
        if mkt > 0:
            n_mkt += 1
    return {
        "label": cfg["label"], "sheet": ws.title,
        "sklad": sklad_val, "mkt": mkt_val,
        "n_sklad": n_sklad, "n_mkt": n_mkt,
        "use_marketplace": cfg["use_marketplace"],
    }


def build_art_cost_map(path: Path, cost_col=20, art_cols=(2, 3, 4)) -> dict:
    """Артикул ВБ -> себестоимость из ручной таблицы (для заполнения пропусков)."""
    ws = pick_sheet(openpyxl.load_workbook(path, data_only=True), None)
    m = {}
    for row in ws.iter_rows(values_only=True):
        cost = num(row[cost_col]) if cost_col < len(row) else 0.0
        if cost <= 0:
            continue
        for ac in art_cols:
            a = row[ac] if ac < len(row) else None
            if a is not None and is_num(a):
                m[str(int(float(a)))] = cost
    return m


def analyze_wb_api(cfg: dict) -> dict:
    """Стоимость остатка на ВБ из живого отчёта, с заполнением пропусков цены."""
    import xlrd  # локальный импорт: нужен только для .xls отчёта ВБ
    sh = xlrd.open_workbook(cfg["path"]).sheet_by_index(0)
    cost_map = build_art_cost_map(cfg["cost_source"])
    val = 0.0
    units_total = units_missing = 0.0
    for r in range(sh.nrows):
        a = sh.cell_value(r, cfg["art_col"])
        # строка-заголовок кабинета -> пропускаем
        if isinstance(a, str) and a.strip() in cfg["cabinet_names"]:
            continue
        if not (is_num(a) and str(a).strip()):
            continue
        stock = num(sh.cell_value(r, cfg["stock_col"]))
        unit_cost = num(sh.cell_value(r, cfg["unit_cost_col"]))
        cost_sum = num(sh.cell_value(r, cfg["cost_sum_col"]))
        units_total += stock
        if unit_cost > 0:
            val += cost_sum
        else:
            c = cost_map.get(str(int(float(a))))
            if c:
                val += c * stock
            else:
                units_missing += stock
    return {"label": cfg["label"], "mkt": val,
            "units_total": units_total, "units_missing": units_missing}


def order_to_rub(o: dict) -> float:
    if o["currency"] == "USD":
        return o["amount"] * o["rub_per_usd"]
    if o["currency"] == "KRW":
        return o["amount"] / o["krw_per_usd"] * o["rub_per_usd"]
    raise ValueError(f"неизвестная валюта {o['currency']}")


def rub(x: float) -> str:
    return f"{x:,.0f}".replace(",", " ") + " ₽"


def build(cfg: dict) -> dict:
    stocks = [analyze_stock(sf, None) for sf in cfg["stock_files"]]
    wb_api = analyze_wb_api(cfg["wb_api"])

    abroad = sum(order_to_rub(o) for o in cfg["abroad_orders"])
    sklad = sum(s["sklad"] for s in stocks)
    mkt = wb_api["mkt"] + sum(s["mkt"] for s in stocks if s["use_marketplace"])
    total = abroad + sklad + mkt

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
        "abroad": abroad, "sklad": sklad, "mkt": mkt, "total": total,
        "limit": limit, "free": free, "flag": flag,
        "cash_bank": cfg["cash_bank"], "wb_liquid": wb_liquid, "cash_live": cash_live,
        "coverage": cash_live / total if total else 0.0,
        "stocks": stocks, "wb_api": wb_api, "orders": cfg["abroad_orders"],
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
    a(f"| 🛒 На маркетах (товар) | {rub(r['mkt'])} |")
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

    a("## Разбивка «в товаре»\n")
    a("| Источник | Склад | Маркеты |")
    a("|---|---|---|")
    for s in r["stocks"]:
        mkt = rub(s["mkt"]) if s["use_marketplace"] else "— (см. отчёт ВБ)"
        a(f"| {s['label']} ({s['sheet']}) | {rub(s['sklad'])} | {mkt} |")
    w = r["wb_api"]
    a(f"| {w['label']} | — | {rub(w['mkt'])} |")
    a("")
    if w["units_missing"]:
        a(f"> ⚠️ В отчёте ВБ у {w['units_missing']:,.0f} ед. нет себестоимости "
          f"(не сопоставлено по Артикулу ВБ) — не учтены. "
          f"Кабинет Скляров может отсутствовать в выгрузке.\n".replace(",", " "))

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
    args = ap.parse_args()

    r = build(CONFIG)
    text = render(r)
    print(text)

    if args.report:
        OUT_DIR.mkdir(exist_ok=True)
        path = OUT_DIR / f"investment_monitor_{dt.date.today().isoformat()}.md"
        path.write_text(text, encoding="utf-8")
        print(f"\n[отчёт записан] {path}")


if __name__ == "__main__":
    main()
