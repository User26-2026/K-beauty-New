import io
import json
import math
import os
import re
import zipfile
from pathlib import Path

import numpy as np
import pandas as pd


DL = Path("/Users/nikita/Downloads")
OUT = Path("/Users/nikita/Desktop/Новая папка 2/outputs/wb_audit")
OUT.mkdir(parents=True, exist_ok=True)


FILES = {
    "cost": DL / "!Корея-Китай-ЮНИТ -ШАБЛОН.xlsx",
    "ad_clusters": DL / "preset-stat-36823214-238415286_2026_5_1-2026_5_27.xlsx",
    "ad_stats": DL / "Статистика.xlsx",
    "weekly_article": DL / "report_2026_5_28.xlsx.XLSX",
    "supplier_goods": DL / "supplier-goods-127314-2026-05-01-2026-05-28-exvxkgdeb.XLSX",
    "geo_sales": DL / "report_2026_5_28.xlsx (1).XLSX",
    "financial": DL / "report 2026-5-28.xlsx",
    "paid_storage": DL / "Отчет по платному хранению за 2026-05-01 - 2026-05-28.xlsx",
    "warehouse_current": DL / "report_2026_5_28.xlsx",
}

ZIPS = {
    "funnel": DL / "воронка с 28.04.2026 по 28.05.2026.zip",
    "stock": DL / "остатки с 28.04.2026 по 28.05.2026.zip",
    "warehouses": DL / "склады с 28.04.2026 по 28.05.2026.zip",
    "localization": DL / "ИЛ с 28.04.2026 по 28.05.2026.zip",
    "audience": DL / "аудитория с 01.05.2026 по 28.05.2026.zip",
    "search": DL / "запросы с 01.05.2026 по 28.05.2026.zip",
    "volume": DL / "report_2026_5_28.zip",
    "stock_index": DL / "ио с 01.05.2026 по 27.05.2026.zip",
}


def zip_xlsx_bytes(path: Path) -> io.BytesIO:
    with zipfile.ZipFile(path) as zf:
        names = [i.filename for i in zf.infolist() if not i.is_dir()]
        if not names:
            raise RuntimeError(f"No files in {path}")
        return io.BytesIO(zf.read(names[0]))


def read_xlsx(source, sheet, header=1):
    return pd.read_excel(source, sheet_name=sheet, header=header)


def as_num(series):
    if getattr(series, "dtype", None) == object:
        series = series.astype(str).str.replace("\u00a0", "", regex=False).str.replace(" ", "", regex=False).str.replace(",", ".", regex=False)
    return pd.to_numeric(series, errors="coerce").fillna(0)


def money(v):
    if pd.isna(v):
        return None
    return round(float(v), 2)


def clean_nm(v):
    if pd.isna(v):
        return np.nan
    s = str(v).strip()
    if not s:
        return np.nan
    m = re.search(r"\d{6,12}", s.replace(".0", ""))
    return int(m.group(0)) if m else np.nan


def load_cost():
    frames = []
    for sheet in ["КОРЕЯ", "КИТАЙ"]:
        df = pd.read_excel(FILES["cost"], sheet_name=sheet, header=8)
        df["source_sheet"] = sheet
        df["nm"] = df.get("Код WB Коротич").map(clean_nm)
        cols = [
            "source_sheet",
            "nm",
            "Наименование",
            "Себ-ть",
        ]
        cols = [c for c in cols if c in df.columns]
        frames.append(df[cols])
    cost = pd.concat(frames, ignore_index=True)
    cost = cost[cost["nm"].notna()].copy()
    cost["nm"] = cost["nm"].astype(int)
    return cost.drop_duplicates("nm", keep="first")


def top(df, by, n=12, cols=None, ascending=False):
    if df.empty:
        return []
    d = df.sort_values(by, ascending=ascending).head(n)
    if cols:
        d = d[cols]
    return d.replace({np.nan: None}).to_dict("records")


summary = {"opened": [], "errors": []}

for key, path in FILES.items():
    try:
        pd.ExcelFile(path).sheet_names
        summary["opened"].append(str(path))
    except Exception as e:
        summary["errors"].append({"file": str(path), "error": str(e)})

zip_sources = {}
for key, path in ZIPS.items():
    try:
        zip_sources[key] = zip_xlsx_bytes(path)
        pd.ExcelFile(zip_sources[key]).sheet_names
        summary["opened"].append(str(path))
        zip_sources[key].seek(0)
    except Exception as e:
        summary["errors"].append({"file": str(path), "error": str(e)})


cost = load_cost()

funnel = read_xlsx(zip_sources["funnel"], "Товары", header=1)
zip_sources["funnel"].seek(0)
funnel_total = read_xlsx(zip_sources["funnel"], "Фильтры", header=1).iloc[0]
funnel["nm"] = as_num(funnel["Артикул WB"]).astype(int)
for c in [
    "Показы",
    "Показы (предыдущий период)",
    "Переходы в карточку",
    "Переходы в карточку (предыдущий период)",
    "Положили в корзину",
    "Заказали товаров, шт",
    "Заказали товаров, шт (предыдущий период)",
    "Выкупили, шт",
    "Отменили, шт",
    "Заказали на сумму, ₽",
    "Выкупили на сумму, ₽",
    "Средняя цена, ₽",
    "Остатки «Склад WB», шт",
    "Сумма остатков на складах, ₽",
    "Локальные заказы, %",
    "Среднее время доставки",
]:
    if c in funnel.columns:
        funnel[c] = as_num(funnel[c])

funnel["order_share"] = funnel["Заказали товаров, шт"] / max(funnel["Заказали товаров, шт"].sum(), 1)
funnel["rev_share"] = funnel["Заказали на сумму, ₽"] / max(funnel["Заказали на сумму, ₽"].sum(), 1)

stock = read_xlsx(zip_sources["stock"], "Детальная информация", header=1)
stock["nm"] = as_num(stock["Артикул WB"]).astype(int)
for c in [
    "Заказали, шт",
    "Выкупили, шт",
    "Остатки на текущий день, шт",
    "Оборачиваемость текущих остатков",
    "Время отсутствия товара на складе",
    "Упущенные заказы, шт",
    "Сумма упущенных заказов, ₽",
    "Текущая цена с учетом скидок, ₽",
]:
    if c in stock.columns:
        stock[c] = as_num(stock[c])
stock_by_nm = stock.groupby("nm", as_index=False).agg(
    stock_qty=("Остатки на текущий день, шт", "sum"),
    missed_orders=("Упущенные заказы, шт", "sum"),
    missed_amount=("Сумма упущенных заказов, ₽", "sum"),
    absent_time=("Время отсутствия товара на складе", "sum"),
    avg_turnover=("Оборачиваемость текущих остатков", "mean"),
    stock_orders=("Заказали, шт", "sum"),
)

il = read_xlsx(zip_sources["localization"], "Детальные данные", header=1)
il["nm"] = as_num(il["Артикул WB"]).astype(int)
for c in [
    "Итого заказов, шт",
    "Итого заказов по товарам локально, шт",
    "Итого заказов по товарам не локально, шт",
    "Остатки склад ВБ, шт",
    "Время доставки",
]:
    il[c] = as_num(il[c])
il_by_nm = il.groupby("nm", as_index=False).agg(
    il_orders=("Итого заказов, шт", "sum"),
    local_orders=("Итого заказов по товарам локально, шт", "sum"),
    nonlocal_orders=("Итого заказов по товарам не локально, шт", "sum"),
    il_stock=("Остатки склад ВБ, шт", "sum"),
    avg_delivery=("Время доставки", "mean"),
)
il_by_nm["local_pct"] = np.where(il_by_nm["il_orders"] > 0, il_by_nm["local_orders"] / il_by_nm["il_orders"] * 100, 0)

il_general = read_xlsx(zip_sources["localization"], "Общее", header=1)
for c in ["Локальные заказы, %", "Не локальные заказы, %", "Доля региона в заказах, %", "Время доставки"]:
    il_general[c] = as_num(il_general[c])

aud = read_xlsx(zip_sources["audience"], "Детализация по артикулам", header=1)
aud["nm"] = as_num(aud["Артикул ВБ"]).astype(int)
for c in ["Показы", "Переходы в карточку", "Добавления в корзину", "Заказы"]:
    aud[c] = as_num(aud[c])
aud_entry = read_xlsx(zip_sources["audience"], "Детализация по точкам входа", header=1)
for c in ["Показы", "Переходы в карточку", "Добавления в корзину", "Заказы"]:
    aud_entry[c] = as_num(aud_entry[c])

search = read_xlsx(zip_sources["search"], "Детальные показатели", header=1)
search["nm"] = as_num(search["Артикул WB"]).astype(int)
for c in [
    "Количество запросов",
    "Видимость, %",
    "Средняя позиция",
    "Медианная позиция",
    "Переходы в карточку",
    "Положили в корзину",
    "Заказали, шт",
    "Конверсия в заказ, %",
]:
    search[c] = as_num(search[c])

search_total = read_xlsx(zip_sources["search"], "Фильтры", header=1).iloc[0]
search_common = read_xlsx(zip_sources["search"], "Общие показатели", header=1)

stock_index = read_xlsx(zip_sources["stock_index"], "Детальная информация", header=1)
stock_index["nm"] = as_num(stock_index["Артикул WB"]).astype(int)
for c in ["Индекс остатка", "O1", "O2"]:
    stock_index[c] = as_num(stock_index[c])
stock_index_by_nm = stock_index.groupby("nm", as_index=False).agg(
    Тип=("Тип", "first"),
    Индекс_остатка=("Индекс остатка", "min"),
    O1=("O1", "sum"),
    O2=("O2", "sum"),
)

supplier = pd.read_excel(FILES["supplier_goods"], sheet_name="Sheet1", header=1)
supplier["nm"] = as_num(supplier["Артикул WB"]).astype(int)
for c in ["шт.", "Сумма заказов минус комиссия WB, руб.", "Выкупили, шт.", "К перечислению за товар, руб.", "Текущий остаток, шт."]:
    supplier[c] = as_num(supplier[c])
supplier_by_nm = supplier.groupby("nm", as_index=False).agg(
    supplier_orders=("шт.", "sum"),
    supplier_order_net_after_commission=("Сумма заказов минус комиссия WB, руб.", "sum"),
    supplier_buyouts=("Выкупили, шт.", "sum"),
    supplier_payout_goods=("К перечислению за товар, руб.", "sum"),
    supplier_current_stock=("Текущий остаток, шт.", "sum"),
)

current_wh = pd.read_excel(FILES["warehouse_current"], sheet_name="Sheet1")
for c in current_wh.columns:
    if c not in ["Бренд", "Предмет"]:
        current_wh[c] = as_num(current_wh[c])

paid_raw = pd.read_excel(FILES["paid_storage"], header=None)
paid_rows = []
for i, row in paid_raw.iterrows():
    if str(row.iloc[1]).strip() == "Рубли":
        paid_rows.append(
            {
                "warehouse": row.iloc[0],
                "total_rub": as_num(pd.Series([row.iloc[2]])).iloc[0],
                "below_base_rub": as_num(pd.Series([row.iloc[4]])).iloc[0],
                "above_base_rub": as_num(pd.Series([row.iloc[5]])).iloc[0],
            }
        )
paid_storage = pd.DataFrame(paid_rows)

ad_clusters = pd.read_excel(FILES["ad_clusters"], sheet_name="Топ поисковых кластеров")
for c in ["Средняя позиция", "Просмотры", "Клики", "CTR", "Добавления в корзину", "Заказов с этим товаром", "Заказанные товары", "Затраты", "CPC"]:
    ad_clusters[c] = as_num(ad_clusters[c])
ad_clusters["CPO_calc"] = np.where(ad_clusters["Заказов с этим товаром"] > 0, ad_clusters["Затраты"] / ad_clusters["Заказов с этим товаром"], np.nan)
ad_stats = pd.read_excel(FILES["ad_stats"], sheet_name="Статистика")

combined = (
    funnel.merge(cost, on="nm", how="left", suffixes=("", "_cost"))
    .merge(supplier_by_nm, on="nm", how="left")
    .merge(stock_by_nm, on="nm", how="left")
    .merge(il_by_nm, on="nm", how="left")
    .merge(stock_index_by_nm, on="nm", how="left")
)
for c in ["stock_qty", "missed_orders", "missed_amount", "local_pct", "avg_delivery", "Индекс_остатка"]:
    if c in combined.columns:
        combined[c] = as_num(combined[c])
for c in ["Себ-ть", "supplier_buyouts", "supplier_payout_goods", "supplier_order_net_after_commission", "supplier_orders"]:
    if c in combined.columns:
        combined[c] = as_num(combined[c])
combined["unit_payout_after_commission"] = np.where(
    combined["supplier_buyouts"] > 0,
    combined["supplier_payout_goods"] / combined["supplier_buyouts"],
    np.nan,
)
combined["cogs_buyouts"] = combined["Себ-ть"] * combined["supplier_buyouts"]
combined["contribution_before_log_storage_ads"] = combined["supplier_payout_goods"] - combined["cogs_buyouts"]
combined["contribution_per_buyout_before_log_storage_ads"] = np.where(
    combined["supplier_buyouts"] > 0,
    combined["contribution_before_log_storage_ads"] / combined["supplier_buyouts"],
    np.nan,
)

cosmetic_subjects = set(
    [
        "Кремы",
        "Тоники",
        "Пенки",
        "Сыворотки",
        "Гидрофильные масла",
        "Патчи",
        "Маски",
        "Косметические наборы для ухода",
        "Корректоры",
        "Скрабы",
        "Кремы для рук",
        "Лосьоны",
        "Эмульсии",
        "Эссенции",
        "Солнцезащитные средства",
    ]
)
combined["is_cosmetic"] = combined["Предмет"].astype(str).str.contains(
    "Крем|Тони|Пенк|Сывор|масл|Патч|Маск|набор|Корректор|Скраб|Лосьон|Эмульс|Эссен|Солнц",
    case=False,
    na=False,
)

overall = {
    "funnel_products": int(len(funnel)),
    "orders": money(funnel_total["Заказали товаров, шт"]),
    "orders_prev": money(funnel_total["Заказали товаров, шт (предыдущий период)"]),
    "order_amount": money(funnel_total["Заказали на сумму, ₽"]),
    "order_amount_prev": money(funnel_total["Заказали на сумму, ₽ (предыдущий период)"]),
    "views": money(funnel_total["Показы"]),
    "views_prev": money(funnel_total["Показы (предыдущий период)"]),
    "card_visits": money(funnel_total["Переходы в карточку"]),
    "card_visits_prev": money(funnel_total["Переходы в карточку (предыдущий период)"]),
    "ctr": money(funnel_total["CTR"]),
    "cart_conv": money(funnel_total["Конверсия в корзину, %"]),
    "order_conv": money(funnel_total["Конверсия в заказ, %"]),
    "buyout_pct": money(funnel_total["Процент выкупа"]),
    "wb_stock_qty": money(funnel["Остатки «Склад WB», шт"].sum()),
    "wb_stock_amount": money(funnel["Сумма остатков на складах, ₽"].sum()),
    "search_views": money(search_total["Показы"]),
    "search_views_prev": money(search_total["Показы (предыдущий период)"]),
    "search_orders": money(search_total["Заказали, шт"]),
    "search_orders_prev": money(search_total["Заказали, шт (предыдущий период)"]),
    "search_avg_position": money(search_common.loc[search_common["Параметр"].eq("Средняя позиция"), "Значение за текущий период"].iloc[0]),
    "search_avg_position_prev": money(search_common.loc[search_common["Параметр"].eq("Средняя позиция"), "Значение за предыдущий период"].iloc[0]),
    "search_visibility": money(search_common.loc[search_common["Параметр"].eq("Видимость, %"), "Значение за текущий период"].iloc[0]),
    "search_visibility_prev": money(search_common.loc[search_common["Параметр"].eq("Видимость, %"), "Значение за предыдущий период"].iloc[0]),
}

overall["local_orders_pct_weighted"] = money(il_by_nm["local_orders"].sum() / max(il_by_nm["il_orders"].sum(), 1) * 100)
overall["il_orders"] = money(il_by_nm["il_orders"].sum())
overall["local_orders"] = money(il_by_nm["local_orders"].sum())
overall["nonlocal_orders"] = money(il_by_nm["nonlocal_orders"].sum())
overall["paid_storage_total"] = money(paid_storage.loc[paid_storage["warehouse"].eq("Все"), "total_rub"].iloc[0])
overall["paid_storage_above_base"] = money(paid_storage.loc[paid_storage["warehouse"].eq("Все"), "above_base_rub"].iloc[0])

ad_overall = {
    "cluster_rows": int(len(ad_clusters)),
    "cluster_spend": money(ad_clusters["Затраты"].sum()),
    "cluster_views": money(ad_clusters["Просмотры"].sum()),
    "cluster_clicks": money(ad_clusters["Клики"].sum()),
    "cluster_orders": money(ad_clusters["Заказов с этим товаром"].sum()),
    "main_campaign_spend": money(pd.to_numeric(ad_stats.loc[ad_stats["Название"].astype(str).str.contains("Всего", na=False), "Затраты, RUB"], errors="coerce").sum()),
}

margin_join = combined[combined["Себ-ть"].notna()].copy()
for c in ["Себ-ть", "supplier_buyouts", "supplier_payout_goods", "supplier_order_net_after_commission", "supplier_orders"]:
    if c in margin_join.columns:
        margin_join[c] = as_num(margin_join[c])
margin_join["unit_payout_after_commission"] = np.where(
    margin_join["supplier_buyouts"] > 0,
    margin_join["supplier_payout_goods"] / margin_join["supplier_buyouts"],
    np.nan,
)
margin_join["cogs_buyouts"] = margin_join["Себ-ть"] * margin_join["supplier_buyouts"]
margin_join["contribution_before_log_storage_ads"] = margin_join["supplier_payout_goods"] - margin_join["cogs_buyouts"]
margin_join["contribution_per_buyout_before_log_storage_ads"] = np.where(
    margin_join["supplier_buyouts"] > 0,
    margin_join["contribution_before_log_storage_ads"] / margin_join["supplier_buyouts"],
    np.nan,
)

low_margin = margin_join[
    (margin_join["supplier_buyouts"] > 0)
    & (margin_join["contribution_per_buyout_before_log_storage_ads"] < 150)
].copy()
negative_margin = margin_join[
    (margin_join["supplier_buyouts"] > 0)
    & (margin_join["contribution_before_log_storage_ads"] < 0)
].copy()
missing_cost_sold = combined[(combined["Заказали товаров, шт"] > 0) & (combined["Себ-ть"].isna())].copy()

problem_local = combined[(combined["Заказали товаров, шт"] >= 10) & (combined["local_pct"] < 40)].copy()
overstock = combined[(combined["Заказали товаров, шт"] <= 2) & (combined["Остатки «Склад WB», шт"] >= 100)].copy()
stockout = combined[(combined["missed_orders"] > 0) | (combined["absent_time"] > 0)].copy()

summary.update(
    {
        "overall": overall,
        "ad_overall": ad_overall,
        "top_products_by_orders": top(
            combined,
            "Заказали товаров, шт",
            20,
            [
                "nm",
                "Артикул продавца",
                "Название",
                "Предмет",
                "Бренд",
                "Показы",
                "CTR",
                "Переходы в карточку",
                "Положили в корзину",
                "Заказали товаров, шт",
                "Заказали на сумму, ₽",
                "Конверсия в заказ, %",
                "Процент выкупа",
                "Остатки «Склад WB», шт",
                "local_pct",
                "Среднее время доставки",
                "Себ-ть",
                "supplier_buyouts",
                "supplier_payout_goods",
                "contribution_per_buyout_before_log_storage_ads",
            ],
        ),
        "top_revenue_products": top(
            combined,
            "Заказали на сумму, ₽",
            15,
            [
                "nm",
                "Артикул продавца",
                "Название",
                "Предмет",
                "Бренд",
                "Заказали товаров, шт",
                "Заказали на сумму, ₽",
                "Выкупили, шт",
                "Средняя цена, ₽",
                "Остатки «Склад WB», шт",
                "local_pct",
                "Себ-ть",
                "supplier_buyouts",
                "supplier_payout_goods",
                "contribution_per_buyout_before_log_storage_ads",
            ],
        ),
        "low_margin_sold": top(
            low_margin,
            "supplier_buyouts",
            20,
            [
                "nm",
                "Артикул продавца",
                "Название",
                "Заказали товаров, шт",
                "Заказали на сумму, ₽",
                "Себ-ть",
                "supplier_buyouts",
                "supplier_payout_goods",
                "unit_payout_after_commission",
                "contribution_per_buyout_before_log_storage_ads",
                "contribution_before_log_storage_ads",
                "Остатки «Склад WB», шт",
            ],
        ),
        "negative_margin_sold": top(
            negative_margin,
            "supplier_buyouts",
            20,
            [
                "nm",
                "Артикул продавца",
                "Название",
                "Заказали товаров, шт",
                "supplier_buyouts",
                "supplier_payout_goods",
                "Себ-ть",
                "unit_payout_after_commission",
                "contribution_per_buyout_before_log_storage_ads",
                "contribution_before_log_storage_ads",
            ],
        ),
        "missing_cost_sold_count": int(len(missing_cost_sold)),
        "problem_local": top(
            problem_local,
            "Заказали товаров, шт",
            20,
            [
                "nm",
                "Артикул продавца",
                "Название",
                "Предмет",
                "Бренд",
                "Заказали товаров, шт",
                "Остатки «Склад WB», шт",
                "local_pct",
                "avg_delivery",
            ],
        ),
        "overstock": top(
            overstock,
            "Остатки «Склад WB», шт",
            20,
            [
                "nm",
                "Артикул продавца",
                "Название",
                "Предмет",
                "Бренд",
                "Заказали товаров, шт",
                "Остатки «Склад WB», шт",
                "Сумма остатков на складах, ₽",
                "Индекс_остатка",
            ],
        ),
        "stockout": top(
            stockout,
            "missed_orders",
            20,
            [
                "nm",
                "Артикул продавца",
                "Название",
                "Заказали товаров, шт",
                "Остатки «Склад WB», шт",
                "missed_orders",
                "missed_amount",
                "absent_time",
            ],
        ),
        "top_search_queries": top(
            search,
            "Заказали, шт",
            30,
            [
                "Артикул WB",
                "Артикул продавца",
                "Название",
                "Бренд",
                "Поисковый запрос",
                "Количество запросов",
                "Видимость, %",
                "Средняя позиция",
                "Переходы в карточку",
                "Положили в корзину",
                "Заказали, шт",
                "Конверсия в заказ, %",
            ],
        ),
        "top_search_missed": top(
            search[(search["Количество запросов"] >= 500) & (search["Заказали, шт"] <= 2)],
            "Количество запросов",
            25,
            [
                "Артикул WB",
                "Артикул продавца",
                "Название",
                "Бренд",
                "Поисковый запрос",
                "Количество запросов",
                "Видимость, %",
                "Средняя позиция",
                "Переходы в карточку",
                "Положили в корзину",
                "Заказали, шт",
                "Конверсия в заказ, %",
            ],
        ),
        "audience_entry": top(
            aud_entry,
            "Заказы",
            20,
            [
                "Раздел",
                "Точка входа",
                "Показы",
                "Переходы в карточку",
                "CTR",
                "Добавления в корзину",
                "Конверсия в корзину",
                "Заказы",
                "Конверсия в заказ",
            ],
        ),
        "il_regions": top(
            il_general,
            "Доля региона в заказах, %",
            15,
            ["Регион", "Локальные заказы, %", "Не локальные заказы, %", "Время доставки", "Доля региона в заказах, %", "Доля всех заказов ВБ, %"],
        ),
        "paid_storage": top(paid_storage[paid_storage["warehouse"].ne("Все")], "total_rub", 20),
        "ad_best_clusters": top(
            ad_clusters[ad_clusters["Заказов с этим товаром"] > 0],
            "Заказов с этим товаром",
            15,
            ["Кластер", "Статус и тип ставки", "Средняя позиция", "Просмотры", "Клики", "CTR", "Добавления в корзину", "Заказов с этим товаром", "Затраты", "CPC", "CPO_calc"],
        ),
        "ad_bad_clusters": top(
            ad_clusters[(ad_clusters["Клики"] >= 5) & (ad_clusters["Заказов с этим товаром"] == 0)],
            "Затраты",
            15,
            ["Кластер", "Статус и тип ставки", "Средняя позиция", "Просмотры", "Клики", "CTR", "Добавления в корзину", "Заказов с этим товаром", "Затраты", "CPC"],
        ),
    }
)

with (OUT / "wb_audit_summary.json").open("w", encoding="utf-8") as f:
    json.dump(summary, f, ensure_ascii=False, indent=2, default=str)

print(json.dumps(summary["overall"], ensure_ascii=False, indent=2))
print("\\nAD", json.dumps(summary["ad_overall"], ensure_ascii=False, indent=2))
print("\\nTOP_PRODUCTS")
for row in summary["top_products_by_orders"][:10]:
    print(row)
print("\\nLOW_MARGIN")
for row in summary["low_margin_sold"][:10]:
    print(row)
print("\\nPROBLEM_LOCAL")
for row in summary["problem_local"][:10]:
    print(row)
print("\\nOVERSTOCK")
for row in summary["overstock"][:10]:
    print(row)
print("\\nSEARCH_MISSED")
for row in summary["top_search_missed"][:10]:
    print(row)
print("\\nAD_BAD")
for row in summary["ad_bad_clusters"][:10]:
    print(row)
