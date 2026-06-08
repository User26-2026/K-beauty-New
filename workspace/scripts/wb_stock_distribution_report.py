import io
import math
import zipfile
from pathlib import Path

import numpy as np
import pandas as pd


DL = Path("/Users/nikita/Downloads")
OUT = Path("/Users/nikita/Desktop/Новая папка 2/outputs/wb_audit")
OUT.mkdir(parents=True, exist_ok=True)

STOCK_ZIP = DL / "остатки с 28.04.2026 по 28.05.2026.zip"
IL_ZIP = DL / "ИЛ с 28.04.2026 по 28.05.2026.zip"
PROFIT_CSV = OUT / "sales_profit_with_cogs_2026-04-01_2026-05-24.csv"
STORAGE_CSV = OUT / "paid_storage_by_sku_exact_2026-04-01_2026-05-28.csv"
IRP_CSV = OUT / "irp_overpay_by_sku_actual_0_60pct.csv"

PERIOD_DAYS = 31
TARGET_COVER_DAYS = 30
SAFETY_FACTOR = 1.2


def zip_xlsx_bytes(path: Path) -> io.BytesIO:
    with zipfile.ZipFile(path) as zf:
        files = [i.filename for i in zf.infolist() if not i.is_dir()]
        if not files:
            raise RuntimeError(f"No files in {path}")
        return io.BytesIO(zf.read(files[0]))


def num(s):
    return pd.to_numeric(s, errors="coerce").fillna(0)


def money(v):
    if pd.isna(v):
        return ""
    return f"{v:,.0f}".replace(",", " ")


def qty(v):
    if pd.isna(v):
        return ""
    return f"{v:,.0f}".replace(",", " ")


def main():
    stock_src = zip_xlsx_bytes(STOCK_ZIP)
    stock = pd.read_excel(stock_src, sheet_name="Детальная информация", header=1)
    stock["nm"] = num(stock["Артикул WB"]).astype("Int64")
    for c in [
        "Заказали, шт",
        "Выкупили, шт",
        "Остатки на текущий день, шт",
        "Стоимость остатков на текущий день, ₽",
        "Упущенные заказы, шт",
        "Сумма упущенных заказов, ₽",
        "Упущенные выкупы, шт",
        "Сумма упущенных выкупов, ₽",
    ]:
        stock[c] = num(stock[c])

    # WB repeats the same missed-buyout estimate on many warehouse rows.
    # Use one SKU-level value to avoid multiplying one lost-sales estimate by warehouse count.
    stock_sku = (
        stock.groupby(["nm", "Артикул продавца", "Название", "Предмет", "Бренд"], dropna=False)
        .agg(
            ordered_units=("Заказали, шт", "sum"),
            bought_units=("Выкупили, шт", "sum"),
            current_stock=("Остатки на текущий день, шт", "sum"),
            missed_buyouts_raw_sum=("Упущенные выкупы, шт", "sum"),
            missed_buyouts_dedup=("Упущенные выкупы, шт", "max"),
            missed_buyouts_amount_dedup=("Сумма упущенных выкупов, ₽", "max"),
            deficit_warehouse_rows=("Склад", "size"),
        )
        .reset_index()
    )

    il_src = zip_xlsx_bytes(IL_ZIP)
    il = pd.read_excel(il_src, sheet_name="Детальные данные", header=1)
    il["nm"] = num(il["Артикул WB"]).astype("Int64")
    for c in [
        "Итого заказов, шт",
        "Итого заказов по товарам локально, шт",
        "Итого заказов по товарам не локально, шт",
        "Остатки склад ВБ, шт",
    ]:
        il[c] = num(il[c])

    profit = pd.read_csv(PROFIT_CSV)
    profit["nm"] = num(profit["nm"]).astype("Int64")
    profit = profit[
        [
            "nm",
            "cost",
            "net_units",
            "profit_per_unit_after_cogs_before_storage_ads",
            "profit_after_cogs_before_storage_ads",
        ]
    ].drop_duplicates("nm")

    storage = pd.read_csv(STORAGE_CSV)
    storage["nm"] = num(storage["Артикул WB"]).astype("Int64")
    storage = storage[["nm", "Категория", "Хранение_период", "Хранение_месяц30"]].drop_duplicates("nm")

    irp = pd.read_csv(IRP_CSV)
    irp["nm"] = num(irp["Артикул WB"]).astype("Int64")
    irp = irp[["nm", "Косметика", "Переплата_ИРП_период", "Переплата_ИРП_месяц"]].drop_duplicates("nm")

    sku = stock_sku.merge(profit, on="nm", how="left").merge(storage, on="nm", how="left").merge(irp, on="nm", how="left")
    sku["storage_per_sold_unit"] = np.where(
        sku["net_units"].fillna(0) > 0,
        sku["Хранение_период"].fillna(0) / sku["net_units"].replace(0, np.nan),
        0,
    )
    sku["profit_per_unit_after_storage_before_ads"] = (
        sku["profit_per_unit_after_cogs_before_storage_ads"] - sku["storage_per_sold_unit"]
    )
    sku["missed_profit_dedup"] = (
        sku["missed_buyouts_dedup"] * sku["profit_per_unit_after_storage_before_ads"]
    )
    sku["known_positive_economy"] = sku["profit_per_unit_after_storage_before_ads"].gt(0)
    sku.to_csv(OUT / "stock_opportunity_by_sku_dedup_2026-04-28_2026-05-28.csv", index=False)

    region = il.merge(
        sku[
            [
                "nm",
                "Артикул продавца",
                "Название",
                "Предмет",
                "Бренд",
                "Категория",
                "current_stock",
                "missed_buyouts_dedup",
                "missed_buyouts_amount_dedup",
                "profit_per_unit_after_storage_before_ads",
                "missed_profit_dedup",
                "Переплата_ИРП_месяц",
                "Косметика",
                "known_positive_economy",
            ]
        ],
        on="nm",
        how="left",
        suffixes=("", "_sku"),
    )

    region["local_pct"] = np.where(
        region["Итого заказов, шт"] > 0,
        region["Итого заказов по товарам локально, шт"] / region["Итого заказов, шт"] * 100,
        0,
    )
    region["nonlocal_pct"] = 100 - region["local_pct"]
    region["target_stock_30d"] = np.ceil(
        region["Итого заказов, шт"] / PERIOD_DAYS * TARGET_COVER_DAYS * SAFETY_FACTOR
    )
    region["recommended_supply_qty"] = np.maximum(
        region["target_stock_30d"] - region["Остатки склад ВБ, шт"],
        0,
    )
    # If the region had almost no demand, do not create noise from tiny international rows.
    region.loc[region["Итого заказов, шт"] < 10, "recommended_supply_qty"] = 0

    nonlocal_total = region.groupby("nm")["Итого заказов по товарам не локально, шт"].transform("sum")
    demand_total = region.groupby("nm")["Итого заказов, шт"].transform("sum")
    region["nonlocal_share"] = np.where(nonlocal_total > 0, region["Итого заказов по товарам не локально, шт"] / nonlocal_total, 0)
    region["demand_share"] = np.where(demand_total > 0, region["Итого заказов, шт"] / demand_total, 0)
    region["allocated_missed_buyouts"] = region["missed_buyouts_dedup"].fillna(0) * region["demand_share"]
    region["allocated_lost_profit"] = (
        region["allocated_missed_buyouts"] * region["profit_per_unit_after_storage_before_ads"]
    )
    region["profit_if_recommended_qty_sells"] = (
        region["recommended_supply_qty"] * region["profit_per_unit_after_storage_before_ads"]
    )
    region["first_wave_14d_qty"] = np.ceil(region["recommended_supply_qty"] * 14 / TARGET_COVER_DAYS)
    region["allocated_irp_saving_month"] = region["Переплата_ИРП_месяц"].fillna(0) * region["nonlocal_share"]
    region["priority_score"] = (
        region["Итого заказов по товарам не локально, шт"].fillna(0) * 2
        + region["Итого заказов, шт"].fillna(0)
        + region["recommended_supply_qty"].fillna(0) * 0.5
    )

    # Keep cosmetics and any SKU with known positive economics. Non-cosmetics stay out of owner-facing supply plan.
    cosmetics_mask = region["Категория"].eq("Красота") | region["Косметика"].fillna(False).astype(bool)
    recommended = region[
        cosmetics_mask
        & (region["recommended_supply_qty"] > 0)
        & (region["Итого заказов по товарам не локально, шт"] >= 10)
    ].copy()
    recommended = recommended.sort_values(["priority_score", "allocated_lost_profit"], ascending=False)

    cols = [
        "nm",
        "Артикул продавца",
        "Название_sku",
        "Предмет_sku",
        "Бренд_sku",
        "Регион",
        "Итого заказов, шт",
        "Итого заказов по товарам локально, шт",
        "Итого заказов по товарам не локально, шт",
        "local_pct",
        "Остатки склад ВБ, шт",
        "target_stock_30d",
        "recommended_supply_qty",
        "first_wave_14d_qty",
        "profit_per_unit_after_storage_before_ads",
        "profit_if_recommended_qty_sells",
        "allocated_lost_profit",
        "allocated_irp_saving_month",
        "missed_profit_dedup",
        "missed_buyouts_dedup",
        "missed_buyouts_amount_dedup",
        "Переплата_ИРП_месяц",
        "Время доставки",
    ]
    recommended[cols].rename(
        columns={
            "nm": "Артикул WB",
            "Название_sku": "Название",
            "Предмет_sku": "Предмет",
            "Бренд_sku": "Бренд",
            "local_pct": "Локальность_региона_%",
            "target_stock_30d": "Целевой_остаток_30дней_с_буфером",
            "recommended_supply_qty": "Рекомендовано_поставить_шт",
            "first_wave_14d_qty": "Первая_волна_14дней_шт",
            "profit_per_unit_after_storage_before_ads": "Прибыль_на_шт_до_рекламы",
            "profit_if_recommended_qty_sells": "Прибыль_если_рекомендованный_объем_продастся",
            "allocated_lost_profit": "Потенциал_прибыли_от_возврата_спроса",
            "allocated_irp_saving_month": "Потенциал_экономии_ИРП_мес",
            "missed_profit_dedup": "Потенциал_прибыли_SKU_по_упущенным_выкупам",
            "missed_buyouts_dedup": "Упущенные_выкупы_SKU_очищено",
            "missed_buyouts_amount_dedup": "Упущенные_выкупы_руб_SKU_очищено",
            "Переплата_ИРП_месяц": "Переплата_ИРП_SKU_мес",
        }
    ).to_csv(OUT / "supply_recommendations_by_sku_region_2026-05-28.csv", index=False)

    by_sku = (
        recommended.groupby(["nm", "Артикул продавца", "Название_sku", "Предмет_sku", "Бренд_sku"], dropna=False)
        .agg(
            recommended_supply_qty=("recommended_supply_qty", "sum"),
            first_wave_14d_qty=("first_wave_14d_qty", "sum"),
            regions=("Регион", lambda s: ", ".join(s.head(5))),
            nonlocal_orders=("Итого заказов по товарам не локально, шт", "sum"),
            current_wb_stock_regions=("Остатки склад ВБ, шт", "sum"),
            allocated_lost_profit=("allocated_lost_profit", "sum"),
            profit_if_recommended_qty_sells=("profit_if_recommended_qty_sells", "sum"),
            allocated_irp_saving_month=("allocated_irp_saving_month", "sum"),
            profit_per_unit=("profit_per_unit_after_storage_before_ads", "first"),
            missed_profit_sku=("missed_profit_dedup", "first"),
            missed_buyouts_sku=("missed_buyouts_dedup", "first"),
            missed_buyouts_amount_sku=("missed_buyouts_amount_dedup", "first"),
            irp_saving_sku_month=("Переплата_ИРП_месяц", "first"),
        )
        .reset_index()
        .sort_values(["recommended_supply_qty", "allocated_lost_profit"], ascending=False)
    )
    by_sku.to_csv(OUT / "supply_recommendations_by_sku_summary_2026-05-28.csv", index=False)

    top_sku = by_sku.head(12).copy()
    top_region = recommended.head(25).copy()
    known = sku[(sku["Категория"].eq("Красота") | sku["Косметика"].fillna(False).astype(bool)) & sku["known_positive_economy"]]
    unknown = sku[(sku["Категория"].eq("Красота") | sku["Косметика"].fillna(False).astype(bool)) & sku["profit_per_unit_after_storage_before_ads"].isna()]

    md = []
    md.append("# Отчет: куда поставлять товар из-за дефицита по регионам")
    md.append("")
    md.append("Период источников: остатки и ИЛ за 28.04.2026-28.05.2026, экономика за 01.04.2026-24.05.2026.")
    md.append("")
    md.append("## Методика")
    md.append("")
    md.append("- `Упущенные выкупы` в отчете WB по складам повторяются одинаковой суммой на нескольких складах одного SKU. Для оценки денег берем очищенный SKU-потенциал один раз, а не сумму всех складских строк.")
    md.append(f"- Целевой остаток: спрос региона за период, пересчитанный на {TARGET_COVER_DAYS} дней, с буфером {round((SAFETY_FACTOR - 1) * 100)}%.")
    md.append("- Прибыль на штуку: после комиссии WB, эквайринга, логистики, себестоимости и хранения, до рекламы.")
    md.append("- Экономия ИРП показана отдельно: это не выручка, а потенциальное снижение логистических переплат при улучшении локальности.")
    md.append("")
    md.append("## Ключевые цифры")
    md.append("")
    md.append(f"- Известная положительная экономика по косметике: {len(known)} SKU.")
    md.append(f"- Очищенный потенциал прибыли по упущенным выкупам: {money(known['missed_profit_dedup'].clip(lower=0).sum())} ₽ за период.")
    md.append(f"- Рекомендованный объем региональной дозагрузки по строкам с явным нелокальным спросом: {qty(recommended['recommended_supply_qty'].sum())} шт.")
    md.append(f"- Первая аккуратная волна на 14 дней: {qty(recommended['first_wave_14d_qty'].sum())} шт.")
    md.append(f"- Потенциальная чистая прибыль при продаже рекомендованного объема: {money(recommended['profit_if_recommended_qty_sells'].clip(lower=0).sum())} ₽ до рекламы.")
    md.append(f"- Потенциальная экономия ИРП в месяц по рекомендованным строкам: {money(recommended['allocated_irp_saving_month'].sum())} ₽/мес.")
    md.append("")
    if not unknown.empty:
        md.append("## Где не хватает себестоимости")
        md.append("")
        md.append("По этим SKU есть спрос/дефицит, но чистую прибыль посчитать нельзя без себестоимости:")
        unknown_with_missed = unknown[unknown["missed_buyouts_amount_dedup"].fillna(0) > 0]
        for _, r in unknown_with_missed.sort_values("missed_buyouts_amount_dedup", ascending=False).head(8).iterrows():
            md.append(f"- {r['Артикул продавца']}, nm {int(r['nm'])}: упущенные выкупы WB {qty(r['missed_buyouts_dedup'])} шт / {money(r['missed_buyouts_amount_dedup'])} ₽.")
        md.append("")
    md.append("## Топ SKU для дозагрузки")
    md.append("")
    md.append("| SKU | Товар | Куда в первую очередь | 30 дней, шт | 14 дней, шт | Прибыль/шт | Прибыль при продаже объема | ИРП, ₽/мес |")
    md.append("|---:|---|---|---:|---:|---:|---:|---:|")
    for _, r in top_sku.iterrows():
        md.append(
            f"| {int(r['nm'])} | {r['Название_sku']} | {r['regions']} | {qty(r['recommended_supply_qty'])} | "
            f"{qty(r['first_wave_14d_qty'])} | {money(r['profit_per_unit'])} | {money(max(r['profit_if_recommended_qty_sells'], 0) if pd.notna(r['profit_if_recommended_qty_sells']) else np.nan)} | {money(r['irp_saving_sku_month'])} |"
        )
    md.append("")
    md.append("## Топ региональных действий")
    md.append("")
    md.append("| SKU | Регион | Заказы | Нелокальные | Остаток WB | 30 дней, шт | 14 дней, шт | Прибыль при продаже | Экономия ИРП/мес |")
    md.append("|---:|---|---:|---:|---:|---:|---:|---:|---:|")
    for _, r in top_region.iterrows():
        md.append(
            f"| {int(r['nm'])} | {r['Регион']} | {qty(r['Итого заказов, шт'])} | "
            f"{qty(r['Итого заказов по товарам не локально, шт'])} | {qty(r['Остатки склад ВБ, шт'])} | "
            f"{qty(r['recommended_supply_qty'])} | {qty(r['first_wave_14d_qty'])} | {money(max(r['profit_if_recommended_qty_sells'], 0) if pd.notna(r['profit_if_recommended_qty_sells']) else np.nan)} | "
            f"{money(r['allocated_irp_saving_month'])} |"
        )
    md.append("")
    md.append("## Управленческий вывод")
    md.append("")
    md.append("Главная задача — не просто увеличить общий остаток, а переложить запас в регионы с высоким нелокальным спросом: Дальневосточный и Сибирский, Южный и Северо-Кавказский, Приволжский, Северо-Западный и Уральский. Для товаров с хорошей маржой дозагрузка должна идти до запуска масштабной рекламы, иначе реклама поднимет спрос и снова упрется в локальный дефицит.")

    report_path = OUT / "stock_distribution_supply_report_2026-05-28.md"
    report_path.write_text("\n".join(md), encoding="utf-8")
    print(report_path)
    print(OUT / "supply_recommendations_by_sku_region_2026-05-28.csv")
    print(OUT / "supply_recommendations_by_sku_summary_2026-05-28.csv")


if __name__ == "__main__":
    main()
