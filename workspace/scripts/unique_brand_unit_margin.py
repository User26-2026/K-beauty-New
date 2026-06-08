import csv
import io
import subprocess
from pathlib import Path

import openpyxl
import pandas as pd


BASE = Path("/Users/nikita/Desktop/Новая папка 2")
PYTHON = "/Users/nikita/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3"
CELIMAX_UNIT = Path("/Users/nikita/Downloads/Celimax.xlsx")
KRW_PER_1000_RUB = 47.5303

CONFIGS = [
    {
        "brand": "DERMA FACTORY",
        "source": Path("/Users/nikita/Downloads/2026-06-01_Поисковый_запрос_дерма фактори_Общая выдача_1.xlsx"),
        "script": BASE / "scripts/match_dermafactory_barcodes.py",
    },
    {
        "brand": "AXIS-Y",
        "source": Path("/Users/nikita/Downloads/2026-06-01_Поисковый_запрос_axis-y_Общая выдача_1.xlsx"),
        "script": BASE / "scripts/match_axisy_barcodes.py",
    },
]


def clean(value):
    return " ".join(str(value or "").split())


def load_constants():
    wb = openpyxl.load_workbook(CELIMAX_UNIT, data_only=True)
    unit = wb["Юнит"]
    return {
        "handling": float(unit["B1"].value),
        "ads": float(unit["B2"].value),
        "storage_per_liter_day": float(unit["B6"].value),
        "storage_days": float(unit["B7"].value),
        "tax": float(unit["B8"].value),
        "reverse_logistics": float(unit["S3"].value),
        "commission": float(unit["S6"].value),
        "acquiring": float(unit["S7"].value),
        "spp": 0.30,
        "logistics": 90.0,
    }


def volume_liters(name, subject):
    low = f"{name} {subject}".lower()
    if any(x in low for x in ["тонер", "тоник", "пенк", "cleanser", "маска", "mist", "мист"]):
        return 1.0
    return 0.4


def buyout_percent(name, subject):
    low = f"{name} {subject}".lower()
    if any(x in low for x in ["пенк", "cleanser", "мист", "тонер", "тоник"]):
        return 96.0
    return 92.19


def calc_margin(price_spp, supply_krw, name, subject, c):
    try:
        price_spp = float(price_spp)
        supply_krw = float(str(supply_krw).replace(",", "."))
    except Exception:
        return None, None
    if price_spp <= 0 or supply_krw <= 0:
        return None, None

    cost_rub = round(supply_krw / 1000 * KRW_PER_1000_RUB) * 1.4
    price_wb = round(price_spp / (1 - c["spp"]))
    price_spp_calc = price_wb * (1 - c["spp"])
    tax = price_spp_calc * c["tax"]
    commission = price_wb * c["commission"]
    acquiring = price_spp_calc * c["acquiring"]
    reverse = c["reverse_logistics"] * ((100 - buyout_percent(name, subject)) / 100)
    storage = c["storage_per_liter_day"] * c["storage_days"] * volume_liters(name, subject)
    ads = price_wb * c["ads"]
    margin = (
        price_wb
        - c["handling"]
        - tax
        - c["logistics"]
        - commission
        - acquiring
        - reverse
        - storage
        - ads
        - cost_rub
    )
    margin_pct = margin / price_wb if price_wb else None
    return margin, margin_pct


def aggregate(config, constants):
    matched_raw = subprocess.check_output([PYTHON, str(config["script"])], cwd=BASE, text=True)
    matched = pd.DataFrame(list(csv.DictReader(io.StringIO(matched_raw), delimiter="\t")))
    source = pd.read_excel(config["source"])
    source = source[source["Бренд"].astype(str).str.upper().eq(config["brand"])].copy()
    source["Артикул WB"] = source["Артикул"].astype(int).astype(str)
    source["Продажи карточки"] = pd.to_numeric(source["Расч. заказы"], errors="coerce").fillna(0).astype(int)
    source["Цена карточки"] = pd.to_numeric(source["Цена (с СПП)"], errors="coerce")

    merged = matched.merge(
        source[["Артикул WB", "Товар", "Предмет", "Продажи карточки", "Цена карточки"]],
        on="Артикул WB",
        how="left",
    )
    merged = merged[merged["Barcode из прайса"].astype(str).str.len() > 0].copy()

    rows = []
    for barcode, group in merged.groupby("Barcode из прайса", sort=False):
        group = group.sort_values("Продажи карточки", ascending=False)
        top = group.iloc[0]
        margin, margin_pct = calc_margin(
            top["Цена карточки"],
            top["Supply KRW"],
            top["Товар"],
            top["Предмет"],
            constants,
        )
        rows.append(
            {
                "Бренд": config["brand"],
                "Артикул WB": top["Артикул WB"],
                "Название": clean(top["Товар"]),
                "Баркод из прайса": barcode,
                "Продажи топ карточки, шт/мес": int(top["Продажи карточки"]),
                "Маржинальность, %": "" if margin_pct is None else round(margin_pct * 100, 1),
                "Маржа/шт, руб.": "" if margin is None else round(margin),
                "Продажи всего товара, шт/мес": int(group["Продажи карточки"].sum()),
            }
        )
    return pd.DataFrame(rows).sort_values("Продажи топ карточки, шт/мес", ascending=False)


def main():
    constants = load_constants()
    result = pd.concat([aggregate(cfg, constants) for cfg in CONFIGS], ignore_index=True)
    cols = [
        "Бренд",
        "Артикул WB",
        "Название",
        "Баркод из прайса",
        "Продажи топ карточки, шт/мес",
        "Маржинальность, %",
        "Маржа/шт, руб.",
    ]
    print("\t".join(cols))
    for _, row in result[cols].iterrows():
        print("\t".join(clean(row[col]) for col in cols))


if __name__ == "__main__":
    main()
