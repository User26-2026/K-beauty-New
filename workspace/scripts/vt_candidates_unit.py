from __future__ import annotations

import math
import re
from pathlib import Path

import openpyxl
import pandas as pd


SOURCE = Path("/Users/nikita/Downloads/2026-06-02_Поисковый_запрос_VT Cosmetics_Общая выдача_1.xlsx")
PRICE = Path("/Users/nikita/Downloads/VT PRICE LIST(-VAT)_25.05.xlsx")
KRW_PER_1000_RUB = 47.5303


def clean_text(value) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def norm(value) -> str:
    value = clean_text(value).lower().replace("ё", "е")
    return re.sub(r"\s+", " ", re.sub(r"[^a-zа-я0-9.%+]+", " ", value)).strip()


def num(value) -> float:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).replace(" ", "").replace("\xa0", "").replace("—", "")
    if not s:
        return 0.0
    try:
        return float(s.replace(",", "."))
    except ValueError:
        return 0.0


def load_price_catalog() -> dict[str, dict]:
    ws = openpyxl.load_workbook(PRICE, data_only=True, read_only=True)["VT_COSMETICS"]
    catalog: dict[str, dict] = {}
    for row in ws.iter_rows(min_row=5, values_only=True):
        barcode = row[3]
        en_name = clean_text(row[7])
        supply = num(row[16])
        if not en_name or not barcode or not supply:
            continue
        catalog[en_name.upper()] = {
            "barcode": str(int(float(barcode))),
            "price_name": en_name,
            "supply_krw": supply,
        }
    return catalog


PRICE_CATALOG = load_price_catalog()


def price_item(en_name: str) -> dict | None:
    return PRICE_CATALOG.get(en_name.upper())


def classify(title: str) -> tuple[str, str | None]:
    t = norm(title)

    if "spring box" in t or "набор уходовой" in t:
        return "VT Spring box set", None
    if "набор 45" in t:
        return "VT Reedle Shot serum set 45", None

    if "pdrn" in t or "пдрн" in t:
        if "capsule cream" in t or "капсульн" in t:
            return "VT PDRN CAPSULE CREAM 100", "VT PDRN CAPSULE CREAM 100"
        if "stick balm" in t or "стик бальзам" in t or "стик бальзам" in t:
            return "VT PDRN ESSENCE STICK BALM", "VT PDRN ESSENCE STICK BALM"
        if "eye lifter" in t:
            return "VT PDRN REEDLE SHOT EYE LIFTER", "VT PDRN REEDLE SHOT EYE LIFTER"
        if "cleansing balm" in t or "гидрофильный бальзам" in t:
            return "PDRN GRINDING CLEANSING BALM", "PDRN GRINDING CLEANSING BALM"
        if "mint" in t or "hyaluronic" in t or "гиалурон" in t:
            return "VT PDRN HYALURONIC ACID MINT MICRO BUBBLE SERUM", "VT PDRN HYALURONIC ACID MINT MICRO BUBBLE SERUM"
        if "маска" in t or "mask" in t:
            return "VT PDRN HYDROGEL MASK", "VT PDRN HYDROGEL MASK"
        if "glow" in t or "glass skin" in t or "мист" in t:
            return "VT PDRN GLOW AMPOULE", "VT PDRN GLOW AMPOULE"
        if "reedle" in t or "микроиг" in t:
            if "300" in t:
                return "VT PDRN REEDLE SHOT 300", "VT PDRN REEDLE SHOT 300"
            return "VT PDRN REEDLE SHOT 100", "VT PDRN REEDLE SHOT 100"
        if "essence" in t or "эссенц" in t or "сыворот" in t:
            return "VT PDRN ESSENCE 100", "VT PDRN ESSENCE 100"

    if "r5 firming" in t:
        return "VT R5 FIRMING AMPOULE", "VT R5 FIRMING AMPOULE"
    if "retinol collagen pink" in t or ("пузырьков" in t and "ретинол" in t):
        return "VT RETINOL COLLAGEN PINK MICRO BUBBLE SERUM", "VT RETINOL COLLAGEN PINK MICRO BUBBLE SERUM"
    if "lip plumper expert" in t or ("плампер" in t and "expert" in t):
        return "VT REEDLE S LIP PLUMPER EXPERT", "VT REEDLE S LIP PLUMPER EXPERT"
    if "garlic" in t or "чеснок" in t:
        return "VT GARLIC AC REEDLE SHOT SPOT CREAM", "VT GARLIC AC REEDLE SHOT SPOT CREAM"
    if "azela" in t or "азелаин" in t:
        return "VT AZ AHA RED MICRO BUBBLE SERUM", "VT AZ AHA RED MICRO BUBBLE SERUM"
    if "дракон" in t or "red booster" in t:
        return "VT RED BOOSTER REEDLE SHOT 300", "VT RED BOOSTER REEDLE SHOT 300"

    if "cica reti a cream" in t or ("reti a cream" in t and "0.05" in t) or ("ретинол" in t and "cream 0.05" in t):
        return "VT CICA RETI-A CREAM 0.05", "VT CICA RETI-A CREAM 0.05"
    if "cica reti a essence 0.5" in t or ("reti a essence" in t and "0.5" in t):
        return "VT CICA RETI-A ESSENCE 0.5", "VT CICA RETI-A ESSENCE 0.5"
    if "cica collagen eye cream" in t or ("массажный" in t and "коллаген" in t and "век" in t):
        return "VT CICA COLLAGEN EYE CREAM", "VT CICA COLLAGEN EYE CREAM"
    if "cica collagen cream" in t or ("увлажняющий с коллагеном" in t and "cream" in t):
        return "VT CICA COLLAGEN CREAM", "VT CICA COLLAGEN CREAM"

    if "крем для век" in t and "витамином b3" in t:
        return "VT REEDLE SHOT VITA-LIGHT EYE CREAM", "VT REEDLE SHOT VITA-LIGHT EYE CREAM"
    if "крем для век" in t and ("reedle shot" in t or "микроиг" in t):
        return "VT REEDLE SHOT LIFTING EYE CREAM", "VT REEDLE SHOT LIFTING EYE CREAM"

    if "pro cica reedle shot 100" in t or ("спикула" in t and "центелл" in t):
        return "VT PRO CICA REEDLE SHOT 100", "VT PRO CICA REEDLE SHOT 100"
    if "коллагеном" in t and "микроиг" in t:
        return "VT COLLAGEN REEDLE SHOT 100", "VT COLLAGEN REEDLE SHOT 100"
    if "витамином с" in t and "микроиг" in t:
        return "VT VITA-LIGHT REEDLE SHOT 100", "VT VITA-LIGHT REEDLE SHOT 100"
    if ("reedle shot 300" in t or ("микроиг" in t and "300" in t)) and ("30 мл" in t or "30мл" in t or "30 ml" in t):
        return "VT REEDLE SHOT 300 30ML", None
    if "reedle shot 300" in t or ("микроиг" in t and "300" in t):
        return "VT REEDLE SHOT 300", "VT REEDLE SHOT 300"
    if ("reedle shot 100" in t or ("микроиг" in t and "100" in t)) and ("30 мл" in t or "30мл" in t or "30 ml" in t):
        return "VT REEDLE SHOT 100 30ML", None
    if "reedle shot 100" in t or ("микроиг" in t and "100" in t):
        return "VT REEDLE SHOT 100", "VT REEDLE SHOT 100"
    if ("reedle shot 50" in t or "микроиглами 50" in t or "микроиглы 50" in t) and ("30 мл" in t or "30мл" in t or "30 ml" in t):
        return "VT REEDLE SHOT 50 30ML", None
    if "reedle shot 50" in t or (("микроиглами 50" in t or "микроиглы 50" in t) and "50 мл" not in t and "50мл" not in t):
        return "VT REEDLE SHOT 50", "VT REEDLE SHOT 50"

    return clean_text(title), None


def cost_rub_from_krw(supply_krw: float) -> int:
    return int(round((supply_krw / 1000) * KRW_PER_1000_RUB * 1.4))


def buyout_percent(name: str) -> float:
    t = norm(name)
    if any(word in t for word in ["пенк", "cleanser", "бальзам", "cleansing", "мист"]):
        return 96.0
    return 92.19


def unit_calc(price_spp: float, cost_rub: int, name: str) -> tuple[float, int, float]:
    price_wb = round(price_spp / 0.70)
    price_spp_calc = price_wb * 0.70
    buyout = buyout_percent(name)
    margin = (
        price_wb
        - 60
        - price_spp_calc * 0.08
        - 90
        - price_wb * 0.325
        - price_spp_calc * 0.0217
        - 50 * ((100 - buyout) / 100)
        - 4.8
        - cost_rub
    )
    roi = margin / cost_rub * 100 if cost_rub else None
    roi10 = (margin - price_wb * 0.10) / cost_rub * 100 if cost_rub else None
    return round(roi, 1), int(round(margin)), round(roi10, 1)


def fmt(value) -> str:
    if value is None or value == "" or (isinstance(value, float) and math.isnan(value)):
        return ""
    if isinstance(value, float):
        if abs(value - round(value)) < 1e-9:
            return str(int(round(value)))
        return str(value).replace(".", ",")
    return str(value)


def main() -> None:
    df = pd.read_excel(SOURCE)
    df["revenue"] = df["Расч. выручка\nпо заказам (с СПП)"].apply(num)
    df["orders"] = df["Расч. заказы"].apply(num)
    df["price_spp"] = df["Цена (с СПП)"].apply(num)
    df = df[df["Бренд"].astype(str).str.upper().str.contains("VT")]
    df = df[df["revenue"] >= 100_000].copy()
    df[["canonical", "price_name"]] = df["Товар"].apply(lambda x: pd.Series(classify(str(x))))
    df = df.sort_values("revenue", ascending=False).drop_duplicates("canonical", keep="first")

    rows = []
    for _, row in df.iterrows():
        price_name = row["price_name"]
        price = price_item(price_name) if price_name else None
        barcode = price["barcode"] if price else ""
        cost = cost_rub_from_krw(price["supply_krw"]) if price else ""
        roi = margin = roi10 = ""
        if cost and row["price_spp"]:
            roi, margin, roi10 = unit_calc(row["price_spp"], cost, row["Товар"])
        rows.append(
            [
                "VT COSMETICS",
                int(row["Артикул"]),
                clean_text(row["Товар"]),
                barcode,
                int(round(row["revenue"])),
                int(round(row["orders"])),
                roi,
                margin,
                int(round(row["price_spp"])),
                cost,
                roi10,
            ]
        )

    rows.sort(key=lambda r: r[4], reverse=True)
    headers = [
        "Бренд",
        "Ссылка/артикул",
        "Название",
        "Баркод из прайса",
        "Расч. выручка топ карточки, руб/мес",
        "Расч. заказы топ карточки, шт/мес",
        "ROI без ДРР, %",
        "Маржа/шт без ДРР, руб.",
        "Цена с СПП топ карточки, ₽",
        "Себестоимость, ₽",
        "ROI при ДРР 10%, %",
    ]
    print("\t".join(headers))
    for row in rows:
        print("\t".join(fmt(v) for v in row))


if __name__ == "__main__":
    main()
