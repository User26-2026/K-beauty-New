import csv
import io
import subprocess
from pathlib import Path

import pandas as pd


BASE = Path("/Users/nikita/Desktop/Новая папка 2")
PYTHON = "/Users/nikita/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3"


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


def aggregate(config):
    raw = subprocess.check_output([PYTHON, str(config["script"])], cwd=BASE, text=True)
    matched = pd.DataFrame(list(csv.DictReader(io.StringIO(raw), delimiter="\t")))
    if matched.empty:
        return pd.DataFrame()

    source = pd.read_excel(config["source"])
    source = source[source["Бренд"].astype(str).str.upper().eq(config["brand"])]
    source = source.copy()
    source["Артикул WB"] = source["Артикул"].astype(int).astype(str)
    source["Продажи карточки"] = pd.to_numeric(source["Расч. заказы"], errors="coerce").fillna(0).astype(int)
    source = source[["Артикул WB", "Продажи карточки"]]

    data = matched.merge(source, on="Артикул WB", how="left")
    data["Продажи карточки"] = data["Продажи карточки"].fillna(0).astype(int)
    data = data[data["Barcode из прайса"].astype(str).str.len() > 0].copy()

    rows = []
    for barcode, group in data.groupby("Barcode из прайса", sort=False):
        group = group.sort_values("Продажи карточки", ascending=False)
        statuses = set(clean(v) for v in group["Сопоставление"].tolist())
        if all(s.startswith("точно") for s in statuses):
            status = "точно"
        elif any("вероят" in s for s in statuses):
            status = "есть вероятные совпадения"
        else:
            status = "; ".join(sorted(statuses))[:120]
        rows.append(
            {
                "Бренд": config["brand"],
                "Barcode": barcode,
                "Уникальный товар": clean(group.iloc[0]["Название в прайсе"]),
                "Supply KRW": group.iloc[0]["Supply KRW"],
                "Продажи, шт/мес": int(group["Продажи карточки"].sum()),
                "Топ артикул WB": group.iloc[0]["Артикул WB"],
                "Карточек в выдаче": int(group["Артикул WB"].nunique()),
                "Сопоставление": status,
            }
        )
    return pd.DataFrame(rows).sort_values("Продажи, шт/мес", ascending=False)


def main():
    frames = [aggregate(config) for config in CONFIGS]
    result = pd.concat(frames, ignore_index=True)
    cols = [
        "Бренд",
        "Barcode",
        "Уникальный товар",
        "Supply KRW",
        "Продажи, шт/мес",
        "Топ артикул WB",
        "Карточек в выдаче",
        "Сопоставление",
    ]
    print("\t".join(cols))
    for _, row in result[cols].iterrows():
        print("\t".join(clean(row[col]) for col in cols))


if __name__ == "__main__":
    main()
