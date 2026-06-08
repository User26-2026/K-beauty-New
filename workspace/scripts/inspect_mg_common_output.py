import os
import pandas as pd

files = [
    "/Users/nikita/Downloads/2026-05-29_Поисковый_запрос_пэды celimax_Общая выдача_1.xlsx",
    "/Users/nikita/Downloads/2026-05-29_Категория_Корейские бренды_Общая выдача_1.xlsx",
]

articles = [
    712024737,
    712022766,
    155947140,
    228473128,
    187281601,
    394475962,
    611417712,
    712016848,
    583311001,
    712019009,
    200907964,
]

for file_path in files:
    print("FILE", os.path.basename(file_path))
    xl = pd.ExcelFile(file_path)
    print("sheets", xl.sheet_names)
    df = pd.read_excel(file_path, sheet_name=0)
    print("cols", list(df.columns))
    print(df.head(8).to_string())
    artcol = None
    for col in df.columns:
        low = str(col).lower()
        if "артикул" in low or low in {"nm", "id", "nm id"}:
            artcol = col
            break
    print("artcol", artcol)
    if artcol:
        normalized = df[artcol].astype(str).str.replace(".0", "", regex=False)
        for article in articles:
            sub = df[normalized == str(article)]
            if len(sub):
                row = sub.iloc[0].to_dict()
                compact = {
                    key: row[key]
                    for key in row
                    if any(token in str(key).lower() for token in ["товар", "артикул", "упущ", "выруч", "продаж", "заказ", "цена"])
                }
                print("match", article, compact)
