import glob
import os
import warnings

import pandas as pd

warnings.filterwarnings("ignore")

articles = [
    712024737,
    712022766,
    155947140,
    228473128,
    187281601,
    394475962,
    263730548,
    240014011,
    583311001,
    712019009,
]

files = glob.glob("/Users/nikita/Downloads/*.xlsx") + glob.glob("/Users/nikita/Downloads/*.XLSX")

for article in articles:
    print(f"ARTICLE {article}")
    for file_path in files:
        try:
            xl = pd.ExcelFile(file_path)
        except Exception:
            continue
        for sheet_name in xl.sheet_names[:10]:
            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
            except Exception:
                continue
            article_cols = [
                col
                for col in df.columns
                if "артикул" in str(col).lower() or str(col).lower() in {"nm", "nm id", "id"}
            ]
            if not article_cols:
                continue
            for article_col in article_cols[:2]:
                normalized = df[article_col].astype(str).str.replace(".0", "", regex=False)
                sub = df[normalized == str(article)]
                if sub.empty:
                    continue
                row = sub.iloc[0]
                values = {}
                for col in df.columns:
                    name = str(col)
                    low = name.lower()
                    if any(token in low for token in ["товар", "выруч", "упущ", "улучш", "налич", "заказ", "цена"]):
                        values[name] = row[col]
                print(os.path.basename(file_path), sheet_name, values)
