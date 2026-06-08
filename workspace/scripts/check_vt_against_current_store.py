from __future__ import annotations

import json
import re
import tempfile
import zipfile
from pathlib import Path

import openpyxl


DOWNLOADS = Path("/Users/nikita/Downloads")
OUT = Path("/Users/nikita/Desktop/Новая папка 2/outputs/wb_product_research")

SOURCE_FILES = [
    DOWNLOADS / "supplier-goods-127314-2026-05-01-2026-05-28-exvxkgdeb.XLSX",
    DOWNLOADS / "28-5-2026 продажи с 01-04-2026 по 24-05-2026.xlsx",
    DOWNLOADS / "продажи с 01.04.2026 по 24.05.2026.zip",
    DOWNLOADS / "юнитка с 01.04.2026 по 24.05.2026.zip",
    DOWNLOADS / "воронка с 28.04.2026 по 28.05.2026.zip",
    DOWNLOADS / "остатки с 28.04.2026 по 28.05.2026.zip",
    DOWNLOADS / "склады с 28.04.2026 по 28.05.2026.zip",
    DOWNLOADS / "!Корея-Китай-ЮНИТ -ШАБЛОН.xlsx",
    DOWNLOADS / "Статистика.xlsx",
    DOWNLOADS / "report_2026_5_28.xlsx",
    DOWNLOADS / "report 2026-5-28.xlsx",
]

TERMS = [
    "vt cosmetics",
    " vt ",
    "re edle",
    "reedle",
    "ридл",
    "микроигл",
    "pdrn",
    "пдрн",
    "cica reti",
    "reti-a",
    "рети",
    "r5 firming",
    "pdrn capsule",
    "pdrn essence",
    "pdrn reedle",
    "lip plumper",
    "cica collagen",
    "grinding cleansing balm",
    "az aha",
    "vita-light",
    "glow ampoule",
    "micro bubble",
]


def clean(value) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).replace("\n", " ")).strip()


def cell_text(row) -> str:
    return " ".join(clean(v) for v in row if clean(v))


def search_xlsx(path: Path) -> list[dict]:
    hits = []
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception:
        return hits
    for ws in wb.worksheets:
        for idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
            text = cell_text(row)
            low = f" {text.lower()} "
            if any(term in low for term in TERMS):
                hits.append(
                    {
                        "file": str(path),
                        "sheet": ws.title,
                        "row": idx,
                        "text": text[:1200],
                    }
                )
                if len(hits) > 300:
                    return hits
    return hits


def iter_input_xlsx() -> list[Path]:
    files = []
    tmp_root = Path(tempfile.mkdtemp(prefix="wb_current_store_"))
    for path in SOURCE_FILES:
        if not path.exists():
            continue
        if path.suffix.lower() in {".xlsx", ".xlsm", ".xls"}:
            files.append(path)
        elif path.suffix.lower() == ".zip":
            out_dir = tmp_root / path.stem
            out_dir.mkdir(parents=True, exist_ok=True)
            try:
                with zipfile.ZipFile(path) as zf:
                    zf.extractall(out_dir)
            except Exception:
                continue
            files.extend(p for p in out_dir.rglob("*") if p.suffix.lower() in {".xlsx", ".xlsm", ".xls"})
    return files


def main() -> None:
    hits = []
    for path in iter_input_xlsx():
        hits.extend(search_xlsx(path))
    OUT.mkdir(parents=True, exist_ok=True)
    (OUT / "vt_current_store_hits.json").write_text(json.dumps(hits, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"hits={len(hits)}")
    for hit in hits[:80]:
        print(f"{Path(hit['file']).name}\t{hit['sheet']}\t{hit['row']}\t{hit['text']}")


if __name__ == "__main__":
    main()
