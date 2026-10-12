#!/usr/bin/env python3
"""Схема загрузки сорокафутового контейнера, вид сбоку и сверху.

Рисуем то, что посчитано в build_container_spb: клинкеты Ду200 и выше
одним слоем на настил, поверх и рядом ящики с мелочью, фильтры верхом.
Масштаб честный, 100 пикселей на метр, поэтому по картинке видно, где
на самом деле остается место.

Результат: outputs/marine_equipment/Схема_загрузки_40ф.svg
"""

import pathlib
import sys

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
from build_container_spb import (  # noqa: E402
    BOX_KG, FLOOR_AREA, FLOOR_L, FLOOR_W, asrow, collect,
)
from build_shipment_spb import item_kind, tier  # noqa: E402

OUT = pathlib.Path(__file__).resolve().parents[2] / "outputs" / "marine_equipment"

# внутренняя высота сорокафутового
HEIGHT = 2.393
BOX_L, BOX_W, BOX_H = 1.2, 1.0, 0.91
BOARD = 0.025

PX = 100          # пикселей на метр
PAD = 70          # поля вокруг рисунка
GAP = 120         # между видами

INK = "#1F3864"
MUTED = "#6b7280"
GATE = "#C9A227"      # клинкеты
GATE_D = "#9c7d12"
BOX1 = "#8FB996"      # ящики нижнего яруса
BOX2 = "#BFD8C4"      # ящики верхнего яруса
FILT = "#E4A672"      # фильтры
WOOD = "#B08968"      # доска
BG = "#ffffff"


def esc(t):
    return (t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def text(x, y, s, size=13, fill=INK, anchor="start", weight="normal"):
    return (f'<text x="{x:.0f}" y="{y:.0f}" font-family="Arial, Helvetica, '
            f'sans-serif" font-size="{size}" fill="{fill}" '
            f'text-anchor="{anchor}" font-weight="{weight}">{esc(s)}</text>')


def rect(x, y, w, h, fill, stroke=INK, sw=1.2, dash=""):
    d = f' stroke-dasharray="{dash}"' if dash else ""
    return (f'<rect x="{x:.1f}" y="{y:.1f}" width="{w:.1f}" '
            f'height="{h:.1f}" fill="{fill}" stroke="{stroke}" '
            f'stroke-width="{sw}"{d}/>')


def line(x1, y1, x2, y2, stroke=MUTED, sw=1, dash=""):
    d = f' stroke-dasharray="{dash}"' if dash else ""
    return (f'<line x1="{x1:.1f}" y1="{y1:.1f}" x2="{x2:.1f}" '
            f'y2="{y2:.1f}" stroke="{stroke}" stroke-width="{sw}"{d}/>')


def plan():
    """Числа для схемы из того же расчета, что и таблицы."""
    a, b, rest, kg, m3 = collect()
    items = a + b
    gates = [r for r in items
             if item_kind(asrow(r)) == "клинкет (задвижка)"
             and (r["Ду"] or 0) >= 200]
    gate_area = sum((((r["Ду"] or 100) / 1000 + 0.20) ** 2) * r["Штук"]
                    for r in gates)
    gate_len = gate_area / FLOOR_W           # длина зоны клинкетов
    gate_h = max(((r["Ду"] or 100) / 1000 + 0.15) for r in gates)

    boxed = [r for r in items if tier(asrow(r)) in (3, 4)]
    filters = [r for r in boxed
               if item_kind(asrow(r)) == "фильтр забортной воды"]
    boxes = int(-(-sum(r["вес"] for r in boxed) // BOX_KG))
    return dict(
        kg=kg, m3=m3, items=len(items),
        units=sum(r["Штук"] for r in items),
        gate_units=sum(r["Штук"] for r in gates),
        gate_kg=sum(r["вес"] for r in gates),
        gate_area=gate_area, gate_len=gate_len, gate_h=gate_h,
        box_count=boxes, box_kg=sum(r["вес"] for r in boxed),
        filt_kg=sum(r["вес"] for r in filters),
        rest_kg=sum(r["вес"] for r in rest),
    )


def side_view(p, x0, y0):
    """Вид сбоку: разрез по длине контейнера."""
    out = []
    w, h = FLOOR_L * PX, HEIGHT * PX
    out.append(text(x0, y0 - 34, "Вид сбоку, разрез по длине", 17,
                    INK, weight="bold"))
    out.append(text(x0, y0 - 14,
                    f"Внутри {FLOOR_L:.2f} x {HEIGHT:.2f} м. "
                    "Масштаб честный: где пусто на рисунке, там пусто и "
                    "в контейнере", 12, MUTED))
    out.append(rect(x0, y0, w, h, BG, INK, 2))

    gl = p["gate_len"] * PX
    base = y0 + h                      # пол

    # настил по полу
    out.append(rect(x0, base - BOARD * PX, w, BOARD * PX, WOOD, INK, 0.8))

    # клинкеты одним слоем
    gy = base - (BOARD + p["gate_h"]) * PX
    out.append(rect(x0, gy, gl, p["gate_h"] * PX, GATE))
    step = gl / max(1, int(p["gate_len"] / 0.5))
    x = x0 + step
    while x < x0 + gl - 2:
        out.append(line(x, gy + 3, x, base - BOARD * PX - 3, GATE_D, 0.8))
        x += step
    out.append(text(x0 + gl / 2, gy + p["gate_h"] * PX / 2 + 5,
                    f"клинкеты Ду200-350, {p['gate_units']:.0f} шт, "
                    f"{p['gate_kg'] / 1000:.1f} т", 13, "#3b2f00", "middle",
                    "bold"))

    # настил поверх клинкетов
    out.append(rect(x0, gy - BOARD * PX, gl, BOARD * PX, WOOD, INK, 0.8))

    # ящики: над клинкетами один ярус, на свободном полу два
    bh = BOX_H * PX
    over_y = gy - BOARD * PX - bh
    n_over = int(p["gate_len"] / BOX_L)
    for i in range(n_over):
        out.append(rect(x0 + i * BOX_L * PX, over_y, BOX_L * PX - 2, bh,
                        BOX1))
    out.append(text(x0 + gl / 2, over_y + bh / 2 + 5,
                    "ящики на клинкетах", 12, "#1f3a28", "middle"))

    free_x = x0 + gl
    free_len = FLOOR_L - p["gate_len"]
    n_free = int(free_len / BOX_L)
    for lvl, (yy, col) in enumerate(((base - BOARD * PX - bh, BOX1),
                                     (base - BOARD * PX - 2 * bh, BOX2))):
        for i in range(n_free):
            out.append(rect(free_x + i * BOX_L * PX + 1, yy,
                            BOX_L * PX - 2, bh, col))
    out.append(text(free_x + n_free * BOX_L * PX / 2,
                    base - BOARD * PX - bh / 2 + 5,
                    f"ящики, нижний ярус", 12, "#1f3a28", "middle"))
    out.append(text(free_x + n_free * BOX_L * PX / 2,
                    base - BOARD * PX - bh - bh / 2 + 5,
                    "ящики, верхний ярус", 12, "#1f3a28", "middle"))

    # фильтры верхом над свободной зоной
    fy = base - BOARD * PX - 2 * bh - 0.45 * PX
    out.append(rect(free_x + 1, fy, min(3 * BOX_L * PX, n_free * BOX_L * PX),
                    0.45 * PX, FILT))
    out.append(text(free_x + 90, fy + 18,
                    f"фильтры, {p['filt_kg'] / 1000:.1f} т", 12,
                    "#5a3519"))

    # высотные отметки
    for val, label in ((BOARD, "настил 25 мм"),
                       (BOARD + p["gate_h"], "верх клинкетов"),
                       (BOARD + p["gate_h"] + BOARD + BOX_H,
                        "верх ящика над клинкетами"),
                       (BOARD + 2 * BOX_H, "верх второго яруса ящиков")):
        yy = base - val * PX
        out.append(line(x0 - 8, yy, x0 + w, yy, MUTED, 0.6, "4 3"))
        out.append(text(x0 + w + 10, yy + 4, f"{val:.2f} м — {label}", 11,
                        MUTED))
    out.append(text(x0 + w + 10, y0 + 12, f"{HEIGHT:.2f} м — потолок", 11,
                    MUTED))
    return "\n".join(out), h


def top_view(p, x0, y0):
    """Вид сверху: что где стоит на полу."""
    out = []
    w, h = FLOOR_L * PX, FLOOR_W * PX
    out.append(text(x0, y0 - 34, "Вид сверху, план пола", 17, INK,
                    weight="bold"))
    out.append(text(x0, y0 - 14,
                    f"Пол {FLOOR_L:.2f} x {FLOOR_W:.2f} = "
                    f"{FLOOR_AREA:.1f} кв. м. Клинкеты занимают "
                    f"{p['gate_area']:.1f} кв. м, это "
                    f"{p['gate_area'] / FLOOR_AREA * 100:.0f}% пола",
                    12, MUTED))
    out.append(rect(x0, y0, w, h, BG, INK, 2))

    gl = p["gate_len"] * PX
    out.append(rect(x0, y0, gl, h, GATE))
    # сетка по клинкетам
    step = 0.5 * PX
    x = x0 + step
    while x < x0 + gl - 2:
        out.append(line(x, y0 + 2, x, y0 + h - 2, GATE_D, 0.7))
        x += step
    out.append(line(x0, y0 + h / 2, x0 + gl, y0 + h / 2, GATE_D, 0.7))
    out.append(text(x0 + gl / 2, y0 + h / 2 + 5,
                    f"клинкеты одним слоем, {p['gate_kg'] / 1000:.1f} т",
                    13, "#3b2f00", "middle", "bold"))

    free_x = x0 + gl
    n_free = int((FLOOR_L - p["gate_len"]) / BOX_L)
    for i in range(n_free):
        for j in range(2):
            out.append(rect(free_x + i * BOX_L * PX + 1,
                            y0 + j * BOX_W * PX + 1,
                            BOX_L * PX - 2, BOX_W * PX - 2, BOX1))
    out.append(text(free_x + n_free * BOX_L * PX / 2, y0 + h / 2 + 5,
                    f"ящики на полу: {n_free} ряда по 2 = {n_free * 2} мест "
                    "в ярус", 13, "#1f3a28", "middle", "bold"))

    out.append(line(x0, y0 + h + 16, x0 + gl, y0 + h + 16, INK, 1.2))
    out.append(text(x0 + gl / 2, y0 + h + 32,
                    f"{p['gate_len']:.2f} м", 12, INK, "middle"))
    out.append(line(free_x, y0 + h + 16, x0 + w, y0 + h + 16, INK, 1.2))
    out.append(text(free_x + (w - gl) / 2, y0 + h + 32,
                    f"{FLOOR_L - p['gate_len']:.2f} м", 12, INK, "middle"))
    return "\n".join(out), h


def legend(p, x0, y0):
    out = [text(x0, y0, "Порядок погрузки", 17, INK, weight="bold")]
    steps = [
        (WOOD, "1. Настил необрезной доской 25 мм по всему полу"),
        (GATE, f"2. Клинкеты Ду200-350 лежа, один слой: "
               f"{p['gate_units']:.0f} шт, {p['gate_kg'] / 1000:.1f} т"),
        (WOOD, "3. Настил доской поверх клинкетов, над корпусами, "
               "а не над маховиками"),
        (BOX1, f"4. Ящики нижнего яруса, самые тяжелые: всего ящиков "
               f"{p['box_count']}, {p['box_kg'] / 1000:.1f} т"),
        (BOX2, "5. Ящики верхнего яруса, что полегче"),
        (FILT, f"6. Фильтры сверху, {p['filt_kg'] / 1000:.1f} т, "
               "над углами ящиков"),
    ]
    y = y0 + 26
    for color, label in steps:
        out.append(rect(x0, y - 11, 16, 14, color))
        out.append(text(x0 + 26, y, label, 13))
        y += 24
    y += 6
    for label in [
        f"Всего в контейнере: {p['items']} позиций, {p['units']:.0f} штук, "
        f"{p['kg'] / 1000:.2f} т, {p['m3']:.1f} из 67,2 куб. м",
        f"Остается во Владивостоке {p['rest_kg'] / 1000:.2f} т",
        "Ящик не выше 0,91 м, иначе второй ярус не встает под потолок",
        "Доска: 68 штук по 4 м, это 1,36 куб. м с запасом на обзол",
    ]:
        out.append(text(x0, y, "— " + label, 12, MUTED))
        y += 20
    return "\n".join(out), y - y0


def main():
    p = plan()
    x0 = PAD
    y_side = PAD + 40
    side, h_side = side_view(p, x0, y_side)
    y_top = y_side + h_side + GAP
    top, h_top = top_view(p, x0, y_top)
    y_leg = y_top + h_top + GAP - 40
    leg, h_leg = legend(p, x0, y_leg)

    width = int(x0 + FLOOR_L * PX + 240)
    height = int(y_leg + h_leg + PAD)

    svg = [
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" '
        f'height="{height}" viewBox="0 0 {width} {height}">',
        f'<rect width="{width}" height="{height}" fill="{BG}"/>',
        text(x0, PAD - 10,
             "Загрузка сорокафутового контейнера, Владивосток — Петербург",
             22, INK, weight="bold"),
        side, top, leg, "</svg>",
    ]
    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / "Схема_загрузки_40ф.svg"
    path.write_text("\n".join(svg), encoding="utf-8")
    print(f"готово: {path}")
    print(f"клинкеты {p['gate_units']:.0f} шт на {p['gate_area']:.1f} кв. м "
          f"({p['gate_len']:.2f} м длины), ящиков {p['box_count']}")


if __name__ == "__main__":
    main()
