#!/usr/bin/env python3
"""Геометрия клинкетов на полу контейнера.

Площадь под задвижки раньше считалась формулой от диаметра. Формула
занижает: задвижка лежит на боку и занимает пол строительной длиной на
полную высоту с маховиком, а поперек контейнера их входит ровно
столько, сколько влезает целых штук, без дробей.

Укладка принята такая: задвижка на боку, шток горизонтально поперек
контейнера, маховик нагрузки не несет. Высота стопки при этом равна
диаметру фланца, а не полной высоте задвижки.

Габариты справочные, по типовым судовым бронзовым клинкетам. Замерить
при погрузке: справочник уже ошибся на массах почти вдвое.
"""

from collections import defaultdict

# внутри двадцатифутового
FLOOR_L, FLOOR_W, HEIGHT = 5.898, 2.352, 2.393
BOARD = 0.025
CRATE_L, CRATE_W, CRATE_H = 1.2, 1.0, 0.91

# Ду: строительная длина, наружный диаметр фланца, полная высота, мм
DIM = {
    350: (500, 520, 1150),
    300: (450, 460, 1000),
    250: (400, 405, 850),
    200: (350, 340, 700),
    150: (280, 280, 560),
    125: (250, 250, 500),
    100: (230, 220, 440),
    80: (200, 195, 380),
}


def across(du):
    """Сколько задвижек этого диаметра встает поперек контейнера."""
    if du not in DIM:
        return 2
    return max(1, int(FLOOR_W * 1000 // DIM[du][2]))


def depth(du):
    """Какую глубину ряда занимает этот диаметр, м."""
    return (DIM[du][0] if du in DIM else 400) / 1000


def stack(du):
    """Высота стопки в один ярус, м: это диаметр фланца."""
    return (DIM[du][1] if du in DIM else 400) / 1000


def rows_of(counts):
    """Ряды снизу вверх. Крупное первым: оно и тяжелее, и выше."""
    out = []
    for du in sorted(counts, reverse=True):
        left = counts[du]
        n = across(du)
        while left > 0:
            qty = min(n, left)
            left -= qty
            out.append(dict(du=du, qty=qty, across=n, depth=depth(du),
                            stack=stack(du)))
    return out


def place(rows, floor_l=FLOOR_L):
    """Разложить ряды по ярусам: что не влезло в длину, идет вторым."""
    layer, used = 1, 0.0
    for r in rows:
        if used + r["depth"] > floor_l:
            layer += 1
            used = 0.0
        r["layer"] = layer
        used += r["depth"]
        r["used"] = used
    return layer


def gate_length(counts):
    """Сколько метров пола просят клинкеты, если класть в один ярус."""
    return sum(r["depth"] for r in rows_of(counts))


def crate_places(counts, floor_l=FLOOR_L):
    """Мест под ящики: на свободном полу и поверх клинкетов.

    На свободном полу ящики встают в два яруса, поверх клинкетов —
    столько, сколько остается высоты до потолка.
    """
    rows = rows_of(counts)
    layers = place(rows, floor_l)
    if not rows:
        return int(floor_l / CRATE_L) * 2 * 2, 0, 0.0, 0.0
    taken = max((r["used"] for r in rows if r["layer"] == layers), default=0.0)
    gl = min(gate_length(counts), floor_l) if layers == 1 else floor_l
    free_l = floor_l - (taken if layers == 1 else floor_l)
    h_gates = BOARD + sum(
        max((r["stack"] for r in rows if r["layer"] == i), default=0.0)
        + BOARD for i in range(1, layers + 1))
    tiers_over = int((HEIGHT - h_gates) / CRATE_H)
    free = int(free_l / CRATE_L) * 2 * 2
    over = int(gl / CRATE_L) * 2 * max(0, tiers_over)
    return free, over, h_gates, free_l
