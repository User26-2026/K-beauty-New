#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Юнит-экономика WB по шорт-листу японских товаров.
Логика — tools/calc_unit.py (вкладка Юнит из Celimax_unit_model.xlsx).
Цены с СПП — рыночные ориентиры из разбора спроса (WB/нишевые японские магазины).
"""
import sys
from openpyxl import Workbook
sys.path.insert(0, "tools")
from calc_unit import calc

# (Товар, себест_руб, цена_с_СПП, объём_л, комментарий по спросу)
SHORTLIST = [
    ("Bifesta мицеллярная вода 400мл",   460,  990, 0.40, "категория-гигант, бренд узнаваем, на WB слабо представлен"),
    ("Bifesta мицеллярная вода 360мл",   376,  850, 0.36, "то же, меньший объём/чек"),
    ("Bifesta салфетки мицеллярные 46",  223,  590, 0.15, "ходовой импульс, низкий чек"),
    ("Bifesta пенка Foaming Whip 200г",  269,  690, 0.20, "трендовый формат пенки-мусса"),
    ("LUCIDO-L аргановое масло 60мл",    498, 1090, 0.06, "узнаваемое имя, конкуренция с корейцами"),
    ("Doshisha пилинг-носочки",          200,  520, 0.05, "спрос огромный, но ценовая яма (лидеры ~196-400р)"),
    ("Doshisha патчи ретинол",           307,  720, 0.05, "трендовый ингредиент, чек выше"),
    ("KOR тканевая маска 30шт",          299,  890, 0.35, "высокая маржа; проверить состав stem cell/exosome под ЕАЭС"),
]

wb = Workbook(); ws = wb.active; ws.title = "Юнитка WB"
ws.append(["Товар","Себест,р","Цена с СПП,р","Цена WB до СПП,р","Комиссия,р",
           "Логистика,р","Хранение,р","Реклама(ДРР10%),р","Маржа/шт,р","ROI %","Выкуп %","Спрос-комментарий"])
print(f"{'Товар':<34}{'Себест':>7}{'СПП':>6}{'Маржа':>7}{'ROI':>6}")
for name, cost, psp, vol, note in SHORTLIST:
    r = calc(price_spp=psp, cost=cost, name=name, drr=0.10, volume_liter=vol)
    ws.append([name, cost, psp, r["price_wb_before_spp"], round(r["commission"]),
               round(r["logistics"]), round(r["storage"],1), round(r["ads"]),
               round(r["margin_after_drr"]), round(r["roi_after_drr_pct"],1), round(r["buyout_pct"],1), note])
    print(f"{name:<34}{cost:>7}{psp:>6}{r['margin_after_drr']:>7.0f}{r['roi_after_drr_pct']:>6.0f}")

out = "outputs/japan_unitka_WB_2026-07-27.xlsx"
wb.save(out)
print("Файл:", out)
print("Параметры: СПП 30%, ДРР 10%, комиссия 32.5%, логистика 90р (из модели Celimax)")
