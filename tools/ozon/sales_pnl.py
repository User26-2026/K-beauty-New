#!/usr/bin/env python3
"""P&L по SKU из отчёта продаж OZON (.xls, лист TDSheet с ABC-разбивкой).

Отчёт содержит готовую строку «Маржа, руб» по каждому SKU — это чистая
прибыль, её и берём (не пересчитываем). Колонки ищем по тексту заголовка
в строках 1-2, чтобы пережить сдвиги.

Использование:
  python tools/ozon/sales_pnl.py <отчёт.xls> [outputs/ozon_pnl/pnl.xlsx]
"""
import sys, os, xlrd, openpyxl
from openpyxl.styles import Font, PatternFill

WANT = {  # ключ: подстрока заголовка (строка 1 отчёта)
    'прод_шт':'Продажи, шт', 'сум_прод':'Сумма продаж, руб', 'комиссия':'Комиссия, руб',
    'логистика':'Логистика полная', 'реклама':'Реклама ОЗОН (с уч', 'себест':'Себестоимость, руб',
    'налог':'Налог сумма, руб', 'маржа':'Маржа, руб', 'маржа_ед':'Маржа на ед', 'roi':'ROI, %'}

def num(v):
    try: return float(v)
    except Exception: return 0.0

def main(path, out=None):
    sh = xlrd.open_workbook(path).sheet_by_index(0)
    h1 = [str(sh.cell_value(1, c)) for c in range(sh.ncols)]
    idx = {}
    for k, sub in WANT.items():
        idx[k] = next((c for c, h in enumerate(h1) if sub.lower() in h.lower()), None)
    # строки товаров: есть Артикул(0) и SKU(1), не служебные
    skip = {'обычная продажа', 'выкуп маркетплэйсом', 'koreadom', 'итого'}
    prods = []
    for r in range(5, sh.nrows):
        name = str(sh.cell_value(r, 0)).strip(); sku = str(sh.cell_value(r, 1)).strip()
        if not name or not sku or name.lower() in skip: continue
        d = {k: num(sh.cell_value(r, idx[k])) if idx[k] is not None else 0.0 for k in WANT}
        d['name'], d['sku'] = name, sku
        d['drr'] = round(d['реклама']/d['сум_прод']*100, 1) if d['сум_прод'] else 0
        prods.append(d)
    tot = {k: sum(p[k] for p in prods) for k in ('прод_шт','сум_прод','себест','реклама','маржа','логистика','налог','комиссия')}
    out = out or f"outputs/ozon_pnl/ozon_pnl_{os.path.splitext(os.path.basename(path))[0]}.xlsx"
    os.makedirs(os.path.dirname(out), exist_ok=True)
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'P&L по SKU'
    bold = Font(bold=True)
    head = ['Товар','SKU','Продажи шт','Выручка ₽','Себест. ₽','Комиссия+услуги ₽','Логистика ₽','Реклама ₽','ДРР %','Налог ₽','Маржа ₽','Маржа/шт ₽','ROI %']
    ws.append(head)
    for c in ws[1]: c.font = bold; c.fill = PatternFill('solid', fgColor='D9E1F2')
    for p in sorted(prods, key=lambda x: -x['маржа']):
        ws.append([p['name'], p['sku'], round(p['прод_шт']), round(p['сум_прод']), round(p['себест']),
                   round(p['комиссия']), round(p['логистика']), round(p['реклама']), p['drr'],
                   round(p['налог']), round(p['маржа']), round(p['маржа_ед']), round(p['roi'], 1)])
        ws.cell(ws.max_row, 11).fill = PatternFill('solid', fgColor='C6EFCE' if p['маржа'] > 0 else 'FFC7CE')
    ws.append([])
    ws.append(['ИТОГО','', round(tot['прод_шт']), round(tot['сум_прод']), round(tot['себест']),
               round(tot['комиссия']), round(tot['логистика']), round(tot['реклама']),
               round(tot['реклама']/tot['сум_прод']*100,1) if tot['сум_прод'] else 0,
               round(tot['налог']), round(tot['маржа']), '', ''])
    for c in ws[ws.max_row]: c.font = bold
    ws.freeze_panes = 'A2'
    wb.save(out)
    print(f"SKU: {len(prods)} | Выручка: {tot['сум_прод']:,.0f} | Прибыль: {tot['маржа']:,.0f} | "
          f"ДРР: {tot['реклама']/tot['сум_прод']*100:.1f}% | -> {out}")

if __name__ == '__main__':
    if len(sys.argv) < 2: sys.exit(__doc__)
    main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
