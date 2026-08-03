"""Живой Excel-шаблон юнит-экономики Ozon (формулы, а не значения).

Меняешь жёлтые ячейки (цена, параметры, при желании себестоимость/комиссию) —
книга пересчитывается сама средствами Excel. Данные (комиссии, логистика,
объём) берутся из выгрузки prices.json, себестоимость — из COSTS.

Вход:  data/ozon_api/prices.json
Выход: outputs/ozon/ozon_unit_model_live_2026-08-03.xlsx
"""
import json
from pathlib import Path

from calc_unit_ozon import (COSTS, DRR, TAX, ACQUIRING, STORAGE_L_DAY,
                            STORAGE_DAYS, FULFILLMENT, buyout, num)

ROOT = Path(__file__).parent.parent.parent
PRICES = ROOT / 'data' / 'ozon_api' / 'prices.json'
OUT = ROOT / 'outputs' / 'ozon'


def scheme_inputs(item, scheme):
    c = item.get('commissions', {}) or {}
    if scheme == 'fbo':
        return (num(c.get('sales_percent_fbo')) / 100,
                num(c.get('fbo_direct_flow_trans_min_amount')),
                num(c.get('fbo_deliv_to_customer_amount')),
                num(c.get('fbo_return_flow_amount')))
    return (num(c.get('sales_percent_fbs')) / 100,
            num(c.get('fbs_direct_flow_trans_min_amount')),
            num(c.get('fbs_deliv_to_customer_amount')),
            num(c.get('fbs_return_flow_amount')))


HEAD = ['Товар', 'Схема', 'Цена', 'Себест.', 'Фулфилмент', 'База затрат',
        'Комиссия %', 'Комиссия ₽', 'Логистика', 'Посл.миля', 'Эквайринг',
        'Налог', 'Выкуп %', 'Обр.лог база', 'Обр.лог', 'Объём, л', 'Хранение',
        'Реклама (ДРР)', 'Маржа до рекл.', 'Маржа/шт', 'Маржа % цены', 'ROI %',
        'Цена безубыт.']
# буквы колонок по порядку HEAD
COL = {name: chr(ord('A') + i) for i, name in enumerate(HEAD)}


def build():
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    items = json.loads(PRICES.read_text())
    rows = []
    for it in items:
        if it.get('offer_id') not in COSTS:
            continue
        for scheme in ('fbo', 'fbs'):
            comm, logi, lastm, retb = scheme_inputs(it, scheme)
            rows.append({
                'offer': it.get('offer_id'), 'scheme': scheme.upper(),
                'price': num(it.get('price', {}).get('price')),
                'cost': COSTS[it.get('offer_id')],
                'comm': comm, 'logi': logi, 'lastm': lastm, 'retb': retb,
                'buyout': buyout(it.get('offer_id')),
                'vol': num(it.get('volume_weight')),
            })

    wb = Workbook()
    ws = wb.active
    ws.title = 'Юнитка Ozon (live)'

    thin = Side(style='thin', color='D0D0D0')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    head_fill = PatternFill('solid', fgColor='305496')
    in_fill = PatternFill('solid', fgColor='FFF2CC')     # редактируемые (жёлтые)
    res_fill = PatternFill('solid', fgColor='E2EFDA')    # результат
    bold = Font(bold=True)

    # --- Блок параметров (редактируемый) ---
    ws['A1'] = 'ПАРАМЕТРЫ — меняй жёлтые ячейки, всё пересчитается'
    ws['A1'].font = Font(bold=True, size=12)
    params = [
        ('ДРР (доля рекламы)', DRR, '0%'),
        ('Налог с оборота', TAX, '0%'),
        ('Эквайринг', ACQUIRING, '0.0%'),
        ('Фулфилмент, ₽/шт', FULFILLMENT, '#,##0'),
        ('Хранение, ₽/л/день', STORAGE_L_DAY, '0.00'),
        ('Дней хранения', STORAGE_DAYS, '#,##0'),
    ]
    # значения в B3..B8
    for i, (label, val, fmt) in enumerate(params):
        r = 3 + i
        ws[f'A{r}'] = label
        cell = ws[f'B{r}']
        cell.value = val
        cell.number_format = fmt
        cell.fill = in_fill
        cell.border = border
        cell.font = bold
    P = {'drr': '$B$3', 'tax': '$B$4', 'acq': '$B$5',
         'ff': '$B$6', 'stl': '$B$7', 'std': '$B$8'}

    # --- Заголовок таблицы ---
    HROW = 11
    for j, name in enumerate(HEAD):
        c = ws.cell(HROW, j + 1, name)
        c.font = Font(bold=True, color='FFFFFF')
        c.fill = head_fill
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = border
    ws.row_dimensions[HROW].height = 30

    money = '#,##0'
    money2 = '#,##0.00'
    pct = '0.0%'

    input_cols = {'Цена', 'Себест.', 'Комиссия %', 'Логистика', 'Посл.миля',
                  'Выкуп %', 'Обр.лог база', 'Объём, л'}

    for k, d in enumerate(rows):
        rr = HROW + 1 + k
        # ссылки на ячейки текущей строки
        def cr(name):
            return f'{COL[name]}{rr}'
        # входные значения
        vals = {
            'Товар': d['offer'], 'Схема': d['scheme'], 'Цена': d['price'],
            'Себест.': round(d['cost'], 2), 'Комиссия %': d['comm'],
            'Логистика': d['logi'], 'Посл.миля': d['lastm'],
            'Выкуп %': d['buyout'], 'Обр.лог база': d['retb'], 'Объём, л': d['vol'],
        }
        # формулы
        formulas = {
            'Фулфилмент': f'={P["ff"]}',
            'База затрат': f'={cr("Себест.")}+{cr("Фулфилмент")}',
            'Комиссия ₽': f'={cr("Цена")}*{cr("Комиссия %")}',
            'Эквайринг': f'={cr("Цена")}*{P["acq"]}',
            'Налог': f'={cr("Цена")}*{P["tax"]}',
            'Обр.лог': f'={cr("Обр.лог база")}*(1-{cr("Выкуп %")})',
            'Хранение': f'={P["stl"]}*{P["std"]}*{cr("Объём, л")}',
            'Реклама (ДРР)': f'={cr("Цена")}*{P["drr"]}',
            'Маржа до рекл.': (f'={cr("Цена")}-{cr("Комиссия ₽")}-{cr("Логистика")}'
                              f'-{cr("Посл.миля")}-{cr("Эквайринг")}-{cr("Налог")}'
                              f'-{cr("Обр.лог")}-{cr("Хранение")}-{cr("База затрат")}'),
            'Маржа/шт': f'={cr("Маржа до рекл.")}-{cr("Реклама (ДРР)")}',
            'Маржа % цены': f'={cr("Маржа/шт")}/{cr("Цена")}',
            'ROI %': f'={cr("Маржа/шт")}/{cr("База затрат")}',
            'Цена безубыт.': (f'=({cr("Логистика")}+{cr("Посл.миля")}+{cr("Обр.лог")}'
                             f'+{cr("Хранение")}+{cr("База затрат")})'
                             f'/(1-({cr("Комиссия %")}+{P["acq"]}+{P["tax"]}+{P["drr"]}))'),
        }

        for name in HEAD:
            cell = ws.cell(rr, HEAD.index(name) + 1)
            cell.value = vals.get(name, formulas.get(name))
            cell.border = border
            # форматы
            if name in ('Комиссия %', 'Выкуп %', 'Маржа % цены', 'ROI %'):
                cell.number_format = pct
            elif name in ('Себест.', 'Фулфилмент', 'База затрат'):
                cell.number_format = money2
            elif name == 'Объём, л':
                cell.number_format = '0.00'
            elif name not in ('Товар', 'Схема'):
                cell.number_format = money
            # заливка
            if name in input_cols:
                cell.fill = in_fill
            if name in ('Маржа/шт', 'ROI %'):
                cell.fill = res_fill
                cell.font = bold

    # ширины
    widths = {'Товар': 42, 'Схема': 7, 'База затрат': 12, 'Обр.лог база': 12,
              'Реклама (ДРР)': 12, 'Маржа до рекл.': 13, 'Маржа % цены': 11,
              'Цена безубыт.': 12}
    for name in HEAD:
        ws.column_dimensions[COL[name]].width = widths.get(name, 10)
    ws.freeze_panes = f'C{HROW + 1}'

    # подсказка
    note = ws.cell(HROW + len(rows) + 3, 1,
                   'Жёлтые ячейки — вход (цена, себест., комиссия, логистика, '
                   'выкуп, объём) и параметры B3:B8. Остальное — формулы. '
                   'Меняешь цену в колонке C → маржа и ROI пересчитываются.')
    note.font = Font(italic=True, color='808080')

    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / 'ozon_unit_model_live_2026-08-03.xlsx'
    wb.save(path)
    print(f'Сохранено: {path}  (строк: {len(rows)})')


if __name__ == '__main__':
    build()
