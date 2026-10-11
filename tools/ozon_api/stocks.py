"""Ozon — остатки на складах + алерт по заканчивающимся позициям.

Логика:
  1. Тянем остатки FBO/FBS по всем SKU.
  2. Тянем продажи за 30 дней (для оценки скорости).
  3. Считаем "дней запаса" = остаток / (продажи/день).
  4. Если дней запаса меньше порога LOW_STOCK_DAYS — в алерт.

Пути сверяй с докой Ozon (разделы Analytics / Stocks).
"""
import json
import datetime as dt
from tools.ozon_api import client
from tools.common import config, notify


def stocks_fbo_fbs():
    """Остатки по складам. POST /v1/analytics/stock_on_warehouses или /v4/product/info/stocks."""
    out = []
    last_id = ''
    while True:
        res = client.post('/v4/product/info/stocks',
                          {'filter': {'visibility': 'ALL'}, 'limit': 1000, 'last_id': last_id}) or {}
        items = (res.get('result') or {}).get('items', [])
        for it in items:
            total = sum(s.get('present', 0) for s in it.get('stocks', []))
            out.append({'sku': it.get('product_id'), 'offer_id': it.get('offer_id'), 'stock': total})
        last_id = (res.get('result') or {}).get('last_id', '')
        if not last_id or not items:
            break
    return out


def sales_30d():
    """Продажи за 30 дней по SKU через аналитику.
    POST /v1/analytics/data  metrics=[ordered_units] dimension=[sku]"""
    today = dt.date.today()
    body = {
        'date_from': (today - dt.timedelta(days=30)).isoformat(),
        'date_to': today.isoformat(),
        'metrics': ['ordered_units'],
        'dimension': ['sku'],
        'limit': 1000, 'offset': 0,
    }
    sold = {}
    while True:
        res = client.post('/v1/analytics/data', body) or {}
        rows = (res.get('result') or {}).get('data', [])
        for row in rows:
            key = row['dimensions'][0].get('id')
            val = row['metrics'][0] if row.get('metrics') else 0
            sold[key] = sold.get(key, 0) + val
        if len(rows) < body['limit']:
            break
        body['offset'] += body['limit']
    return sold


def run():
    today = dt.date.today().isoformat()
    stocks = stocks_fbo_fbs()
    try:
        sold = sales_30d()
    except Exception:
        sold = {}

    report = []
    for s in stocks:
        v30 = sold.get(str(s['sku']), 0) or sold.get(s['sku'], 0)
        per_day = v30 / 30 if v30 else 0
        days_left = round(s['stock'] / per_day, 1) if per_day else None
        report.append({**s, 'sold_30d': v30, 'days_left': days_left})

    out = config.DATA_OZON / f'stocks_{today}.json'
    out.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')

    low = [r for r in report if r['days_left'] is not None and r['days_left'] < config.LOW_STOCK_DAYS]
    msg = [f'<b>Ozon остатки</b> {today}: позиций {len(report)}, заканчивается {len(low)} (меньше {config.LOW_STOCK_DAYS} дн).']
    for r in sorted(low, key=lambda x: x['days_left'])[:20]:
        msg.append(f"{r['offer_id']}: остаток {r['stock']} шт, хватит на {r['days_left']} дн (продажи 30д: {r['sold_30d']})")
    notify.send('\n'.join(msg))
    return report


if __name__ == '__main__':
    run()
