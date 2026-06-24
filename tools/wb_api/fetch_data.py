"""Сбор данных из WB API (только чтение) → data/wb_api/.

Все запросы — GET, без изменений в кабинете WB.
"""
import os, json, time, datetime, requests
from pathlib import Path

WB_TOKEN = os.getenv('WB_TOKEN', '')

BASE_STAT   = 'https://statistics-api.wildberries.ru'
BASE_SUPPLY = 'https://suppliers-api.wildberries.ru'

OUT = Path(__file__).parent.parent.parent / 'data' / 'wb_api'
OUT.mkdir(parents=True, exist_ok=True)


def headers():
    if not WB_TOKEN:
        raise ValueError('WB_TOKEN не задан')
    return {'Authorization': WB_TOKEN, 'Content-Type': 'application/json'}


def get(base, path, params=None, retries=3):
    url = base.rstrip('/') + '/' + path.lstrip('/')
    for attempt in range(retries):
        try:
            r = requests.get(url, headers=headers(), params=params, timeout=30)
            if r.status_code == 429:
                time.sleep(2 ** attempt)
                continue
            r.raise_for_status()
            return r.json()
        except requests.exceptions.RequestException:
            if attempt == retries - 1:
                raise
            time.sleep(2 ** attempt)


def save(name, data):
    path = OUT / f'{name}.json'
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2))
    print(f'  ok {path.name} — {len(data) if isinstance(data, list) else "ok"}')


TASKS = {
    'warehouses':    lambda: get(BASE_SUPPLY, '/api/v3/offices'),
    'orders':        lambda: get(BASE_SUPPLY, '/api/v3/orders',
                         params={'dateFrom': (datetime.date.today()-datetime.timedelta(days=90)).isoformat(), 'next': 0}),
    'sales':         lambda: get(BASE_STAT, '/api/v1/supplier/sales',
                         params={'dateFrom': (datetime.date.today()-datetime.timedelta(days=90)).isoformat(), 'flag': 0}),
    'report_detail': lambda: get(BASE_STAT, '/api/v1/supplier/reportDetailByPeriod',
                         params={'dateFrom': (datetime.date.today()-datetime.timedelta(days=30)).isoformat(),
                                 'dateTo': datetime.date.today().isoformat(), 'limit': 100000}),
    'incomes':       lambda: get(BASE_STAT, '/api/v1/supplier/incomes',
                         params={'dateFrom': (datetime.date.today()-datetime.timedelta(days=90)).isoformat()}),
    'stocks_stat':   lambda: get(BASE_STAT, '/api/v1/supplier/stocks',
                         params={'dateFrom': (datetime.date.today()-datetime.timedelta(days=7)).isoformat()}),
}

if __name__ == '__main__':
    print(f'WB API fetch — {datetime.date.today()}')
    if not WB_TOKEN:
        raise SystemExit('WB_TOKEN не задан — выход')

    errors = []
    for name, fn in TASKS.items():
        print(f'-> {name}')
        try:
            save(name, fn())
        except Exception as e:
            print(f'  err {e}')
            errors.append((name, str(e)))

    meta = {'fetched_at': datetime.datetime.utcnow().isoformat()+'Z', 'errors': errors}
    (OUT / 'fetch_meta.json').write_text(json.dumps(meta, ensure_ascii=False, indent=2))

    if errors:
        raise SystemExit(f'{len(errors)} errors')
    print('Done.')
