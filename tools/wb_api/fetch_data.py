"""Сбор данных из WB API (только чтение) → data/wb_api/.

Доступен из GitHub Actions только statistics-api.wildberries.ru.
suppliers-api.wildberries.ru не резолвится из облачных runners GitHub.
"""
import os, json, time, datetime, requests
from pathlib import Path

WB_TOKEN = os.getenv('WB_TOKEN', '')

BASE_STAT = 'https://statistics-api.wildberries.ru'

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
            r = requests.get(url, headers=headers(), params=params, timeout=60)
            if r.status_code == 429:
                time.sleep(60)
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
    size = len(data) if isinstance(data, list) else path.stat().st_size
    print(f'  ok {path.name} — {size}')


def days_ago(n):
    return (datetime.date.today() - datetime.timedelta(days=n)).isoformat()


def fetch_sales():
    """Продажи и возвраты за 90 дней."""
    data = get(BASE_STAT, '/api/v1/supplier/sales',
               params={'dateFrom': days_ago(90), 'flag': 0})
    save('sales', data)


def fetch_stocks():
    """Остатки на складах."""
    data = get(BASE_STAT, '/api/v1/supplier/stocks',
               params={'dateFrom': days_ago(3)})
    save('stocks', data)


def fetch_orders():
    """Заказы за 90 дней (Statistics API)."""
    data = get(BASE_STAT, '/api/v1/supplier/orders',
               params={'dateFrom': days_ago(90), 'flag': 0})
    save('orders', data)


def fetch_report_detail():
    """Детализированный финансовый отчёт за 30 дней.
    rrdid=0 — начать с первой записи.
    """
    data = get(BASE_STAT, '/api/v1/supplier/reportDetailByPeriod',
               params={
                   'dateFrom': days_ago(30),
                   'dateTo': datetime.date.today().isoformat(),
                   'rrdid': 0,
               })
    save('report_detail', data)


def fetch_incomes():
    """Поставки (FBO приходы) за 90 дней."""
    data = get(BASE_STAT, '/api/v1/supplier/incomes',
               params={'dateFrom': days_ago(90)})
    save('incomes', data)


TASKS = [
    ('sales',         fetch_sales),
    ('stocks',        fetch_stocks),
    ('orders',        fetch_orders),
    ('report_detail', fetch_report_detail),
    ('incomes',       fetch_incomes),
]


if __name__ == '__main__':
    print(f'WB API fetch — {datetime.date.today()}')
    if not WB_TOKEN:
        raise SystemExit('WB_TOKEN не задан — выход')

    errors = []
    for name, fn in TASKS:
        print(f'-> {name}')
        try:
            fn()
        except Exception as e:
            print(f'  err {e}')
            errors.append((name, str(e)))

    meta = {
        'fetched_at': datetime.datetime.utcnow().isoformat() + 'Z',
        'errors': errors,
        'tasks_ok': [t[0] for t in TASKS if t[0] not in {e[0] for e in errors}],
    }
    (OUT / 'fetch_meta.json').write_text(json.dumps(meta, ensure_ascii=False, indent=2))

    if errors:
        print(f'\n{len(errors)} ошибок: {[e[0] for e in errors]}')
        raise SystemExit(1)
    print('Done.')
