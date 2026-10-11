"""Сбор данных из WB API (только чтение) → data/wb_api/.

Запускается из GitHub Actions или локально:
    python tools/wb_api/fetch_data.py

Все запросы — GET, без изменений в кабинете WB.
"""
import os, json, time, datetime, requests
from pathlib import Path
from dotenv import load_dotenv

load_dotenv(Path(__file__).parent.parent.parent / '.env')

WB_TOKEN = os.getenv('WB_TOKEN', '')

BASE_STAT   = 'https://statistics-api.wildberries.ru'
BASE_SUPPLY = 'https://suppliers-api.wildberries.ru'
BASE_ANALYT = 'https://seller-analytics-api.wildberries.ru'

OUT = Path(__file__).parent.parent.parent / 'data' / 'wb_api'
OUT.mkdir(parents=True, exist_ok=True)


def headers():
    if not WB_TOKEN:
        raise ValueError('WB_TOKEN не задан в .env / GitHub Secrets')
    return {'Authorization': WB_TOKEN, 'Content-Type': 'application/json'}


def get(base: str, path: str, params: dict = None, retries: int = 3):
    url = base.rstrip('/') + '/' + path.lstrip('/')
    for attempt in range(retries):
        try:
            r = requests.get(url, headers=headers(), params=params, timeout=30)
            if r.status_code == 429:
                time.sleep(2 ** attempt)
                continue
            r.raise_for_status()
            return r.json()
        except requests.exceptions.RequestException as e:
            if attempt == retries - 1:
                raise
            time.sleep(2 ** attempt)


def save(name: str, data):
    path = OUT / f'{name}.json'
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f'  ✓ {path.name} — {len(data) if isinstance(data, list) else "ok"}')


def fetch_stocks():
    """Текущие остатки FBO по всем складам."""
    today = datetime.date.today().isoformat()
    # dateSupplies — дата начала поставки; ставим старую дату, чтобы получить всё
    data = get(BASE_SUPPLY, '/api/v3/stocks/all', params={'dateFrom': '2019-01-01'})
    save('stocks_fbo', data)


def fetch_orders(days_back: int = 90):
    """Заказы за последние N дней."""
    date_from = (datetime.date.today() - datetime.timedelta(days=days_back)).isoformat()
    data = get(BASE_SUPPLY, '/api/v3/orders', params={'dateFrom': date_from, 'next': 0})
    save('orders', data)


def fetch_sales(days_back: int = 90):
    """Продажи (детальный отчёт Statistics API)."""
    date_from = (datetime.date.today() - datetime.timedelta(days=days_back)).isoformat()
    date_to = datetime.date.today().isoformat()
    data = get(BASE_STAT, '/api/v1/supplier/sales',
               params={'dateFrom': date_from, 'flag': 0})
    save('sales', data)


def fetch_report_detail(days_back: int = 30):
    """Детализированный финансовый отчёт (Statistics API)."""
    date_from = (datetime.date.today() - datetime.timedelta(days=days_back)).isoformat()
    date_to = datetime.date.today().isoformat()
    data = get(BASE_STAT, '/api/v1/supplier/reportDetailByPeriod',
               params={'dateFrom': date_from, 'dateTo': date_to, 'limit': 100000})
    save('report_detail', data)


def fetch_warehouses():
    """Список складов WB."""
    data = get(BASE_SUPPLY, '/api/v3/offices')
    save('warehouses', data)


def fetch_incomes(days_back: int = 90):
    """Поставки (приходы на склад)."""
    date_from = (datetime.date.today() - datetime.timedelta(days=days_back)).isoformat()
    data = get(BASE_STAT, '/api/v1/supplier/incomes', params={'dateFrom': date_from})
    save('incomes', data)


def fetch_stocks_stat(days_back: int = 7):
    """Остатки (Statistics API — история по дням)."""
    date_from = (datetime.date.today() - datetime.timedelta(days=days_back)).isoformat()
    data = get(BASE_STAT, '/api/v1/supplier/stocks', params={'dateFrom': date_from})
    save('stocks_stat', data)


TASKS = [
    ('warehouses',    fetch_warehouses),
    ('stocks_fbo',    fetch_stocks),
    ('orders',        fetch_orders),
    ('sales',         fetch_sales),
    ('report_detail', fetch_report_detail),
    ('incomes',       fetch_incomes),
    ('stocks_stat',   fetch_stocks_stat),
]


if __name__ == '__main__':
    print(f'WB API fetch — {datetime.date.today()}')
    print(f'Token: {"OK" if WB_TOKEN else "НЕ ЗАДАН — выход"}')
    if not WB_TOKEN:
        raise SystemExit(1)

    errors = []
    for name, fn in TASKS:
        print(f'→ {name}')
        try:
            fn()
        except Exception as e:
            print(f'  ✗ {e}')
            errors.append((name, str(e)))

    # Сохраняем мета-информацию о последнем запуске
    meta = {
        'fetched_at': datetime.datetime.utcnow().isoformat() + 'Z',
        'errors': errors,
        'tasks': [t[0] for t in TASKS],
    }
    (OUT / 'fetch_meta.json').write_text(json.dumps(meta, ensure_ascii=False, indent=2))

    if errors:
        print(f'\n{len(errors)} ошибок: {[e[0] for e in errors]}')
        raise SystemExit(1)
    print('\nГотово.')
