"""WB API client — базовый клиент для Wildberries Supplier API."""
import os, time, json, requests
from pathlib import Path
from dotenv import load_dotenv

load_dotenv(Path(__file__).parent.parent.parent / '.env')

BASE_STAT   = 'https://statistics-api.wildberries.ru'
BASE_SUPPLY = 'https://suppliers-api.wildberries.ru'
BASE_CONTENT= 'https://content-api.wildberries.ru'
BASE_ANALYT = 'https://seller-analytics-api.wildberries.ru'
BASE_PROMO  = 'https://advert-api.wildberries.ru'

# WB выдаёт ОДИН токен со всеми категориями. Поддержка старых раздельных
# переменных оставлена как fallback.
WB_TOKEN = (os.getenv('WB_TOKEN', '')
            or os.getenv('WB_TOKEN_STATISTICS', ''))

def _headers(token_type: str = None) -> dict:
    t = WB_TOKEN
    if not t:
        raise ValueError('WB_TOKEN не задан в .env')
    return {'Authorization': t, 'Content-Type': 'application/json'}

def get(base: str, path: str, token_type: str, params: dict = None, retries: int = 3):
    url = base.rstrip('/') + '/' + path.lstrip('/')
    for attempt in range(retries):
        try:
            r = requests.get(url, headers=_headers(token_type), params=params, timeout=30)
            if r.status_code == 429:
                time.sleep(2 ** attempt)
                continue
            r.raise_for_status()
            return r.json()
        except requests.exceptions.RequestException as e:
            if attempt == retries - 1:
                raise
            time.sleep(2 ** attempt)

def post(base: str, path: str, token_type: str, body: dict = None, retries: int = 3):
    url = base.rstrip('/') + '/' + path.lstrip('/')
    for attempt in range(retries):
        try:
            r = requests.post(url, headers=_headers(token_type), json=body or {}, timeout=30)
            if r.status_code == 429:
                time.sleep(2 ** attempt)
                continue
            r.raise_for_status()
            return r.json()
        except requests.exceptions.RequestException as e:
            if attempt == retries - 1:
                raise
            time.sleep(2 ** attempt)

# ── Быстрые методы ──────────────────────────────────────────────────────────

def ping():
    """Проверка доступности API и токена Statistics."""
    return get(BASE_STAT, '/api/v1/supplier/reportDetailByPeriod',
               'statistics', params={'dateFrom': '2026-06-23', 'dateTo': '2026-06-23', 'limit': 1})

def get_stocks(warehouse_id: int = None):
    """Текущие остатки FBO."""
    params = {'dateSupplies': '2019-06-20'}
    return get(BASE_SUPPLY, '/api/v3/stocks/Коледино' if not warehouse_id else f'/api/v3/stocks/{warehouse_id}',
               'marketplace', params)

def get_orders(date_from: str, status: int = 0):
    """Заказы с указанной даты (status=0 все новые)."""
    return get(BASE_SUPPLY, '/api/v3/orders', 'marketplace',
               params={'dateFrom': date_from, 'status': status})

def get_warehouses():
    """Список складов WB."""
    return get(BASE_SUPPLY, '/api/v3/offices', 'marketplace')

def get_nm_report(nm_ids: list):
    """Статистика по nmId (Content API)."""
    return post(BASE_CONTENT, '/content/v2/nm-report/detail', 'content',
                body={'nmIDs': nm_ids, 'period': {'begin': '2026-06-01', 'end': '2026-06-23'},
                      'page': 1})

if __name__ == '__main__':
    print('WB_TOKEN задан:', bool(WB_TOKEN))
    print('Тест ping Statistics API...')
    try:
        r = ping()
        print('OK:', type(r), str(r)[:200])
    except Exception as e:
        print('ОШИБКА:', e)
