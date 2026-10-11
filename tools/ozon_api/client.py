"""Ozon Seller API — базовый клиент (api-seller.ozon.ru).

Аутентификация: заголовки Client-Id и Api-Key (из .env).
Доку сверяй: https://docs.ozon.ru/api/seller/
"""
import time
import requests
from tools.common import config

BASE = 'https://api-seller.ozon.ru'


def _headers():
    config.require('OZON_CLIENT_ID', 'OZON_API_KEY')
    return {
        'Client-Id': config.OZON_CLIENT_ID,
        'Api-Key': config.OZON_API_KEY,
        'Content-Type': 'application/json',
    }


def post(path, body=None, retries=4):
    url = BASE + path
    for attempt in range(retries):
        try:
            r = requests.post(url, headers=_headers(), json=body or {}, timeout=40)
            if r.status_code == 429:
                time.sleep(2 ** attempt); continue
            r.raise_for_status()
            return r.json()
        except requests.exceptions.RequestException:
            if attempt == retries - 1:
                raise
            time.sleep(2 ** attempt)


def ping():
    """Проверка токена: список товаров (1 шт)."""
    return post('/v3/product/list', {'filter': {}, 'limit': 1})
