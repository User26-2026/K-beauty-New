"""Ozon — прогрузка/сбор поисковых кластеров (ключевых фраз) по товарам.

Назначение: регулярно тянуть статистику поисковых запросов/кластеров Ozon
по нашим SKU, складывать в data/ozon_api/ и использовать для SEO и рекламы.

ВАЖНО про эндпоинты: доступ к поисковым запросам у Ozon есть в аналитике
(раздел «Что покупают / Поисковые запросы»). Точный путь и доступность зависят
от тарифа кабинета (часть отчётов — Premium Plus). Ниже — рабочий каркас:
подставь актуальный путь из своей доки Ozon, остальная логика (пагинация,
сохранение, сравнение с прошлым срезом) готова.
"""
import json
import datetime as dt
from tools.ozon_api import client
from tools.common import config, notify

# TODO: подтвердить путь по доке Ozon для твоего тарифа.
SEARCH_QUERIES_PATH = '/v1/analytics/data'   # заменить на реальный эндпоинт поисковых запросов


def pull_clusters(date_from=None, date_to=None):
    """Тянет статистику поисковых запросов по SKU за период.
    Возвращает список {sku, query, frequency, clicks, orders, ...}."""
    today = dt.date.today()
    date_from = date_from or (today - dt.timedelta(days=30)).isoformat()
    date_to = date_to or today.isoformat()

    body = {
        'date_from': date_from, 'date_to': date_to,
        'metrics': ['ordered_units', 'revenue'],
        'dimension': ['sku'],          # заменить на срез поисковых запросов, когда известен путь
        'limit': 1000, 'offset': 0,
    }
    rows = []
    while True:
        res = client.post(SEARCH_QUERIES_PATH, body) or {}
        data = (res.get('result') or {}).get('data', [])
        rows.extend(data)
        if len(data) < body['limit']:
            break
        body['offset'] += body['limit']
    return rows


def run():
    today = dt.date.today().isoformat()
    try:
        rows = pull_clusters()
    except Exception as e:
        notify.send(f'<b>Ozon кластеры</b>: ошибка выгрузки — {e}. Проверь путь SEARCH_QUERIES_PATH и тариф кабинета.')
        return []
    out = config.DATA_OZON / f'clusters_{today}.json'
    out.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding='utf-8')
    notify.send(f'<b>Ozon кластеры</b> {today}: выгружено записей {len(rows)} → {out.name}')
    return rows


if __name__ == '__main__':
    run()
