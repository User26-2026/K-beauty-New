"""Сбор данных из WB API (только чтение) → data/wb_api/.

Из GitHub Actions резолвятся:
  - statistics-api.wildberries.ru (продажи, заказы, остатки, фин.отчёт)
  - advert-api.wildberries.ru   (рекламные кампании, ставки, статистика)
suppliers-api.wildberries.ru не резолвится из облачных runners GitHub.
"""
import os, json, time, datetime, requests
from pathlib import Path

WB_TOKEN = os.getenv('WB_TOKEN', '')

BASE_STAT  = 'https://statistics-api.wildberries.ru'
BASE_ADV   = 'https://advert-api.wildberries.ru'

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


def post(base, path, payload, retries=3):
    url = base.rstrip('/') + '/' + path.lstrip('/')
    for attempt in range(retries):
        try:
            r = requests.post(url, headers=headers(), json=payload, timeout=60)
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


def fetch_adv_campaigns():
    """Список рекламных кампаний (id, тип, статус).
    GET /adv/v1/promotion/count → группы по типам и статусам.
    Сохраняем сырой ответ и плоский список id для дальнейших запросов.
    """
    data = get(BASE_ADV, '/adv/v1/promotion/count')
    save('adv_campaigns', data)

    # Для статистики берём только кампании со статусами, у которых есть
    # открутка: 9 — активна, 11 — пауза. Иначе 1700+ кампаний × лимит
    # 1 запрос/мин превращают прогон в десятки минут.
    STAT_STATUSES = {9, 11}
    ids = []
    for grp in (data or {}).get('adverts', []) or []:
        if grp.get('status') not in STAT_STATUSES:
            continue
        for adv in grp.get('advert_list', []) or []:
            if 'advertId' in adv:
                ids.append(adv['advertId'])
    return ids


def fetch_adv_details(ids):
    """Детали кампаний по id (ставки, бюджеты, предметы/номенклатуры).
    POST /adv/v1/promotion/adverts — тело это массив id (макс 50 за запрос).
    """
    if not ids:
        print('  пропуск adv_details — нет id кампаний')
        return
    out = []
    for i in range(0, len(ids), 50):
        chunk = ids[i:i + 50]
        part = post(BASE_ADV, '/adv/v1/promotion/adverts', chunk)
        if isinstance(part, list):
            out.extend(part)
        time.sleep(1)  # лимит advert-api
    save('adv_details', out)


def fetch_adv_fullstats(ids):
    """Статистика по кампаниям за 30 дней (показы, клики, CTR, CPC, ДРР, расход).
    Актуальный эндпоинт — GET /adv/v3/fullstats (v2 POST устарел → 404,
    v3 принимает только GET → POST даёт 405). Параметры: ids (через запятую,
    макс 100), from/to (даты). Не чаще 1 запроса/мин.
    """
    if not ids:
        print('  пропуск adv_fullstats — нет активных id кампаний')
        return
    date_from = days_ago(30)
    date_to = datetime.date.today().isoformat()
    out = []
    for i in range(0, len(ids), 100):
        chunk = ids[i:i + 100]
        part = get(BASE_ADV, '/adv/v3/fullstats',
                   params={'ids': ','.join(str(c) for c in chunk),
                           'from': date_from, 'to': date_to})
        if isinstance(part, list):
            out.extend(part)
        elif part:
            out.append(part)
        if i + 100 < len(ids):
            time.sleep(60)  # жёсткий лимит 1 запрос/мин
    save('adv_fullstats', out)


def fetch_adv_balance():
    """Баланс рекламного счёта (баланс, бонусы, кошелёк)."""
    data = get(BASE_ADV, '/adv/v1/balance')
    save('adv_balance', data)


def fetch_adv_upd():
    """История списаний по рекламе за 30 дней."""
    data = get(BASE_ADV, '/adv/v1/upd',
               params={'from': days_ago(30), 'to': datetime.date.today().isoformat()})
    save('adv_upd', data)


def fetch_advertising():
    """Единая задача по рекламе: кампании → детали → статистика → баланс/списания.
    Под-запросы независимы: падение одного (напр. detail) не должно блокировать
    остальные, особенно fullstats со статистикой ДРР/CPC/расхода.
    """
    ids = fetch_adv_campaigns()
    print(f'  кампаний найдено: {len(ids)}')
    sub_errors = []
    for label, fn in [
        ('adv_details',   lambda: fetch_adv_details(ids)),
        ('adv_balance',   fetch_adv_balance),
        ('adv_upd',       fetch_adv_upd),
        ('adv_fullstats', lambda: fetch_adv_fullstats(ids)),  # последним — паузы по лимиту
    ]:
        try:
            fn()
        except Exception as e:
            print(f'  err {label}: {e}')
            sub_errors.append((label, str(e)))
    if sub_errors:
        # Сигнализируем наверх, но только если упало вообще всё рекламное.
        if len(sub_errors) == 4:
            raise RuntimeError(f'все рекламные под-запросы упали: {sub_errors}')
        print(f'  реклама: частичные ошибки {[s[0] for s in sub_errors]}')


TASKS = [
    ('sales',         fetch_sales),
    ('stocks',        fetch_stocks),
    ('orders',        fetch_orders),
    ('report_detail', fetch_report_detail),
    ('incomes',       fetch_incomes),
    ('advertising',   fetch_advertising),
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
        # Частичные ошибки не валят весь прогон: то, что скачалось, должно
        # закоммититься. Падаем только если не удалось вообще ничего.
        print(f'\n{len(errors)} ошибок: {[e[0] for e in errors]}')
        if len(errors) == len(TASKS):
            raise SystemExit('Все задачи упали — выход с ошибкой')
    print('Done.')
