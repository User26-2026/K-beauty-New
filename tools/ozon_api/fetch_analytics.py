"""Собственная аналитика кабинета из Ozon Seller API (только чтение).

Метод POST /v1/analytics/data отдаёт метрики ПО НАШИМ карточкам: показы,
добавления в корзину, конверсия, заказы, выручка, возвраты, отмены. Это не
данные конкурентов (их публичный API не отдаёт), а мониторинг наших продаж —
чтобы после запуска видеть динамику рядом с тарифами.

Ключи из окружения (в GitHub Actions — из secrets):
  OZON_CLIENT_ID, OZON_API_KEY

Результат:
  data/ozon_api/analytics_by_sku.json / analytics_by_day.json  — сырые ответы
  outputs/ozon/ozon_analytics_sku_<дата>.tsv/.xlsx             — по SKU за 28 дней
  outputs/ozon/ozon_analytics_day_<дата>.tsv/.xlsx            — по дням (тренд)
"""
import os, json, time, datetime, requests
from pathlib import Path

CLIENT_ID = os.getenv('OZON_CLIENT_ID', '')
API_KEY   = os.getenv('OZON_API_KEY', '')

BASE = 'https://api-seller.ozon.ru'

ROOT = Path(__file__).parent.parent.parent
OUT_DATA = ROOT / 'data' / 'ozon_api'
OUT_XLS  = ROOT / 'outputs' / 'ozon'
OUT_DATA.mkdir(parents=True, exist_ok=True)
OUT_XLS.mkdir(parents=True, exist_ok=True)

TODAY = datetime.date.today()
TODAY_S = TODAY.isoformat()
# Данные аналитики приходят с задержкой ~1 день: берём окно [-29; -1] день.
DATE_TO   = (TODAY - datetime.timedelta(days=1)).isoformat()
DATE_FROM = (TODAY - datetime.timedelta(days=28)).isoformat()

# До 14 метрик за запрос. Порядок сохраняем для колонок.
METRICS = [
    ('revenue',          'Выручка, ₽'),
    ('ordered_units',    'Заказано, шт'),
    ('delivered_units',  'Доставлено, шт'),
    ('returns',          'Возвраты'),
    ('cancellations',    'Отмены'),
    ('hits_view',        'Показы'),
    ('hits_tocart',      'В корзину'),
    ('session_view',     'Сессии с товаром'),
    ('conv_tocart',      'Конв. в корзину'),
    ('position_category','Позиция в категории'),
]
METRIC_KEYS = [k for k, _ in METRICS]


def headers():
    if not (CLIENT_ID and API_KEY):
        raise ValueError('OZON_CLIENT_ID / OZON_API_KEY не заданы')
    return {'Client-Id': CLIENT_ID, 'Api-Key': API_KEY,
            'Content-Type': 'application/json'}


def post(path, payload, retries=4):
    url = BASE.rstrip('/') + '/' + path.lstrip('/')
    for attempt in range(retries):
        try:
            r = requests.post(url, headers=headers(), json=payload, timeout=60)
            if r.status_code == 429:
                time.sleep(2 ** attempt)
                continue
            if r.status_code >= 400:
                print(f'  [ДИАГ] {path} -> HTTP {r.status_code}: {r.text[:300]}')
            r.raise_for_status()
            return r.json()
        except requests.exceptions.RequestException:
            if attempt == retries - 1:
                raise
            time.sleep(2 ** attempt)


def save_json(name, data):
    path = OUT_DATA / f'{name}.json'
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2))
    print(f'  ok {path.name}')


def fetch_analytics(dimension):
    """Аналитика по нашим товарам за 28 дней в разрезе `dimension`
    (['sku'] — по карточкам, ['day'] — тренд). Пагинация offset/limit.
    """
    rows, offset = [], 0
    while True:
        body = {
            'date_from': DATE_FROM, 'date_to': DATE_TO,
            'metrics': METRIC_KEYS, 'dimension': dimension,
            'filters': [], 'sort': [{'key': 'revenue', 'order': 'DESC'}],
            'limit': 1000, 'offset': offset,
        }
        data = (post('/v1/analytics/data', body) or {}).get('result', {}) or {}
        page = data.get('data', []) or []
        rows.extend(page)
        print(f'  analytics {dimension}: +{len(page)} (всего {len(rows)})')
        if len(page) < 1000:
            break
        offset += 1000
        time.sleep(0.3)
    return rows


def rows_to_table(raw):
    """Плоские строки: [id, имя, метрика1..N] из ответа analytics/data."""
    out = []
    for r in raw:
        dims = r.get('dimensions', []) or []
        did = dims[0].get('id', '') if dims else ''
        dname = dims[0].get('name', '') if dims else ''
        vals = r.get('metrics', []) or []
        out.append([did, dname] + [vals[i] if i < len(vals) else ''
                                   for i in range(len(METRIC_KEYS))])
    return out


def write_tables(tag, dim_head, rows):
    head = [dim_head, 'Наименование'] + [h for _, h in METRICS]
    tsv = OUT_XLS / f'ozon_analytics_{tag}_{TODAY_S}.tsv'
    lines = ['\t'.join(head)]
    for r in rows:
        lines.append('\t'.join('' if v is None else str(v) for v in r))
    tsv.write_text('\n'.join(lines), encoding='utf-8')
    print(f'  ok {tsv.name} — {len(rows)} строк')
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font
        wb = Workbook(); ws = wb.active; ws.title = f'Аналитика {tag}'
        ws.append(head)
        for c in ws[1]:
            c.font = Font(bold=True)
        for r in rows:
            ws.append(['' if v is None else v for v in r])
        ws.freeze_panes = 'A2'
        wb.save(OUT_XLS / f'ozon_analytics_{tag}_{TODAY_S}.xlsx')
        print(f'  ok ozon_analytics_{tag}_{TODAY_S}.xlsx')
    except ImportError:
        print('  openpyxl нет — пропускаю xlsx')


if __name__ == '__main__':
    print(f'Ozon analytics fetch — окно {DATE_FROM}..{DATE_TO}')
    if not (CLIENT_ID and API_KEY):
        raise SystemExit('OZON_CLIENT_ID / OZON_API_KEY не заданы — выход')

    print('-> analytics по SKU')
    try:
        by_sku = fetch_analytics(['sku'])
        save_json('analytics_by_sku', by_sku)
        write_tables('sku', 'SKU', rows_to_table(by_sku))
    except Exception as e:
        print(f'  analytics по SKU: {e}')
        by_sku = []

    print('-> analytics по дням')
    try:
        by_day = fetch_analytics(['day'])
        save_json('analytics_by_day', by_day)
        write_tables('day', 'Дата', rows_to_table(by_day))
    except Exception as e:
        print(f'  analytics по дням: {e}')

    (OUT_DATA / 'analytics_meta.json').write_text(json.dumps({
        'fetched_at': datetime.datetime.utcnow().isoformat() + 'Z',
        'date_from': DATE_FROM, 'date_to': DATE_TO,
        'sku_rows': len(by_sku),
    }, ensure_ascii=False, indent=2))

    if not by_sku:
        print('Внимание: аналитика пустая — вероятно, продаж/показов за период '
              'ещё нет (кабинет до запуска) или у ключа нет доступа к аналитике.')
    print('Done.')
