"""Сбор комиссий, тарифов и логистики из Ozon Seller API (только чтение).

Источник ставок — метод цен товаров: по каждому SKU Ozon возвращает блок
`commissions` с комиссией продажи (FBO/FBS) и всеми плечами логистики
(магистраль, последняя миля, обработка, обратная логистика). Это точные
цифры под наши артикулы, а не средние по рынку.

Ключи берём из окружения (в GitHub Actions — из secrets):
  OZON_CLIENT_ID, OZON_API_KEY

Результат:
  data/ozon_api/*.json           — сырые ответы API
  outputs/ozon/ozon_tariffs_<дата>.tsv   — таблица для Google Sheets
  outputs/ozon/ozon_tariffs_<дата>.xlsx  — та же таблица книгой Excel
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

TODAY = datetime.date.today().isoformat()


def headers():
    if not (CLIENT_ID and API_KEY):
        raise ValueError('OZON_CLIENT_ID / OZON_API_KEY не заданы')
    return {
        'Client-Id': CLIENT_ID,
        'Api-Key': API_KEY,
        'Content-Type': 'application/json',
    }


def post(path, payload, retries=4):
    url = BASE.rstrip('/') + '/' + path.lstrip('/')
    for attempt in range(retries):
        try:
            r = requests.post(url, headers=headers(), json=payload, timeout=60)
            if r.status_code == 429:  # лимит запросов — ждём и повторяем
                time.sleep(2 ** attempt)
                continue
            r.raise_for_status()
            return r.json()
        except requests.exceptions.RequestException:
            if attempt == retries - 1:
                raise
            time.sleep(2 ** attempt)


def save_json(name, data):
    path = OUT_DATA / f'{name}.json'
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2))
    n = len(data) if isinstance(data, list) else path.stat().st_size
    print(f'  ok {path.name} — {n}')


# --- Сбор данных --------------------------------------------------------

def fetch_prices():
    """Цены и комиссии по всем товарам (POST /v5/product/info/prices).

    Пагинация курсором: пустой cursor — первая страница, дальше берём
    cursor из ответа, пока он не опустеет. limit 1000 — максимум Ozon.
    """
    items, cursor = [], ''
    while True:
        body = {'cursor': cursor, 'limit': 1000,
                'filter': {'visibility': 'ALL'}}
        data = post('/v5/product/info/prices', body) or {}
        page = data.get('items', []) or []
        items.extend(page)
        cursor = data.get('cursor', '') or ''
        print(f'  prices: +{len(page)} (всего {len(items)})')
        if not cursor or not page:
            break
        time.sleep(0.3)
    save_json('prices', items)
    return items


def num(v, default=0.0):
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


def _liters_from_dims(depth, width, height, unit):
    """Объём в литрах из габаритов: (Д×Ш×В в см) / 1000. Единицы — mm/cm/m."""
    f = {'mm': 0.1, 'cm': 1.0, 'дм': 10.0, 'м': 100.0, 'm': 100.0}.get(
        (unit or 'mm').lower(), 0.1)
    d, w, h = num(depth), num(width), num(height)
    if d > 0 and w > 0 and h > 0:
        return round((d * f) * (w * f) * (h * f) / 1000.0, 3)
    return None


def fetch_product_names(product_ids):
    """Имена, категория и габариты товаров (POST /v3/product/info/list).

    Из габаритов (depth/width/height/dimension_unit) считаем объём в литрах.
    Ошибки не критичны — тогда идём без имён/категорий/литража.
    """
    info = {}
    ids = [i for i in product_ids if i]
    for i in range(0, len(ids), 1000):
        chunk = ids[i:i + 1000]
        try:
            data = post('/v3/product/info/list',
                        {'product_id': chunk, 'offer_id': [], 'sku': []}) or {}
            for it in data.get('items', []) or []:
                info[it.get('id')] = {
                    'name': it.get('name', ''),
                    'category_id': it.get('description_category_id', ''),
                    'liters': _liters_from_dims(it.get('depth'), it.get('width'),
                                                it.get('height'),
                                                it.get('dimension_unit')),
                }
        except Exception as e:
            print(f'  info/list чанк {i}: {e}')
        time.sleep(0.3)
    save_json('product_info', info)
    return info


def fetch_dimensions(product_ids, info):
    """Габариты через /v4/product/info/attributes — добираем литраж там, где
    info/list не отдал габариты. Пагинация через last_id.
    """
    ids = [i for i in product_ids if i]
    for i in range(0, len(ids), 1000):
        chunk = ids[i:i + 1000]
        last = ''
        while True:
            try:
                data = post('/v4/product/info/attributes', {
                    'filter': {'product_id': [str(x) for x in chunk],
                               'visibility': 'ALL'},
                    'limit': 1000, 'last_id': last}) or {}
            except Exception as e:
                print(f'  attributes чанк {i}: {e}')
                break
            for it in data.get('result', []) or []:
                pid = it.get('id')
                lit = _liters_from_dims(it.get('depth'), it.get('width'),
                                        it.get('height'), it.get('dimension_unit'))
                if not lit:
                    continue
                if pid in info and not info[pid].get('liters'):
                    info[pid]['liters'] = lit
                elif pid not in info:
                    info[pid] = {'liters': lit}
            last = data.get('last_id', '') or ''
            if not last:
                break
            time.sleep(0.3)
    save_json('product_info', info)
    return info


def fetch_category_names():
    """Дерево категорий (POST /v1/description-category/tree) → {id: имя}.

    Best-effort: если не отдалось, таблица просто будет без названий категорий.
    """
    id2name = {}
    try:
        data = post('/v1/description-category/tree', {'language': 'RU'}) or {}

        def walk(nodes):
            for n in nodes or []:
                cid = n.get('description_category_id')
                if cid is not None:
                    id2name[cid] = n.get('category_name', '')
                walk(n.get('children'))

        walk(data.get('result', []))
        save_json('category_tree_flat', id2name)
    except Exception as e:
        print(f'  category tree: {e}')
    return id2name


# --- Формирование таблицы ----------------------------------------------

# (ключ в ответе commissions/price, заголовок колонки)
COMMISSION_COLS = [
    ('sales_percent_fbo',                'Комиссия FBO, %'),
    ('sales_percent_fbs',                'Комиссия FBS, %'),
    ('fbo_fulfillment_amount',           'FBO обработка, ₽'),
    ('fbo_direct_flow_trans_min_amount', 'FBO логистика (мин), ₽'),
    ('fbo_direct_flow_trans_max_amount', 'FBO логистика (макс), ₽'),
    ('fbo_deliv_to_customer_amount',     'FBO последняя миля, ₽'),
    ('fbo_return_flow_amount',           'FBO обратная логистика, ₽'),
    ('fbs_first_mile_min_amount',        'FBS первая миля, ₽'),
    ('fbs_direct_flow_trans_min_amount', 'FBS логистика (мин), ₽'),
    ('fbs_direct_flow_trans_max_amount', 'FBS логистика (макс), ₽'),
    ('fbs_deliv_to_customer_amount',     'FBS последняя миля, ₽'),
    ('fbs_return_flow_amount',           'FBS обратная логистика, ₽'),
]

HEAD = ['Артикул (offer_id)', 'product_id', 'Наименование', 'Категория',
        'Цена, ₽', 'Объёмный вес, л'] + [h for _, h in COMMISSION_COLS]


def build_rows(prices, info, cat_names):
    rows = []
    for it in prices:
        pid = it.get('product_id')
        meta = info.get(pid, {})
        cat_id = meta.get('category_id', '')
        comm = it.get('commissions', {}) or {}
        price_block = it.get('price', {}) or {}
        row = [
            it.get('offer_id', ''),
            pid,
            meta.get('name', ''),
            cat_names.get(cat_id, cat_id or ''),
            price_block.get('price', ''),
            it.get('volume_weight', ''),
        ]
        row += [comm.get(k, '') for k, _ in COMMISSION_COLS]
        rows.append(row)
    rows.sort(key=lambda r: (str(r[3]), str(r[0])))  # по категории, потом артикулу
    return rows


def write_tsv(rows):
    path = OUT_XLS / f'ozon_tariffs_{TODAY}.tsv'
    lines = ['\t'.join(HEAD)]
    for r in rows:
        lines.append('\t'.join('' if v is None else str(v) for v in r))
    path.write_text('\n'.join(lines), encoding='utf-8')
    print(f'  ok {path.name} — {len(rows)} строк')


def write_xlsx(rows):
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font
    except ImportError:
        print('  openpyxl нет — пропускаю xlsx')
        return
    wb = Workbook()
    ws = wb.active
    ws.title = 'Ozon тарифы'
    ws.append(HEAD)
    for c in ws[1]:
        c.font = Font(bold=True)
    for r in rows:
        ws.append(['' if v is None else v for v in r])
    ws.freeze_panes = 'A2'
    path = OUT_XLS / f'ozon_tariffs_{TODAY}.xlsx'
    wb.save(path)
    print(f'  ok {path.name}')


# --- main ---------------------------------------------------------------

if __name__ == '__main__':
    print(f'Ozon API fetch — {TODAY}')
    if not (CLIENT_ID and API_KEY):
        raise SystemExit('OZON_CLIENT_ID / OZON_API_KEY не заданы — выход')

    print('-> prices (комиссии + логистика)')
    prices = fetch_prices()

    print('-> product_info (имена, категории, габариты→литры)')
    pids = [p.get('product_id') for p in prices]
    info = fetch_product_names(pids)
    print('-> dimensions (attributes) — добор литража')
    info = fetch_dimensions(pids, info)

    print('-> category tree')
    cat_names = fetch_category_names()

    rows = build_rows(prices, info, cat_names)
    write_tsv(rows)
    write_xlsx(rows)

    meta = {
        'fetched_at': datetime.datetime.utcnow().isoformat() + 'Z',
        'products': len(prices),
    }
    (OUT_DATA / 'fetch_meta.json').write_text(
        json.dumps(meta, ensure_ascii=False, indent=2))

    if not prices:
        print('Внимание: товаров в кабинете не найдено — таблица пустая '
              '(проверь права API-ключа: нужен доступ к методам товаров).')
    print('Done.')
