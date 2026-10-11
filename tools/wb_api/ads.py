"""WB — автоматизация рекламы (advert-api.wildberries.ru).

Что делает:
  1. Тянет список активных рекламных кампаний.
  2. Забирает статистику (показы, клики, затраты, заказы, выручка).
  3. Считает ДРР и CPO по каждой кампании.
  4. Помечает неэффективные (ДРР выше порога) и — в режиме APPLY_CHANGES —
     может ставить их на паузу. По умолчанию ничего не меняет, только отчёт.

ВАЖНО: пути API сверяй с актуальной докой WB (раздел «Продвижение»).
Токен — один, с галочкой «Продвижение». Заголовок: Authorization: <token> (без Bearer).
"""
import time
import json
import datetime as dt
import requests

from tools.common import config, notify

BASE = 'https://advert-api.wildberries.ru'
STATUSES_ACTIVE = (9, 11)   # 9 — идёт, 11 — на паузе (пример; сверь с докой)


def _headers():
    config.require('WB_TOKEN')
    return {'Authorization': config.WB_TOKEN, 'Content-Type': 'application/json'}


def _req(method, path, **kw):
    url = BASE + path
    for attempt in range(4):
        try:
            r = requests.request(method, url, headers=_headers(), timeout=30, **kw)
            if r.status_code == 429:
                time.sleep(2 ** attempt); continue
            r.raise_for_status()
            return r.json() if r.text.strip() else {}
        except requests.exceptions.RequestException:
            if attempt == 3:
                raise
            time.sleep(2 ** attempt)


def list_campaigns():
    """Список кампаний, сгруппированный по типу/статусу.
    GET /adv/v1/promotion/count -> {"adverts":[{"type":..,"status":..,"advert_list":[{"advertId":..}]}]}"""
    data = _req('GET', '/adv/v1/promotion/count') or {}
    ids = []
    for grp in data.get('adverts', []) or []:
        for a in grp.get('advert_list', []) or []:
            ids.append({'id': a.get('advertId'), 'type': grp.get('type'), 'status': grp.get('status')})
    return ids


def fullstats(campaign_ids, date_from, date_to):
    """Статистика кампаний за период.
    POST /adv/v2/fullstats  body=[{"id":123,"interval":{"begin":"YYYY-MM-DD","end":"YYYY-MM-DD"}}]"""
    if not campaign_ids:
        return []
    body = [{'id': cid, 'interval': {'begin': date_from, 'end': date_to}} for cid in campaign_ids]
    # WB ограничивает размер запроса — бьём по 50
    out = []
    for i in range(0, len(body), 50):
        chunk = body[i:i + 50]
        res = _req('POST', '/adv/v2/fullstats', data=json.dumps(chunk))
        if isinstance(res, list):
            out.extend(res)
        time.sleep(1)
    return out


def analyze(stats):
    """Считает ДРР и CPO по каждой кампании. Возвращает список строк отчёта."""
    rows = []
    for s in stats:
        spend   = s.get('sum', 0) or 0
        orders  = s.get('orders', 0) or 0
        revenue = s.get('sum_price', 0) or s.get('orders_sum', 0) or 0
        drr = round(spend / revenue * 100, 1) if revenue else None
        cpo = round(spend / orders, 1) if orders else None
        rows.append({
            'id': s.get('advertId') or s.get('id'),
            'views': s.get('views', 0), 'clicks': s.get('clicks', 0),
            'ctr': s.get('ctr', 0), 'spend': round(spend, 1),
            'orders': orders, 'revenue': round(revenue, 1),
            'drr': drr, 'cpo': cpo,
            'verdict': ('НЕЭФФЕКТИВНА' if (drr is not None and drr > config.MAX_DRR)
                        else ('НЕТ ЗАКАЗОВ' if orders == 0 and spend > 0 else 'ок')),
        })
    return rows


def pause_campaign(cid):
    """Поставить кампанию на паузу. GET /adv/v0/pause?id=<cid>"""
    return _req('GET', '/adv/v0/pause', params={'id': cid})


def run(days=7):
    """Основной прогон: собрать статистику за N дней, отчёт, алерт, (опц.) паузу убыточных."""
    today = dt.date.today()
    date_from = (today - dt.timedelta(days=days)).isoformat()
    date_to = today.isoformat()

    camps = list_campaigns()
    active_ids = [c['id'] for c in camps if c['status'] in STATUSES_ACTIVE and c['id']]
    stats = fullstats(active_ids, date_from, date_to)
    rows = analyze(stats)

    # Сохраняем выгрузку
    out = config.DATA_WB / f'ads_report_{date_to}.json'
    out.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding='utf-8')

    bad = [r for r in rows if r['verdict'] != 'ок']
    msg = [f'<b>WB реклама</b> {date_from}…{date_to}: кампаний {len(rows)}, проблемных {len(bad)}.']
    for r in sorted(bad, key=lambda x: -(x['spend'] or 0))[:15]:
        msg.append(f"#{r['id']}: ДРР {r['drr']}% CPO {r['cpo']}₽ затраты {r['spend']}₽ заказы {r['orders']} — {r['verdict']}")

    paused = []
    if config.APPLY_CHANGES:
        for r in bad:
            if r['verdict'] == 'НЕТ ЗАКАЗОВ':   # самый безопасный кейс для автопаузы
                try:
                    pause_campaign(r['id']); paused.append(r['id'])
                except Exception as e:
                    msg.append(f"не удалось поставить на паузу #{r['id']}: {e}")
        if paused:
            msg.append(f'Поставлены на паузу (нет заказов): {paused}')
    else:
        msg.append('Режим отчёта (APPLY_CHANGES=false) — ничего не меняли.')

    notify.send('\n'.join(msg))
    return rows


if __name__ == '__main__':
    run()
