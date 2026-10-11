"""Коннектор WB под конкретного клиента (токен берётся из его реестра)."""
import time
import json
import requests

ADV = 'https://advert-api.wildberries.ru'
STAT = 'https://statistics-api.wildberries.ru'
CONTENT = 'https://content-api.wildberries.ru'
DISCO = 'https://discounts-prices-api.wildberries.ru'


class WB:
    def __init__(self, client):
        env = client.marketplaces.get('wb', {}).get('token_env', 'WB_TOKEN')
        self.token = client.token(env)

    def _h(self):
        if not self.token:
            raise RuntimeError('WB токен не задан для клиента')
        return {'Authorization': self.token, 'Content-Type': 'application/json'}

    def _req(self, method, base, path, **kw):
        for a in range(4):
            try:
                r = requests.request(method, base + path, headers=self._h(), timeout=30, **kw)
                if r.status_code == 429:
                    time.sleep(2 ** a); continue
                r.raise_for_status()
                return r.json() if r.text.strip() else {}
            except requests.exceptions.RequestException:
                if a == 3:
                    raise
                time.sleep(2 ** a)

    # ── Реклама ──
    def campaigns(self):
        d = self._req('GET', ADV, '/adv/v1/promotion/count') or {}
        return [{'id': a.get('advertId'), 'type': g.get('type'), 'status': g.get('status')}
                for g in d.get('adverts', []) or [] for a in g.get('advert_list', []) or []]

    def fullstats(self, ids, begin, end):
        if not ids:
            return []
        body = [{'id': i, 'interval': {'begin': begin, 'end': end}} for i in ids]
        return self._req('POST', ADV, '/adv/v2/fullstats', data=json.dumps(body)) or []

    def pause(self, cid):
        return self._req('GET', ADV, '/adv/v0/pause', params={'id': cid})

    def set_bid(self, cid, cpm, type_, param):
        # POST /adv/v0/cpm  — ставка. Параметры сверь с докой.
        return self._req('POST', ADV, '/adv/v0/cpm',
                         data=json.dumps({'advertId': cid, 'type': type_, 'cpm': cpm, 'param': param}))

    # ── Цены ──
    def set_price(self, nm_id, price):
        # POST /api/v2/upload/task — изменение цены. Сверь путь с докой «Цены и скидки».
        body = {'data': [{'nmID': nm_id, 'price': int(price)}]}
        return self._req('POST', DISCO, '/api/v2/upload/task', data=json.dumps(body))

    # ── Контент (SEO) ──
    def update_card(self, payload):
        # POST /content/v2/cards/update — обновление карточки (название/описание).
        return self._req('POST', CONTENT, '/content/v2/cards/update', data=json.dumps(payload))
