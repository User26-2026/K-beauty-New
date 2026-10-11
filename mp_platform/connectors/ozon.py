"""Коннектор Ozon под конкретного клиента (Client-Id + Api-Key из его реестра)."""
import time
import requests

BASE = 'https://api-seller.ozon.ru'


class Ozon:
    def __init__(self, client):
        mp = client.marketplaces.get('ozon', {})
        self.client_id = client.token(mp.get('client_id_env', 'OZON_CLIENT_ID'))
        self.api_key = client.token(mp.get('api_key_env', 'OZON_API_KEY'))

    def _h(self):
        if not (self.client_id and self.api_key):
            raise RuntimeError('Ozon токены не заданы для клиента')
        return {'Client-Id': self.client_id, 'Api-Key': self.api_key, 'Content-Type': 'application/json'}

    def post(self, path, body=None):
        for a in range(4):
            try:
                r = requests.post(BASE + path, headers=self._h(), json=body or {}, timeout=40)
                if r.status_code == 429:
                    time.sleep(2 ** a); continue
                r.raise_for_status()
                return r.json()
            except requests.exceptions.RequestException:
                if a == 3:
                    raise
                time.sleep(2 ** a)

    def stocks(self):
        out, last = [], ''
        while True:
            res = self.post('/v4/product/info/stocks',
                           {'filter': {'visibility': 'ALL'}, 'limit': 1000, 'last_id': last}) or {}
            r = res.get('result') or {}
            for it in r.get('items', []):
                out.append({'sku': it.get('product_id'), 'offer_id': it.get('offer_id'),
                            'stock': sum(s.get('present', 0) for s in it.get('stocks', []))})
            last = r.get('last_id', '')
            if not last or not r.get('items'):
                break
        return out

    def set_price(self, offer_id, price):
        # POST /v1/product/import/prices
        return self.post('/v1/product/import/prices',
                         {'prices': [{'offer_id': offer_id, 'price': str(int(price)),
                                      'auto_action_enabled': 'UNKNOWN'}]})
