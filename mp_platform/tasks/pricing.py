"""Задача ЦЕНООБРАЗОВАНИЕ. Держит цену под целевой ROI в рамках guardrails.

Вход: себестоимость и текущая цена по SKU (из снапшотов/выгрузок клиента).
Логика: если ROI ниже минимума — поднять цену до целевого ROI, но не больше
лимита изменения за раз. Исполнение — set_price (WB/Ozon) в режиме auto.
"""
from .base import Task
from ..core import db, guardrails
from ..engine import unit_economics as ue
from ..connectors.wb import WB
from ..connectors.ozon import Ozon


class PricingTask(Task):
    name = 'pricing'

    def collect(self):
        # Ожидается список {sku, marketplace, cost, price, rates?} из данных клиента.
        # Здесь — подключение к источнику цен/себеса клиента (выгрузки или API).
        # Возвращаем пусто, если источник ещё не настроен.
        return {'items': self._load_items()}

    def _load_items(self):
        # TODO: привязать к реальному источнику (снапшоты цен + себес клиента).
        # Для эталонного магазина — читать из data/ (себес по последнему приходу + цены).
        return []

    def analyze(self, data):
        g = (self.client.guardrails or {}).get('pricing', {})
        target = g.get('min_roi_pct', 30)
        props = []
        for it in data['items']:
            cost = it.get('cost') or 0
            price = it.get('price') or 0
            r = ue.Rates(**it.get('rates', {})) if it.get('rates') else ue.Rates()
            cur_roi = ue.roi(price, cost, r)
            if cur_roi is None or cur_roi >= target:
                continue
            new_price = ue.price_for_roi(cost, target, r)
            new_roi = ue.roi(new_price, cost, r)
            ok, reason = guardrails.check_price_change(self.client, price, new_price, new_roi)
            props.append({'marketplace': it['marketplace'], 'sku': it['sku'],
                          'action': 'set_price', 'allowed': ok,
                          'params': {'old': price, 'new': new_price, 'roi': new_roi},
                          'why': f'ROI {cur_roi}%→{new_roi}% ({reason})'})
        return props

    def execute(self, p):
        try:
            if p['marketplace'] == 'wb':
                WB(self.client).set_price(p['sku'], p['params']['new'])
            elif p['marketplace'] == 'ozon':
                Ozon(self.client).set_price(p['sku'], p['params']['new'])
            return 'applied', f"цена → {p['params']['new']}₽"
        except Exception as e:
            return 'failed', str(e)
