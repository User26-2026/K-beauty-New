"""Задача РЕКЛАМА (WB). Автоматически: пауза убыточных, снижение ставок при высоком ДРР."""
import datetime as dt
from .base import Task
from ..core import db, guardrails
from ..connectors.wb import WB


class AdsTask(Task):
    name = 'ads'

    def collect(self):
        wb = WB(self.client)
        today = dt.date.today()
        camps = wb.campaigns()
        active = [c['id'] for c in camps if c['id']]
        stats = wb.fullstats(active, (today - dt.timedelta(days=7)).isoformat(), today.isoformat())
        db.save_snapshot(self.client.id, 'wb', 'ads', stats)
        return {'wb': wb, 'stats': stats}

    def analyze(self, data):
        props = []
        for s in data['stats']:
            spend = s.get('sum', 0) or 0
            orders = s.get('orders', 0) or 0
            revenue = s.get('sum_price', 0) or s.get('orders_sum', 0) or 0
            drr = round(spend / revenue * 100, 1) if revenue else None
            cid = s.get('advertId') or s.get('id')
            # Нет заказов, но есть расход — пауза
            if orders == 0 and spend > 0:
                props.append({'marketplace': 'wb', 'sku': cid, 'action': 'pause_campaign',
                              'params': {'id': cid}, 'allowed': True,
                              'why': f'нет заказов, потрачено {round(spend)}₽'})
                continue
            flag, _ = guardrails.check_ad_drr(self.client, drr)
            if flag:
                props.append({'marketplace': 'wb', 'sku': cid, 'action': 'reduce_bid',
                              'params': {'id': cid, 'drr': drr}, 'allowed': True,
                              'why': f'ДРР {drr}% выше лимита'})
        return props

    def execute(self, p):
        wb = WB(self.client)
        try:
            if p['action'] == 'pause_campaign':
                wb.pause(p['params']['id'])
                return 'applied', 'кампания на паузе'
            if p['action'] == 'reduce_bid':
                # шаг снижения ставки в рамках лимита — тут безопасная заглушка:
                # реальное снижение через wb.set_bid(...) после подтверждения текущей ставки.
                return 'awaiting_approval', 'снижение ставки требует текущей ставки/типа кампании'
        except Exception as e:
            return 'failed', str(e)
        return 'skipped', ''
