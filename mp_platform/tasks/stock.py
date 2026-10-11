"""Задача ОСТАТКИ/ПОСТАВКИ. Прогноз запаса, сигнал на поставку/переброску.

Авто-часть: считает «дней запаса» и формирует список на дозаказ/переброску.
Реальную поставку оформляет человек (это физический процесс), поэтому действия
идут как awaiting_approval + алерт, даже в режиме auto.
"""
import datetime as dt
from .base import Task
from ..core import db, guardrails
from ..connectors.ozon import Ozon


class StockTask(Task):
    name = 'stock'

    def collect(self):
        oz = Ozon(self.client)
        try:
            stocks = oz.stocks()
        except Exception:
            stocks = []
        db.save_snapshot(self.client.id, 'ozon', 'stocks', stocks)
        return {'stocks': stocks}

    def analyze(self, data):
        days = guardrails.low_stock_days(self.client)
        props = []
        for s in data['stocks']:
            # скорость продаж подставляется из снапшотов продаж клиента (TODO: привязать)
            per_day = s.get('per_day', 0)
            left = round(s['stock'] / per_day, 1) if per_day else None
            if left is not None and left < days:
                props.append({'marketplace': 'ozon', 'sku': s.get('offer_id'),
                              'action': 'replenish', 'allowed': False,  # поставка — ручной процесс
                              'params': {'stock': s['stock'], 'days_left': left},
                              'why': f'хватит на {left} дн (<{days})'})
        return props

    def execute(self, p):
        # Поставку не исполняем автоматически — только сигнал.
        return 'awaiting_approval', 'нужна ручная поставка/переброска'
