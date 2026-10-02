"""Задача SEO. ИИ генерирует название/описание из кластеров; публикация — через подтверждение.

По умолчанию mode=suggest: ИИ готовит SEO, кладёт в журнал как awaiting_approval.
Публикация карточки (update_card) исполняется только когда клиент разрешил auto И
действие подтверждено (контент меняет восприятие бренда — осторожно).
"""
import json
from .base import Task
from ..core import db, ai
from ..connectors.wb import WB


class SeoTask(Task):
    name = 'seo'

    def collect(self):
        # Ожидается список {sku, name, clusters:[{phrase,freq}], composition} из данных клиента.
        return {'items': self._load_items()}

    def _load_items(self):
        # TODO: привязать к кластерам (Ozon/WB) и составам товаров клиента.
        return []

    def analyze(self, data):
        props = []
        for it in data['items']:
            clusters = ', '.join(f"{c['phrase']} ({c.get('freq','')})" for c in it.get('clusters', [])[:25])
            prompt = (f"Товар: {it['name']}. Состав: {it.get('composition','')}. "
                      f"Кластеры из топа: {clusters}. "
                      "Сделай SEO-название (≤60 симв, без бренда, главный товарный ключ) "
                      "и SEO-описание (900-1400 знаков, кластеры естественно). Верни JSON {title, description}.")
            out = ai.ask(prompt, system=ai.SEO_SYSTEM)
            props.append({'marketplace': it.get('marketplace', 'wb'), 'sku': it['sku'],
                          'action': 'update_card', 'allowed': False,  # публикация только после подтверждения
                          'params': {'seo': out}, 'why': 'сгенерировано ИИ, на проверку'})
        return props

    def execute(self, p):
        # Публикуем только если клиент явно разрешил и контент подтверждён.
        return 'awaiting_approval', 'SEO готово, нужна проверка перед публикацией'
