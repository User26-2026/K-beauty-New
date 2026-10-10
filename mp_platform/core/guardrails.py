"""Ограничители автоматических действий. Без них автомат опасен на чужих магазинах.

Каждое действие ИИ проверяется против рамок клиента (guardrails в реестре).
Возвращает (ok, reason). Если ok=False — действие не исполняется, идёт в отчёт.
"""


def check_price_change(client, old_price, new_price, new_roi_pct):
    g = (client.guardrails or {}).get('pricing', {})
    min_roi = g.get('min_roi_pct', 30)
    max_chg = g.get('max_price_change_pct', 15)
    if new_roi_pct is not None and new_roi_pct < min_roi:
        return False, f'ROI {new_roi_pct}% ниже минимума {min_roi}%'
    if old_price:
        chg = abs(new_price - old_price) / old_price * 100
        if chg > max_chg:
            return False, f'Изменение цены {chg:.0f}% больше лимита {max_chg}%'
    return True, 'ok'


def check_bid_change(client, old_bid, new_bid):
    g = (client.guardrails or {}).get('ads', {})
    max_chg = g.get('max_bid_change_pct', 20)
    if old_bid:
        chg = abs(new_bid - old_bid) / old_bid * 100
        if chg > max_chg:
            return False, f'Изменение ставки {chg:.0f}% больше лимита {max_chg}%'
    return True, 'ok'


def check_ad_drr(client, drr_pct):
    g = (client.guardrails or {}).get('ads', {})
    cap = g.get('max_drr_pct', 20)
    if drr_pct is not None and drr_pct > cap:
        return True, f'ДРР {drr_pct}% выше {cap}% — кандидат на пауза/снижение ставки'
    return False, 'в норме'


def low_stock_days(client):
    return (client.guardrails or {}).get('stock', {}).get('low_stock_days', 14)
