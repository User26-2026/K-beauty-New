"""Движок юнит-экономики (методика проекта). Используется ценообразованием и рекламой.

Формула (цена до СПП = P):
  маржа = P − комиссия − эквайринг − налог − логистика − хранение − себес − реклама
  ROI = маржа / себес * 100
Ставки по возможности брать фактические (Ракета), иначе — дефолты.
"""
from dataclasses import dataclass


@dataclass
class Rates:
    commission: float = 0.415   # комиссия WB, доля
    acquiring: float = 0.0235   # эквайринг
    tax: float = 0.06           # налог УСН/НДС
    drr: float = 0.10           # реклама, доля
    logistics: float = 90.0     # логистика ₽/ед
    storage: float = 0.0        # хранение ₽/ед
    spp: float = 0.30           # СПП (скидка покупателю)


def margin(price_before_spp: float, cost: float, r: Rates):
    """Маржа ₽/ед при цене до СПП."""
    P = price_before_spp
    g = 1 - r.commission - r.acquiring - r.tax - r.drr
    m = g * P - r.logistics - r.storage - cost
    return round(m, 2)


def roi(price_before_spp: float, cost: float, r: Rates):
    if not cost:
        return None
    return round(margin(price_before_spp, cost, r) / cost * 100, 1)


def price_for_roi(cost: float, target_roi_pct: float, r: Rates):
    """Цена до СПП, дающая заданный ROI."""
    g = 1 - r.commission - r.acquiring - r.tax - r.drr
    P = ((1 + target_roi_pct / 100) * cost + r.logistics + r.storage) / g
    return round(P)


def price_with_spp(price_before_spp: float, r: Rates):
    return round(price_before_spp * (1 - r.spp))
