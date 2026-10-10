# Автоматизация WB + Ozon (24/7)

Скрипты для постоянной автоматизации: **WB — реклама**, **Ozon — остатки и кластеры**.
Работают локально и на сервере (Hetzner) по расписанию (cron).

## Что делает

| Задача | Модуль | Что собирает / делает |
|---|---|---|
| WB реклама | `tools/wb_api/ads.py` | Кампании → ДРР/CPO → помечает убыточные, (опц.) ставит на паузу |
| Ozon остатки | `tools/ozon_api/stocks.py` | Остатки FBO/FBS + «дней запаса» → алерт по заканчивающимся |
| Ozon кластеры | `tools/ozon_api/clusters.py` | Поисковые запросы/кластеры по SKU → в `data/ozon_api/` |

Все выгрузки складываются в `data/wb_api/` и `data/ozon_api/`, алерты — в Telegram.

## Установка

```bash
cd <проект>
python3 -m venv .venv && source .venv/bin/activate
pip install requests python-dotenv openpyxl
cp .env.example .env          # впиши токены WB и Ozon
```

## Запуск (локально — проверка)

```bash
# проверить токены
python -c "from tools.ozon_api import client; print(client.ping())"

# по отдельности
python -m tools.automation.run_all stocks     # остатки Ozon
python -m tools.automation.run_all clusters    # кластеры Ozon
python -m tools.automation.run_all ads         # реклама WB
python -m tools.automation.run_all all         # всё сразу
```

По умолчанию `APPLY_CHANGES=false` — скрипты только **собирают данные и шлют алерты**,
ничего не меняя. Когда убедишься, что всё верно, ставь `APPLY_CHANGES=true`
для реальных действий (автопауза убыточной рекламы и т.п.).

## Сервер (Hetzner) — 24/7

1. Арендуй VPS (CX22/CPX11, Ubuntu 24.04).
2. Склонируй репозиторий в `/opt/k-beauty`, создай `.venv`, положи `.env`.
3. `mkdir logs`
4. `crontab -e` → вставь строки из `tools/automation/crontab.example`.

Готово — задачи крутятся по расписанию без запущенной сессии.

## Безопасность

- Токены только в `.env` (права `chmod 600 .env`), **не в git**.
- Автодействия включаем по одному, начиная с самого безопасного
  (пауза кампаний без заказов). Остальное — после проверки на отчётах.

## Что сверить с доками перед боем

- WB: пути `/adv/v1/promotion/count`, `/adv/v2/fullstats`, `/adv/v0/pause` — раздел «Продвижение».
- Ozon: `stocks.py` — `/v4/product/info/stocks`, `/v1/analytics/data`;
  `clusters.py` — путь поисковых запросов зависит от тарифа (часть — Premium Plus),
  подставь актуальный в `SEARCH_QUERIES_PATH`.
