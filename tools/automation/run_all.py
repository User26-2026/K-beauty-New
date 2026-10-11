"""Оркестратор автоматизации — запускает все задачи по очереди.

Локально:   python -m tools.automation.run_all
На сервере: вешается на cron (см. tools/automation/crontab.example).

Задачи:
  WB   — анализ рекламы (ДРР/CPO, пауза убыточных при APPLY_CHANGES)
  Ozon — остатки (+ алерт по заканчивающимся) и прогрузка кластеров
"""
import sys
import traceback

from tools.common import notify


def safe(name, fn):
    try:
        fn()
        print(f'[ok] {name}')
    except Exception as e:
        print(f'[FAIL] {name}: {e}')
        traceback.print_exc()
        notify.send(f'⚠️ Задача «{name}» упала: {e}')


def main(which='all'):
    if which in ('all', 'wb', 'ads'):
        from tools.wb_api import ads
        safe('WB реклама', lambda: ads.run(days=7))
    if which in ('all', 'ozon', 'stocks'):
        from tools.ozon_api import stocks
        safe('Ozon остатки', stocks.run)
    if which in ('all', 'ozon', 'clusters'):
        from tools.ozon_api import clusters
        safe('Ozon кластеры', clusters.run)


if __name__ == '__main__':
    main(sys.argv[1] if len(sys.argv) > 1 else 'all')
