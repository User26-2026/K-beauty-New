"""Главный прогон платформы: по каждому активному клиенту — все задачи.

Локально:   python -m mp_platform.runner [client_id] [task]
На сервере: cron (см. mp_platform/crontab.example).
"""
import sys
import traceback

from . import config
from .core import notify
from .tasks.ads import AdsTask
from .tasks.pricing import PricingTask
from .tasks.stock import StockTask
from .tasks.seo import SeoTask

TASKS = {'ads': AdsTask, 'pricing': PricingTask, 'stock': StockTask, 'seo': SeoTask}


def run_client(client, only=None):
    for name, cls in TASKS.items():
        if only and name != only:
            continue
        try:
            cls(client).run()
            print(f'[ok] {client.id}/{name} ({client.mode(name)})')
        except Exception as e:
            print(f'[FAIL] {client.id}/{name}: {e}')
            traceback.print_exc()
            notify.send(f'⚠️ {client.name}: задача {name} упала — {e}')


def main():
    only_client = sys.argv[1] if len(sys.argv) > 1 and sys.argv[1] != 'all' else None
    only_task = sys.argv[2] if len(sys.argv) > 2 else None
    for client in config.load_clients():
        if only_client and client.id != only_client:
            continue
        run_client(client, only_task)


if __name__ == '__main__':
    main()
