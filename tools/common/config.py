"""Единая загрузка настроек из .env для всех скриптов автоматизации.

Токены и секреты НИКОГДА не коммитим — только в .env (он в .gitignore).
Пример значений — в .env.example.
"""
import os
from pathlib import Path
from dotenv import load_dotenv

ROOT = Path(__file__).resolve().parent.parent.parent
load_dotenv(ROOT / '.env')

# ── Wildberries ──────────────────────────────────────────────────────────────
WB_TOKEN = os.getenv('WB_TOKEN', '')            # один токен со всеми категориями

# ── Ozon ─────────────────────────────────────────────────────────────────────
OZON_CLIENT_ID = os.getenv('OZON_CLIENT_ID', '')
OZON_API_KEY   = os.getenv('OZON_API_KEY', '')

# ── Telegram-уведомления (опционально) ───────────────────────────────────────
TG_BOT_TOKEN = os.getenv('TG_BOT_TOKEN', '')
TG_CHAT_ID   = os.getenv('TG_CHAT_ID', '')

# ── Параметры автоматизации ──────────────────────────────────────────────────
# Порог остатка (дней запаса), ниже которого шлём алерт "скоро закончится".
LOW_STOCK_DAYS = int(os.getenv('LOW_STOCK_DAYS', '14'))
# Порог ДРР (%), выше которого реклама считается неэффективной.
MAX_DRR = float(os.getenv('MAX_DRR', '20'))
# Режим реальных изменений. По умолчанию FALSE — только отчёт/алерт, ничего не меняем.
APPLY_CHANGES = os.getenv('APPLY_CHANGES', 'false').lower() == 'true'

# Куда складывать выгрузки
DATA_WB   = ROOT / 'data' / 'wb_api'
DATA_OZON = ROOT / 'data' / 'ozon_api'
DATA_WB.mkdir(parents=True, exist_ok=True)
DATA_OZON.mkdir(parents=True, exist_ok=True)


def require(*names: str):
    """Проверка, что нужные переменные заданы; иначе понятная ошибка."""
    missing = [n for n in names if not globals().get(n)]
    if missing:
        raise SystemExit(f'Не заданы в .env: {", ".join(missing)}. См. .env.example')
