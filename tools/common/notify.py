"""Уведомления в Telegram. Если токен не задан — просто печатаем в консоль/лог."""
import requests
from . import config


def send(text: str):
    """Отправить сообщение в Telegram (или вывести в консоль, если бот не настроен)."""
    print(text)  # всегда дублируем в лог
    if not (config.TG_BOT_TOKEN and config.TG_CHAT_ID):
        return
    try:
        requests.post(
            f'https://api.telegram.org/bot{config.TG_BOT_TOKEN}/sendMessage',
            json={'chat_id': config.TG_CHAT_ID, 'text': text, 'parse_mode': 'HTML'},
            timeout=15,
        )
    except requests.exceptions.RequestException as e:
        print(f'[notify] не удалось отправить в Telegram: {e}')
