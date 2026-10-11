"""ИИ-слой платформы — вызовы Claude API (это «мозг», работающий на сервере).

Используется там, где нужно рассуждение, а не формула: генерация SEO,
приоритизация действий, объяснение отчётов. Рутину (расчёты) делают движки.

Ключ: ANTHROPIC_API_KEY в .env. Модель — из настроек платформы.
"""
import os
import json
import requests
from .. import config

API_URL = 'https://api.anthropic.com/v1/messages'


def ask(prompt: str, system: str = '', heavy: bool = False, max_tokens: int = 1500) -> str:
    key = os.getenv('ANTHROPIC_API_KEY', '')
    if not key:
        return '[ai] ANTHROPIC_API_KEY не задан — пропускаю ИИ-шаг'
    ps = config.platform_settings()
    model = ps.get('ai_model_heavy' if heavy else 'ai_model', 'claude-sonnet-5-5')
    body = {
        'model': model,
        'max_tokens': max_tokens,
        'messages': [{'role': 'user', 'content': prompt}],
    }
    if system:
        body['system'] = system
    try:
        r = requests.post(API_URL, timeout=60,
                          headers={'x-api-key': key, 'anthropic-version': '2023-06-01',
                                   'content-type': 'application/json'},
                          data=json.dumps(body))
        r.raise_for_status()
        data = r.json()
        return ''.join(b.get('text', '') for b in data.get('content', []))
    except requests.exceptions.RequestException as e:
        return f'[ai] ошибка вызова: {e}'


# Методичка проекта — встраивается как system-подсказка в ИИ-задачи (SEO и т.п.)
SEO_SYSTEM = (
    'Ты SEO-аналитик для карточек Wildberries/Ozon. Правила: бренд НЕ в SEO-название; '
    'название ≤60 символов с главным товарным ключом из топа; описание 900-1400 знаков, '
    'кластеры встроены естественно; не брать нерелевантные ингредиенты/свойства; '
    'составы косметики писать по-русски. Отвечай кратко и по делу.'
)
