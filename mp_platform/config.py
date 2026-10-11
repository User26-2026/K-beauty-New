"""Загрузка реестра клиентов и глобальных настроек платформы.

Клиенты описаны в mp_platform/clients/clients.yaml (в .gitignore).
Токены берутся из переменных окружения по именам из реестра — в файле их нет.
"""
import os
from pathlib import Path
from dataclasses import dataclass, field

import yaml
from dotenv import load_dotenv

ROOT = Path(__file__).resolve().parent.parent
load_dotenv(ROOT / '.env')

CLIENTS_FILE = Path(__file__).parent / 'clients' / 'clients.yaml'
EXAMPLE_FILE = Path(__file__).parent / 'clients' / 'clients.example.yaml'
DATA_DIR = ROOT / 'data' / 'platform'
DATA_DIR.mkdir(parents=True, exist_ok=True)


@dataclass
class Client:
    id: str
    name: str
    active: bool
    marketplaces: dict
    automation: dict
    guardrails: dict
    notify: dict = field(default_factory=dict)

    def token(self, key_env: str):
        """Достаёт секрет из окружения по имени переменной из реестра."""
        return os.getenv(key_env, '')

    def data_dir(self) -> Path:
        d = DATA_DIR / self.id
        d.mkdir(parents=True, exist_ok=True)
        return d

    def mode(self, task: str) -> str:
        """auto | suggest | off для конкретной задачи."""
        return (self.automation or {}).get(task, 'off')


def _load():
    path = CLIENTS_FILE if CLIENTS_FILE.exists() else EXAMPLE_FILE
    with open(path, encoding='utf-8') as f:
        return yaml.safe_load(f)


def load_clients():
    data = _load()
    return [Client(**c) for c in data.get('clients', []) if c.get('active', True)]


def platform_settings():
    return _load().get('platform', {})
