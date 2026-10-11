"""Хранилище платформы. SQLite сейчас (ноль зависимостей), Postgres-ready потом.

Таблицы:
  snapshots — сырые срезы данных по клиенту/маркетплейсу/типу (JSON).
  actions   — журнал действий ИИ (что предложено/исполнено, с результатом).
Данные изолированы по client_id.
"""
import sqlite3
import json
import datetime as dt
from pathlib import Path
from .. import config

DB_PATH = config.DATA_DIR / 'platform.db'


def conn():
    c = sqlite3.connect(DB_PATH)
    c.row_factory = sqlite3.Row
    return c


def init():
    with conn() as c:
        c.executescript('''
        CREATE TABLE IF NOT EXISTS snapshots(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            client_id TEXT, marketplace TEXT, kind TEXT,
            ts TEXT, payload TEXT
        );
        CREATE INDEX IF NOT EXISTS ix_snap ON snapshots(client_id, kind, ts);
        CREATE TABLE IF NOT EXISTS actions(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            client_id TEXT, task TEXT, marketplace TEXT,
            sku TEXT, action TEXT, params TEXT,
            mode TEXT,              -- auto | suggest
            status TEXT,            -- proposed | applied | skipped | failed | awaiting_approval
            result TEXT, ts TEXT
        );
        CREATE INDEX IF NOT EXISTS ix_act ON actions(client_id, task, ts);
        ''')


def save_snapshot(client_id, marketplace, kind, payload):
    with conn() as c:
        c.execute('INSERT INTO snapshots(client_id,marketplace,kind,ts,payload) VALUES(?,?,?,?,?)',
                  (client_id, marketplace, kind, dt.datetime.utcnow().isoformat(),
                   json.dumps(payload, ensure_ascii=False)))


def log_action(client_id, task, marketplace, sku, action, params, mode, status, result=''):
    with conn() as c:
        c.execute('''INSERT INTO actions(client_id,task,marketplace,sku,action,params,mode,status,result,ts)
                     VALUES(?,?,?,?,?,?,?,?,?,?)''',
                  (client_id, task, marketplace, str(sku), action,
                   json.dumps(params, ensure_ascii=False), mode, status, result,
                   dt.datetime.utcnow().isoformat()))


def recent_actions(client_id, limit=50):
    with conn() as c:
        return [dict(r) for r in c.execute(
            'SELECT * FROM actions WHERE client_id=? ORDER BY id DESC LIMIT ?',
            (client_id, limit))]


init()
