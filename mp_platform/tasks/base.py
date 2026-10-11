"""Базовый класс задачи. Каждая задача: собрать → проанализировать → (авто) исполнить.

mode клиента: auto (исполняем в рамках guardrails) | suggest (только предложения) | off.
Все действия пишутся в журнал (db.actions) с аудитом.
"""
from ..core import db, notify


class Task:
    name = 'base'

    def __init__(self, client):
        self.client = client
        self.mode = client.mode(self.name)

    def run(self):
        if self.mode == 'off':
            return []
        data = self.collect()
        proposals = self.analyze(data)      # список предложенных действий
        applied = []
        for p in proposals:
            if self.mode == 'auto' and p.get('allowed'):
                status, result = self.execute(p)
            else:
                status, result = ('awaiting_approval' if self.mode == 'suggest' else 'skipped'), ''
            db.log_action(self.client.id, self.name, p.get('marketplace', ''), p.get('sku', ''),
                          p.get('action', ''), p.get('params', {}), self.mode, status, result)
            applied.append({**p, 'status': status})
        self.report(applied)
        return applied

    # Переопределяются в наследниках
    def collect(self):
        return {}

    def analyze(self, data):
        return []

    def execute(self, proposal):
        return 'skipped', 'не реализовано'

    def report(self, applied):
        if not applied:
            return
        tochange = [a for a in applied if a['status'] in ('applied', 'awaiting_approval')]
        if tochange:
            lines = [f'<b>{self.name}</b> [{self.client.name}] — действий: {len(tochange)}']
            for a in tochange[:15]:
                lines.append(f"{a.get('sku','')}: {a.get('action')} {a.get('why','')} [{a['status']}]")
            notify.send('\n'.join(lines))
