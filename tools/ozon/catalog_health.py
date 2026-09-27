#!/usr/bin/env python3
"""Health-check каталога OZON из выгрузки товаров (CSV, разделитель ';').

Находит товары «Не продается» и группирует по причине блокировки с
рекомендуемым действием; отдельным листом — предупреждения. Колонки ищем по
имени заголовка (первая строка), заголовки OZON бывают многострочные.

Использование:
  python tools/ozon/catalog_health.py <каталог.csv> [out.xlsx]
"""
import sys, os, csv, openpyxl
from openpyxl.styles import Font, PatternFill

def gi(hdr, name):
    return next((i for i, h in enumerate(hdr) if name.lower() in h.lower()), None)

def classify(pr, err, art, name):
    t = (pr + ' ' + err).lower(); a = art.lower()
    if any(x in a for x in ('оаррп','рапор','оплгп')) or not name.strip():
        return 'Тестовая/мусорная карточка', 'Удалить карточку'
    if 'копи' in t or 'реплик' in t:
        return 'Копии/реплики запрещены', 'Документы на бренд + апелляция в поддержку OZON'
    if 'документы на бренд' in t:
        return 'Нет документов на бренд', 'Загрузить документы бренда (сертификат/декларация/письмо-разрешение)'
    if 'документ' in t:
        return 'Нет документов качества', 'Загрузить декларацию соответствия / СГР'
    if 'фото' in t or 'изображени' in t:
        return 'Нет/некорректное фото', 'Добавить главное фото; убрать надписи (кэшбэк/розыгрыш) с изображений'
    return 'Другое', pr or err

def main(path, out=None):
    rows = list(csv.reader(open(path, encoding='utf-8-sig'), delimiter=';'))
    hdr = [h.replace('\n', ' ').strip() for h in rows[0]]
    data = [r for r in rows[1:] if any(c.strip() for c in r)]
    I = {k: gi(hdr, k) for k in ('Артикул','Название','Статус','Категория','Причины','Ошибки','Предупреждения')}
    g = lambda r, k: r[I[k]].strip() if I[k] is not None and I[k] < len(r) else ''
    notsell = [r for r in data if g(r, 'Статус') == 'Не продается']
    out = out or f"outputs/ozon_catalog_health/ozon_blockers_{os.path.splitext(os.path.basename(path))[0]}.xlsx"
    os.makedirs(os.path.dirname(out), exist_ok=True)
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'Не продается'
    bold = Font(bold=True)
    ws.append(['Артикул','Название','Категория','Причина блокировки','Что делать'])
    for c in ws[1]: c.font = bold; c.fill = PatternFill('solid', fgColor='FFC7CE')
    from collections import Counter
    cnt = Counter()
    for r in sorted(notsell, key=lambda x: classify(g(x,'Причины'), g(x,'Ошибки'), g(x,'Артикул'), g(x,'Название'))[0]):
        pr, act = classify(g(r,'Причины'), g(r,'Ошибки'), g(r,'Артикул'), g(r,'Название'))
        cnt[pr] += 1
        ws.append([g(r,'Артикул'), g(r,'Название'), g(r,'Категория'), pr, act])
    ws2 = wb.create_sheet('Предупреждения'); ws2.append(['Артикул','Название','Категория','Предупреждение'])
    for c in ws2[1]: c.font = bold; c.fill = PatternFill('solid', fgColor='FFF2CC')
    for r in data:
        if g(r, 'Предупреждения'): ws2.append([g(r,'Артикул'), g(r,'Название'), g(r,'Категория'), g(r,'Предупреждения')])
    wb.save(out)
    from collections import Counter as C2
    st = C2(g(r, 'Статус') for r in data)
    print(f"Товаров: {len(data)} | статусы: {dict(st)} | не продаётся по причинам: {dict(cnt)} | -> {out}")

if __name__ == '__main__':
    if len(sys.argv) < 2: sys.exit(__doc__)
    main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
