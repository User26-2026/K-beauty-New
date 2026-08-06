# -*- coding: utf-8 -*-
import json, openpyxl, re, os
SP=os.path.dirname(os.path.abspath(__file__))
recs={str(r['Артикул WB']).strip():r for r in json.load(open(os.path.join(SP,'base_recs.json')))}
order=['207466866']
PRODUCER={'207466866':'Daleaf Co., Ltd.'}  # изготовитель/владелец Heimish
BRAND_CASE={'heimish':'Heimish'}

def title(r):
    name=str(r.get('Наименование') or '').strip()
    brand=str(r.get('Бренд') or '').strip(); brand=BRAND_CASE.get(brand.lower(),brand)
    vol=r.get('Объем товара')
    m=re.search(r',?\s*\d+\s*(мл|ml)\s*$', name, re.I)
    base=name[:m.start()].rstrip(' ,') if m else name
    t=base
    if brand and brand.lower() not in base.lower(): t=f"{base} {brand}"
    try: t=f"{t}, {int(float(vol))} мл"
    except: pass
    return t[:200]

tpl=os.path.join(SP,'cat4_tpl.xlsx')
wb=openpyxl.load_workbook(tpl); ws=wb['Шаблон']; val=wb['validation']
TNVED=[val.cell(r,11).value for r in range(1,val.max_row+1) if val.cell(r,11).value and str(val.cell(r,11).value).startswith('3304990000')][0]
def setc(row,col,v): ws.cell(row=row,column=col,value=v)
def mm(x):
    try: return int(round(float(x)*10))
    except: return None

row=5
for i,aw in enumerate(order,1):
    r=recs[aw]
    photos=[p for p in str(r.get('Фото') or '').split(';') if p.strip()]
    w_kg=r.get('Вес с упаковкой (кг)')
    setc(row,1,i); setc(row,2,r.get('Артикул продавца')); setc(row,3,title(r))
    setc(row,6,7); setc(row,7,'Нет'); setc(row,8,'Нет')
    setc(row,10,r.get('Баркод')); setc(row,11,TNVED)
    setc(row,12,int(round(float(w_kg)*1000)) if w_kg else None)
    setc(row,13,mm(r.get('Ширина упаковки'))); setc(row,14,mm(r.get('Высота упаковки'))); setc(row,15,mm(r.get('Длина упаковки')))
    setc(row,16,photos[0] if photos else None)
    if len(photos)>1: setc(row,17,';'.join(photos[1:]))
    setc(row,19,'Основа под макияж'); setc(row,20,r.get('Бренд')); setc(row,21,r.get('Артикул продавца'))
    setc(row,24,r.get('Объем товара')); setc(row,25,'Нет'); setc(row,26,1095)
    setc(row,29,PRODUCER.get(aw) or r.get('Бренд')); setc(row,32,r.get('Описание'))
    setc(row,35,'SPF 50+')                   # SPF
    setc(row,36,r.get('Состав'))             # состав (уже EN)
    setc(row,39,'Взрослая')                  # целевая аудитория
    setc(row,40,1)                           # единиц в товаре
    setc(row,43,'Южная Корея'); setc(row,44,1)
    row+=1

os.makedirs('/home/user/K-beauty-New/outputs/ozon_makeupbase_cards',exist_ok=True)
outp='/home/user/K-beauty-New/outputs/ozon_makeupbase_cards/ozon_makeupbase_2026-08-06.xlsx'
wb.save(outp); print('SAVED',outp,'| rows:',len(order))
