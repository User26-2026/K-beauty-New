# -*- coding: utf-8 -*-
import json, openpyxl, re, os
SP=os.path.dirname(os.path.abspath(__file__))
recs={str(r['Артикул WB']).strip():r for r in json.load(open(os.path.join(SP,'scrub_recs.json')))}
order=['151313163','154160085']  # оба в наличии

INCI='Sodium Bicarbonate; Water; Glycerin; Silica; Hydrated Silica; Microcrystalline Cellulose; PEG-7 Glyceryl Cocoate; PEG-60 Hydrogenated Castor Oil; Sodium Methyl Cocoyl Taurate; Fragrance (Parfum); Cellulose Gum; Sodium Sulfite; Menthoxypropanediol; Limonene; Hexyl Cinnamal; Sodium Chloride; Disodium EDTA; Linalool; Tetrasodium EDTA'
SOSTAV_EN={'151313163':INCI,'154160085':INCI}  # состав из ВБ был на русском -> заменяем англ. INCI
SPOSOB={
 '151313163':'После умывания возьмите умеренное количество скраба и распределите по слегка влажной коже лица, избегая области вокруг глаз и губ. Мягко массируйте 1–2 минуты, затем тщательно смойте водой. Использовать 1–2 раза в неделю.',
 '154160085':'Вскройте саше и выдавите содержимое на ладонь. После умывания распределите по слегка влажной коже лица, избегая области вокруг глаз и губ, и мягко массируйте 1–2 минуты, затем смойте водой. Одно саше рассчитано на одно применение; использовать 1–2 раза в неделю.',
}
PRODUCER={'151313163':'Amorepacific Corporation','154160085':'Amorepacific Corporation'}
CUR={
 '151313163':{'skin':['Для всех типов кожи','Для комбинированной кожи','Для жирной кожи'],'eff':['Очищение','Сужение пор']},
 '154160085':{'skin':['Для всех типов кожи','Для комбинированной кожи','Для жирной кожи'],'eff':['Очищение','Сужение пор']},
}
BRAND_CASE={'etude house':'Etude House'}

def title(r):
    name=str(r.get('Наименование') or '').strip()
    brand=str(r.get('Бренд') or '').strip(); brand=BRAND_CASE.get(brand.lower(),brand)
    vol=r.get('Объем товара')
    m=re.search(r',?\s*\d+\s*(мл|ml|г|g)\s*$', name, re.I)
    base=name[:m.start()].rstrip(' ,') if m else name
    t=base
    if brand and brand.lower() not in base.lower(): t=f"{base} {brand}"
    try: t=f"{t}, {int(float(vol))} мл"
    except: pass
    return t[:200]

tpl=os.path.join(SP,'cat2_tpl.xlsx')
wb=openpyxl.load_workbook(tpl); ws=wb['Шаблон']; val=wb['validation']
TNVED=[val.cell(r,11).value for r in range(1,val.max_row+1) if val.cell(r,11).value and str(val.cell(r,11).value).startswith('3304990000')][0]
def setc(row,col,v): ws.cell(row=row,column=col,value=v)
def mm(x):
    try: return int(round(float(x)*10))
    except: return None

row=5
for i,aw in enumerate(order,1):
    r=recs[aw]; cur=CUR[aw]
    photos=[p for p in str(r.get('Фото') or '').split(';') if p.strip()]
    w_kg=r.get('Вес с упаковкой (кг)')
    setc(row,1,i); setc(row,2,r.get('Артикул продавца')); setc(row,3,title(r))
    setc(row,6,7); setc(row,7,'Нет'); setc(row,8,'Нет')
    setc(row,10,r.get('Баркод')); setc(row,11,TNVED)
    setc(row,12,int(round(float(w_kg)*1000)) if w_kg else None)
    setc(row,13,mm(r.get('Ширина упаковки'))); setc(row,14,mm(r.get('Высота упаковки'))); setc(row,15,mm(r.get('Длина упаковки')))
    setc(row,16,photos[0] if photos else None)
    if len(photos)>1: setc(row,17,';'.join(photos[1:]))
    setc(row,19,'Скраб'); setc(row,20,r.get('Бренд')); setc(row,21,r.get('Артикул продавца'))
    setc(row,23,r.get('Объем товара')); setc(row,24,SOSTAV_EN.get(aw) or r.get('Состав'))
    setc(row,25,'Нет'); setc(row,26,1095)
    setc(row,29,PRODUCER.get(aw,r.get('Бренд'))); setc(row,32,r.get('Описание'))
    setc(row,36,SPOSOB.get(aw))
    setc(row,38,';'.join(cur['skin'])); setc(row,40,';'.join(cur['eff']))
    setc(row,41,'Лицо'); setc(row,43,1)
    setc(row,45,'Южная Корея'); setc(row,46,1)
    row+=1

os.makedirs('/home/user/K-beauty-New/outputs/ozon_scrub_cards',exist_ok=True)
outp='/home/user/K-beauty-New/outputs/ozon_scrub_cards/ozon_scrubs_2026-08-06.xlsx'
wb.save(outp); print('SAVED',outp,'| rows:',len(order))
