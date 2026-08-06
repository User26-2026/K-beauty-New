# -*- coding: utf-8 -*-
import json, openpyxl, re, os
SP=os.path.dirname(os.path.abspath(__file__))
recs={str(r['Артикул WB']).strip():r for r in json.load(open(os.path.join(SP,'maskB_recs.json')))}
order=['226114831','226114935','226115141','226115248','226115324','226115478','226117746']
PRODUCER_ALL='Barunson Co., Ltd.'  # OEM Petitfee (держатель — NS Retail Co., Ltd.)
SPOSOB='На очищенную кожу лица наложите гидрогелевую маску (верхнюю и нижнюю части), плотно прижмите и оставьте на 20–30 минут. Снимите маску и мягко вбейте остатки эссенции подушечками пальцев до впитывания; смывать не нужно. Используйте 2–3 раза в неделю.'
WEIGHT_DEFAULT=290  # г, вес упаковки 5-шт набора (по аналогам линейки)
CUR={
 '226114831':{'eff':['Увлажнение','Восстановление'],'ing':['Растительный экстракт'],'app':['Дневной уход','Ночной уход'],'feat':[]},
 '226114935':{'eff':['Антивозрастной уход','Увлажнение'],'ing':[],'app':['Дневной уход','Ночной уход'],'feat':['С золотом']},
 '226115141':{'eff':['Антивозрастной уход','Восстановление'],'ing':['Муцин улитки'],'app':['Дневной уход','Ночной уход'],'feat':['С золотом']},
 '226115248':{'eff':['Тонизирование','Увлажнение'],'ing':['Кофеин','Растительный экстракт'],'app':['Дневной уход','Ночной уход'],'feat':[]},
 '226115324':{'eff':['Противоотечный','Увлажнение'],'ing':['Растительный экстракт'],'app':['Дневной уход','Ночной уход'],'feat':[]},
 '226115478':{'eff':['Антивозрастной уход','Увлажнение'],'ing':[],'app':['Дневной уход','Ночной уход'],'feat':['С золотом']},
 '226117746':{'eff':['Осветление','Увлажнение'],'ing':['Растительный экстракт'],'app':['Дневной уход','От пигментации'],'feat':[]},
}
BRAND_CASE={'petitfee':'Petitfee'}
def title(r):
    name=str(r.get('Наименование') or '').strip()
    brand=str(r.get('Бренд') or '').strip(); brand=BRAND_CASE.get(brand.lower(),brand)
    if brand and brand.lower() not in name.lower(): name=f"{name} {brand}"
    return name[:200]

tpl=os.path.join(SP,'cat6_a.xlsx')
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
    weight_g=int(round(float(w_kg)*1000)) if w_kg else WEIGHT_DEFAULT
    bc=r.get('Баркод'); bc=None if (bc in (0,'0',None) or str(bc).strip() in ('0','')) else bc
    setc(row,1,i); setc(row,2,r.get('Артикул продавца')); setc(row,3,title(r))
    setc(row,6,7); setc(row,7,'Нет'); setc(row,8,'Нет')
    if bc: setc(row,10,bc)
    setc(row,11,TNVED); setc(row,12,weight_g)
    setc(row,13,mm(r.get('Ширина упаковки'))); setc(row,14,mm(r.get('Высота упаковки'))); setc(row,15,mm(r.get('Длина упаковки')))
    setc(row,16,photos[0] if photos else None)
    if len(photos)>1: setc(row,17,';'.join(photos[1:]))
    setc(row,19,'Маска косметическая'); setc(row,20,r.get('Бренд')); setc(row,21,r.get('Артикул продавца'))
    setc(row,22,5)                              # единиц в товаре (набор 5 шт)
    setc(row,25,r.get('Состав')); setc(row,26,'Нет')
    setc(row,29,PRODUCER_ALL); setc(row,32,r.get('Описание'))
    setc(row,35,SPOSOB)
    ings=cur.get('ing') or []
    setc(row,36,';'.join(ings[:5]) if ings else None)
    setc(row,37,'Гидрогелевая')                 # вид маски
    setc(row,38,'Для всех типов кожи'); setc(row,39,';'.join(cur['eff'])); setc(row,40,';'.join(cur['app'])); setc(row,41,'Лицо')
    if cur.get('feat'): setc(row,42,';'.join(cur['feat']))
    setc(row,43,'Взрослая'); setc(row,44,'Женский'); setc(row,45,'Для любого возраста')
    setc(row,47,'Южная Корея'); setc(row,48,1); setc(row,50,1095)
    row+=1

os.makedirs('/home/user/K-beauty-New/outputs/ozon_facemask_cards',exist_ok=True)
outp='/home/user/K-beauty-New/outputs/ozon_facemask_cards/ozon_facemasks_2026-08-06.xlsx'
wb.save(outp); print('SAVED',outp,'| rows:',len(order))
