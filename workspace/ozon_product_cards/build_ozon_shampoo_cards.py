# -*- coding: utf-8 -*-
import json, openpyxl, re, os
SP=os.path.dirname(os.path.abspath(__file__))
recs={str(r['Артикул WB']).strip():r for r in json.load(open(os.path.join(SP,'shampoo_recs.json')))}
order=['207472742','207475898','207476309','214305722','608624295','608837066','608841921']

DEO_INCI='Water; Sodium Lauryl Sulfate; Sodium Laureth Sulfate; Glycerin; Sodium Chloride; Cocamide DEA; Cocamidopropyl Betaine; Butylene Glycol; Allium Sativum (Garlic) Bulb Extract; Lawsonia Inermis (Henna) Extract; Sodium Hyaluronate; Panthenol; Hydrolyzed Collagen; Thiamine HCL; Tocopherol; Styrene/VP Copolymer; Tocopheryl Acetate; Diaminopyrimidine Oxide; Copper Tripeptide-1; Glycol Distearate; Citric Acid; Polyquaternium-7; Dimethicone; Laureth-23; Laureth-3; Alcohol; Hydroxyethylcellulose; Methylparaben; Climbazole; Disodium EDTA; Fragrance'
SOSTAV_EN={'214305722':DEO_INCI,'608624295':DEO_INCI,'608837066':DEO_INCI,'608841921':DEO_INCI}
PRODUCER={
 '207472742':'Myungin Cosmetics Co., Ltd.','207475898':'Myungin Cosmetics Co., Ltd.','207476309':'Myungin Cosmetics Co., Ltd.',
 '214305722':'Greencos Co., Ltd.','608624295':'Greencos Co., Ltd.','608837066':'Greencos Co., Ltd.','608841921':'Greencos Co., Ltd.',
}
SPOSOB_SH='Нанесите шампунь на влажные волосы, вспеньте лёгкими массирующими движениями по коже головы и длине, затем тщательно смойте тёплой водой. При необходимости повторите.'
SPOSOB_SB='Нанесите шампунь на влажные волосы, вспеньте и помассируйте кожу головы, смойте тёплой водой. Затем нанесите бальзам на длину волос, оставьте на 1–2 минуты и смойте. При необходимости повторите нанесение шампуня.'
CUR={
 '207472742':{'hair':['Для всех типов волос','Поврежденные','Жирные'],'eff':['Лечение кожи головы','Восстановление','Очищение'],'ing':['Растительный экстракт'],'sb':False},
 '207475898':{'hair':['Для всех типов волос','Поврежденные','Сухие'],'eff':['Восстановление','Блеск','Мягкость'],'ing':['Коллаген'],'sb':False},
 '207476309':{'hair':['Для всех типов волос','Тонкие'],'eff':['Объем','Блеск','Мягкость'],'ing':['Масло'],'sb':False},
 '214305722':{'hair':['Для всех типов волос','Поврежденные','Жирные'],'eff':['Лечение кожи головы','Восстановление','Очищение'],'ing':['Растительный экстракт','Коллаген'],'sb':True},
 '608624295':{'hair':['Для всех типов волос','Поврежденные','Жирные'],'eff':['Лечение кожи головы','Восстановление','Очищение'],'ing':['Растительный экстракт','Коллаген'],'sb':False},
 '608837066':{'hair':['Для всех типов волос','Поврежденные','Жирные'],'eff':['Лечение кожи головы','Восстановление','Очищение'],'ing':['Растительный экстракт','Коллаген'],'sb':False},
 '608841921':{'hair':['Для всех типов волос','Поврежденные','Жирные'],'eff':['Лечение кожи головы','Восстановление','Очищение'],'ing':['Растительный экстракт','Коллаген'],'sb':False},
}
BRAND_CASE={'farmstay':'FARMSTAY','deoproce':'DEOPROCE'}
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

tpl=os.path.join(SP,'cat7_tpl.xlsx')
wb=openpyxl.load_workbook(tpl); ws=wb['Шаблон']; val=wb['validation']
TNVED=[val.cell(r,11).value for r in range(1,val.max_row+1) if val.cell(r,11).value and str(val.cell(r,11).value).startswith('3305100000')][0]
def setc(row,col,v): ws.cell(row=row,column=col,value=v)
def mm(x):
    try: return int(round(float(x)*10))
    except: return None

row=5
for i,aw in enumerate(order,1):
    r=recs[aw]; cur=CUR[aw]
    photos=[p for p in str(r.get('Фото') or '').split(';') if p.strip()]
    w_kg=r.get('Вес с упаковкой (кг)')
    bc=r.get('Баркод'); bc=None if (bc in (0,'0',None) or str(bc).strip() in ('0','')) else bc
    sostav = SOSTAV_EN.get(aw) or r.get('Состав')
    tnved_prod = str(r.get('ТНВЭД') or '').strip()
    setc(row,1,i); setc(row,2,r.get('Артикул продавца')); setc(row,3,title(r))
    setc(row,6,7); setc(row,7,'Нет'); setc(row,8,'Нет')
    if bc: setc(row,10,bc)
    setc(row,11, TNVED)
    setc(row,12,int(round(float(w_kg)*1000)) if w_kg else None)
    setc(row,13,mm(r.get('Ширина упаковки'))); setc(row,14,mm(r.get('Высота упаковки'))); setc(row,15,mm(r.get('Длина упаковки')))
    setc(row,16,photos[0] if photos else None)
    if len(photos)>1: setc(row,17,';'.join(photos[1:]))
    setc(row,19,'Шампунь для волос'); setc(row,20,r.get('Бренд')); setc(row,21,r.get('Артикул продавца'))
    setc(row,23,1); setc(row,24,r.get('Объем товара'))
    setc(row,25,'Нет'); setc(row,26,1095); setc(row,27,'Не опасен')
    setc(row,30,PRODUCER.get(aw) or r.get('Бренд')); setc(row,33,r.get('Описание'))
    setc(row,36,'Нет')                                   # проф средство
    setc(row,38,sostav)
    setc(row,39, SPOSOB_SB if cur['sb'] else SPOSOB_SH)
    ings=cur.get('ing') or []
    setc(row,40,';'.join(ings[:5]) if ings else None)
    setc(row,41,';'.join(cur['hair'])); setc(row,42,';'.join(cur['eff']))
    setc(row,44,'Взрослая'); setc(row,45,'Женский'); setc(row,47,'Флакон')
    setc(row,49,'Южная Корея'); setc(row,50,1)
    row+=1

os.makedirs('/home/user/K-beauty-New/outputs/ozon_shampoo_cards',exist_ok=True)
outp='/home/user/K-beauty-New/outputs/ozon_shampoo_cards/ozon_shampoo_2026-08-11.xlsx'
wb.save(outp); print('SAVED',outp,'| rows:',len(order))
