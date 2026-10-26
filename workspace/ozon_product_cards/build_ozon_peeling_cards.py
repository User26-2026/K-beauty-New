# -*- coding: utf-8 -*-
import json, openpyxl, re, os
SP=os.path.dirname(os.path.abspath(__file__))
recs={str(r['Артикул WB']).strip():r for r in json.load(open(os.path.join(SP,'peel_recs.json')))}
order=['551540924','551600265','551607905']

SOST_BRIGHT='Water; Dipropylene Glycol; Butylene Glycol; Isopentyldiol; Niacinamide; Propanediol; Glycerin; Hydroxyacetophenone; Caprylyl Glycol; Tranexamic Acid; Ethylhexylglycerin; Xanthan Gum; Carica Papaya (Papaya) Fruit Extract; Citric Acid; Prunus Mume Fruit Extract; Pyrus Malus (Apple) Fruit Extract; Vitis Vinifera (Grape) Fruit Extract; 1,2-Hexanediol; Sodium Phytate; Eriobotrya Japonica Leaf Extract; Sodium Hyaluronate; Mentha Viridis (Spearmint) Leaf Extract; Arbutin; Dipotassium Glycyrrhizate; Hydroxypropyltrimonium Hyaluronate; Tocopherol; Hydrolyzed Hyaluronic Acid; Sodium Acetylated Hyaluronate; Hyaluronic Acid; Sodium Hyaluronate Crosspolymer; Hydrolyzed Sodium Hyaluronate; Potassium Hyaluronate'
SOST_CICA='Water; Butylene Glycol; Alcohol Denat.; Camellia Japonica Flower Extract; 1,2-Hexanediol; Arginine; Salicylic Acid; Hydroxyacetophenone; C12-14 Pareth-12; Lactobionic Acid; Ethylhexylglycerin; Glycerin; Bifida Ferment Lysate; Centella Asiatica Extract; Disodium EDTA; Lactose; Melaleuca Alternifolia (Tea Tree) Leaf Oil; Houttuynia Cordata Extract; Portulaca Oleracea Extract; Perilla Ocymoides Seed Extract; Panthenol; Lonicera Caprifolium (Honeysuckle) Extract; Pelargonium Graveolens Extract; Uncaria Tomentosa Extract; Agrimonia Eupatoria Extract; Corydalis Turtschaninovii Root Extract; Cupressus Sempervirens Leaf Extract; Nepeta Cataria Extract; Madecassoside; Capryloyl Salicylic Acid'
SPOSOB_B='После очищения протрите лицо пропитанным пэдом по массажным линиям, избегая области вокруг глаз, уделяя внимание участкам с пигментацией и расширенными порами. Оставшуюся эссенцию вбейте подушечками пальцев; не смывайте. Днём завершайте уход солнцезащитным кремом.'
SPOSOB_C='После очищения протрите лицо пропитанным пэдом по массажным линиям, избегая области вокруг глаз; текстурной стороной мягко пройдитесь по проблемным зонам. Оставшуюся эссенцию вбейте подушечками пальцев; не смывайте. Из-за кислот в составе днём обязательно используйте солнцезащитный крем.'
CFG={
 '551540924':{'vol':100,'units':40,'sost':SOST_BRIGHT,'sposob':SPOSOB_B,
   'skin':['Для всех типов кожи','Для комбинированной кожи','Для жирной кожи'],
   'eff':['Осветление','Сужение пор','Очищение'],'ing':['Ниацинамид','Альфа-арбутин','Гиалуроновая кислота']},
 '551600265':{'vol':30,'units':10,'sost':SOST_CICA,'sposob':SPOSOB_C,
   'skin':['Для проблемной кожи','Для чувствительной кожи','Для комбинированной кожи'],
   'eff':['Очищение','Устранение высыпаний','Сужение пор'],'ing':['BHA-кислота','Центелла','Лизат молочнокислых бактерий']},
 '551607905':{'vol':170,'units':60,'sost':SOST_CICA,'sposob':SPOSOB_C,
   'skin':['Для проблемной кожи','Для чувствительной кожи','Для комбинированной кожи'],
   'eff':['Очищение','Устранение высыпаний','Сужение пор'],'ing':['BHA-кислота','Центелла','Лизат молочнокислых бактерий']},
}
PRODUCER='Cosmecca Korea Co., Ltd.'
def title(r):
    name=str(r.get('Наименование') or '').strip()
    brand='Celimax'
    if brand.lower() not in name.lower(): name=f"{name} {brand}"
    return name[:200]

tpl=os.path.join(SP,'cat8_tpl.xlsx')
wb=openpyxl.load_workbook(tpl); ws=wb['Шаблон']; val=wb['validation']
TNVED=[val.cell(r,11).value for r in range(1,val.max_row+1) if val.cell(r,11).value and str(val.cell(r,11).value).startswith('3304990000')][0]
def setc(row,col,v): ws.cell(row=row,column=col,value=v)
def mm(x):
    try: return int(round(float(x)*10))
    except: return None
row=5
for i,aw in enumerate(order,1):
    r=recs[aw]; c=CFG[aw]
    photos=[p for p in str(r.get('Фото') or '').split(';') if p.strip()]
    w_kg=r.get('Вес с упаковкой (кг)')
    bc=r.get('Баркод'); bc=None if (bc in (0,'0',None) or str(bc).strip() in ('0','')) else bc
    setc(row,1,i); setc(row,2,r.get('Артикул продавца')); setc(row,3,title(r))
    setc(row,6,7); setc(row,7,'Нет'); setc(row,8,'Нет')
    if bc: setc(row,10,bc)
    setc(row,11,TNVED); setc(row,12,int(round(float(w_kg)*1000)) if w_kg else None)
    setc(row,13,mm(r.get('Ширина упаковки'))); setc(row,14,mm(r.get('Высота упаковки'))); setc(row,15,mm(r.get('Длина упаковки')))
    setc(row,16,photos[0] if photos else None)
    if len(photos)>1: setc(row,17,';'.join(photos[1:]))
    setc(row,19,'Пилинг'); setc(row,20,'Celimax'); setc(row,21,r.get('Артикул продавца'))
    setc(row,22,c['units']); setc(row,23,c['vol']); setc(row,24,c['sost'])
    setc(row,25,'Нет'); setc(row,26,1095)
    setc(row,29,PRODUCER); setc(row,32,r.get('Описание'))
    setc(row,36,c['sposob']); setc(row,37,';'.join(c['ing'][:5]))
    setc(row,38,'Пилинг-пэды'); setc(row,39,';'.join(c['skin'])); setc(row,41,';'.join(c['eff'])); setc(row,42,'Лицо')
    setc(row,44,'Для любого возраста'); setc(row,45,'Южная Корея'); setc(row,46,1)
    row+=1

os.makedirs('/home/user/K-beauty-New/outputs/ozon_peeling_cards',exist_ok=True)
outp='/home/user/K-beauty-New/outputs/ozon_peeling_cards/ozon_peeling_2026-08-17.xlsx'
wb.save(outp); print('SAVED',outp,'| rows:',len(order))
