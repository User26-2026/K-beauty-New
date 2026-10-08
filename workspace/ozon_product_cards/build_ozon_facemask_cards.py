# -*- coding: utf-8 -*-
import json, openpyxl, re, os
SP=os.path.dirname(os.path.abspath(__file__))
recs={}
for fn in ['maskB_recs.json','add1_recs.json']:
    for r in json.load(open(os.path.join(SP,fn))):
        recs[str(r['Артикул WB']).strip()]=r

# 7 одиночных гидрогелевых Petitfee + 3 набора Petitfee (5/6/7 шт) + 3 тканевые (JMSolution ×2, Jigott)
order=['226114831','226114935','226115141','226115248','226115324','226115478','226117746',
       '664171648','664171647','255574854',
       '214964091','214977901','255574453']
NABOR={'664171648','664171647','255574854'}  # ассорти-наборы: состав репрезентативный

SPOSOB_HYDRO='На очищенную кожу лица наложите гидрогелевую маску (верхнюю и нижнюю части), плотно прижмите и оставьте на 20–30 минут. Снимите маску и мягко вбейте остатки эссенции подушечками пальцев до впитывания; смывать не нужно. Используйте 2–3 раза в неделю.'
SPOSOB_SHEET='После очищения и тонизирования достаньте маску, разверните и наложите на лицо, совместив вырезы для глаз, носа и губ. Разгладьте и оставьте на 15–20 минут, затем снимите и мягко вбейте остатки эссенции подушечками пальцев до впитывания; смывать не нужно.'
WEIGHT_DEFAULT=290

# Составы EN для тканевых (репрезентативные, ассорти) — заполнит ресёрч
SOSTAV_EN={'214964091':'Water; Methylpropanediol; Glycerin; Glycereth-26; PEG/PPG-17/6 Copolymer; Sodium Hyaluronate; Hydrolyzed Hyaluronic Acid; Hyaluronic Acid; Collagen Extract; Allantoin; Trehalose; Betaine; Ethylhexylglycerin; Xanthan Gum; Hydroxyethylcellulose; Dipropylene Glycol; Butylene Glycol; Disodium Phosphate; Pentylene Glycol; Sodium Phosphate; Disodium EDTA; Glycine; Serine; Glutamic Acid; Aspartic Acid; Leucine; Alanine; Lysine; Proline; Threonine; Valine; 1,2-Hexanediol; PEG-60 Hydrogenated Castor Oil; Rosa Damascena Flower Water; Rosa Damascena Flower Oil; Phenoxyethanol', '214977901':'Water; Glycerin; Dipropylene Glycol; Sea Water; Butylene Glycol; 1,2-Hexanediol; Hydrolyzed Collagen; Hydrolyzed Conchiolin Protein; Pearl Extract; Undaria Pinnatifida Extract; Codium Fragile Extract; Enteromorpha Compressa Extract; Laminaria Japonica Extract; Salicornia Herbacea Extract; Sodium Hyaluronate; Hydrolyzed Hyaluronic Acid; Sodium Acetylated Hyaluronate; Olive Oil; Tocopheryl Acetate; Polysorbate 20; Ethylhexylglycerin; Carbomer; Tromethamine; Disodium EDTA; Phenoxyethanol; Fragrance', '255574453':'Water; Glycerin; Butylene Glycol; Betaine; Carbomer; Triethanolamine; Disodium EDTA; Hydroxyethylcellulose; Polysorbate 20; Sodium Hyaluronate; Tocopheryl Acetate; Allantoin; Panthenol; Portulaca Oleracea Extract; 1,2-Hexanediol; Phenoxyethanol; Fragrance'}
PRODUCER={  # изготовители
 '226114831':'Barunson Co., Ltd.','226114935':'Barunson Co., Ltd.','226115141':'Barunson Co., Ltd.',
 '226115248':'Barunson Co., Ltd.','226115324':'Barunson Co., Ltd.','226115478':'Barunson Co., Ltd.','226117746':'Barunson Co., Ltd.',
 '664171648':'Barunson Co., Ltd.','664171647':'Barunson Co., Ltd.','255574854':'Barunson Co., Ltd.',
 '214964091':'GP Club Co., Ltd.','214977901':'GP Club Co., Ltd.','255574453':'Skinine Cosmetic Co., Ltd.',
}
CUR={
 '226114831':{'vid':'Гидрогелевая','units':5,'eff':['Увлажнение','Восстановление'],'ing':['Растительный экстракт'],'app':['Дневной уход','Ночной уход'],'feat':[]},
 '226114935':{'vid':'Гидрогелевая','units':5,'eff':['Антивозрастной уход','Увлажнение'],'ing':[],'app':['Дневной уход','Ночной уход'],'feat':['С золотом']},
 '226115141':{'vid':'Гидрогелевая','units':5,'eff':['Антивозрастной уход','Восстановление'],'ing':['Муцин улитки'],'app':['Дневной уход','Ночной уход'],'feat':['С золотом']},
 '226115248':{'vid':'Гидрогелевая','units':5,'eff':['Тонизирование','Увлажнение'],'ing':['Кофеин','Растительный экстракт'],'app':['Дневной уход','Ночной уход'],'feat':[]},
 '226115324':{'vid':'Гидрогелевая','units':5,'eff':['Противоотечный','Увлажнение'],'ing':['Растительный экстракт'],'app':['Дневной уход','Ночной уход'],'feat':[]},
 '226115478':{'vid':'Гидрогелевая','units':5,'eff':['Антивозрастной уход','Увлажнение'],'ing':[],'app':['Дневной уход','Ночной уход'],'feat':['С золотом']},
 '226117746':{'vid':'Гидрогелевая','units':5,'eff':['Осветление','Увлажнение'],'ing':['Растительный экстракт'],'app':['Дневной уход','От пигментации'],'feat':[]},
 '664171648':{'vid':'Гидрогелевая','units':5,'eff':['Увлажнение','Восстановление'],'ing':[],'app':['Дневной уход','Ночной уход'],'feat':[]},
 '664171647':{'vid':'Гидрогелевая','units':6,'eff':['Увлажнение','Восстановление'],'ing':[],'app':['Дневной уход','Ночной уход'],'feat':[]},
 '255574854':{'vid':'Гидрогелевая','units':7,'eff':['Увлажнение','Восстановление'],'ing':[],'app':['Дневной уход','Ночной уход'],'feat':[]},
 '214964091':{'vid':'Тканевая','units':10,'eff':['Увлажнение','Питание'],'ing':['Гиалуроновая кислота','Коллаген'],'app':['Дневной уход','Ночной уход'],'feat':[]},
 '214977901':{'vid':'Тканевая','units':10,'eff':['Увлажнение'],'ing':['Коллаген','Гиалуроновая кислота'],'app':['Дневной уход','Ночной уход'],'feat':[]},
 '255574453':{'vid':'Тканевая','units':15,'eff':['Увлажнение'],'ing':['Гиалуроновая кислота'],'app':['Дневной уход','Ночной уход'],'feat':[]},
}
BRAND_CASE={'petitfee':'Petitfee','jmsolution':'JMSolution','jigott':'Jigott'}
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

missing=[]
row=5
for i,aw in enumerate(order,1):
    r=recs[aw]; cur=CUR[aw]
    photos=[p for p in str(r.get('Фото') or '').split(';') if p.strip()]
    w_kg=r.get('Вес с упаковкой (кг)')
    weight_g=int(round(float(w_kg)*1000)) if w_kg else (350 if aw in NABOR else WEIGHT_DEFAULT)
    bc=r.get('Баркод'); bc=None if (bc in (0,'0',None) or str(bc).strip() in ('0','')) else bc
    if aw in NABOR:
        sostav = recs['226114831'].get('Состав')      # репрезентативный состав (ассорти-набор гидрогелевых масок)
    elif aw in SOSTAV_EN:
        sostav = SOSTAV_EN[aw]
    else:
        sostav = r.get('Состав')
    if not sostav: missing.append(aw)
    sposob = SPOSOB_SHEET if cur['vid']=='Тканевая' else SPOSOB_HYDRO
    setc(row,1,i); setc(row,2,r.get('Артикул продавца')); setc(row,3,title(r))
    setc(row,6,7); setc(row,7,'Нет'); setc(row,8,'Нет')
    if bc: setc(row,10,bc)
    setc(row,11,TNVED); setc(row,12,weight_g)
    setc(row,13,mm(r.get('Ширина упаковки'))); setc(row,14,mm(r.get('Высота упаковки'))); setc(row,15,mm(r.get('Длина упаковки')))
    setc(row,16,photos[0] if photos else None)
    if len(photos)>1: setc(row,17,';'.join(photos[1:]))
    setc(row,19,'Маска косметическая'); setc(row,20,r.get('Бренд')); setc(row,21,r.get('Артикул продавца'))
    setc(row,22,cur['units'])
    setc(row,25,sostav); setc(row,26,'Нет')
    setc(row,29,PRODUCER.get(aw) or r.get('Бренд')); setc(row,32,r.get('Описание'))
    setc(row,35,sposob)
    ings=cur.get('ing') or []
    setc(row,36,';'.join(ings[:5]) if ings else None)
    setc(row,37,cur['vid'])
    setc(row,38,'Для всех типов кожи'); setc(row,39,';'.join(cur['eff'])); setc(row,40,';'.join(cur['app'])); setc(row,41,'Лицо')
    if cur.get('feat'): setc(row,42,';'.join(cur['feat']))
    setc(row,43,'Взрослая'); setc(row,44,'Женский'); setc(row,45,'Для любого возраста')
    setc(row,47,'Южная Корея'); setc(row,48,1); setc(row,50,1095)
    row+=1

os.makedirs('/home/user/K-beauty-New/outputs/ozon_facemask_cards',exist_ok=True)
outp='/home/user/K-beauty-New/outputs/ozon_facemask_cards/ozon_facemasks_2026-08-06.xlsx'
wb.save(outp); print('SAVED',outp,'| rows:',len(order))
print('MISSING состав:', missing)
