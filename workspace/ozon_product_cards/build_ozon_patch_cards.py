# -*- coding: utf-8 -*-
import json, openpyxl, re, os
SP=os.path.dirname(os.path.abspath(__file__))
recs={str(r['Артикул WB']).strip():r for r in json.load(open(os.path.join(SP,'patch_recs.json')))}
order=['151175877','151175880','179611207','179611231','179611278','179611304','179611318','180711041','180711104','180987509','180988755','180992464']

# Составы EN для позиций, где в ВБ русский/нет (заполняется ресёрчем)
SOSTAV_EN={'151175880':'Water; Glycerin; Calcium Chloride; Butylene Glycol; Ceratonia Siliqua Gum; Xanthan Gum; Ethyl Hexanediol; PEG-60 Hydrogenated Castor Oil; Carrageenan; Glucomannan; Dextrin; Betaine; Allantoin; Scutellaria Baicalensis Root Extract; Citrus Grandis (Grapefruit) Seed Extract; Synthetic Fluorphlogopite; Acacia Senegal Gum; 1,2-Hexanediol; Alcohol; Chamaecyparis Obtusa Water; Camellia Sinensis Leaf Extract; Houttuynia Cordata Extract; Artemisia Princeps Leaf Extract; Citrus Junos Fruit Extract; Caprylyl Glycol; Titanium Dioxide; Disodium EDTA; Bambusa Textilis Stem Extract; Pinus Palustris Leaf Extract; Peat Water; Paraffinum Liquidum; Hydroxypropyl Methylcellulose; Citric Acid; Sodium Hyaluronate; Cucumis Sativus (Cucumber) Fruit Extract; Aloe Barbadensis Leaf Extract; Potassium Sorbate; Ethylhexylglycerin; Rosa Centifolia Flower Water; Pearl Extract; Gold; Parfum', '179611207':'Water; Glycerin; Calcium Chloride; Butylene Glycol; Ceratonia Siliqua Gum; Xanthan Gum; Chondrus Crispus (Carrageenan); Snail Secretion Filtrate; Hydroxyethylcellulose; PEG-60 Hydrogenated Castor Oil; Alcohol; Synthetic Fluorphlogopite; Tin Oxide; Iron Oxide; Titanium Dioxide; Gold; Disodium EDTA; Scutellaria Baicalensis Root Extract; Camellia Sinensis Leaf Extract; Houttuynia Cordata Extract; Artemisia Vulgaris Extract; Citrus Junos Fruit Extract; 1,2-Hexanediol; Caprylyl Glycol; Ethyl Hexanediol; Citrus Grandis (Grapefruit) Seed Extract; Bambusa Textilis Stem Extract; Pinus Palustris Leaf Extract; Chlorphenesin; Phenoxyethanol; Methylparaben; Fragrance', '180987509':'Water; Butylene Glycol; Glycerin; Polyglycerin-3; Niacinamide; Methyl Gluceth-20; Camellia Sinensis Leaf Extract; Hippophae Rhamnoides Extract; Hamamelis Virginiana (Witch Hazel) Extract; Rose Extract; Boerhavia Diffusa Root Extract; Portulaca Oleracea Extract; Brassica Oleracea Gongylodes (Kohlrabi) Extract; Beta Vulgaris (Beet) Root Extract; Raphanus Sativus (Radish) Root Extract; Brassica Oleracea Capitata (Cabbage) Leaf Extract; Allium Cepa (Onion) Root Extract; Trehalose; Betaine; Allantoin; Erythritol; Sodium Hyaluronate; Panthenol; Pantolactone; Dipotassium Glycyrrhizate; Dipropylene Glycol; Hydroxyacetophenone; Caprylyl Glycol; Glycereth-25 PCA Isostearate; Cyanocobalamin; Carbomer; Arginine; 1,2-Hexanediol; Disodium EDTA; Ethylhexylglycerin; Fragrance', '180992464':'Water; Glycerin; Dipropylene Glycol; Niacinamide; 1,2-Hexanediol; Ceratonia Siliqua (Carob) Gum; Chondrus Crispus Powder; Chondrus Crispus Extract; Panthenol; Cellulose Gum; Butylene Glycol; Paeonia Suffruticosa Root Extract; Centella Asiatica Extract; Potassium Chloride; Sodium Polyacrylate; Algin; Hydrogenated Polydecene; Chamomilla Recutita (Matricaria) Flower Extract; Glyceryl Caprylate; Allantoin; Xanthan Gum; Polysorbate 20; Sucrose; Adenosine; Ethylhexylglycerin; Trideceth-6; Silica; Dextrin; Pantolactone; Opal Powder; Hydrolyzed Extensin; Pearl Powder; Diamond Powder; Hydrolyzed Conchiolin Protein; Glycine Soja (Soybean) Protein; Hydrolyzed Lupine Protein; Hydrolyzed Sweet Almond Protein; Hydrolyzed Vegetable Protein; Hydrolyzed Rice Protein; Hydrolyzed Corn Protein; Hydrolyzed Soy Protein; Disodium EDTA; Fragrance/Parfum; Titanium Dioxide (CI 77891)', '180711041':'Water; Glycerin; Calcium Chloride; Butylene Glycol; Ceratonia Siliqua Gum; Xanthan Gum; Chondrus Crispus (Carrageenan); Ethyl Hexanediol; Citrus Grandis (Grapefruit) Seed Extract; Bambusa Textilis Stem Extract; Pinus Palustris Leaf Extract; PEG-60 Hydrogenated Castor Oil; Scutellaria Baicalensis Root Extract; Camellia Sinensis Leaf Extract; Houttuynia Cordata Extract; Artemisia Vulgaris Extract; Citrus Junos Fruit Extract; 1,2-Hexanediol; Caprylyl Glycol; Phenoxyethanol; Methylparaben; Chlorphenesin; Fragrance; Synthetic Fluorphlogopite; Titanium Dioxide; Iron Oxides; Gold; Alcohol; Aloe Barbadensis Leaf Extract; Panax Ginseng Root Extract; Acacia Seyal Gum Extract; Caviar Extract; Sodium Hyaluronate; Betaine; sh-Oligopeptide-1; Caffeoyl Tripeptide-1; Lavandula Angustifolia (Lavender) Flower Extract; Monarda Didyma Leaf Extract; Mentha Piperita (Peppermint) Leaf Extract; Freesia Alba Flower Extract; Chamomilla Recutita Flower/Leaf Extract; Rosmarinus Officinalis (Rosemary) Leaf Extract'}
PRODUCER_ALL='Barunson Co., Ltd.'  # OEM-изготовитель Petitfee (держатель — NS Retail Co., Ltd.)

SPOSOB_HYDRO='На очищенную кожу вокруг глаз наложите патчи широкой стороной вниз, плотно прижмите и оставьте на 20–30 минут. Снимите патчи и мягко вбейте остатки эссенции подушечками пальцев до впитывания; смывать не нужно. Используйте 2–3 раза в неделю или по необходимости. Для охлаждающего эффекта можно хранить в холодильнике.'
SPOSOB_FABRIC='На очищенную кожу под глаза наложите тканевые патчи, плотно прижмите к коже и оставьте на 20–30 минут. Снимите патчи и мягко вбейте остатки эссенции подушечками пальцев до впитывания; смывать не нужно. Используйте 2–3 раза в неделю или по необходимости.'

CUR={
 '151175877':{'base':'Гидрогелевая','eff':['Противоотечный','Увлажнение'],'app':['Дневной уход','Ночной уход'],'ing':['Растительный экстракт'],'feat':[]},
 '151175880':{'base':'Гидрогелевая','eff':['Антивозрастной уход','Увлажнение'],'app':['Дневной уход','Ночной уход'],'ing':[],'feat':['С золотом']},
 '179611207':{'base':'Гидрогелевая','eff':['Антивозрастной уход','Восстановление'],'app':['Дневной уход','Ночной уход'],'ing':['Муцин улитки'],'feat':['С золотом']},
 '179611231':{'base':'Гидрогелевая','eff':['Антивозрастной уход','Противоотечный','Увлажнение'],'app':['Дневной уход','Ночной уход'],'ing':['Коллаген','Коэнзим Q10'],'feat':[]},
 '179611278':{'base':'Гидрогелевая','eff':['Увлажнение','Восстановление'],'app':['Дневной уход','Ночной уход','От покраснений'],'ing':['Растительный экстракт','Аллантоин'],'feat':[]},
 '179611304':{'base':'Гидрогелевая','eff':['Тонизирование','Увлажнение'],'app':['Дневной уход','Ночной уход'],'ing':['Кофеин','Растительный экстракт'],'feat':[]},
 '179611318':{'base':'Гидрогелевая','eff':['Осветление','Противоотечный','Увлажнение'],'app':['Дневной уход','От пигментации'],'ing':['Растительный экстракт'],'feat':[]},
 '180711041':{'base':'Гидрогелевая','eff':['Антивозрастной уход','Восстановление'],'app':['Дневной уход','Ночной уход'],'ing':['Пептиды'],'feat':['С золотом']},
 '180711104':{'base':'Гидрогелевая','eff':['Антивозрастной уход','Увлажнение'],'app':['Дневной уход','Ночной уход'],'ing':[],'feat':['С золотом']},
 '180987509':{'base':'Тканевая','eff':['Осветление','Увлажнение'],'app':['Дневной уход','От пигментации'],'ing':['Витамин С'],'feat':[]},
 '180988755':{'base':'Тканевая','eff':['Увлажнение','Антивозрастной уход'],'app':['Дневной уход','Ночной уход'],'ing':[],'feat':[]},
 '180992464':{'base':'Гидрогелевая','eff':['Противоотечный','Увлажнение'],'app':['Дневной уход','Ночной уход'],'ing':[],'feat':[]},
}
BRAND_CASE={'petitfee':'Petitfee'}

def title(r):
    name=str(r.get('Наименование') or '').strip()
    brand=str(r.get('Бренд') or '').strip(); brand=BRAND_CASE.get(brand.lower(),brand)
    m=re.search(r',?\s*\d+\s*(мл|ml|шт)\s*$', name, re.I)
    base=name[:m.start()].rstrip(' ,') if m else name
    t=base
    if brand and brand.lower() not in base.lower(): t=f"{base} {brand}"
    return t[:200]

tpl=os.path.join(SP,'cat5_tpl.xlsx')
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
    sostav = SOSTAV_EN.get(aw) if aw in SOSTAV_EN else r.get('Состав')
    if aw not in SOSTAV_EN: sostav=r.get('Состав')
    if not sostav: missing.append(('состав',aw))
    setc(row,1,i); setc(row,2,r.get('Артикул продавца')); setc(row,3,title(r))
    setc(row,6,7); setc(row,7,'Нет'); setc(row,8,'Нет')
    setc(row,10,r.get('Баркод')); setc(row,11,TNVED)
    setc(row,12,int(round(float(w_kg)*1000)) if w_kg else None)
    setc(row,13,mm(r.get('Ширина упаковки'))); setc(row,14,mm(r.get('Высота упаковки'))); setc(row,15,mm(r.get('Длина упаковки')))
    setc(row,16,photos[0] if photos else None)
    if len(photos)>1: setc(row,17,';'.join(photos[1:]))
    setc(row,19,'Патчи'); setc(row,20,r.get('Бренд')); setc(row,21,r.get('Артикул продавца'))
    setc(row,22,1); setc(row,25,sostav)                       # 22 единиц; 25 Состав
    setc(row,26,'Нет')                                        # код маркировки
    setc(row,29,PRODUCER_ALL or r.get('Бренд')); setc(row,32,r.get('Описание'))
    setc(row,35,SPOSOB_FABRIC if cur['base']=='Тканевая' else SPOSOB_HYDRO)
    ings=cur.get('ing') or []
    setc(row,36,';'.join(ings[:5]) if ings else None)
    setc(row,37,cur['base'])                                  # основа патчей
    setc(row,38,';'.join(cur['eff'])); setc(row,39,';'.join(cur['app'])); setc(row,40,'Область вокруг глаз')
    if cur.get('feat'): setc(row,41,';'.join(cur['feat']))
    setc(row,42,'Взрослая'); setc(row,43,'Женский'); setc(row,44,'Для любого возраста')
    setc(row,46,'Южная Корея'); setc(row,47,1); setc(row,49,1095)
    row+=1

os.makedirs('/home/user/K-beauty-New/outputs/ozon_patch_cards',exist_ok=True)
outp='/home/user/K-beauty-New/outputs/ozon_patch_cards/ozon_patches_2026-08-06.xlsx'
wb.save(outp); print('SAVED',outp,'| rows:',len(order))
print('MISSING состав:', missing)
