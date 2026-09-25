# -*- coding: utf-8 -*-
import json, openpyxl, re, os
SP=os.path.dirname(os.path.abspath(__file__))
recs={str(r['Артикул WB']).strip():r for r in json.load(open(os.path.join(SP,'tonik_recs.json')))}
order=['198143251','199427475','212757456','212759698','212760119','212883090','551518955','1123574113']

# Состав EN (для позиций, где в ВБ был русский/пусто). Manyo — уже EN в выгрузке (не переопределяем).
SOSTAV_EN={
 '199427475':'Water; Niacinamide; Propanediol; Pentylene Glycol; Melaleuca Alternifolia (Tea Tree) Leaf Extract; Portulaca Oleracea Extract; Centella Asiatica Extract; Ficus Carica (Fig) Fruit Extract; Caprylic/Capric Triglyceride; Gloiopeltis Furcata Extract; Aloe Barbadensis Leaf Extract; Anthemis Nobilis Flower Extract; Camellia Sinensis Leaf Extract; Glycyrrhiza Glabra (Licorice) Rhizome/Root; Hamamelis Virginiana (Witch Hazel) Extract; Sodium Citrate; Cinnamomum Zeylanicum Bark Extract; Pinus Pinaster Bark Extract; Melaleuca Alternifolia (Tea Tree) Leaf Oil; Ethylhexylglycerin; Citric Acid; Disodium EDTA; Hydrogenated Lecithin; Capryloyl Glycine; Hexylene Glycol; Octyldodeceth-16; Allantoin; Sarcosine; Ceramide NP; Camphor; Carthamus Tinctorius (Safflower) Seed Oil; 1,2-Hexanediol; Silica; Glycine Soja (Soybean) Sterols; Linoleic Acid; Phospholipids; 4-Terpineol; Methyl Methacrylate Crosspolymer; Salicylic Acid',
 '551518955':'Morinda Citrifolia Fruit Extract; Butylene Glycol; Water; 1,2-Hexanediol; Hydroxyethyl Urea; Glycine; Serine; Glutamic Acid; Aspartic Acid; Leucine; Alanine; Lysine; Arginine; Tyrosine; Phenylalanine; Proline; Threonine; Valine; Isoleucine; Histidine; Cysteine; Methionine; Morinda Citrifolia Seed Oil; Rosmarinus Officinalis (Rosemary) Leaf Oil; Melia Azadirachta Leaf Extract; Melia Azadirachta Flower Extract; Theobroma Cacao (Cocoa) Extract; Betaine; Allantoin; Dextrin; Caprylic/Capric Triglyceride; C12-13 Pareth-9; Disodium EDTA; Chlorphenesin',
 '1123574113':'Water; Butylene Glycol; Glycerin; Pentylene Glycol; Propanediol; Chondrus Crispus Extract; Saccharum Officinarum (Sugar Cane) Extract; Sea Water; 1,2-Hexanediol; Protease; Betaine; Panthenol; Ethylhexylglycerin; Allantoin; Xanthan Gum; Disodium EDTA',
 # Jigott ×4 — заполняется после ресёрча
 '212757456':'Water; Glycerin; Propylene Glycol; Alcohol Denat.; Niacinamide; 1,2-Hexanediol; Methylpropanediol; Sodium Hyaluronate; Polysorbate 80; Carbomer; Triethanolamine; Allantoin; Centella Asiatica Extract (2,000ppm); Camellia Sinensis Leaf Extract; Chamomilla Recutita (Matricaria) Flower Extract; Hamamelis Virginiana (Witch Hazel) Extract; Rosa Centifolia Flower Extract; Lactobacillus/Soybean Ferment Extract; Salix Alba (Willow) Bark Extract; Cinnamomum Cassia Bark Extract; Chamaecyparis Obtusa Leaf Extract; Origanum Vulgare Leaf Extract; Scutellaria Baicalensis Root Extract; Portulaca Oleracea Extract; Adenosine; Disodium EDTA; Fragrance (Parfum)', '212759698':'Water (Aqua); Glycerin; Propylene Glycol; Alcohol Denat.; Niacinamide; 1,2-Hexanediol; Methylpropanediol; Sodium Hyaluronate; Polysorbate 80; Carbomer; Triethanolamine; Allantoin; Aloe Barbadensis Leaf Extract (2,000ppm); Camellia Sinensis Leaf Extract; Chamomilla Recutita (Matricaria) Flower Extract; Centella Asiatica Extract; Hamamelis Virginiana (Witch Hazel) Extract; Rosa Centifolia Flower Extract; Lactobacillus/Soybean Ferment Extract; Salix Alba (Willow) Bark Extract; Cinnamomum Cassia Bark Extract; Chamaecyparis Obtusa Leaf Extract; Origanum Vulgare Leaf Extract; Scutellaria Baicalensis Root Extract; Portulaca Oleracea Extract; Adenosine; Disodium EDTA; Fragrance (Parfum)', '212760119':'Water (Aqua); Glycerin; Propylene Glycol; Alcohol Denat.; Niacinamide; 1,2-Hexanediol; Methylpropanediol; Sodium Hyaluronate; Polysorbate 80; Carbomer; Triethanolamine; Allantoin; Hydrolyzed Collagen (2,000ppm); Camellia Sinensis Leaf Extract; Chamomilla Recutita (Matricaria) Flower Extract; Centella Asiatica Extract; Hamamelis Virginiana (Witch Hazel) Extract; Rosa Centifolia Flower Extract; Lactobacillus/Soybean Ferment Extract; Salix Alba (Willow) Bark Extract; Cinnamomum Cassia Bark Extract; Chamaecyparis Obtusa Leaf Extract; Origanum Vulgare Leaf Extract; Scutellaria Baicalensis Root Extract; Portulaca Oleracea Extract; Adenosine; Disodium EDTA; Fragrance (Parfum)', '212883090':'Water (Aqua); Glycerin; Propylene Glycol; Alcohol Denat.; Niacinamide; 1,2-Hexanediol; Methylpropanediol; Sodium Hyaluronate (10,000ppm); Polysorbate 80; Carbomer; Triethanolamine; Allantoin; Hydroxypropyltrimonium Hyaluronate; Sodium Acetylated Hyaluronate; Hydrolyzed Hyaluronic Acid; Hyaluronic Acid; Sodium Hyaluronate Crosspolymer; Hydrolyzed Sodium Hyaluronate; Potassium Hyaluronate; Chamomilla Recutita (Matricaria) Flower Extract; Centella Asiatica Extract; Hamamelis Virginiana (Witch Hazel) Extract; Rosa Centifolia Flower Extract; Camellia Sinensis Leaf Extract; Lactobacillus/Soybean Ferment Extract; Salix Alba (Willow) Bark Extract; Cinnamomum Cassia Bark Extract; Chamaecyparis Obtusa Leaf Extract; Origanum Vulgare Leaf Extract; Scutellaria Baicalensis Root Extract; Portulaca Oleracea Extract; Adenosine; Disodium EDTA; Fragrance (Parfum)',
}
SPOSOB={
 '198143251':'После очищения нанесите тонер-лосьон на ватный диск и мягко протрите лицо, либо распределите ладонями и вбейте похлопывающими движениями до впитывания. Используйте утром и вечером как первый увлажняющий шаг перед сывороткой и кремом.',
 '199427475':'После умывания нанесите небольшое количество тонера на ватный диск и мягко протрите лицо, уделяя внимание проблемным и жирным зонам, либо вбейте тонер ладонями до впитывания. Применяйте утром и вечером перед последующим уходом.',
 '551518955':'На очищенную кожу нанесите тонер ватным диском или ладонями и мягко вбейте до полного впитывания. Используйте утром и вечером как увлажняющий шаг перед сывороткой и кремом.',
 '1123574113':'На очищенную кожу нанесите достаточное количество тонера ладонями или ватным диском, мягко распределите похлопывающими движениями до впитывания. Используйте утром и вечером как первый увлажняющий шаг после умывания, перед сывороткой и кремом.',
 '212757456':'После умывания нанесите тонер на лицо ладонями или ватным диском, мягко распределяя по массажным линиям и избегая области вокруг глаз. Слегка вбейте подушечками пальцев до впитывания, затем нанесите эмульсию или крем.', '212759698':'На очищенную кожу нанесите тонер ладонями или ватным диском, распределяя по массажным линиям и избегая зоны вокруг глаз. Дайте впитаться, при желании повторите нанесение для усиленного увлажнения, затем нанесите эмульсию или крем.', '212760119':'После очищения нанесите тонер ладонями или ватным диском по массажным линиям, избегая области вокруг глаз, и вбейте подушечками пальцев до впитывания. Для антивозрастного ухода нанесите следом эмульсию, сыворотку и крем.', '212883090':'На очищенную кожу нанесите тонер ладонями или ватным диском, распределяя по массажным линиям и избегая зоны вокруг глаз. Слегка вбейте до впитывания; для глубокого увлажнения можно нанести в 2 слоя, затем сыворотку и крем.',
}
PRODUCER={
 '198143251':'Cosmax Co., Ltd.',            # OEM подтверждён
 '199427475':'Have & Be Co., Ltd.',         # держатель (OEM не подтверждён)
 '551518955':'Cosmecca Korea Co., Ltd.',    # OEM подтверждён
 '1123574113':'Cosmax Co., Ltd.',           # изготовитель подтверждён (Hwahae)
 '212757456':'Skinine Co., Ltd.','212759698':'Skinine Co., Ltd.','212760119':'Skinine Co., Ltd.','212883090':'Skinine Co., Ltd.',
}
CUR={
 '198143251':{'skin':['Для всех типов кожи','Для чувствительной кожи'],'eff':['Увлажнение','Восстановление','Тонизирование'],'app':['Дневной уход','Ночной уход'],'ing':['Лизат молочнокислых бактерий','Гиалуроновая кислота','Ниацинамид','Пребиотики']},
 '199427475':{'skin':['Для проблемной кожи','Для жирной кожи','Для комбинированной кожи'],'eff':['Очищение','Тонизирование','Устранение высыпаний'],'app':['Дневной уход','Ночной уход','От прыщей (анти-акне)','От покраснений'],'ing':['Ниацинамид','BHA-кислота','Центелла','Аллантоин','Растительный экстракт']},
 '551518955':{'skin':['Для всех типов кожи','Для сухой кожи'],'eff':['Увлажнение','Тонизирование','Питание'],'app':['Дневной уход','Ночной уход'],'ing':['Растительный экстракт','Аллантоин']},
 '1123574113':{'skin':['Для всех типов кожи','Для чувствительной кожи'],'eff':['Увлажнение','Тонизирование','Восстановление'],'app':['Дневной уход','Ночной уход'],'ing':['Пантенол','Аллантоин']},
 '212757456':{'skin':['Для всех типов кожи','Для чувствительной кожи'],'eff':['Увлажнение','Тонизирование','Восстановление'],'app':['Дневной уход','Ночной уход','От покраснений'],'ing':['Центелла']},
 '212759698':{'skin':['Для всех типов кожи','Для чувствительной кожи'],'eff':['Увлажнение','Тонизирование'],'app':['Дневной уход','Ночной уход','От покраснений'],'ing':['Растительный экстракт']},
 '212760119':{'skin':['Для всех типов кожи','Для зрелой кожи'],'eff':['Антивозрастной уход','Увлажнение','Тонизирование'],'app':['Дневной уход','Ночной уход'],'ing':['Коллаген']},
 '212883090':{'skin':['Для всех типов кожи','Для сухой кожи'],'eff':['Увлажнение','Тонизирование'],'app':['Дневной уход','Ночной уход'],'ing':['Гиалуроновая кислота']},
}
BRAND_CASE={'manyo':'Manyo','dr.jart+':'Dr.Jart+','jigott':'Jigott','celimax':'Celimax','round lab':'ROUND LAB'}

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

tpl=os.path.join(SP,'cat3_tpl.xlsx')
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
    sostav=SOSTAV_EN.get(aw) or r.get('Состав')
    if not sostav: missing.append(('состав',aw))
    if not SPOSOB.get(aw): missing.append(('способ',aw))
    setc(row,1,i); setc(row,2,r.get('Артикул продавца')); setc(row,3,title(r))
    setc(row,6,7); setc(row,7,'Нет'); setc(row,8,'Нет')
    setc(row,10,r.get('Баркод')); setc(row,11,TNVED)
    setc(row,12,int(round(float(w_kg)*1000)) if w_kg else None)
    setc(row,13,mm(r.get('Ширина упаковки'))); setc(row,14,mm(r.get('Высота упаковки'))); setc(row,15,mm(r.get('Длина упаковки')))
    setc(row,16,photos[0] if photos else None)
    if len(photos)>1: setc(row,17,';'.join(photos[1:]))
    setc(row,19,'Тоник для ухода за кожей'); setc(row,20,r.get('Бренд')); setc(row,21,r.get('Артикул продавца'))
    setc(row,22,1); setc(row,23,r.get('Объем товара')); setc(row,24,sostav)
    setc(row,25,'Нет'); setc(row,26,1095)
    setc(row,29,PRODUCER.get(aw) or r.get('Бренд')); setc(row,32,r.get('Описание'))
    setc(row,36,SPOSOB.get(aw))
    ings=cur.get('ing') or []
    setc(row,37,';'.join(ings[:5]) if ings else None)
    setc(row,38,';'.join(cur['skin'])); setc(row,40,';'.join(cur['eff'])); setc(row,41,';'.join(cur.get('app',[])))
    setc(row,42,'Лицо'); setc(row,44,'Для любого возраста'); setc(row,45,'Южная Корея'); setc(row,46,1)
    row+=1

os.makedirs('/home/user/K-beauty-New/outputs/ozon_toner_cards',exist_ok=True)
outp='/home/user/K-beauty-New/outputs/ozon_toner_cards/ozon_toners_2026-08-06.xlsx'
wb.save(outp); print('SAVED',outp,'| rows:',len(order))
print('MISSING (Jigott, ждём ресёрч):', missing)
