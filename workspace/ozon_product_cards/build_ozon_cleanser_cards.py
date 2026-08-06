# -*- coding: utf-8 -*-
import json, openpyxl, re, os
SP=os.path.dirname(os.path.abspath(__file__))
recs={str(r['Артикул WB']).strip():r for r in json.load(open(os.path.join(SP,'penki_recs.json')))}
order=['204258828','307073881','868575374']  # in-stock only

SOSTAV_EN={
 '868575374':'Purified Water; Sodium C14-16 Olefin Sulfonate; Lauryl Betaine; Glycerin; Sodium Chloride; Lactobacillus/Panax Ginseng Root Ferment Filtrate; Lactobacillus/Soymilk Ferment Filtrate; Butylene Glycol; Sodium Hyaluronate; Hydroxypropyltrimonium Hyaluronate; Hydrolyzed Hyaluronic Acid; Sodium Acetylated Hyaluronate; Hyaluronic Acid; Sodium Hyaluronate Crosspolymer; Potassium Hyaluronate; Hydrogenated Lecithin; Centella Asiatica Extract; Camellia Sinensis Leaf Extract; 1,2-Hexanediol; Rosemary Leaf Extract; Allantoin; Xylitol; Panax Ginseng Berry Extract; Saponaria Officinalis Leaf Extract; Guar Hydroxypropyltrimonium Chloride; Maltodextrin; Hexylene Glycol; Propylene Glycol Laurate; Ceramide NP; Bifida Ferment Lysate; Glucose; Sodium Bicarbonate; Phytosteryl/Octyldodecyl Lauroyl Glutamate; Papain; Xylitylglucoside; Anhydroxylitol; Fructan; Ethylhexylglycerin; Gluconolactone; Caprylic/Capric Triglyceride',
}
SPOSOB={
 '204258828':'Выдавите небольшое количество средства на ладонь, добавьте воды и вспеньте. Мягкими массирующими движениями нанесите пену на влажную кожу лица и шеи, избегая области вокруг глаз; в Т-зоне проработайте по массажным линиям. Тщательно смойте тёплой водой. Начинать лучше 1 раз в день, вечером.',
 '307073881':'Намочите лицо тёплой водой. Небольшое количество пенки вспеньте во влажных руках до плотной пены и нанесите на кожу, мягко массируя всё лицо, уделяя внимание Т-зоне и порам. Тщательно смойте тёплой водой. Использовать утром и вечером, после снятия макияжа.',
 '868575374':'Выдавите небольшое количество геля на влажные ладони и вспеньте, при необходимости добавив немного воды. Нанесите мягкую пену на влажную кожу лица круговыми массирующими движениями, избегая области вокруг глаз. Тщательно смойте тёплой водой. Подходит для ежедневного применения утром и вечером.',
}
PRODUCER={
 '204258828':'COSRX INC.',                    # держатель бренда (OEM не подтверждён)
 '307073881':'Amorepacific Corporation',      # изготовитель (на упаковке)
 '868575374':'Manyo Factory Co., Ltd.',       # держатель бренда (OEM не подтверждён)
}
# per-item multiselects
CUR={
 '204258828':{'vid':'пенка','skin':['Для проблемной кожи','Для жирной кожи','Для комбинированной кожи'],
   'eff':['Очищение','Устранение высыпаний','Сужение пор'],'ing':['BHA-кислота'],
   'feat':['Подходит для ежедневного применения']},
 '307073881':{'vid':'пенка','skin':['Для всех типов кожи','Для комбинированной кожи','Для жирной кожи'],
   'eff':['Очищение','Сужение пор'],'ing':[],
   'feat':['Подходит для ежедневного применения']},
 '868575374':{'vid':'гель','skin':['Для всех типов кожи','Для чувствительной кожи'],
   'eff':['Очищение','Увлажнение','Восстановление'],
   'ing':['Лизат молочнокислых бактерий','Гиалуроновая кислота','Центелла','Пребиотики'],
   'feat':['Подходит для ежедневного применения']},
}
BRAND_CASE={'cosrx':'COSRX','etude house':'Etude House','manyo':'Manyo'}

def title(r):
    name=str(r.get('Наименование') or '').strip()
    brand=str(r.get('Бренд') or '').strip()
    brand=BRAND_CASE.get(brand.lower(),brand)
    vol=r.get('Объем товара')
    m=re.search(r',?\s*\d+\s*(мл|ml)\s*$', name, re.I)
    base=name[:m.start()].rstrip(' ,') if m else name
    t=base
    if brand and brand.lower() not in base.lower(): t=f"{base} {brand}"
    try: t=f"{t}, {int(float(vol))} мл"
    except: pass
    return t[:200]

tpl=os.path.join(SP,'new_a.xlsx')
wb=openpyxl.load_workbook(tpl)
ws=wb['Шаблон']
val=wb['validation']
TNVED=[val.cell(r,11).value for r in range(1,val.max_row+1) if val.cell(r,11).value and str(val.cell(r,11).value).startswith('3304990000')][0]

def setc(row,col,v): ws.cell(row=row,column=col,value=v)
def mm(x):
    try: return int(round(float(x)*10))
    except: return None

row=5
for i,aw in enumerate(order,1):
    r=recs[aw]; cur=CUR[aw]
    photos=[p for p in str(r.get('Фото') or '').split(';') if p.strip()]
    sostav=r.get('Состав') or SOSTAV_EN.get(aw)
    w_kg=r.get('Вес с упаковкой (кг)')
    setc(row,1,i)
    setc(row,2,r.get('Артикул продавца'))
    setc(row,3,title(r))
    # 4 цена — пусто
    setc(row,6,7)
    setc(row,7,'Нет')                # рассрочка
    setc(row,8,'Нет')               # ускоренный сбор отзывов
    setc(row,10,r.get('Баркод'))
    setc(row,11,TNVED)
    setc(row,12,int(round(float(w_kg)*1000)) if w_kg else None)
    setc(row,13,mm(r.get('Ширина упаковки')))
    setc(row,14,mm(r.get('Высота упаковки')))
    setc(row,15,mm(r.get('Длина упаковки')))
    setc(row,16,photos[0] if photos else None)
    if len(photos)>1: setc(row,17,';'.join(photos[1:]))
    setc(row,19,'Средство для умывания')   # Тип*
    setc(row,20,r.get('Бренд'))
    setc(row,21,r.get('Артикул продавца'))  # модель
    setc(row,22,1)                          # единиц
    setc(row,23,r.get('Объем товара'))
    setc(row,24,sostav)
    setc(row,25,'Нет')                      # код маркировки
    setc(row,26,1095)                       # срок годности дней
    setc(row,29,PRODUCER.get(aw,r.get('Бренд')))  # производитель
    setc(row,32,r.get('Описание'))          # аннотация
    setc(row,36,SPOSOB.get(aw))             # способ применения
    ings=cur.get('ing') or []
    setc(row,37,';'.join(ings[:5]) if ings else None)  # бьюти-ингредиент
    setc(row,38,cur['vid'])                 # вид средства для ухода
    setc(row,39,';'.join(cur['skin']))      # тип кожи
    setc(row,41,';'.join(cur['eff']))       # эффект
    setc(row,42,'Лицо')                     # область использования
    if cur.get('feat'): setc(row,43,';'.join(cur['feat']))  # особенности состава
    setc(row,44,'Взрослая')                 # целевая аудитория
    setc(row,45,'Для любого возраста')      # возрастной диапазон
    setc(row,46,'Южная Корея')              # страна-изготовитель
    setc(row,47,1)                          # кол-во упаковок
    setc(row,49,'Не опасен')                # класс опасности
    row+=1

os.makedirs('/home/user/K-beauty-New/outputs/ozon_cleanser_cards',exist_ok=True)
outp='/home/user/K-beauty-New/outputs/ozon_cleanser_cards/ozon_cleansers_2026-08-06.xlsx'
wb.save(outp)
print('SAVED',outp,'| rows:',len(order))
