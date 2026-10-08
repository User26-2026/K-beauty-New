# -*- coding: utf-8 -*-
import json, openpyxl, re, os

SP=os.path.dirname(os.path.abspath(__file__))
data=json.load(open(os.path.join(SP,'serums_instock.json')))
data={str(r['Артикул WB']):r for r in data}

def srok_days(s):
    if not s: return 1080
    s=str(s).lower()
    if '3 года' in s or '3 год' in s: return 1080
    m=re.search(r'(\d+)\s*мес', s)
    if m: return int(m.group(1))*30
    return 1080

# ingredient keyword -> OZON value
ING={
 'hyaluron':'Гиалуроновая кислота','гиалурон':'Гиалуроновая кислота',
 'niacinamide':'Ниацинамид','ниацинамид':'Ниацинамид',
 'collagen':'Коллаген','коллаген':'Коллаген',
 'snail':'Муцин улитки','муцин':'Муцин улитки','escargot':'Муцин улитки','улит':'Муцин улитки',
 'retinol':'Ретинол','ретинол':'Ретинол',
 'pdrn':'PDRN',
 'centella':'Центелла','центелл':'Центелла','cica':'Центелла','цика':'Центелла',
 'peptide':'Пептиды','пептид':'Пептиды','argirel':'Пептиды','hexapeptide':'Пептиды','adenosine':'Пептиды','аденозин':'Пептиды',
 'galactom':'Галактомисис','галактом':'Галактомисис',
 'bifida':'Лизат молочнокислых бактерий','бифид':'Лизат молочнокислых бактерий','lactobacillus':'Лизат молочнокислых бактерий','lactococcus':'Лизат молочнокислых бактерий',
 'vitamin c':'Витамин С','витамин с':'Витамин С','ascorb':'Витамин С','vita c':'Витамин С','vita-c':'Витамин С',
 'noni':'Растительный экстракт','нони':'Растительный экстракт',
 'aha':'AHA-кислота','bha':'BHA-кислота','lha':'LHA-кислота',
 'macadamia':'Масло','макадам':'Масло','glycerin':'Глицерин',
}

# per-item curated multiselects (skin, effect, application, features, area, type override, ingredients extra)
CUR={
 '168956391':{'eff':['Увлажнение','Антивозрастной уход'],'app':['Дневной уход','Ночной уход'],'skin':['Для всех типов кожи'],'ing':['Коллаген','Гиалуроновая кислота','Ниацинамид']},
 '198147185':{'eff':['Восстановление','Увлажнение'],'app':['Дневной уход','Ночной уход'],'skin':['Для всех типов кожи','Для чувствительной кожи'],'ing':['Лизат молочнокислых бактерий','Пептиды','Ниацинамид','Пребиотики']},
 '198156472':{'eff':['Осветление','Увлажнение'],'app':['Дневной уход','От пигментации'],'skin':['Для всех типов кожи'],'ing':['Витамин С','Галактомисис','Ниацинамид']},
 '206058262':{'eff':['Увлажнение','Восстановление'],'app':['Дневной уход','Ночной уход'],'skin':['Для всех типов кожи'],'ing':['Галактомисис']},
 '212885914':{'eff':['Увлажнение','Противоотечный'],'app':['Дневной уход','От темных кругов под глазами'],'skin':['Для всех типов кожи'],'area':'Область вокруг глаз','ing':['Коллаген','Гиалуроновая кислота']},
 '214311963':{'eff':['Увлажнение','Восстановление'],'app':['Дневной уход','От покраснений'],'skin':['Для всех типов кожи','Для чувствительной кожи'],'ing':['Центелла']},
 '214312385':{'eff':['Антивозрастной уход'],'app':['Дневной уход','Ночной уход'],'skin':['Для всех типов кожи','Для зрелой кожи'],'ing':['Пептиды'],'country':'Канада'},
 '419630142':{'eff':['Увлажнение','Восстановление'],'app':['Дневной уход','Ночной уход'],'skin':['Для всех типов кожи'],'ing':['Муцин улитки','Ниацинамид']},
 '437776835':{'eff':['Увлажнение','Сужение пор','Восстановление'],'app':['Дневной уход','От покраснений'],'skin':['Для жирной кожи','Для проблемной кожи','Для комбинированной кожи'],'ing':['Центелла','LHA-кислота','Гиалуроновая кислота']},
 '551465879':{'eff':['Антивозрастной уход','Сужение пор'],'app':['Ночной уход'],'skin':['Для всех типов кожи','Для зрелой кожи'],'ing':['Ретинол','Микроиглы']},
 '551472457':{'eff':['Восстановление','Увлажнение','Тонизирование'],'app':['Дневной уход','Ночной уход'],'skin':['Для всех типов кожи','Для сухой кожи','Для комбинированной кожи'],'ing':['Растительный экстракт','Ниацинамид']},
 '551482493':{'eff':['Осветление','Сужение пор'],'app':['Дневной уход','От пигментации'],'skin':['Для всех типов кожи'],'ing':['Ниацинамид']},
 '868700077':{'eff':['Восстановление','Увлажнение'],'app':['Ночной уход'],'skin':['Для всех типов кожи','Для зрелой кожи'],'ing':['Лизат молочнокислых бактерий','Микроиглы','Пептиды']},
 '1119249556':{'eff':['Восстановление','Увлажнение','Антивозрастной уход'],'app':['Дневной уход','Ночной уход'],'skin':['Для всех типов кожи'],'type':'Эссенция для ухода за кожей','ing':['PDRN','Ниацинамид'],'vol':30,'country':'Южная Корея','srok':1080},
 '1123494802':{'eff':['Восстановление','Антивозрастной уход'],'app':['Ночной уход'],'skin':['Для проблемной кожи','Для зрелой кожи'],'ing':['Микроиглы','Ниацинамид','Масло']},
}

def detect_ing(rec, extra):
    blob=' '.join(str(rec.get(k) or '') for k in ['Состав','Наименование','Артикул продавца','Назначение косметического средства']).lower()
    found=[]
    for kw,val in ING.items():
        if kw in blob and val not in found:
            found.append(val)
    # merge curated first (priority), keep order, max 5
    res=[]
    for v in (extra or []): 
        if v not in res: res.append(v)
    for v in found:
        if v not in res: res.append(v)
    return res[:5]

BRAND_CASE={'farmstay':'FARMSTAY','manyo':'Manyo','elizavecca':'Elizavecca','skin1004':'SKIN1004',
 'the ordinary':'The Ordinary','round lab':'ROUND LAB','celimax':'Celimax','vt cosmetics':'VT Cosmetics'}
def title(rec):
    name=str(rec.get('Наименование') or '').strip()
    brand=str(rec.get('Бренд') or '').strip()
    brand=BRAND_CASE.get(brand.lower(), brand)
    vol=rec.get('Объем товара')
    m=re.search(r',?\s*\d+\s*мл\s*$', name)
    base=name[:m.start()].rstrip(' ,') if m else name
    t=base
    if brand and brand.lower() not in base.lower():
        t=f"{base} {brand}"
    try:
        v=int(float(vol))
        t=f"{t}, {v} мл"
    except Exception:
        pass
    return t[:200]

# --- Состав EN (INCI) для 4 позиций без состава в выгрузке ВБ (с сайтов брендов, проверить по упаковке партии) ---
SOSTAV_EN={
 '212885914':'Water; Glycereth-26; Glycerin; Dimethicone; Dipropylene Glycol; Niacinamide; Pentaerythrityl Tetraisostearate; Sea Water; Trehalose; Cyclopentasiloxane; Ammonium Acryloyldimethyltaurate/VP Copolymer; Squalane; Sodium Hyaluronate; Hydrolyzed Collagen; Vegetable Amino Acids; Oat Amino Acids; Butylene Glycol; Dimethicone/Vinyl Dimethicone Crosspolymer; Dimethiconol; Hydroxyethyl Acrylate/Sodium Acryloyldimethyl Taurate Copolymer; Polysorbate 60; Carbomer; Arginine; Ethylhexylglycerin; 1,2-Hexanediol; Hydroxyacetophenone; Adenosine; Disodium EDTA; Fragrance (Parfum); Hexyl Cinnamal; Citronellol; Geraniol; Alpha-Isomethyl Ionone; Linalool; Limonene',
 '214311963':'Water; Glycerin; Butylene Glycol; Centella Asiatica Extract; 1,2-Hexanediol; Cellulose Gum; Ethylhexylglycerin',
 '868700077':'Bifida Ferment Lysate; Methylpropanediol; Dipropylene Glycol; Butylene Glycol; 1,2-Hexanediol; Niacinamide; Water; Glycerin; Pentylene Glycol; Chlorella Vulgaris Extract; Hydrolyzed Sponge; Theobroma Cacao (Cocoa) Extract; Sodium Hyaluronate; Glyceryl Polymethacrylate; Polyglyceryl-10 Laurate; Ammonium Acryloyldimethyltaurate/VP Copolymer; Glucose; Acrylates/C10-30 Alkyl Acrylate Crosspolymer; Diethoxyethyl Succinate; Hydrogenated Lecithin; Fructooligosaccharides; Fructose; Hydrogenated Polyisobutene; Panthenol; Tromethamine; Ethylhexylglycerin; Allantoin; Adenosine; Dextrin; Gluconolactone; Sodium Phytate; Caprylic/Capric Triglyceride; Tocopherol; Ceramide NP; Hyaluronic Acid/Polyisopropylacrylamide Copolymer; Sodium Stearoyl Glutamate',
 '1119249556':'Water; Isopropyl Myristate; Panax Ginseng Root Extract; Caprylic/Capric Triglyceride; Cetyl Ethylhexanoate; Glycerin; Propanediol; Niacinamide; Polyglyceryl-2 Dipolyhydroxystearate; PEG-30 Dipolyhydroxystearate; Polyglyceryl-3 Diisostearate; 1,2-Hexanediol; Tocopherol; Hydroxyacetophenone; Helianthus Annuus (Sunflower) Seed Oil; Acrylates/C10-30 Alkyl Acrylate Crosspolymer; Arginine; Adenosine; Ceramide NP; Adansonia Digitata Seed Oil; Disodium EDTA; Hydrogenated Lecithin; Allantoin; Squalane; Polyglucuronic Acid; Ceramide NS; Ceramide AS; Ceramide EOP; Ceramide AP; Sodium DNA',
}
# --- Способ применения RU для всех 15 (с сайтов брендов/маркетплейсов) ---
SPOSOB={
 '168956391':'Нанесите ампулу на очищенную и тонизированную кожу лица (при желании — шеи и зоны декольте) утром и/или вечером. Распределите по коже лёгкими массирующими движениями до впитывания, затем закройте уход кремом.',
 '198147185':'После очищения и тонизирования нанесите небольшое количество ампулы на лицо и распределите по коже, мягко вбивая подушечками пальцев до полного впитывания. Используйте утром и вечером, затем нанесите крем. При чувствительной коже можно наносить точечно на проблемные зоны.',
 '198156472':'После умывания и тоника нанесите эссенцию на лицо ладонями, мягко вбивая и распределяя до полного впитывания; можно также нанести на ватный диск и протереть лицо. Применяйте утром и вечером перед кремом. Если параллельно используете средства с витамином C, разнесите их по времени или сделайте паузу около 30 минут.',
 '206058262':'Нанесите несколько капель сыворотки на очищенную кожу лица и шеи и распределите лёгкими похлопывающими движениями до впитывания. Можно использовать как маску (пропитать ватные диски и приложить к сухим участкам) или добавить несколько капель в привычный крем. Наносится до более плотных кремов.',
 '212885914':'На завершающем этапе ухода прокатайте металлическим роликом-аппликатором по коже вокруг глаз, слегка нажимая, чтобы нанести сыворотку тонким слоем. Двигайтесь от внутреннего уголка глаза к внешнему; излишки вбейте подушечками пальцев до полного впитывания. Прохладный ролик помогает снять отёчность.',
 '214311963':'Нанесите 2–3 капли ампулы на очищенную кожу после тонера и мягко вбейте подушечками пальцев до впитывания. Используйте утром и вечером перед кремом; на сухие участки можно нанести повторно. Завершите уход увлажняющим кремом, а днём солнцезащитным средством.',
 '214312385':'Нанесите несколько капель сыворотки на очищенную кожу всего лица или на проблемные зоны утром и вечером, вбивая ладонями в кожу. Используйте на этапе ухода после очищения и до нанесения крема. Не сочетайте в одном уходе с прямыми кислотами или чистым витамином C.',
 '419630142':'Используйте дважды в день после умывания и тоника. Нанесите нужное количество ампулы с муцином улитки на кожу лица и шеи кончиками пальцев и равномерно распределите до впитывания. Средство с лёгкой гелевой текстурой быстро впитывается и не оставляет липкости.',
 '437776835':'Применяйте на этапе эссенции/сыворотки, после тоника, утром и/или вечером. Нанесите 1–2 капли на чистую кожу и мягко вбейте подушечками пальцев по всему лицу до полного впитывания, затем нанесите крем.',
 '551465879':'Используйте в вечернем уходе после очищения и тонизирования. Нанесите умеренное количество сыворотки на сухую кожу и мягко вбейте подушечками пальцев до впитывания, затем закрепите увлажняющим кремом. В начале применяйте через день, постепенно доводя до ежедневного использования; утром обязательно наносите солнцезащитный крем.',
 '551472457':'Применяйте утром и вечером после тоника/эссенции и перед кремом. Нанесите 1–2 нажатия ампулы равномерно по лицу подушечками пальцев, избегая области вокруг глаз, и мягко вбейте до полного впитывания. Лучше наносить на слегка влажную кожу для усиления увлажнения.',
 '551482493':'Используйте утром и вечером после очищения и тонизирования, перед кремом. Нанесите умеренное количество равномерно по лицу, уделяя внимание участкам с пигментацией и расширенными порами, и мягко вбейте до впитывания. Днём завершайте уход солнцезащитным кремом.',
 '868700077':'Хорошо встряхните флакон перед применением. После очищения и тоника нанесите ампулу на лицо и мягко вбейте подушечками пальцев до впитывания — вы можете почувствовать лёгкое покалывание от микроспикул, это нормально. Используйте вечером 2–3 раза в неделю, начиная с малого количества, затем нанесите увлажняющий крем; в дни применения обязательно пользуйтесь SPF днём.',
 '1119249556':'После умывания нанесите небольшое количество эссенции на лицо и распределите от центра к периферии, слегка вдавливая подушечками пальцев до полного впитывания. В вечернем уходе используйте после бустеров с микроиглами, затем завершите уход кремом. Подходит для всех типов кожи.',
 '1123494802':'В вечернем уходе нанесите средство на очищенную кожу, распределите пальцами и дайте полностью впитаться, затем продолжите уход тонером и кремом. В начале применяйте раз в 3 дня, постепенно доводя до ежедневного использования; лёгкое покалывание при первых применениях считается нормой. Днём обязательно используйте солнцезащитный крем SPF50 PA++++.',
}
# --- Производитель (изготовитель, не бренд) ---
PRODUCER={
 '168956391':'Myungin Cosmetics Co., Ltd.',   # OEM подтверждён (собств. завод бренда)
 '212885914':'Myungin Cosmetics Co., Ltd.',   # -//-
 '419630142':'Myungin Cosmetics Co., Ltd.',   # -//-
 '206058262':'Miz Trade Inc.',                # изготовитель/держатель на упаковке
 '198147185':'Cosmax Co., Ltd.',              # OEM (по корейским маркетплейсам)
 '198156472':'Manyo Co., Ltd.',               # OEM не подтверждён -> юрлицо-держатель
 '868700077':'Manyo Co., Ltd.',               # OEM не подтверждён -> юрлицо-держатель
 '437776835':'COSMOCOS Co., Ltd.',            # OEM подтверждён
 '551465879':'Absolvlab Inc.',                # OEM не подтверждён -> юрлицо-держатель Celimax
 '551472457':'Absolvlab Inc.',                # -//-
 '551482493':'Absolvlab Inc.',                # -//-
 '1119249556':'Hanacos Co., Ltd.',            # OEM (по Olive Young)
 '1123494802':'VT Cosmetics Inc.',            # OEM не подтверждён -> юрлицо-держатель
 '214311963':'Craver Corporation',            # OEM не подтверждён -> юрлицо-держатель SKIN1004
 '214312385':'DECIEM Inc.',                   # изготовитель (Канада)
}

# open template
tpl=os.path.join(SP,'ozon_serum_template.xlsx')
wb=openpyxl.load_workbook(tpl)
ws=wb['Шаблон']

order=['168956391','198147185','198156472','206058262','212885914','214311963','214312385','419630142','437776835','551465879','551472457','551482493','868700077','1119249556','1123494802']

def setc(row,col,val):
    ws.cell(row=row,column=col,value=val)

missing_sostav=[]
row=5
for i,aw in enumerate(order,1):
    r=data[aw]
    cur=CUR[aw]
    vol=cur.get('vol') or r.get('Объем товара')
    photos=[p for p in str(r.get('Фото') or '').split(';') if p.strip()]
    main_photo=photos[0] if photos else None
    add_photos=';'.join(photos[1:]) if len(photos)>1 else None
    sostav=r.get('Состав') or SOSTAV_EN.get(aw)
    if not sostav: missing_sostav.append(aw)
    barcode=r.get('Баркод')
    w_kg=r.get('Вес с упаковкой (кг)')
    weight_g=int(round(float(w_kg)*1000)) if w_kg else None
    shir=r.get('Ширина упаковки'); vys=r.get('Высота упаковки'); dl=r.get('Длина упаковки')
    def mm(x):
        try: return int(round(float(x)*10))
        except: return None
    srok=cur.get('srok') or srok_days(r.get('Срок годности'))
    country=cur.get('country') or ('Южная Корея' if r.get('Страна производства') else 'Южная Корея')

    setc(row,1,i)                                   # №
    setc(row,2,r.get('Артикул продавца'))           # Артикул*
    setc(row,3,title(r))                            # Название*
    # D price left EMPTY (pending pricing rule)
    setc(row,6,7)                                   # НДС %*
    setc(row,7,'Нет')                               # Рассрочка
    setc(row,8,'Нет')                               # ускоренный сбор отзывов
    setc(row,10,barcode)                            # Штрихкод
    setc(row,11,'3304990000 - Прочие косметические средства или средства для макияжа и средства для ухода за кожей (кроме лекарственных), включая средства против загара или для загара; средства для маникюра или педикюра')  # ТН ВЭД*
    setc(row,12,weight_g)                           # вес упак г*
    setc(row,13,mm(shir))                           # ширина мм*
    setc(row,14,mm(vys))                            # высота мм*
    setc(row,15,mm(dl))                             # длина мм*
    setc(row,16,main_photo)                         # главное фото*
    if add_photos: setc(row,17,add_photos)          # доп фото
    setc(row,19,cur.get('type','Сыворотка для ухода за кожей'))  # Тип*
    setc(row,20,r.get('Бренд'))                     # Бренд*
    setc(row,21,r.get('Артикул продавца'))          # Название модели*
    setc(row,22,1)                                  # единиц в товаре
    setc(row,23,vol)                                # объём мл
    setc(row,24,sostav)                             # Состав*
    setc(row,25,cur.get('area','Лицо'))             # Область использования*
    setc(row,26,'Нет')                              # нужен код маркировки*
    setc(row,27,1095)                               # срок годности дней* (фикс 1095)
    setc(row,30,PRODUCER.get(aw, r.get('Бренд')))   # производитель (изготовитель)
    setc(row,33,r.get('Описание'))                  # аннотация
    setc(row,38,SPOSOB.get(aw))                     # способ применения (RU)
    ings=detect_ing(r,cur.get('ing'))
    setc(row,39,';'.join(ings) if ings else None)   # бьюти-ингредиент
    setc(row,40,';'.join(cur.get('skin',['Для всех типов кожи'])))  # тип кожи
    setc(row,42,';'.join(cur.get('eff',[])))        # эффект
    setc(row,43,';'.join(cur.get('app',[])))        # применение
    setc(row,45,'Для любого возраста')              # возрастной диапазон
    setc(row,46,country)                            # страна-изготовитель
    setc(row,47,1)                                  # количество заводских упаковок
    row+=1

os.makedirs('/home/user/K-beauty-New/outputs/ozon_serum_cards',exist_ok=True)
outp='/home/user/K-beauty-New/outputs/ozon_serum_cards/ozon_serums_2026-08-06.xlsx'
wb.save(outp)
print('SAVED', outp)
print('rows filled:', len(order))
print('MISSING sostav (need to add):', missing_sostav)
