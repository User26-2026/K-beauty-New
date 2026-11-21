# tools/ozon — аналитика OZON (переиспользуемые скрипты)

Все скрипты читают выгрузки из кабинета OZON Seller и складывают результат в
`outputs/ozon_*`. Зависимости: `openpyxl`, `xlrd`, `python-calamine`.

## sales_pnl.py — прибыль по SKU
Из отчёта продаж OZON (.xls, лист TDSheet с ABC). Берёт готовую строку
«Маржа, руб» (чистая прибыль), считает ДРР и итог.
```
python tools/ozon/sales_pnl.py <отчёт_продаж.xls> [outputs/ozon_pnl/pnl.xlsx]
```

## funnel.py — воронка по SKU + узкое место
Из аналитики «По товарам» (.xlsx). Конверсии показ→карточка→заказ→выкуп,
классификация узкого места (видимость / CTR / карточка / выкуп) и действие.
```
python tools/ozon/funnel.py <аналитика_по_товарам.xlsx> [out.xlsx]
```

## catalog_health.py — блокеры каталога
Из выгрузки товаров (CSV, разделитель `;`). Группирует «Не продается» по
причине (документы бренда/качества, фото, копии, мусорные карточки) +
предупреждения.
```
python tools/ozon/catalog_health.py <каталог.csv> [out.xlsx]
```

## Генераторы карточек
Лежат в `workspace/ozon_product_cards/build_ozon_*_cards.py` (по одной
категории). Правила заполнения — в корневом `CLAUDE.md`, раздел «OZON».

## API
`tools/ozon_api/fetch_tariffs.py` — тарифы/комиссии из Ozon Seller API
(ключи `OZON_CLIENT_ID`, `OZON_API_KEY`).
