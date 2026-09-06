# Источники по судовому оборудованию

Данные выгружены из Google Drive, папка `судовое оборудование`
(`1RMkSIT-20Edzd-V5vJ95LqOczFK9IPpu`).

| CSV | Исходный файл на Drive | ID файла |
| --- | --- | --- |
| `armatura_ulica_foye.csv` | `Арматура.xlsx` | `1h07EP3gPYuwpmRgFoonozSBMEBTJZtqg` |
| `sklad_1a_1b_2_3_armatura.csv` | `Склад 1А,1Б, 2,3 арматура.xlsx` | `1V56XVH8dcKRG9bglpSGuOOaJFkSlP9D-` |
| `elektrika.csv` | `Электрика.xlsx` | `19eiEjfvZcrRjYhLAmAmBBe4cI00e0CHa` |
| `dveri_kryshki_illuminatory.csv` | `Двери , крышки , иллюминаторы.xlsx` | `1UhoLH0iqB0BO2TuwfW-Zx6Tkw7kGEYqV` |
| `brashpil_nasos_lebedki.csv` | `Брашпиль, насос, лебедки.xlsx` | `1VtpUbUnFJqu2gk2ExsBZ2vTurFSzUA53` |

Отдельно на Drive лежат:

- `отчет кп суд.оборудование.xlsx` (`1lsWvrpdTHiQiY4cveoklCnkLGyxVbFoy`) — журнал
  обзвона покупателей. В репозиторий не копируем: файл состоит из личных
  контактов сотрудников сторонних компаний.
- папка `Фото` (`1FL3PM-pKy5rr0DBr70bm-1xI9Q3hfl_q`) — фото позиций для КП.

## Что в данных есть и чего нет

Есть: место хранения, наименование, номер чертежа, количество.
Нет: цен, единиц измерения, веса и габаритов, года выпуска, состояния
по большинству позиций.

Цена по позиции проставляется в `outputs/marine_equipment/sudovoe_price_template_*.tsv`
из прайсов поставщиков или из факта сделок, исходные CSV при этом не меняем.
