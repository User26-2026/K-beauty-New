# Как перенести этот проект в Claude Code / Claude Web с Git

## Вариант А. Если Claude работает через web + подключенный Git

Это твой случай. Нужно загрузить содержимое этой папки в отдельный приватный GitHub/GitLab репозиторий.

### Что важно

- `CLAUDE.md` должен лежать в корне репозитория.
- Все файлы должны лежать внутри репозитория, а не ссылаться на `/Users/nikita/...`.
- В Claude Web открываешь подключенный Git-репозиторий и даешь стартовый промпт ниже.
- Если сервис не читает большие `.xlsx` напрямую, попроси агента сначала посмотреть `CLAUDE.md`, потом нужные CSV/JSON/XLSX из `workspace/`.

### Как загрузить

1. Создай приватный репозиторий, например `wb-kbeauty-agent`.
2. Загрузи в него содержимое папки `claude_code_transfer`, не саму zip-архивную папку.
3. Проверь, что в корне репозитория видны:

```text
CLAUDE.md
README_IMPORT.md
requirements.txt
tools/
data/
workspace/
```

4. В Claude Web выбери этот репозиторий как источник проекта.

### Стартовый промпт для Claude Web

```text
Изучи CLAUDE.md и структуру репозитория. Работай как агент по Wildberries для запуска магазина корейской косметики.

Если я прошу юнитку, считай по нашей логике из CLAUDE.md и tools/calc_unit.py.
Основные данные лежат в data/ и workspace/.
Если просится таблица для Google Sheets, выдавай TSV.
Если нужно обновить PDF, редактируй workspace/wb_kbeauty_launch/build_roadmap_pdf.py и пересобирай PDF.
Не используй локальные пути /Users/nikita, работай только с файлами репозитория.
```

### Если нужно добавить новые файлы

Новые выгрузки WB, прайсы или юнитки добавляй в репозиторий:

```text
data/new_uploads/
```

И в сообщении агенту пиши:

```text
Я добавил новый файл в data/new_uploads/<название>.xlsx, прочитай его и продолжи по нашей логике.
```

## Вариант Б. Если Claude Code установлен локально

По официальной инструкции Anthropic установка идет через npm:

```bash
npm install -g @anthropic-ai/claude-code
```

Нужен Node.js 18+.

## Открыть проект локально

Скопируй папку `claude_code_transfer` в удобное место и зайди в нее:

```bash
cd /path/to/claude_code_transfer
claude
```

Claude Code должен автоматически прочитать `CLAUDE.md` в корне проекта.

## Установить библиотеки для расчетов

Если нужно запускать Python-скрипты:

```bash
python3 -m pip install -r requirements.txt
```

## Стартовый промпт для локального Claude Code

Скопируй в Claude Code:

```text
Изучи CLAUDE.md и структуру проекта. Работай как агент по Wildberries для запуска магазина корейской косметики. 
Основные файлы: data/unit_economics/Celimax_unit_model.xlsx, data/stock_costs/ОСТАТКИ_КОРЕЯ_03.06.xlsx, workspace/wb_kbeauty_launch/build_roadmap_pdf.py.
Если я прошу юнитку, считай по нашей логике из CLAUDE.md и tools/calc_unit.py. Если прошу таблицу для Google Sheets, выдавай TSV.
```

## 5. Что уже лежит внутри

- `CLAUDE.md` — память и правила агента.
- `data/` — ключевые Excel-файлы.
- `workspace/` — старые отчеты, PDF и скрипты.
- `tools/calc_unit.py` — быстрый расчет юнитки.

## 6. Файлы, которые можно добавить вручную

Некоторые временные прайсы из WhatsApp могли исчезнуть со старых путей. Если они нужны, положи их сюда:

```text
data/price_lists/
```

Желательные файлы:

- `CELIMAX PRICE LIST(-VAT)_26.03.xlsx`
- `DERMAFACTORY_2601.xlsx`
- `AXIS-Y PRICE LIST(-VAT)_26.01.xlsx`
- `VT PRICE LIST(-VAT)_25.05.xlsx`

Даже без них в папке уже есть рассчитанные отчеты, где часть себестоимости и баркодов сохранена.
