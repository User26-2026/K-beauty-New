import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "/Users/nikita/Desktop/Новая папка 2/outputs/wb_product_research";
const outputPath = `${outputDir}/kbeauty_candidates_clean_2026-05-29.xlsx`;

const theme = {
  navy: "#173653",
  teal: "#0F766E",
  mint: "#E6F4F0",
  paleBlue: "#EEF4FA",
  green: "#DDF3E4",
  gray: "#F3F6FA",
  text: "#1F2937",
  white: "#FFFFFF",
};

const rows = [
  [
    "Добавить",
    "Celimax",
    "Dual Barrier Mild Gel Cleanser 200 ml",
    "Гель для умывания барьерный с церамидами",
    "1 300-1 500 руб.",
    8193,
    10534635,
    "nm 712024737",
    359,
    1315,
    132,
    295,
    "По косвенным признакам нет большой зависимости от внешнего трафика",
    "Молодая карточка уже продает 359 шт/мес; сильный аналог дает 444 шт/мес; товар не найден в текущем ассортименте магазина.",
  ],
  [
    "Добавить",
    "Celimax",
    "Noni Acne Bubble Cleanser",
    "Пенка для умывания пузырьковая с нони от акне",
    "1 050-1 250 руб.",
    8193,
    10534635,
    "nm 712022766",
    182,
    1138,
    121,
    120,
    "По косвенным признакам нет большой зависимости от внешнего трафика",
    "Молодая карточка: 121 день, продажи +333%, 88 поисковых запросов, средняя позиция 4, 1 316 полок ТОП-10.",
  ],
  [
    "Запасной вариант",
    "Celimax",
    "Derma Nature Relief Madecica pH Balancing Foam Cleansing 150 ml",
    "Балансирующая пенка для умывания Madecica pH",
    "1 200-1 400 руб.",
    8193,
    10534635,
    "nm 155947140",
    323,
    1387,
    1103,
    2564,
    "Нет отдельной колонки по внешнему трафику в выгрузке",
    "Продажи есть: 323 шт/мес и 449 761 руб., но карточка старая, с 2 564 отзывами и падающей динамикой продаж -34%.",
  ],
  [
    "Не отдельный товар",
    "Celimax",
    "Dual Barrier Mild Gel Cleanser, аналог nm 228473128",
    "Гель-пенка с комплексом церамидов Dual Barrier Mild",
    "1 300-1 500 руб.",
    8193,
    10534635,
    "nm 228473128",
    82,
    1659,
    736,
    339,
    "Нет отдельной колонки по внешнему трафику в выгрузке",
    "Использовать как подтверждающий аналог к Dual Barrier, но не дублировать отдельной закупочной строкой.",
  ],
];

const analogs = [
  [233144319, "Гель для умывания барьерный с церамидами", "MASKSHOP.RU", 1769, 444, 780860, 735, 327, 67, 39, 20.0, "Сильная взрослая карточка"],
  [712024737, "Гель для умывания лица Celimax Dual Barrier Mild Gel Cleanser 200ml", "Modern Factory", 1315, 359, 480091, 132, 295, 88, 77, 3.0, "Главный ориентир"],
  [712022766, "Пенка для умывания пузырьковая Noni Acne Bubble Cleanser", "Modern Factory", 1138, 182, 206530, 121, 120, 49, 88, 4.0, "Отдельный кандидат"],
  [155947140, "Балансирующая пенка для умывания Derma Nature Relief Madecica pH Balancing Foam Cleansing, 150 мл", "LIFECOSM", 1387, 323, 449761, 1103, 2564, 159, 36, 22.5, "Спрос есть, но карточка старая"],
  [228473128, "Гель-пенка с комплексом церамидов Dual Barrier Mild Корея", "Minari Cosmetics", 1659, 82, 137778, 736, 339, 15, 104, 530.5, "Аналог, не отдельный товар"],
];

const wb = Workbook.create();
wb.comments.setSelf({ displayName: "ChatGPT" });

function setWidths(sheet, widths) {
  widths.forEach((width, index) => {
    sheet.getCell(0, index).format.columnWidthPx = width;
  });
}

function title(sheet, range, text, subtitle) {
  const end = range.split(":")[1].replace(/[0-9]/g, "");
  const r = sheet.getRange(range);
  r.merge();
  r.values = [[text]];
  r.format = { fill: theme.navy, font: { bold: true, color: theme.white }, verticalAlignment: "center" };
  r.format.rowHeightPx = 36;
  const s = sheet.getRange(`A2:${end}2`);
  s.merge();
  s.values = [[subtitle]];
  s.format = { fill: theme.paleBlue, font: { color: theme.text }, wrapText: true, verticalAlignment: "center" };
  s.format.rowHeightPx = 34;
}

function header(range) {
  range.format = { fill: theme.navy, font: { bold: true, color: theme.white }, wrapText: true, horizontalAlignment: "center", verticalAlignment: "center" };
}

function body(range) {
  range.format = { fill: theme.white, font: { color: theme.text }, wrapText: true, verticalAlignment: "top" };
}

const main = wb.worksheets.add("Кандидаты");
main.showGridLines = false;
title(main, "A1:N1", "Кандидаты на заведение: корейская косметика", "Чистая версия без расчетного ДРР и без спорной формулировки landed. ДРР будет рассчитан после фактической юнит-экономики.");
setWidths(main, [150, 105, 250, 260, 145, 125, 135, 140, 120, 105, 105, 105, 260, 360]);
const headers = [
  "Решение",
  "Бренд",
  "Конкретный товар",
  "Что искать у поставщика",
  "Ценовой коридор WB",
  "Продажи кластера, шт/мес",
  "Выручка кластера, руб/мес",
  "Карточка-ориентир",
  "Продажи SKU, шт/мес",
  "Цена SKU, руб",
  "Возраст, дней",
  "Отзывы",
  "Внешний трафик",
  "Обоснование",
];
main.getRange("A4:N4").values = [headers];
main.getRange("A5:N8").values = rows;
header(main.getRange("A4:N4"));
body(main.getRange("A5:N8"));
main.getRange("F5:G8").format.numberFormat = Array.from({ length: 4 }, () => ["#,##0", "#,##0"]);
main.getRange("I5:L8").format.numberFormat = Array.from({ length: 4 }, () => ["#,##0", "#,##0", "#,##0", "#,##0"]);
main.getRange("A4:N4").format.rowHeightPx = 48;
main.getRange("A5:N8").format.rowHeightPx = 74;
main.tables.add("A4:N8", true, "CleanCandidatesTable").style = "TableStyleMedium2";
main.freezePanes.freezeRows(4);

const unit = wb.worksheets.add("Данные для юнитки");
unit.showGridLines = false;
title(unit, "A1:H1", "Что нужно для корректного расчета ДРР", "В этом файле ДРР не считается как проходной показатель, потому что нет фактической себестоимости и полной структуры расходов WB.");
setWidths(unit, [260, 420, 240, 260, 160, 160, 160, 160]);
unit.getRange("A4:D4").values = [["Что нужно", "Зачем", "Где взять", "Статус"]];
unit.getRange("A5:D12").values = [
  ["Закупочная цена товара", "Без нее нельзя посчитать прибыль на штуку и проходной ДРР.", "Прайс поставщика / инвойс", "Нужно получить"],
  ["Доставка до склада / подготовка", "Влияет на себестоимость единицы.", "Логистика поставки / расчет партии", "Нужно получить"],
  ["Комиссия WB по предмету", "Комиссия напрямую снижает маржу.", "Кабинет WB / тарифы категории", "Нужно проверить"],
  ["Логистика WB", "Фактическая логистика зависит от габаритов, склада и локальности.", "Кабинет WB / юнит-отчет", "Нужно проверить"],
  ["Хранение", "Для медленных SKU может съесть прибыль.", "Отчет платного хранения / прогноз оборачиваемости", "Нужно проверить"],
  ["Эквайринг и прочие удержания", "Нужны для чистой прибыли после WB.", "Отчет доходов и расходов", "Нужно проверить"],
  ["Плановая цена продажи", "От нее считается допустимый рекламный расход.", "MarketGuru / WB выдача", "Есть рыночный коридор"],
  ["Рекламные расходы", "Фактический ДРР можно считать только после теста или по аналогам.", "Реклама WB / MarketGuru", "Есть только рыночный ориентир"],
];
header(unit.getRange("A4:D4"));
body(unit.getRange("A5:D12"));
unit.getRange("A4:D12").format.rowHeightPx = 46;
unit.tables.add("A4:D12", true, "UnitDataTable").style = "TableStyleMedium4";

unit.getRange("F4:H8").values = [
  ["Текущий вывод", "", ""],
  ["ДРР в основной таблице убран.", "", ""],
  ["Ранее использовался рыночный ориентир из выгрузки по категории/кластеру, а не расчет проходного ДРР.", "", ""],
  ["Проходной ДРР нужно считать после фактической закупки и всех расходов WB.", "", ""],
  ["До этого момента ДРР нельзя показывать как готовую рекомендацию.", "", ""],
];
unit.getRange("F4:H4").merge();
unit.getRange("F5:H8").merge(true);
header(unit.getRange("F4:H4"));
body(unit.getRange("F5:H8"));
unit.getRange("F5:H8").format = { fill: theme.mint, font: { bold: true, color: theme.text }, wrapText: true, verticalAlignment: "top" };
unit.getRange("F4:H8").format.rowHeightPx = 48;

const analog = wb.worksheets.add("Аналоги в выдаче");
analog.showGridLines = false;
title(analog, "A1:L1", "Карточки-аналоги из выдачи Celimax", "Здесь оставлены карточки, по которым принималось решение: что добавлять, что держать запасным вариантом, а что не дублировать.");
setWidths(analog, [115, 360, 180, 100, 105, 120, 105, 105, 115, 115, 115, 260]);
analog.getRange("A4:L4").values = [["Артикул", "Товар", "Продавец", "Цена, руб", "Продажи, шт", "Выручка, руб", "Возраст, дней", "Отзывы", "Новые отзывы", "Запросы, шт", "Средняя позиция", "Комментарий"]];
analog.getRange(`A5:L${4 + analogs.length}`).values = analogs;
header(analog.getRange("A4:L4"));
body(analog.getRange(`A5:L${4 + analogs.length}`));
analog.getRange(`D5:K${4 + analogs.length}`).format.numberFormat = Array.from({ length: analogs.length }, () => Array(8).fill("#,##0"));
analog.getRange("A4:L4").format.rowHeightPx = 44;
analog.getRange(`A5:L${4 + analogs.length}`).format.rowHeightPx = 58;
analog.tables.add(`A4:L${4 + analogs.length}`, true, "AnalogCardsCleanTable").style = "TableStyleMedium9";
analog.freezePanes.freezeRows(4);

const sources = wb.worksheets.add("Источники");
sources.showGridLines = false;
title(sources, "A1:F1", "Источники и ограничения", "Фиксируем, какие цифры подтверждены выгрузкой, а какие требуют отдельной проверки.");
setWidths(sources, [230, 520, 260, 170, 170, 260]);
sources.getRange("A4:F4").values = [["Блок", "Источник", "Что использовано", "Период", "Статус", "Ограничение"]];
sources.getRange("A5:F9").values = [
  ["Карточная выдача", "2026-04-28—2026-05-27_FBO_Поисковые_запросы (3).xlsx", "Карточки, продажи, цена, отзывы, возраст, позиции", "28.04-27.05.2026", "Использовано", "Данные по внешнему трафику отдельной колонкой не выгружены."],
  ["Частотность", "2026-04-28—2026-05-27_FBO_Поисковые_запросы (1).xlsx", "Запросы Celimax умывалка / пенка / гель", "28.04-27.05.2026", "Использовано", "Нужна карточная выдача по каждому новому запросу."],
  ["Внешний трафик", "Косвенная оценка по поисковым позициям и количеству запросов", "Формулировка: нет признака большой зависимости", "28.04-27.05.2026", "Предварительно", "Для жесткого подтверждения нужна выгрузка источников трафика в MarketGuru/WB."],
  ["ДРР", "Рыночный ориентир из выгрузки, не проходной расчет", "Из основной таблицы убран", "Модель", "Не рассчитан", "Нужны себестоимость, комиссии, логистика, хранение и прочие расходы WB."],
  ["Текущий магазин", "supplier-goods-127314-2026-05-01-2026-05-28-exvxkgdeb.XLSX", "Проверка, что Celimax cleanser/foam/gel не найден", "01.05-28.05.2026", "Использовано", "В магазине есть Round Lab cleansers, но это другой бренд."],
];
header(sources.getRange("A4:F4"));
body(sources.getRange("A5:F9"));
sources.getRange("A4:F4").format.rowHeightPx = 44;
sources.getRange("A5:F9").format.rowHeightPx = 58;
sources.tables.add("A4:F9", true, "SourcesCleanTable").style = "TableStyleMedium2";

await fs.mkdir(outputDir, { recursive: true });

const preview = await wb.inspect({
  kind: "table",
  range: "Кандидаты!A4:N8",
  include: "values,formulas",
  tableMaxRows: 8,
  tableMaxCols: 14,
});
console.log(preview.ndjson);

const errors = await wb.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "formula error scan",
});
console.log(errors.ndjson);

const render = await wb.render({ sheetName: "Кандидаты", autoCrop: "all", scale: 1, format: "png" });
await fs.writeFile(`${outputDir}/kbeauty_candidates_clean_candidates.png`, new Uint8Array(await render.arrayBuffer()));

const output = await SpreadsheetFile.exportXlsx(wb);
await output.save(outputPath);
console.log(outputPath);
