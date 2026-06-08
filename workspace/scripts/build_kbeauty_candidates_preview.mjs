import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "/Users/nikita/Desktop/Новая папка 2/outputs/wb_product_research";
const outputPath = `${outputDir}/kbeauty_candidates_preview_2026-05-29.xlsx`;

const theme = {
  navy: "#173653",
  teal: "#0F766E",
  mint: "#E6F4F0",
  paleBlue: "#EEF4FA",
  amber: "#F8C24A",
  green: "#DDF3E4",
  red: "#FBE4E4",
  gray: "#EEF1F5",
  border: "#D6DEE8",
  text: "#1F2937",
  muted: "#6B7280",
  white: "#FFFFFF",
};

const candidate = {
  priority: 1,
  status: "Подтвержден для проверки закупки",
  brand: "Celimax",
  product: "Dual Barrier Mild Gel Cleanser 200 ml",
  wbName: "Гель для умывания барьерный с церамидами",
  sourceSku: 712024737,
  marketSales: 8193,
  marketRevenue: 10534635,
  weightedPrice: 1286,
  medianPrice: 1192,
  sourceSales: 359,
  sourceRevenue: 480091,
  sourcePrice: 1315,
  sourceAge: 132,
  sourceReviews: 295,
  sourceNewReviews: 88,
  sourcePosition: 3,
  targetDrr: "10-12%",
  testDrr: "до 15%",
  priceCorridor: "1 300-1 500 руб.",
  maxCost: "600-650 руб.",
  externalTraffic: "Нет признака большой зависимости от внешнего трафика",
  launchCondition: "Проверить landed-себестоимость, документы на бренд и возможность значка «Оригинал».",
  risk: "Если landed-себестоимость выше 650 руб., маржа заметно сжимается.",
  source: "2026-04-28—2026-05-27_FBO_Поисковые_запросы (3).xlsx",
};

const candidateRows = [
  [
    1,
    "Подтвержден для проверки закупки",
    "Celimax",
    "Dual Barrier Mild Gel Cleanser 200 ml",
    "Гель для умывания барьерный с церамидами",
    "1 300-1 500 руб.",
    "600-650 руб.",
    "10-12%",
    "до 15%",
    8193,
    10534635,
    "nm 712024737",
    359,
    1315,
    132,
    295,
    "Нет признака большой зависимости от внешнего трафика",
    "Молодая карточка уже продает 359 шт/мес; у сильного аналога 444 шт/мес; рынок по релевантным пенкам/гелям 8 193 шт/мес.",
    "Проверить landed-себестоимость, документы на бренд и возможность значка «Оригинал».",
    "Если landed-себестоимость выше 650 руб., маржа заметно сжимается.",
  ],
  [
    2,
    "Добавить в шорт-лист",
    "Celimax",
    "Noni Acne Bubble Cleanser",
    "Пенка для умывания пузырьковая с нони от акне",
    "1 050-1 250 руб.",
    "470-530 руб.",
    "10-12%",
    "до 15%",
    8193,
    10534635,
    "nm 712022766",
    182,
    1138,
    121,
    120,
    "Нет признака большой зависимости от внешнего трафика",
    "Молодая карточка: 121 день, продажи +333%, 88 поисковых запросов, средняя позиция 4, 1 316 полок ТОП-10.",
    "Проверить закупку, документы и наличие стабильной партии.",
    "Нужно не переплатить за закупку: товар дешевле Dual Barrier, цена входа ниже.",
  ],
  [
    3,
    "Наблюдение / второй эшелон",
    "Celimax",
    "Derma Nature Relief Madecica pH Balancing Foam Cleansing 150 ml",
    "Балансирующая пенка для умывания Madecica pH",
    "1 200-1 400 руб.",
    "540-600 руб.",
    "10-12%",
    "до 15%",
    8193,
    10534635,
    "nm 155947140",
    323,
    1387,
    1103,
    2564,
    "Не подтверждено отдельной колонкой в выгрузке",
    "Продажи есть: 323 шт/мес и 449 761 руб., но карточка старая и динамика падает: продажи -34%, заказы -44%.",
    "Добавлять только при очень хорошей закупке и документах.",
    "Высокий отзывный барьер и отрицательная динамика делают запуск новой карточки рискованным.",
  ],
  [
    4,
    "Не отдельный кандидат / аналог",
    "Celimax",
    "Dual Barrier Mild Gel Cleanser, аналог nm 228473128",
    "Гель-пенка с комплексом церамидов Dual Barrier Mild",
    "1 300-1 500 руб.",
    "600-650 руб.",
    "10-12%",
    "до 15%",
    8193,
    10534635,
    "nm 228473128",
    82,
    1659,
    736,
    339,
    "Не подтверждено отдельной колонкой в выгрузке",
    "Это тот же товарный смысл, что уже выбранный Dual Barrier. Плюс: 82 продажи, рост +4 000%. Минус: средняя позиция 530,5.",
    "Использовать как подтверждающий аналог, не дублировать отдельной закупочной строкой.",
    "Если добавить отдельно, таблица будет задваивать один и тот же товар.",
  ],
];

const analogs = [
  [233144319, "Гель для умывания барьерный с церамидами", "MASKSHOP.RU", 1769, 444, 780860, 33950, 735, 327, 67, 39, 20.0, 7, "Сильная взрослая карточка, показывает потолок спроса"],
  [712024737, "Гель для умывания лица Celimax Dual Barrier Mild Gel Cleanser 200ml", "Modern Factory", 1315, 359, 480091, 0, 132, 295, 88, 77, 3.0, 10, "Главный ориентир: молодая карточка с продажами"],
  [263714947, "Кремовый гель-пенка для умывания с керамидами Dual Barrier Mild Gel Cleanser, 200 мл", "LIFECOSM", 1991, 278, 554274, 0, 602, 2564, 159, 41, 6.0, 8, "Высокая цена и большой отзывный хвост"],
  [325987569, "Гель пенка для умывания Dual Barrier Mild Gel Cleanser", "KoreanBoom", 1280, 151, 193711, 0, 479, 3495, 261, 52, 31.5, 13, "Стабильный спрос, но сильная база отзывов"],
  [378809159, "Гель для умывания лица Dual Barrier Mild Gel Cleanser 200 мл", "ИП Жанасыл Tendernesstore", 1478, 114, 168868, 0, 418, 233, 21, 80, 15.5, 7, "Цена ближе к рабочему коридору"],
  [679215808, "Кремовый гель для умывания с керамидами Dual Barrier Mild Gel Cleanser, 200 мл", "Charming", 1741, 75, 114558, 0, 170, 257, 25, 18, 93.0, 4, "Молодая карточка, продажи есть"],
  [455436352, "Гель для умывания Dual barrier mild gel cleanser, 200 мл", "Магазин 'Подружка'", 1865, 51, 92481, 0, 323, 35, 3, 47, 26.0, 15, "Мало отзывов, но продажи слабее"],
  [740777379, "Гель для умывания лица Dual Barrier Mild Gel Cleanser 200 мл", "ИП Туякбаева", 1422, 38, 54630, 0, 144, 29, 9, 84, 32.0, 5, "Молодая карточка с низким отзывным барьером"],
  [491211113, "Гель для умывания лица Dual Barrier Mild Gel Cleanser", "Beauty Corporation", 1289, 38, 49157, 0, 291, 89, 3, 28, 42.0, 3, "Низкая цена, но слабее позиции"],
];

const workbook = Workbook.create();
workbook.comments.setSelf({ displayName: "ChatGPT" });

function styleTitle(sheet, range, title, subtitle) {
  const titleRange = sheet.getRange(range);
  const endColumn = range.split(":")[1].replace(/[0-9]/g, "");
  titleRange.merge();
  titleRange.values = [[title]];
  titleRange.format = {
    fill: theme.navy,
    font: { bold: true, color: theme.white },
    horizontalAlignment: "left",
    verticalAlignment: "center",
  };
  titleRange.format.rowHeightPx = 42;

  if (subtitle) {
    const sub = sheet.getRange(`A2:${endColumn}2`);
    sub.merge();
    sub.values = [[subtitle]];
    sub.format = {
      fill: theme.paleBlue,
      font: { color: theme.text },
      wrapText: true,
      verticalAlignment: "center",
    };
    sub.format.rowHeightPx = 34;
  }
}

function styleHeader(range) {
  range.format = {
    fill: theme.navy,
    font: { bold: true, color: theme.white },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
  };
}

function styleBody(range) {
  range.format = {
    fill: theme.white,
    font: { color: theme.text },
    verticalAlignment: "top",
    wrapText: true,
  };
}

function setWidths(sheet, widths) {
  widths.forEach((width, index) => {
    sheet.getCell(0, index).format.columnWidthPx = width;
  });
}

const main = workbook.worksheets.add("Кандидаты");
main.showGridLines = false;
styleTitle(main, "A1:T1", "Кандидаты на заведение: корейская косметика", "Пробная структура финального файла. Сейчас заполнены выбранные товары Celimax, чтобы проверить формат до добавления остальных брендов.");
setWidths(main, [82, 175, 105, 220, 260, 150, 135, 115, 115, 120, 130, 150, 110, 105, 110, 110, 245, 245, 260, 230]);

main.getRange("A4:D5").values = [
  ["Кандидат", candidate.product, "Статус", candidate.status],
  ["Приоритет", candidate.priority, "Внешний трафик", candidate.externalTraffic],
];
styleBody(main.getRange("A4:D5"));
main.getRange("A4:A5").format = { fill: theme.mint, font: { bold: true, color: theme.text } };
main.getRange("C4:C5").format = { fill: theme.mint, font: { bold: true, color: theme.text } };
main.getRange("D5").format = { fill: theme.green, font: { bold: true, color: theme.text }, wrapText: true };

const mainHeaders = [
  "Приоритет",
  "Статус",
  "Бренд",
  "Конкретный товар",
  "Что искать у поставщика",
  "Ценовой коридор",
  "Max landed-себестоимость",
  "Целевой ДРР",
  "ДРР на тесте",
  "Продажи кластера, шт/мес",
  "Выручка кластера, руб/мес",
  "Ключевая карточка",
  "Продажи SKU, шт/мес",
  "Цена SKU, руб",
  "Возраст, дней",
  "Отзывы",
  "Внешний трафик",
  "Почему перспективно",
  "Условия запуска",
  "Риски",
];
const firstDataRow = 9;
const lastDataRow = firstDataRow + candidateRows.length - 1;
main.getRange("A8:T8").values = [mainHeaders];
main.getRange(`A${firstDataRow}:T${lastDataRow}`).values = candidateRows;
styleHeader(main.getRange("A8:T8"));
styleBody(main.getRange(`A${firstDataRow}:T${lastDataRow}`));
main.getRange(`J${firstDataRow}:K${lastDataRow}`).format.numberFormat = Array.from({ length: candidateRows.length }, () => ["#,##0", "#,##0"]);
main.getRange(`M${firstDataRow}:P${lastDataRow}`).format.numberFormat = Array.from({ length: candidateRows.length }, () => ["#,##0", "#,##0", "#,##0", "#,##0"]);
main.getRange("A1:T1").format.rowHeightPx = 36;
main.getRange("A2:T2").format.rowHeightPx = 34;
main.getRange("A4:D5").format.rowHeightPx = 42;
main.getRange("A8:T8").format.rowHeightPx = 52;
main.getRange(`A${firstDataRow}:T${lastDataRow}`).format.rowHeightPx = 88;
const mainTable = main.tables.add(`A8:T${lastDataRow}`, true, "CandidatesTable");
mainTable.style = "TableStyleMedium2";
main.freezePanes.freezeRows(8);

const passport = workbook.worksheets.add("Паспорт товара");
passport.showGridLines = false;
styleTitle(passport, "A1:H1", "Паспорт товара: Celimax Dual Barrier Mild Gel Cleanser 200 ml", "Сводка по конкретному SKU-кандидату: спрос, конкуренция, внешние источники трафика, экономика и действия перед закупкой.");
setWidths(passport, [230, 235, 210, 210, 210, 210, 210, 210]);

passport.getRange("A4:B12").values = [
  ["Конкретный товар", candidate.product],
  ["Русское название в выдаче", candidate.wbName],
  ["Бренд", candidate.brand],
  ["Статус", candidate.status],
  ["Ключевой пример карточки", `nm ${candidate.sourceSku}`],
  ["Продажи ключевой карточки", `${candidate.sourceSales} шт/мес`],
  ["Выручка ключевой карточки", `${candidate.sourceRevenue.toLocaleString("ru-RU")} руб/мес`],
  ["Возраст / отзывы", `${candidate.sourceAge} дней / ${candidate.sourceReviews} отзывов`],
  ["Вывод", "Брать в проверку закупки, если landed-себестоимость укладывается в 600-650 руб."],
];
styleBody(passport.getRange("A4:B12"));
passport.getRange("A4:A12").format = { fill: theme.mint, font: { bold: true, color: theme.text }, wrapText: true };

passport.getRange("D4:H8").values = [
  ["Показатель", "Значение", "Комментарий", "", ""],
  ["Рынок пенок/гелей Celimax", "8 193 шт/мес", "Релевантные карточки после очистки выдачи от масла/тонера", "", ""],
  ["Видимая выручка рынка", "10,53 млн руб/мес", "По карточной выдаче MarketGuru/WB", "", ""],
  ["Медианная цена", "1 192 руб.", "Рабочий ориентир для цены входа", "", ""],
  ["Медиана отзывов", "80,5", "Барьер входа не выглядит закрытым", "", ""],
];
passport.getRange("D4:H4").merge(true);
styleHeader(passport.getRange("D4:H4"));
styleBody(passport.getRange("D5:H8"));
passport.getRange("D5:D8").format = { fill: theme.paleBlue, font: { bold: true, color: theme.text }, wrapText: true };

passport.getRange("A15:H20").values = [
  ["Внешний трафик", "", "", "", "", "", "", ""],
  ["Рабочий вывод", "Нет признака большой зависимости от внешнего трафика.", "", "", "", "", "", ""],
  ["Основание", "В карточной выдаче сильный пример nm 712024737 имеет среднюю позицию по поисковым запросам 3, 77 поисковых запросов, 1 529 полок ТОП-10 и 4 347 полок ТОП-100. Это похоже на продажи через поиск/органику, а не на товар, который держится только внешним трафиком.", "", "", "", "", "", ""],
  ["Ограничение", "В самой выгрузке нет отдельной колонки с долей внешнего трафика. Если в MarketGuru есть отдельный показатель источников трафика, его стоит выгрузить и подтвердить перед финальным решением.", "", "", "", "", "", ""],
  ["Действие", "Оставить отметку «нет большого количества внешнего трафика», но перед закупкой проверить источник трафика в карточке/MarketGuru, если доступно.", "", "", "", "", "", ""],
  ["", "", "", "", "", "", "", ""],
];
passport.getRange("A15:H15").merge();
passport.getRange("A16:A19").format = { fill: theme.mint, font: { bold: true, color: theme.text }, wrapText: true };
passport.getRange("B16:H19").merge(true);
styleHeader(passport.getRange("A15:H15"));
styleBody(passport.getRange("A16:H19"));
passport.getRange("A1:H1").format.rowHeightPx = 36;
passport.getRange("A2:H2").format.rowHeightPx = 36;
passport.getRange("A15:H15").format.rowHeightPx = 32;
passport.getRange("A16:H19").format.rowHeightPx = 48;

const analog = workbook.worksheets.add("Аналоги в выдаче");
analog.showGridLines = false;
styleTitle(analog, "A1:N1", "Карточки-аналоги по Dual Barrier Mild Gel Cleanser", "Нужны для проверки, что конкретный товар продается у разных продавцов и есть молодые карточки с продажами.");
setWidths(analog, [110, 310, 170, 92, 92, 105, 105, 92, 92, 92, 105, 105, 82, 270]);
const analogHeaders = [
  "Артикул",
  "Товар",
  "Продавец",
  "Цена, руб",
  "Продажи, шт",
  "Выручка, руб",
  "Упущено, руб",
  "Возраст, дней",
  "Отзывы",
  "Новые отзывы",
  "Запросы, шт",
  "Средняя позиция",
  "Фото",
  "Комментарий",
];
analog.getRange("A4:N4").values = [analogHeaders];
analog.getRange(`A5:N${4 + analogs.length}`).values = analogs;
styleHeader(analog.getRange("A4:N4"));
styleBody(analog.getRange(`A5:N${4 + analogs.length}`));
analog.getRange(`D5:L${4 + analogs.length}`).format.numberFormat = Array.from({ length: analogs.length }, () => Array(9).fill("#,##0"));
const analogTable = analog.tables.add(`A4:N${4 + analogs.length}`, true, "AnalogCardsTable");
analogTable.style = "TableStyleMedium9";
analog.freezePanes.freezeRows(4);

const econ = workbook.worksheets.add("Экономика");
econ.showGridLines = false;
styleTitle(econ, "A1:H1", "Экономика входа и ДРР", "Черновой расчет максимальной landed-себестоимости. В финальном файле сюда можно будет подставлять фактическую закупку.");
setWidths(econ, [210, 130, 130, 130, 130, 130, 200, 260]);

econ.getRange("A4:H4").values = [["Сценарий", "Цена WB, руб", "Расходы WB", "ДРР", "Целевая прибыль", "Max landed, руб", "Оценка", "Комментарий"]];
econ.getRange("A5:H8").values = [
  ["Консервативный", 1300, 0.27, 0.12, 0.18, null, "Покупать осторожно", "Подходит для первого теста, если закупка не дороже расчетного потолка."],
  ["Рабочий", 1400, 0.27, 0.12, 0.15, null, "Основной ориентир", "Самый реалистичный сценарий для проверки поставщика."],
  ["Запас на тест рекламы", 1400, 0.27, 0.15, 0.12, null, "Только на тест", "Можно выдержать повышенный ДРР, но прибыль ниже."],
  ["Верх рынка", 1500, 0.27, 0.12, 0.15, null, "Если цена удерживается", "Работает только если карточка держит цену и конверсию."],
];
econ.getRange("F5:F8").formulas = [["=B5*(1-C5-D5-E5)"], ["=B6*(1-C6-D6-E6)"], ["=B7*(1-C7-D7-E7)"], ["=B8*(1-C8-D8-E8)"]];
styleHeader(econ.getRange("A4:H4"));
styleBody(econ.getRange("A5:H8"));
econ.getRange("B5:B8").format.numberFormat = [["#,##0"], ["#,##0"], ["#,##0"], ["#,##0"]];
econ.getRange("C5:E8").format.numberFormat = Array.from({ length: 4 }, () => ["0%", "0%", "0%"]);
econ.getRange("F5:F8").format.numberFormat = [["#,##0"], ["#,##0"], ["#,##0"], ["#,##0"]];
const econTable = econ.tables.add("A4:H8", true, "EconomicsTable");
econTable.style = "TableStyleMedium4";

econ.getRange("A11:H15").values = [
  ["Как читать расчет", "", "", "", "", "", "", ""],
  ["Max landed-себестоимость", "Это максимальная закупка с доставкой до склада/готовностью к продаже, при которой остается целевая прибыль после WB-расходов и рекламы.", "", "", "", "", "", ""],
  ["Расходы WB", "Черновой агрегированный процент на комиссии, логистику и прочие WB-расходы. Для финального решения заменить на фактическую юнитку.", "", "", "", "", "", ""],
  ["ДРР", "Цель 10-12%; на тесте можно до 15%, если цена/маржа выдерживают.", "", "", "", "", "", ""],
  ["Решение", "Если поставщик дает landed до 600-650 руб. и документы по бренду, товар остается в шорт-листе.", "", "", "", "", "", ""],
];
econ.getRange("A11:H11").merge();
econ.getRange("B12:H15").merge(true);
styleHeader(econ.getRange("A11:H11"));
styleBody(econ.getRange("A12:H15"));
econ.getRange("A12:A15").format = { fill: theme.mint, font: { bold: true, color: theme.text }, wrapText: true };

const sources = workbook.worksheets.add("Источники");
sources.showGridLines = false;
styleTitle(sources, "A1:F1", "Источники и допущения", "Лист нужен, чтобы в финальном файле было понятно, откуда взяты цифры и где есть ограничения.");
setWidths(sources, [220, 520, 260, 180, 180, 180]);
sources.getRange("A4:F4").values = [["Блок", "Источник / допущение", "Что использовано", "Период", "Статус", "Комментарий"]];
sources.getRange("A5:F9").values = [
  ["Карточная выдача", candidate.source, "Карточки, продажи, цена, отзывы, возраст, позиции", "28.04.2026-27.05.2026", "Использовано", "Главный источник для выбора конкретного товара."],
  ["Частотность", "2026-04-28—2026-05-27_FBO_Поисковые_запросы (1).xlsx", "Запросы Celimax умывалка / пенка / гель", "28.04.2026-27.05.2026", "Использовано", "Подтверждает спрос по брендовым запросам."],
  ["Внешний трафик", "Рабочий вывод по выдаче и поисковым позициям", "Нет признака большой зависимости от внешнего трафика", "28.04.2026-27.05.2026", "Нужно подтвердить при возможности", "В исходной выгрузке нет отдельной колонки доли внешнего трафика."],
  ["Экономика", "Черновая модель", "WB-расходы 27%, ДРР 10-15%, прибыль 12-18%", "Модель", "Черновик", "Заменить на фактическую юнит-экономику после цены поставщика."],
  ["Текущий магазин", "supplier-goods-127314-2026-05-01-2026-05-28-exvxkgdeb.XLSX", "Проверка, что Celimax cleanser/foam/gel не найден в текущем ассортименте", "01.05.2026-28.05.2026", "Использовано", "В магазине есть Round Lab cleansers, но это другой бренд."],
];
styleHeader(sources.getRange("A4:F4"));
styleBody(sources.getRange("A5:F9"));
const sourcesTable = sources.tables.add("A4:F9", true, "SourcesTable");
sourcesTable.style = "TableStyleMedium2";

analog.getRange("A1:N1").format.rowHeightPx = 36;
analog.getRange("A2:N2").format.rowHeightPx = 36;
analog.getRange("A4:N4").format.rowHeightPx = 42;
analog.getRange(`A5:N${4 + analogs.length}`).format.rowHeightPx = 58;
econ.getRange("A1:H1").format.rowHeightPx = 36;
econ.getRange("A2:H2").format.rowHeightPx = 36;
econ.getRange("A4:H4").format.rowHeightPx = 42;
econ.getRange("A5:H8").format.rowHeightPx = 42;
econ.getRange("A11:H11").format.rowHeightPx = 34;
econ.getRange("A12:H15").format.rowHeightPx = 50;
sources.getRange("A1:F1").format.rowHeightPx = 36;
sources.getRange("A2:F2").format.rowHeightPx = 36;
sources.getRange("A4:F4").format.rowHeightPx = 42;
sources.getRange("A5:F9").format.rowHeightPx = 58;

await fs.mkdir(outputDir, { recursive: true });

const candidatePreview = await workbook.inspect({
  kind: "table",
  range: `Кандидаты!A8:T${lastDataRow}`,
  include: "values,formulas",
  tableMaxRows: 10,
  tableMaxCols: 20,
});
console.log(candidatePreview.ndjson);

const errorScan = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "formula error scan",
});
console.log(errorScan.ndjson);

const render1 = await workbook.render({ sheetName: "Кандидаты", autoCrop: "all", scale: 1, format: "png" });
await fs.writeFile(`${outputDir}/kbeauty_candidates_preview_candidates.png`, new Uint8Array(await render1.arrayBuffer()));
const render2 = await workbook.render({ sheetName: "Экономика", autoCrop: "all", scale: 1, format: "png" });
await fs.writeFile(`${outputDir}/kbeauty_candidates_preview_economics.png`, new Uint8Array(await render2.arrayBuffer()));

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(outputPath);
