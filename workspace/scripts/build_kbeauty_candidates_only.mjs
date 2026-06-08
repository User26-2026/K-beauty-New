import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "/Users/nikita/Desktop/Новая папка 2/outputs/wb_product_research";
const outputPath = `${outputDir}/kbeauty_candidates_oil_pads_creams_2026-05-29.xlsx`;

const theme = {
  navy: "#173653",
  paleBlue: "#EEF4FA",
  text: "#1F2937",
  white: "#FFFFFF",
};

const rows = [
  [
    "",
    "Celimax",
    "Dual Barrier Mild Gel Cleanser 200 ml",
    "Гель для умывания барьерный с церамидами",
    "1 300-1 500 руб.",
    8193,
    10534635,
    193401,
    "Карточка MarketGuru / скрин",
    "nm 712024737",
    359,
    1315,
    132,
    295,
    "По косвенным признакам нет большой зависимости от внешнего трафика",
    "Молодая карточка уже продает 359 шт/мес; сильный аналог дает 444 шт/мес; товар не найден в текущем ассортименте магазина.",
  ],
  [
    "",
    "Celimax",
    "Noni Acne Bubble Cleanser",
    "Пенка для умывания пузырьковая с нони от акне",
    "1 050-1 250 руб.",
    8193,
    10534635,
    60146,
    "Карточка MarketGuru / скрин",
    "nm 712022766",
    182,
    1138,
    121,
    120,
    "По косвенным признакам нет большой зависимости от внешнего трафика",
    "Молодая карточка: 121 день, продажи +333%, 88 поисковых запросов, средняя позиция 4, 1 316 полок ТОП-10. По карточке товара MarketGuru видно 60 146 руб. упущенной выручки.",
  ],
  [
    "",
    "Celimax",
    "Derma Nature Relief Madecica pH Balancing Foam Cleansing 150 ml",
    "Балансирующая пенка для умывания Madecica pH",
    "1 200-1 400 руб.",
    8193,
    10534635,
    0,
    "Выгрузка FBO: Упущенная выручка",
    "nm 155947140",
    323,
    1387,
    1103,
    2564,
    "Нет отдельной колонки по внешнему трафику в выгрузке",
    "Продажи есть: 323 шт/мес и 449 761 руб., но карточка старая, с 2 564 отзывами и падающей динамикой продаж -34%.",
  ],
  [
    "",
    "Celimax",
    "Dual Barrier Mild Gel Cleanser, аналог nm 228473128",
    "Гель-пенка с комплексом церамидов Dual Barrier Mild",
    "1 300-1 500 руб.",
    8193,
    10534635,
    0,
    "Выгрузка FBO: Упущенная выручка",
    "nm 228473128",
    82,
    1659,
    736,
    339,
    "Нет отдельной колонки по внешнему трафику в выгрузке",
    "Использовать как подтверждающий аналог к Dual Barrier, но не дублировать отдельной закупочной строкой.",
  ],
  [
    "",
    "Celimax",
    "Fresh Blackhead Jojoba Cleansing Oil 150 ml",
    "Гидрофильное масло для умывания с экстрактом жожоба 150 мл",
    "1 100-1 400 руб.",
    3759,
    4023078,
    0,
    "Выгрузка FBO: Упущенная выручка",
    "nm 187281601",
    522,
    1283,
    938,
    729,
    "По косвенным признакам нет большой зависимости от внешнего трафика",
    "Один из лидеров выдачи: 522 продажи/мес, средняя позиция по запросам 2, 895 полок ТОП-10 и 2 162 полки ТОП-100. Товар не найден в текущем ассортименте магазина.",
  ],
  [
    "",
    "Celimax",
    "Fresh Blackhead Jojoba Cleansing Oil 20 ml",
    "Гидрофильное масло для умывания с экстрактом жожоба 20 мл",
    "350-450 руб.",
    3759,
    4023078,
    0,
    "Выгрузка FBO: Упущенная выручка",
    "nm 394475962",
    547,
    380,
    386,
    729,
    "По косвенным признакам нет большой зависимости от внешнего трафика",
    "Продает 547 шт/мес, средняя позиция 3. Лучше рассматривать как мини-формат, набор или допродажу, а не как основной оборотный SKU из-за низкой цены.",
  ],
  [
    "",
    "Celimax",
    "Ji Woo Gae Cica BHA Blemish Toner Pad 60 шт",
    "Успокаивающие пэды с кислотами Cica BHA / Blemish Toner Pad",
    "1 100-1 550 руб.",
    16157,
    19472960,
    404528,
    "Карточка MarketGuru / скрин",
    "nm 263730548",
    986,
    1544,
    590,
    10174,
    "По косвенным признакам нет большой зависимости от внешнего трафика",
    "Подходит к закупу: товарная ниша уже сформирована, лидер выдачи продает 986 шт/мес. По карточке товара MarketGuru видно 404 528 руб. упущенной выручки. Молодой ориентир nm 611417712 дополнительно подтверждает вход: 275 шт/мес, 92 отзыва, рост +1000%.",
  ],
  [
    "",
    "Celimax",
    "Pore+Dark Spot Brightening Pad 40 шт",
    "Осветляющие тонер-пэды Pore+Dark Spot Brightening Pad",
    "1 100-1 500 руб.",
    16157,
    19472960,
    386443,
    "Общая выдача: пэды Celimax",
    "nm 240014011",
    1104,
    1400,
    637,
    10174,
    "По косвенным признакам нет большой зависимости от внешнего трафика",
    "Подходит к закупу: осветляющие пэды закрывают понятный запрос пигментации и ровного тона. Лидер продает 1 104 шт/мес. Молодой ориентир nm 712016848 продает 170 шт/мес при росте +608%.",
  ],
  [
    "",
    "Celimax",
    "The Real Noni Calming Glow Pad 50 шт",
    "Успокаивающие тонер-пэды с нони 50 шт",
    "1 450-1 700 руб.",
    16157,
    19472960,
    23246,
    "Выгрузка FBO: Упущенная выручка",
    "nm 583311001",
    355,
    1579,
    184,
    1445,
    "Нет отдельной колонки по внешнему трафику в выгрузке",
    "Новая для текущего магазина линия: 355 шт/мес, рост продаж +86%, цена 1 579 руб. Барьер отзывов выше, поэтому заводить только после проверки себестоимости и документов.",
  ],
  [
    "",
    "Celimax",
    "Ji Woo Gae Heartleaf BHA Peeling Pad 60 шт",
    "Пэды Heartleaf BHA / мягкий пилинг",
    "1 050-1 250 руб.",
    16157,
    19472960,
    43252,
    "Общая выдача: пэды Celimax",
    "nm 712019009",
    164,
    1103,
    132,
    193,
    "Нет отдельной колонки по внешнему трафику в выгрузке",
    "Молодая карточка продает 164 шт/мес при 193 отзывах, но динамика продаж -43%. Не первая волна, оставить как резервный вариант.",
  ],
  [
    "",
    "Celimax",
    "The Real Noni Energy Repair Cream 50 ml",
    "Восстанавливающий крем для лица с нони 50 мл",
    "1 200-1 600 руб.",
    3242,
    4419293,
    0,
    "Выгрузка FBO: Упущенная выручка",
    "nm 758360096",
    103,
    1271,
    111,
    57,
    "Нет отдельной колонки по внешнему трафику в выгрузке",
    "Подходит к проверке закупки: молодая карточка 111 дней, 103 продажи/мес, 57 отзывов, средняя позиция 13. Ниша Noni-кремов в выдаче дает 3 242 продажи/мес и 4,42 млн руб. выручки. Не выглядит прямым дублем текущих Celimax-кремов магазина.",
  ],
  [
    "",
    "Celimax",
    "Oil Control Moisturizing Cream 80 ml",
    "Себорегулирующий крем для жирной кожи / Oil Control",
    "1 300-1 750 руб.",
    973,
    1465555,
    0,
    "Выгрузка FBO: Упущенная выручка",
    "nm 953742609",
    113,
    1308,
    40,
    5362,
    "Нет отдельной колонки по внешнему трафику в выгрузке",
    "Интересен как отдельная ниша для жирной кожи: молодая карточка 40 дней уже продает 113 шт/мес, рост продаж +495%. Лидер nm 240483875 продает 196 шт/мес с ростом +96%. Проверить себестоимость и барьер отзывов.",
  ],
  [
    "",
    "Celimax",
    "The Real Cica Soothing Cream 50 ml",
    "Успокаивающий крем с центеллой / The Real Cica",
    "1 150-1 350 руб.",
    646,
    812900,
    0,
    "Выгрузка FBO: Упущенная выручка",
    "nm 775939662",
    78,
    1197,
    108,
    17,
    "Нет отдельной колонки по внешнему трафику в выгрузке",
    "Резервный вариант: карточка молодая, 108 дней, 78 продаж/мес, рост +420%, всего 17 отзывов. Общий кластер Cica-кремов меньше, чем Noni/Dual Barrier, поэтому не первая волна.",
  ],
];

const wb = Workbook.create();
wb.comments.setSelf({ displayName: "ChatGPT" });

const sheet = wb.worksheets.add("Кандидаты");
sheet.showGridLines = false;

const widths = [150, 105, 250, 260, 145, 125, 135, 150, 210, 140, 120, 105, 105, 105, 260, 360];
widths.forEach((width, index) => {
  sheet.getCell(0, index).format.columnWidthPx = width;
});

sheet.getRange("A1:P1").merge();
sheet.getRange("A1:P1").values = [["Кандидаты на заведение: корейская косметика"]];
sheet.getRange("A1:P1").format = {
  fill: theme.navy,
  font: { bold: true, color: theme.white },
  verticalAlignment: "center",
};
sheet.getRange("A1:P1").format.rowHeightPx = 36;

sheet.getRange("A2:P2").merge();
sheet.getRange("A2:P2").values = [["Только основная таблица кандидатов. ДРР не рассчитан до получения фактической себестоимости, комиссии, логистики, хранения и прочих расходов WB. По упущенной выручке отдельно указан источник, потому что карточка товара и разные выгрузки MarketGuru могут показывать разные значения."]];
sheet.getRange("A2:P2").format = {
  fill: theme.paleBlue,
  font: { color: theme.text },
  wrapText: true,
  verticalAlignment: "center",
};
sheet.getRange("A2:P2").format.rowHeightPx = 44;

const headers = [
  "Решение",
  "Бренд",
  "Конкретный товар",
  "Что искать у поставщика",
  "Ценовой коридор WB",
  "Продажи кластера, шт/мес",
  "Выручка кластера, руб/мес",
  "Упущенная выручка по карточке, руб/мес",
  "Источник упущенной выручки",
  "Карточка-ориентир",
  "Продажи SKU, шт/мес",
  "Цена SKU, руб",
  "Возраст, дней",
  "Отзывы",
  "Внешний трафик",
  "Обоснование",
];

sheet.getRange("A4:P4").values = [headers];
const lastDataRow = 4 + rows.length;
sheet.getRange(`A5:P${lastDataRow}`).values = rows;

sheet.getRange("A4:P4").format = {
  fill: theme.navy,
  font: { bold: true, color: theme.white },
  wrapText: true,
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
sheet.getRange(`A5:P${lastDataRow}`).format = {
  fill: theme.white,
  font: { color: theme.text },
  wrapText: true,
  verticalAlignment: "top",
};

sheet.getRange(`F5:H${lastDataRow}`).format.numberFormat = Array.from({ length: rows.length }, () => ["#,##0", "#,##0", "#,##0"]);
sheet.getRange(`K5:N${lastDataRow}`).format.numberFormat = Array.from({ length: rows.length }, () => ["#,##0", "#,##0", "#,##0", "#,##0"]);
sheet.getRange("A4:P4").format.rowHeightPx = 48;
sheet.getRange(`A5:P${lastDataRow}`).format.rowHeightPx = 74;

sheet.tables.add(`A4:P${lastDataRow}`, true, "CandidatesOnlyTable").style = "TableStyleMedium2";
sheet.freezePanes.freezeRows(4);

await fs.mkdir(outputDir, { recursive: true });

const preview = await wb.inspect({
  kind: "table",
  range: `Кандидаты!A4:P${lastDataRow}`,
  include: "values,formulas",
  tableMaxRows: 12,
  tableMaxCols: 16,
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
await fs.writeFile(`${outputDir}/kbeauty_candidates_with_oil_pads_candidates.png`, new Uint8Array(await render.arrayBuffer()));

const output = await SpreadsheetFile.exportXlsx(wb);
await output.save(outputPath);
console.log(outputPath);
