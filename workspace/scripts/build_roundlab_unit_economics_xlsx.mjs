import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "/Users/nikita/Desktop/Новая папка 2/outputs/wb_product_research";
const inputPath = `${outputDir}/roundlab_unit_economics_from_celimax_2026-06-01.json`;
const outputPath = `${outputDir}/roundlab_unit_economics_from_celimax_2026-06-01.xlsx`;

const data = JSON.parse(await fs.readFile(inputPath, "utf8"));
const rows = data.rows;

const headers = [
  "Артикул WB",
  "Наименование",
  "Кол-во",
  "ЗАКУП руб.",
  "ИТОГО себес.",
  "Итого прибыль",
  "Себ-ть",
  "Цена продажи на WB",
  "РРЦ плановая без СПП",
  "Кол-во к закупу",
  "% СПП",
  "Цена с СПП",
  "Доп обработка склад",
  "Налог УСН",
  "Стоимость денег",
  "Объем, литр",
  "Логистика WB",
  "Комиссия WB 32,5%",
  "Эквайринг 2,17%",
  "Обратная логистика",
  "% выкупа",
  "Хранение",
  "Рекламные расходы",
  "Брак",
  "Платная приемка",
  "МАРЖА с ед.",
  "ROI %",
  "Проходной ДРР",
  "Расч. выручка",
  "Упущенная выручка",
  "Комментарий",
];

const tableRows = rows.map((r) => [
  r["Артикул WB"],
  r["Товар"],
  1,
  r["Себ-ть"] ?? "",
  r["ИТОГО себес."] ?? "",
  r["Итого прибыль"] ?? "",
  r["Себ-ть"] ?? "",
  r["Цена продажи на WB"] ?? "",
  r["РРЦ плановая без СПП"] ?? "",
  r["Расч. заказы"] ?? "",
  r["% СПП"] ?? "",
  r["Цена с СПП расчетная"] ?? r["Цена с СПП, ₽"] ?? "",
  r["Доп обработка склад"] ?? "",
  r["Налог УСН"] ?? "",
  r["Стоимость денег"] ?? "",
  r["Объем, литр"] ?? "",
  r["Логистика WB"] ?? "",
  r["Комиссия WB"] ?? "",
  r["Эквайринг"] ?? "",
  r["Обратная логистика"] ?? "",
  r["% выкупа"] ?? "",
  r["Хранение"] ?? "",
  r["Рекламные расходы"] ?? "",
  r["Брак"] ?? "",
  r["Платная приемка"] ?? "",
  r["МАРЖА с ед."] ?? "",
  r["ROI %"] ?? "",
  r["Проходной ДРР от WB цены"] ?? "",
  r["Расч. выручка, ₽"] ?? "",
  r["Упущенная выручка, ₽"] ?? "",
  r["Комментарий"] ?? "",
]);

const theme = {
  navy: "#173653",
  white: "#FFFFFF",
  text: "#1F2937",
  paleBlue: "#EEF4FA",
  paleGreen: "#DFF3E7",
  paleRed: "#F8D7DA",
  paleYellow: "#FFF3BF",
  gray: "#EEF0F3",
  border: "#D8E1EA",
};

const wb = Workbook.create();
wb.comments.setSelf({ displayName: "ChatGPT" });
const sheet = wb.worksheets.add("Юнитка Round Lab");
sheet.showGridLines = false;

const lastCol = "AE";
const widths = [
  115, 360, 75, 95, 115, 115, 90, 120, 120, 105, 80, 105, 115, 105, 105, 95,
  95, 120, 110, 115, 90, 85, 115, 85, 110, 110, 85, 105, 120, 135, 360,
];
widths.forEach((width, index) => {
  sheet.getCell(0, index).format.columnWidthPx = width;
});

sheet.getRange(`A1:${lastCol}1`).merge();
sheet.getRange(`A1:${lastCol}1`).values = [["Юнит-экономика Round Lab по логике шаблона Celimax.xlsx"]];
sheet.getRange(`A1:${lastCol}1`).format = {
  fill: theme.navy,
  font: { bold: true, color: theme.white, size: 14 },
  verticalAlignment: "center",
};
sheet.getRange(`A1:${lastCol}1`).format.rowHeightPx = 36;

const c = data.summary.constants;
sheet.getRange(`A2:${lastCol}2`).merge();
sheet.getRange(`A2:${lastCol}2`).values = [[
  `Источник: ${data.summary.source}. Расчетные параметры: СПП 30%, комиссия WB 32,5%, эквайринг 2,17%, УСН 8%, логистика WB 90 руб., обработка 60 руб., хранение 4,8 руб., реклама 10%. Себестоимость взята из прайса Round Lab в рублях после коэффициента 1,4.`,
]];
sheet.getRange(`A2:${lastCol}2`).format = {
  fill: theme.paleBlue,
  font: { color: theme.text },
  wrapText: true,
  verticalAlignment: "center",
};
sheet.getRange(`A2:${lastCol}2`).format.rowHeightPx = 54;

sheet.getRange(`A4:${lastCol}4`).values = [headers];
sheet.getRange(`A5:${lastCol}${tableRows.length + 4}`).values = tableRows;

sheet.getRange(`A4:${lastCol}4`).format = {
  fill: theme.navy,
  font: { bold: true, color: theme.white },
  wrapText: true,
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
sheet.getRange(`A4:${lastCol}4`).format.rowHeightPx = 54;

sheet.getRange(`A5:${lastCol}${tableRows.length + 4}`).format = {
  fill: theme.white,
  font: { color: theme.text },
  wrapText: true,
  verticalAlignment: "top",
};
sheet.getRange(`A5:${lastCol}${tableRows.length + 4}`).format.rowHeightPx = 62;

const numberRanges = [
  ["C", "J", 8],
  ["L", "T", 9],
  ["V", "Z", 5],
  ["AC", "AD", 2],
];
for (const [start, end, width] of numberRanges) {
  sheet.getRange(`${start}5:${end}${tableRows.length + 4}`).format.numberFormat =
    Array.from({ length: tableRows.length }, () => Array.from({ length: width }, () => "#,##0"));
}
sheet.getRange(`K5:K${tableRows.length + 4}`).format.numberFormat = Array.from({ length: tableRows.length }, () => ["0%"]);
sheet.getRange(`U5:U${tableRows.length + 4}`).format.numberFormat = Array.from({ length: tableRows.length }, () => ["0.00"]);
sheet.getRange(`AA5:AB${tableRows.length + 4}`).format.numberFormat = Array.from({ length: tableRows.length }, () => ["0%", "0%"]);

rows.forEach((r, idx) => {
  const excelRow = idx + 5;
  let fill = theme.white;
  if (r["МАРЖА с ед."] == null) fill = theme.gray;
  else if (r["МАРЖА с ед."] < 0) fill = theme.paleRed;
  else if (r["ROI %"] >= 0.5) fill = theme.paleGreen;
  else fill = theme.paleYellow;
  sheet.getRange(`A${excelRow}:${lastCol}${excelRow}`).format = {
    fill,
    font: { color: theme.text },
    wrapText: true,
    verticalAlignment: "top",
  };
});

sheet.tables.add(`A4:${lastCol}${tableRows.length + 4}`, true, "RoundLabUnitEconomics").style = "TableStyleMedium2";
sheet.freezePanes.freezeRows(4);

const summary = wb.worksheets.add("Итог");
summary.showGridLines = false;
summary.getCell(0, 0).format.columnWidthPx = 320;
summary.getCell(0, 1).format.columnWidthPx = 180;
summary.getCell(0, 2).format.columnWidthPx = 360;

summary.getRange("A1:C1").merge();
summary.getRange("A1:C1").values = [["Итог по расчету"]];
summary.getRange("A1:C1").format = {
  fill: theme.navy,
  font: { bold: true, color: theme.white, size: 14 },
  verticalAlignment: "center",
};
summary.getRange("A1:C1").format.rowHeightPx = 36;

const summaryRows = [
  ["Товаров в списке", data.summary.rows_total, ""],
  ["Посчитано с себестоимостью", data.summary.rows_calculated, ""],
  ["Нет себестоимости", data.summary.rows_missing_cost, "Нужен прайс/точное сопоставление"],
  ["Суммарная себестоимость закупки по расчетным заказам", data.summary.total_cost, ""],
  ["Суммарная прибыль при рекламе 10%", data.summary.total_profit, "По текущей логике шаблона"],
  ["Средняя маржа на единицу", data.summary.avg_margin, ""],
];
summary.getRange("A3:C8").values = summaryRows;
summary.getRange("A3:C8").format = {
  font: { color: theme.text },
  wrapText: true,
  verticalAlignment: "center",
};
summary.getRange("A3:C8").format.rowHeightPx = 34;
summary.getRange("B4:B8").format.numberFormat = Array.from({ length: 5 }, () => ["#,##0"]);

await fs.mkdir(outputDir, { recursive: true });

const preview = await wb.inspect({
  kind: "table",
  range: `Юнитка Round Lab!A1:${lastCol}12`,
  include: "values,formulas",
  tableMaxRows: 12,
  tableMaxCols: 31,
});
console.log(preview.ndjson);

const errors = await wb.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "formula error scan",
});
console.log(errors.ndjson);

const render = await wb.render({ sheetName: "Юнитка Round Lab", range: `A1:${lastCol}16`, scale: 1, format: "png" });
await fs.writeFile(`${outputDir}/roundlab_unit_economics_from_celimax_2026-06-01_preview.png`, new Uint8Array(await render.arrayBuffer()));

const output = await SpreadsheetFile.exportXlsx(wb);
await output.save(outputPath);
console.log(outputPath);
