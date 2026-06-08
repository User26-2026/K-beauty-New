import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "/Users/nikita/Desktop/Новая папка 2/outputs/wb_product_research";
const inputPath = `${outputDir}/roundlab_creamshop_candidates_2026-06-01.json`;
const outputPath = `${outputDir}/roundlab_creamshop_candidates_2026-06-01.xlsx`;

const data = JSON.parse(await fs.readFile(inputPath, "utf8"));

const headers = [
  "Позиция",
  "Артикул WB",
  "Товар",
  "Линейка/тип",
  "Предмет",
  "Цена с СПП, ₽",
  "Расч. заказы",
  "Расч. выручка, ₽",
  "Упущенная выручка, ₽",
  "Наличие",
  "Рейтинг",
  "Оценки",
  "Возраст, дней",
  "Полки TOP-10",
  "Полки TOP-100",
  "Себестоимость из прайса, ₽",
  "Базовая себестоимость, ₽",
  "Источник себестоимости",
  "Артикул себестоимости",
  "Точность сопоставления себестоимости",
  "Статус дубля с K-Beauty",
  "Актуальность закупки",
  "Итог",
  "Комментарий по закупке",
  "Комментарий по дублю",
];

const rows = data.rows.map((row) => [
  row["Позиция"],
  row["Артикул WB"],
  row["Товар"],
  row["Линейка/тип"],
  row["Предмет"],
  row["Цена с СПП, ₽"],
  row["Расч. заказы"],
  row["Расч. выручка, ₽"],
  row["Упущенная выручка, ₽"],
  row["Наличие"],
  row["Рейтинг"],
  row["Оценки"],
  row["Возраст товара, дней"],
  row["Полки TOP-10"],
  row["Полки TOP-100"],
  row["Себестоимость из прайса, ₽"],
  row["Базовая себестоимость, ₽"],
  row["Источник себестоимости"],
  row["Артикул себестоимости"],
  row["Точность сопоставления себестоимости"],
  row["Статус дубля с K-Beauty"],
  row["Актуальность закупки"],
  row["Итог"],
  row["Комментарий по закупке"],
  row["Комментарий по дублю"],
]);

const theme = {
  navy: "#173653",
  white: "#FFFFFF",
  text: "#1F2937",
  paleBlue: "#EEF4FA",
  green: "#DFF3E7",
  yellow: "#FFF3BF",
  orange: "#FFE4CC",
  red: "#F8D7DA",
  gray: "#EEF0F3",
};

const wb = Workbook.create();
wb.comments.setSelf({ displayName: "ChatGPT" });
const sheet = wb.worksheets.add("Round Lab");
sheet.showGridLines = false;

const widths = [80, 115, 330, 170, 165, 110, 110, 135, 135, 90, 80, 90, 105, 105, 110, 125, 125, 310, 120, 210, 190, 160, 145, 360, 360];
widths.forEach((width, index) => {
  sheet.getCell(0, index).format.columnWidthPx = width;
});

sheet.getRange("A1:Y1").merge();
sheet.getRange("A1:Y1").values = [["Round Lab у продавца Cream.Shop: проверка дублей с K-Beauty и актуальность закупки"]];
sheet.getRange("A1:Y1").format = {
  fill: theme.navy,
  font: { bold: true, color: theme.white, size: 14 },
  verticalAlignment: "center",
};
sheet.getRange("A1:Y1").format.rowHeightPx = 36;

sheet.getRange("A2:Y2").merge();
sheet.getRange("A2:Y2").values = [[`Источник: ${data.summary.source_file}. Найдено товаров Round Lab у Cream.Shop: ${data.summary.total_rows}. Себестоимость из юнит-файла проставлена по ${data.summary.cost_matched} позициям, где есть надежное сопоставление. Дубль с K-Beauty помечен отдельно; SPF выделен как сезонный риск на 01.06.`]];
sheet.getRange("A2:Y2").format = {
  fill: theme.paleBlue,
  font: { color: theme.text },
  wrapText: true,
  verticalAlignment: "center",
};
sheet.getRange("A2:Y2").format.rowHeightPx = 42;

sheet.getRange("A4:Y4").values = [headers];
sheet.getRange(`A5:Y${rows.length + 4}`).values = rows;

sheet.getRange("A4:Y4").format = {
  fill: theme.navy,
  font: { bold: true, color: theme.white },
  wrapText: true,
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
sheet.getRange("A4:Y4").format.rowHeightPx = 52;
sheet.getRange(`A5:Y${rows.length + 4}`).format = {
  fill: theme.white,
  font: { color: theme.text },
  wrapText: true,
  verticalAlignment: "top",
};
sheet.getRange(`A5:Y${rows.length + 4}`).format.rowHeightPx = 64;

sheet.getRange(`F5:I${rows.length + 4}`).format.numberFormat = rows.map(() => ["#,##0", "#,##0", "#,##0", "#,##0"]);
sheet.getRange(`K5:O${rows.length + 4}`).format.numberFormat = rows.map(() => ["0.0", "#,##0", "#,##0", "#,##0", "#,##0"]);
sheet.getRange(`P5:Q${rows.length + 4}`).format.numberFormat = rows.map(() => ["#,##0.00", "#,##0.00"]);

rows.forEach((row, idx) => {
  const excelRow = idx + 5;
  const relevance = row[21];
  const duplicate = row[20];
  let fill = theme.white;
  if (duplicate === "Дубль подтвержден") fill = theme.red;
  else if (relevance === "Высокая") fill = theme.green;
  else if (relevance === "Средняя") fill = theme.yellow;
  else if (relevance === "Проверить после сезонного решения") fill = theme.orange;
  else if (relevance === "Низкая") fill = theme.gray;
  sheet.getRange(`A${excelRow}:Y${excelRow}`).format = {
    fill,
    font: { color: theme.text },
    wrapText: true,
    verticalAlignment: "top",
  };
});

sheet.tables.add(`A4:Y${rows.length + 4}`, true, "RoundLabCandidatesTable").style = "TableStyleMedium2";
sheet.freezePanes.freezeRows(4);

await fs.mkdir(outputDir, { recursive: true });

const preview = await wb.inspect({
  kind: "table",
  range: "Round Lab!A1:Y12",
  include: "values,formulas",
  tableMaxRows: 12,
  tableMaxCols: 25,
});
console.log(preview.ndjson);

const errors = await wb.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "formula error scan",
});
console.log(errors.ndjson);

const render = await wb.render({ sheetName: "Round Lab", range: "A1:Y14", scale: 1, format: "png" });
await fs.writeFile(`${outputDir}/roundlab_creamshop_candidates_2026-06-01_preview.png`, new Uint8Array(await render.arrayBuffer()));

const output = await SpreadsheetFile.exportXlsx(wb);
await output.save(outputPath);
console.log(outputPath);
