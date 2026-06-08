import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "/Users/nikita/Desktop/Новая папка 2/outputs/wb_product_research";
const inputPath = `${outputDir}/roundlab_creamshop_candidates_2026-06-01.json`;
const outputPath = `${outputDir}/roundlab_candidates_no_dupes_no_spf_2026-06-01.xlsx`;

const data = JSON.parse(await fs.readFile(inputPath, "utf8"));

const rowsClean = data.rows.filter((row) => {
  const line = String(row["Линейка/тип"] || "").toLowerCase();
  const name = String(row["Товар"] || "").toLowerCase();
  const duplicate = String(row["Статус дубля с K-Beauty"] || "").toLowerCase();
  const isSpf = line.includes("spf") || name.includes("солнцезащит") || name.includes("спф") || name.includes("spf");
  const isDuplicate = duplicate.includes("дубль");
  return !isSpf && !isDuplicate;
});

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
  "Источник себестоимости",
  "Точность сопоставления себестоимости",
  "Проверка дубля с K-Beauty",
  "Актуальность закупки",
  "Итог",
  "Комментарий по закупке",
];

const rows = rowsClean.map((row) => [
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
  row["Источник себестоимости"],
  row["Точность сопоставления себестоимости"],
  row["Статус дубля с K-Beauty"],
  row["Актуальность закупки"],
  row["Итог"],
  row["Комментарий по закупке"],
]);

const theme = {
  navy: "#173653",
  white: "#FFFFFF",
  text: "#1F2937",
  paleBlue: "#EEF4FA",
  green: "#DFF3E7",
  yellow: "#FFF3BF",
  gray: "#EEF0F3",
};

const wb = Workbook.create();
wb.comments.setSelf({ displayName: "ChatGPT" });
const sheet = wb.worksheets.add("Кандидаты");
sheet.showGridLines = false;

const lastCol = "V";
const widths = [80, 115, 330, 170, 165, 110, 110, 135, 135, 90, 80, 90, 105, 105, 110, 125, 310, 210, 190, 160, 145, 380];
widths.forEach((width, index) => {
  sheet.getCell(0, index).format.columnWidthPx = width;
});

sheet.getRange(`A1:${lastCol}1`).merge();
sheet.getRange(`A1:${lastCol}1`).values = [["Round Lab: кандидаты без дублей K-Beauty и без солнцезащитных товаров"]];
sheet.getRange(`A1:${lastCol}1`).format = {
  fill: theme.navy,
  font: { bold: true, color: theme.white, size: 14 },
  verticalAlignment: "center",
};
sheet.getRange(`A1:${lastCol}1`).format.rowHeightPx = 36;

sheet.getRange(`A2:${lastCol}2`).merge();
sheet.getRange(`A2:${lastCol}2`).values = [[`Источник: ${data.summary.source_file}. В итоговый список включены только товары Round Lab продавца Cream.Shop без подтвержденных/вероятных дублей K-Beauty и без SPF. Осталось кандидатов: ${rows.length}.`]];
sheet.getRange(`A2:${lastCol}2`).format = {
  fill: theme.paleBlue,
  font: { color: theme.text },
  wrapText: true,
  verticalAlignment: "center",
};
sheet.getRange(`A2:${lastCol}2`).format.rowHeightPx = 48;

sheet.getRange(`A4:${lastCol}4`).values = [headers];
sheet.getRange(`A5:${lastCol}${rows.length + 4}`).values = rows;

sheet.getRange(`A4:${lastCol}4`).format = {
  fill: theme.navy,
  font: { bold: true, color: theme.white },
  wrapText: true,
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
sheet.getRange(`A4:${lastCol}4`).format.rowHeightPx = 52;

rows.forEach((row, idx) => {
  const excelRow = idx + 5;
  const relevance = row[19];
  let fill = theme.white;
  if (relevance === "Высокая") fill = theme.green;
  else if (relevance === "Средняя") fill = theme.yellow;
  else if (relevance === "Низкая") fill = theme.gray;
  sheet.getRange(`A${excelRow}:${lastCol}${excelRow}`).format = {
    fill,
    font: { color: theme.text },
    wrapText: true,
    verticalAlignment: "top",
  };
  sheet.getRange(`A${excelRow}:${lastCol}${excelRow}`).format.rowHeightPx = 64;
});

sheet.getRange(`F5:I${rows.length + 4}`).format.numberFormat = rows.map(() => ["#,##0", "#,##0", "#,##0", "#,##0"]);
sheet.getRange(`K5:O${rows.length + 4}`).format.numberFormat = rows.map(() => ["0.0", "#,##0", "#,##0", "#,##0", "#,##0"]);
sheet.getRange(`P5:P${rows.length + 4}`).format.numberFormat = rows.map(() => ["#,##0.00"]);

sheet.tables.add(`A4:${lastCol}${rows.length + 4}`, true, "RoundLabCandidatesClean").style = "TableStyleMedium2";
sheet.freezePanes.freezeRows(4);

await fs.mkdir(outputDir, { recursive: true });

const preview = await wb.inspect({
  kind: "table",
  range: `Кандидаты!A1:${lastCol}12`,
  include: "values,formulas",
  tableMaxRows: 12,
  tableMaxCols: 22,
});
console.log(preview.ndjson);

const errors = await wb.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "formula error scan",
});
console.log(errors.ndjson);

const render = await wb.render({ sheetName: "Кандидаты", range: `A1:${lastCol}14`, scale: 1, format: "png" });
await fs.writeFile(`${outputDir}/roundlab_candidates_no_dupes_no_spf_2026-06-01_preview.png`, new Uint8Array(await render.arrayBuffer()));

const output = await SpreadsheetFile.exportXlsx(wb);
await output.save(outputPath);
console.log(outputPath);
