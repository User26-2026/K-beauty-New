import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "/Users/nikita/Desktop/Новая папка 2/outputs/wb_product_research";
const dataPath = `${outputDir}/wb_card_comparison_template_data_2026-05-29.json`;
const outputPath = `${outputDir}/wb_card_comparison_google_template_2026-05-29.xlsx`;

const source = JSON.parse(await fs.readFile(dataPath, "utf8"));

const theme = {
  navy: "#173653",
  label: "#D9DCE5",
  header: "#D9DCE5",
  grid: "#A5AAB7",
  green: "#65F246",
  yellow: "#FFFF45",
  light: "#F3F6FA",
  white: "#FFFFFF",
  text: "#111827",
};

const metrics = [
  "Название",
  "Категория",
  "Предмет",
  "Дата создания",
  "Бренд",
  "Рейтинг карточки",
  "Рейтинг по отзывам",
  "Количество отзывов",
  "Минимальная цена со скидкой (по размерам), ₽",
  "Максимальная цена со скидкой (по размерам), ₽",
  "Медианная цена покупателя, ₽",
  "Среднее время доставки",
  "Средняя позиция",
  "Статус продвижения",
  "Показы",
  "Переходы в карточку",
  "CTR",
  "Положили в корзину",
  "Конверсия в корзину, %",
  "Заказали, шт за месяц",
  "Заказали, шт за квартал",
  "Заказали, шт за прошлый период",
  "Динамика заказов к прошлому периоду, %",
  "Конверсия в заказ, %",
  "Выкупили, шт",
  "Процент выкупа, %",
  "Отмены, шт",
  "Период анализа",
  "Предыдущий период",
];

function columnName(zeroBasedIndex) {
  let n = zeroBasedIndex + 1;
  let name = "";
  while (n > 0) {
    const mod = (n - 1) % 26;
    name = String.fromCharCode(65 + mod) + name;
    n = Math.floor((n - mod) / 26);
  }
  return name;
}

function productValue(product, metric) {
  const sourceMetric =
    metric === "Переходы в карточку"
      ? "Переходы в карточку"
      : metric === "CTR"
        ? "CTR, %"
        : metric;
  return product.values[sourceMetric] ?? "";
}

function rowsForProducts(products, blank = false) {
  const rows = [];
  rows.push(["Показатели", ...products.map((p, index) => (blank ? `Артикул WB ${index + 1}` : `Артикул WB ${p.nm}`))]);

  for (const metric of metrics) {
    if (metric === "Период анализа") {
      rows.push([metric, ...products.map(() => (blank ? "" : source.period))]);
      continue;
    }
    if (metric === "Предыдущий период") {
      rows.push([metric, ...products.map(() => (blank ? "" : source.previous_period))]);
      continue;
    }
    if (metric === "Заказали, шт за квартал") {
      rows.push([metric, ...products.map(() => "")]);
      continue;
    }
    if (metric === "Динамика заказов к прошлому периоду, %") {
      rows.push([metric, ...products.map(() => "")]);
      continue;
    }
    rows.push([metric, ...products.map((p) => (blank ? "" : productValue(p, metric)))]);
  }

  return rows;
}

function applyTemplate(sheet, rows, withFormulas = true) {
  sheet.showGridLines = false;
  const productCount = rows[0].length - 1;
  const lastCol = columnName(productCount);
  const lastRow = rows.length;

  sheet.getCell(0, 0).format.columnWidthPx = 360;
  for (let i = 1; i <= productCount; i += 1) {
    sheet.getCell(0, i).format.columnWidthPx = 190;
  }

  sheet.getRange(`A1:${lastCol}${lastRow}`).values = rows;

  if (withFormulas) {
    const currentOrdersRow = metrics.indexOf("Заказали, шт за месяц") + 2;
    const previousOrdersRow = metrics.indexOf("Заказали, шт за прошлый период") + 2;
    const dynamicsRow = metrics.indexOf("Динамика заказов к прошлому периоду, %") + 2;
    const formulas = [];
    for (let i = 1; i <= productCount; i += 1) {
      const col = columnName(i);
      formulas.push(`=IFERROR((${col}${currentOrdersRow}-${col}${previousOrdersRow})/${col}${previousOrdersRow},"")`);
    }
    sheet.getRange(`B${dynamicsRow}:${lastCol}${dynamicsRow}`).formulas = [formulas];
  }

  sheet.getRange(`A1:${lastCol}1`).format = {
    fill: theme.header,
    font: { bold: true, color: "#000000", size: 12 },
    wrapText: true,
    verticalAlignment: "bottom",
  };
  sheet.getRange(`A1:${lastCol}1`).format.rowHeightPx = 46;

  sheet.getRange(`A2:A${lastRow}`).format = {
    fill: theme.label,
    font: { bold: true, color: "#000000" },
    wrapText: true,
    verticalAlignment: "center",
  };
  sheet.getRange(`B2:${lastCol}${lastRow}`).format = {
    fill: theme.white,
    font: { color: theme.text },
    wrapText: true,
    verticalAlignment: "center",
  };

  sheet.getRange(`A2:${lastCol}2`).format.rowHeightPx = 88;
  sheet.getRange(`A3:${lastCol}${lastRow}`).format.rowHeightPx = 28;

  const medianRow = metrics.indexOf("Медианная цена покупателя, ₽") + 2;
  const monthOrdersRow = metrics.indexOf("Заказали, шт за месяц") + 2;
  const quarterOrdersRow = metrics.indexOf("Заказали, шт за квартал") + 2;
  const periodsStartRow = metrics.indexOf("Период анализа") + 2;
  const dynamicsRow = metrics.indexOf("Динамика заказов к прошлому периоду, %") + 2;

  sheet.getRange(`B${medianRow}:${lastCol}${medianRow}`).format = {
    fill: theme.green,
    font: { color: theme.text },
    wrapText: true,
    verticalAlignment: "center",
  };
  sheet.getRange(`B${monthOrdersRow}:${lastCol}${monthOrdersRow}`).format = {
    fill: theme.yellow,
    font: { color: theme.text },
    wrapText: true,
    verticalAlignment: "center",
  };
  sheet.getRange(`B${quarterOrdersRow}:${lastCol}${quarterOrdersRow}`).format = {
    fill: theme.green,
    font: { color: theme.text },
    wrapText: true,
    verticalAlignment: "center",
  };
  sheet.getRange(`A${periodsStartRow}:${lastCol}${lastRow}`).format = {
    fill: theme.light,
    font: { color: theme.text },
    wrapText: true,
    verticalAlignment: "center",
  };

  sheet.getRange(`B${dynamicsRow}:${lastCol}${dynamicsRow}`).format.numberFormat = [
    Array.from({ length: productCount }, () => "0%"),
  ];

  const numericRows = [9, 10, 11, 12, 14, 16, 17, 18, 19, 20, 21, 23, 25, 26, 27, 28];
  for (const row of numericRows) {
    sheet.getRange(`B${row}:${lastCol}${row}`).format.numberFormat = [
      Array.from({ length: productCount }, () => "#,##0"),
    ];
  }

  sheet.freezePanes.freezeRows(1);
  sheet.freezePanes.freezeColumns(1);
  sheet.tables.add(`A1:${lastCol}${lastRow}`, true, `${sheet.name.replace(/\W/g, "")}Table`).style = "TableStyleLight1";
}

const wb = Workbook.create();
wb.comments.setSelf({ displayName: "ChatGPT" });

const dataSheet = wb.worksheets.add("Пенка Celimax");
applyTemplate(dataSheet, rowsForProducts(source.products), true);

const blankProducts = Array.from({ length: 8 }, (_, index) => ({
  nm: `${index + 1}`,
  values: {},
}));
const templateSheet = wb.worksheets.add("Шаблон");
applyTemplate(templateSheet, rowsForProducts(blankProducts, true), true);

await fs.mkdir(outputDir, { recursive: true });

const preview = await wb.inspect({
  kind: "table",
  range: "Пенка Celimax!A1:F29",
  include: "values,formulas",
  tableMaxRows: 29,
  tableMaxCols: 6,
});
console.log(preview.ndjson);

const errors = await wb.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "formula error scan",
});
console.log(errors.ndjson);

const renderData = await wb.render({ sheetName: "Пенка Celimax", range: "A1:F29", scale: 1, format: "png" });
await fs.writeFile(`${outputDir}/wb_card_comparison_google_template_preview.png`, new Uint8Array(await renderData.arrayBuffer()));
const renderTemplate = await wb.render({ sheetName: "Шаблон", range: "A1:I29", scale: 1, format: "png" });
await fs.writeFile(`${outputDir}/wb_card_comparison_google_template_blank_preview.png`, new Uint8Array(await renderTemplate.arrayBuffer()));

const output = await SpreadsheetFile.exportXlsx(wb);
await output.save(outputPath);
console.log(outputPath);
