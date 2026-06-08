import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const ROOT = "/Users/nikita/Desktop/Новая папка 2";
const OUT = path.join(ROOT, "outputs/wb_audit");
const SKU_CSV = path.join(OUT, "supply_recommendations_by_sku_summary_2026-05-28.csv");
const REGION_CSV = path.join(OUT, "supply_recommendations_by_sku_region_2026-05-28.csv");
const MD_REPORT = path.join(OUT, "stock_distribution_supply_report_2026-05-28.md");
const OUTPUT_XLSX = path.join(OUT, "stock_distribution_supply_report_2026-05-28.xlsx");

const COLORS = {
  navy: "#1F2937",
  slate: "#475569",
  blue: "#2563EB",
  green: "#059669",
  amber: "#D97706",
  red: "#DC2626",
  paleBlue: "#EAF2FF",
  paleGreen: "#E9F8F2",
  paleAmber: "#FFF7E6",
  paleRed: "#FEECEC",
  border: "#D9E2EC",
  header: "#EEF4FF",
  white: "#FFFFFF",
};

function parseCsv(text) {
  const rows = [];
  let row = [];
  let cell = "";
  let inQuotes = false;
  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i];
    const next = text[i + 1];
    if (inQuotes) {
      if (ch === '"' && next === '"') {
        cell += '"';
        i += 1;
      } else if (ch === '"') {
        inQuotes = false;
      } else {
        cell += ch;
      }
    } else if (ch === '"') {
      inQuotes = true;
    } else if (ch === ",") {
      row.push(cell);
      cell = "";
    } else if (ch === "\n") {
      row.push(cell);
      rows.push(row);
      row = [];
      cell = "";
    } else if (ch !== "\r") {
      cell += ch;
    }
  }
  if (cell.length || row.length) {
    row.push(cell);
    rows.push(row);
  }
  const headers = rows.shift();
  return rows
    .filter((r) => r.some((v) => String(v).trim() !== ""))
    .map((r) => Object.fromEntries(headers.map((h, idx) => [h, r[idx] ?? ""])));
}

function n(v) {
  if (v === null || v === undefined || v === "") return 0;
  const num = Number(String(v).replace(/\s/g, "").replace(",", "."));
  return Number.isFinite(num) ? num : 0;
}

function money(v) {
  return Math.round(n(v));
}

function fmt(v) {
  return String(Math.round(n(v))).replace(/\B(?=(\d{3})+(?!\d))/g, " ");
}

function shortText(value, max = 54) {
  const s = String(value ?? "");
  return s.length > max ? `${s.slice(0, max - 1)}…` : s;
}

function firstRegions(value, max = 3) {
  return String(value ?? "")
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean)
    .slice(0, max)
    .join("; ");
}

function pct(v) {
  return n(v) / 100;
}

function colName(index) {
  let s = "";
  let n1 = index;
  while (n1 > 0) {
    const m = (n1 - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n1 = Math.floor((n1 - m) / 26);
  }
  return s;
}

function setValues(sheet, startRow, startCol, values) {
  const rowCount = values.length;
  const colCount = values[0]?.length ?? 0;
  const endRow = startRow + rowCount - 1;
  const endCol = startCol + colCount - 1;
  const range = `${colName(startCol)}${startRow}:${colName(endCol)}${endRow}`;
  sheet.getRange(range).values = values;
  return sheet.getRange(range);
}

function styleHeader(range) {
  range.format.fill.color = COLORS.header;
  range.format.font.bold = true;
  range.format.font.color = COLORS.navy;
  range.format.font.size = 11;
}

function styleTitle(range, size = 20) {
  range.format.font.bold = true;
  range.format.font.size = size;
  range.format.font.color = COLORS.navy;
}

function styleSubtle(range) {
  range.format.font.color = COLORS.slate;
  range.format.font.size = 10;
}

function styleKpi(range, fill) {
  range.format.fill.color = fill;
  range.format.font.color = COLORS.navy;
}

function setWidths(sheet, widths) {
  widths.forEach((width, idx) => {
    sheet.getRange(`${colName(idx + 1)}:${colName(idx + 1)}`).columnWidthPx = width;
  });
}

function normalizeSku(rows) {
  return rows.map((r) => ({
    nm: String(r.nm || ""),
    sellerArticle: r["Артикул продавца"] || "",
    name: r["Название_sku"] || "",
    subject: r["Предмет_sku"] || "",
    brand: r["Бренд_sku"] || "",
    qty30: money(r.recommended_supply_qty),
    qty14: money(r.first_wave_14d_qty),
    regions: r.regions || "",
    nonlocalOrders: money(r.nonlocal_orders),
    currentStock: money(r.current_wb_stock_regions),
    profitRecommended: money(r.profit_if_recommended_qty_sells),
    irpSavingMonth: money(r.allocated_irp_saving_month),
    profitPerUnit: r.profit_per_unit ? Math.round(n(r.profit_per_unit)) : null,
    missedProfitSku: money(r.missed_profit_sku),
    missedBuyoutsSku: n(r.missed_buyouts_sku),
    missedBuyoutsAmountSku: money(r.missed_buyouts_amount_sku),
    irpSavingSkuMonth: money(r.irp_saving_sku_month),
    missingCost: !r.profit_per_unit,
  }));
}

function normalizeRegion(rows) {
  return rows.map((r) => ({
    nm: String(r["Артикул WB"] || ""),
    sellerArticle: r["Артикул продавца"] || "",
    name: r["Название"] || "",
    subject: r["Предмет"] || "",
    brand: r["Бренд"] || "",
    region: r["Регион"] || "",
    orders: money(r["Итого заказов, шт"]),
    localOrders: money(r["Итого заказов по товарам локально, шт"]),
    nonlocalOrders: money(r["Итого заказов по товарам не локально, шт"]),
    localPct: pct(r["Локальность_региона_%"]),
    currentStock: money(r["Остатки склад ВБ, шт"]),
    targetStock: money(r["Целевой_остаток_30дней_с_буфером"]),
    qty30: money(r["Рекомендовано_поставить_шт"]),
    qty14: money(r["Первая_волна_14дней_шт"]),
    profitPerUnit: r["Прибыль_на_шт_до_рекламы"] ? Math.round(n(r["Прибыль_на_шт_до_рекламы"])) : null,
    profitRecommended: money(r["Прибыль_если_рекомендованный_объем_продастся"]),
    lostProfit: money(r["Потенциал_прибыли_от_возврата_спроса"]),
    irpSavingMonth: money(r["Потенциал_экономии_ИРП_мес"]),
    deliveryTime: r["Время доставки"] || "",
    missingCost: !r["Прибыль_на_шт_до_рекламы"],
  }));
}

function groupByRegion(regionRows) {
  const map = new Map();
  for (const r of regionRows) {
    if (!map.has(r.region)) {
      map.set(r.region, {
        region: r.region,
        qty30: 0,
        qty14: 0,
        profitRecommended: 0,
        irpSavingMonth: 0,
        nonlocalOrders: 0,
        currentStock: 0,
        skuCount: new Set(),
      });
    }
    const g = map.get(r.region);
    g.qty30 += r.qty30;
    g.qty14 += r.qty14;
    g.profitRecommended += r.profitRecommended;
    g.irpSavingMonth += r.irpSavingMonth;
    g.nonlocalOrders += r.nonlocalOrders;
    g.currentStock += r.currentStock;
    g.skuCount.add(r.nm);
  }
  return [...map.values()]
    .map((g) => ({ ...g, skuCount: g.skuCount.size }))
    .sort((a, b) => b.qty30 - a.qty30);
}

function writeTable(sheet, row, col, headers, rows, options = {}) {
  const values = [headers, ...rows];
  const range = setValues(sheet, row, col, values);
  styleHeader(sheet.getRange(`${colName(col)}${row}:${colName(col + headers.length - 1)}${row}`));
  const dataRange = sheet.getRange(`${colName(col)}${row + 1}:${colName(col + headers.length - 1)}${row + rows.length}`);
  dataRange.format.font.size = options.fontSize ?? 10;
  return range;
}

async function main() {
  const [skuText, regionText, mdText] = await Promise.all([
    fs.readFile(SKU_CSV, "utf8"),
    fs.readFile(REGION_CSV, "utf8"),
    fs.readFile(MD_REPORT, "utf8"),
  ]);

  const skuRows = normalizeSku(parseCsv(skuText));
  const regionRows = normalizeRegion(parseCsv(regionText));
  const regionSummary = groupByRegion(regionRows);
  const topSku = skuRows.slice(0, 15);
  const topRegion = regionRows.slice(0, 30);
  const missingCost = skuRows.filter((r) => r.missingCost && (r.qty30 > 0 || r.missedBuyoutsAmountSku > 0));

  const totalQty30 = regionRows.reduce((s, r) => s + r.qty30, 0);
  const totalQty14 = regionRows.reduce((s, r) => s + r.qty14, 0);
  const totalProfit = regionRows.reduce((s, r) => s + r.profitRecommended, 0);
  const totalIrp = regionRows.reduce((s, r) => s + r.irpSavingMonth, 0);
  const topKnownProfit = skuRows.filter((r) => !r.missingCost).slice(0, 12);

  const workbook = Workbook.create();
  const dash = workbook.worksheets.add("Дашборд");
  const skuSheet = workbook.worksheets.add("SKU план");
  const regionSheet = workbook.worksheets.add("Регионы");
  const regionSummarySheet = workbook.worksheets.add("Свод регионов");
  const missingSheet = workbook.worksheets.add("Нет себестоимости");
  const methodSheet = workbook.worksheets.add("Методика");

  setWidths(dash, [170, 120, 120, 170, 120, 120, 170, 120, 120, 150, 150, 150, 150, 150, 150, 150]);
  setValues(dash, 1, 1, [["Распределение остатков WB: куда поставлять и сколько"]]);
  styleTitle(dash.getRange("A1"), 18);
  dash.getRange("1:1").rowHeightPx = 34;
  setValues(dash, 2, 1, [["K-Beauty Hub, период источников: остатки/ИЛ 28.04-28.05.2026, экономика 01.04-24.05.2026"]]);
  styleSubtle(dash.getRange("A2"));

  const kpis = [
    ["Поставить на 30 дней", totalQty30, "шт", COLORS.paleBlue],
    ["Первая волна на 14 дней", totalQty14, "шт", COLORS.paleGreen],
    ["Потенциальная прибыль", totalProfit, "₽ до рекламы", COLORS.paleAmber],
    ["Экономия ИРП", totalIrp, "₽/мес", COLORS.paleGreen],
    ["SKU в плане", skuRows.length, "шт", COLORS.paleBlue],
    ["Нет себестоимости", missingCost.length, "SKU", COLORS.paleRed],
  ];
  const kpiPositions = [
    ["A4:C6", "A4:C4", "A5:C5", "A6:C6"],
    ["D4:F6", "D4:F4", "D5:F5", "D6:F6"],
    ["G4:I6", "G4:I4", "G5:I5", "G6:I6"],
    ["A8:C10", "A8:C8", "A9:C9", "A10:C10"],
    ["D8:F10", "D8:F8", "D9:F9", "D10:F10"],
    ["G8:I10", "G8:I8", "G9:I9", "G10:I10"],
  ];
  kpis.forEach((k, idx) => {
    const [box, label, value, unit] = kpiPositions[idx];
    styleKpi(dash.getRange(box), k[3]);
    dash.getRange(label).values = [[k[0], "", ""]];
    dash.getRange(value).values = [[fmt(k[1]), "", ""]];
    dash.getRange(unit).values = [[k[2], "", ""]];
    dash.getRange(label).format.font.bold = true;
    dash.getRange(value).format.font.size = 20;
    dash.getRange(value).format.font.bold = true;
    dash.getRange(unit).format.font.color = COLORS.slate;
  });

  const chartData1 = topKnownProfit.slice(0, 10);
  dash.charts.add("bar", {
    title: "Топ SKU по прибыли",
    categories: chartData1.map((r) => `${r.nm} ${r.brand}`),
    series: [{ name: "Прибыль, ₽", values: chartData1.map((r) => r.profitRecommended) }],
    hasLegend: false,
    from: { row: 12, col: 1 },
    extent: { widthPx: 680, heightPx: 330 },
    dataLabels: { showValue: false },
    barOptions: { direction: "bar", grouping: "clustered", gapWidth: 70 },
  });

  const chartData2 = topSku.slice(0, 10);
  dash.charts.add("bar", {
    title: "Топ SKU по объему",
    categories: chartData2.map((r) => `${r.nm} ${r.brand}`),
    series: [
      { name: "30 дней, шт", values: chartData2.map((r) => r.qty30) },
      { name: "14 дней, шт", values: chartData2.map((r) => r.qty14) },
    ],
    hasLegend: true,
    legend: { position: "top" },
    from: { row: 12, col: 7 },
    extent: { widthPx: 680, heightPx: 330 },
    barOptions: { direction: "bar", grouping: "clustered", gapWidth: 75 },
  });

  const chartData3 = regionSummary.slice(0, 10);
  dash.charts.add("ColumnClustered", {
    title: "Куда везти: объем по регионам",
    categories: chartData3.map((r) => r.region),
    series: [{ name: "30 дней, шт", values: chartData3.map((r) => r.qty30) }],
    hasLegend: false,
    from: { row: 31, col: 1 },
    extent: { widthPx: 680, heightPx: 300 },
    barOptions: { direction: "column", grouping: "clustered", gapWidth: 100 },
  });

  dash.charts.add("ColumnClustered", {
    title: "Потенциальная экономия ИРП по регионам",
    categories: chartData3.map((r) => r.region),
    series: [{ name: "ИРП, ₽/мес", values: chartData3.map((r) => r.irpSavingMonth) }],
    hasLegend: false,
    from: { row: 31, col: 7 },
    extent: { widthPx: 680, heightPx: 300 },
    barOptions: { direction: "column", grouping: "clustered", gapWidth: 100 },
  });

  setValues(dash, 49, 1, [["Короткий вывод"]]);
  styleTitle(dash.getRange("A49"), 14);
  const conclusion = [
    ["1", "Основной резерв роста не только в закупке товара, а в правильной региональной раскладке остатков."],
    ["2", "Сначала закрыть дефицит по прибыльным SKU и регионам с нелокальным спросом, затем масштабировать рекламу."],
    ["3", "Для регулярной работы внедрить MPStats / расчет поставки и сверять его с юнит-экономикой."],
  ];
  writeTable(dash, 50, 1, ["№", "Вывод"], conclusion, { fontSize: 11 });
  dash.getRange("A50:B53").format.fill.color = COLORS.white;

  setWidths(skuSheet, [100, 230, 360, 120, 120, 100, 100, 360, 110, 110, 110, 130, 110, 110, 120, 110]);
  setValues(skuSheet, 1, 1, [["SKU план поставок"]]);
  styleTitle(skuSheet.getRange("A1"), 18);
  const skuHeaders = [
    "Артикул WB",
    "Артикул продавца",
    "Товар",
    "Предмет",
    "Бренд",
    "30 дней, шт",
    "14 дней, шт",
    "Куда в первую очередь",
    "Нелокальные",
    "Текущий остаток",
    "Прибыль/шт",
    "Прибыль объема",
    "ИРП/мес",
    "Упущ. выкупы",
    "Упущ. ₽",
    "Нет себест.",
  ];
  const skuTable = skuRows.map((r) => [
    r.nm,
    shortText(r.sellerArticle, 34),
    shortText(r.name, 46),
    r.subject,
    r.brand,
    r.qty30,
    r.qty14,
    firstRegions(r.regions, 4),
    r.nonlocalOrders,
    r.currentStock,
    r.profitPerUnit,
    r.profitRecommended,
    r.irpSavingMonth,
    r.missedBuyoutsSku,
    r.missedBuyoutsAmountSku,
    r.missingCost ? "Да" : "",
  ]);
  writeTable(skuSheet, 3, 1, skuHeaders, skuTable);
  skuSheet.getRange(`A3:P${skuTable.length + 3}`).format.wrapText = false;
  skuSheet.getRange(`A3:P3`).format.wrapText = true;
  for (let row = 4; row <= Math.min(skuTable.length + 3, 80); row += 1) {
    skuSheet.getRange(`${row}:${row}`).rowHeightPx = 30;
  }

  setWidths(regionSheet, [95, 220, 340, 115, 115, 220, 95, 95, 100, 105, 105, 105, 105, 105, 110, 130, 115, 105, 105]);
  setValues(regionSheet, 1, 1, [["Рекомендации по SKU и регионам"]]);
  styleTitle(regionSheet.getRange("A1"), 18);
  const regionHeaders = [
    "Артикул WB",
    "Артикул продавца",
    "Товар",
    "Предмет",
    "Бренд",
    "Регион",
    "Заказы",
    "Локальные",
    "Нелокальные",
    "Локальность",
    "Остаток WB",
    "Цель 30д",
    "30 дней, шт",
    "14 дней, шт",
    "Прибыль/шт",
    "Прибыль объема",
    "ИРП/мес",
    "Доставка",
    "Нет себест.",
  ];
  const regionTable = regionRows.map((r) => [
    r.nm,
    shortText(r.sellerArticle, 34),
    shortText(r.name, 56),
    r.subject,
    r.brand,
    r.region,
    r.orders,
    r.localOrders,
    r.nonlocalOrders,
    r.localPct,
    r.currentStock,
    r.targetStock,
    r.qty30,
    r.qty14,
    r.profitPerUnit,
    r.profitRecommended,
    r.irpSavingMonth,
    r.deliveryTime,
    r.missingCost ? "Да" : "",
  ]);
  writeTable(regionSheet, 3, 1, regionHeaders, regionTable, { fontSize: 9 });
  regionSheet.getRange(`A3:S${regionTable.length + 3}`).format.wrapText = false;
  regionSheet.getRange(`A3:S3`).format.wrapText = true;
  for (let row = 4; row <= Math.min(regionTable.length + 3, 120); row += 1) {
    regionSheet.getRange(`${row}:${row}`).rowHeightPx = 28;
  }

  setWidths(regionSummarySheet, [260, 120, 120, 130, 130, 140, 130, 90, 40, 140, 140, 140, 140, 140]);
  setValues(regionSummarySheet, 1, 1, [["Свод по регионам"]]);
  styleTitle(regionSummarySheet.getRange("A1"), 18);
  const regionSummaryTable = regionSummary.map((r) => [
    r.region,
    r.qty30,
    r.qty14,
    r.nonlocalOrders,
    r.currentStock,
    r.profitRecommended,
    r.irpSavingMonth,
    r.skuCount,
  ]);
  writeTable(
    regionSummarySheet,
    3,
    1,
    ["Регион", "30 дней, шт", "14 дней, шт", "Нелокальные", "Остаток WB", "Прибыль объема", "ИРП/мес", "SKU"],
    regionSummaryTable,
  );
  regionSummarySheet.charts.add("ColumnClustered", {
    title: "Объем поставки по регионам",
    categories: regionSummary.slice(0, 10).map((r) => r.region),
    series: [{ name: "30 дней, шт", values: regionSummary.slice(0, 10).map((r) => r.qty30) }],
    hasLegend: false,
    from: { row: 3, col: 10 },
    extent: { widthPx: 680, heightPx: 300 },
  });

  setWidths(missingSheet, [100, 240, 360, 120, 120, 100, 100, 120, 130, 360]);
  setValues(missingSheet, 1, 1, [["SKU, по которым нужно дозаполнить себестоимость"]]);
  styleTitle(missingSheet.getRange("A1"), 18);
  const missingTable = missingCost.map((r) => [
    r.nm,
    r.sellerArticle,
    r.name,
    r.subject,
    r.brand,
    r.qty30,
    r.qty14,
    r.missedBuyoutsSku,
    r.missedBuyoutsAmountSku,
    r.regions,
  ]);
  writeTable(
    missingSheet,
    3,
    1,
    ["Артикул WB", "Артикул продавца", "Товар", "Предмет", "Бренд", "30 дней", "14 дней", "Упущ. выкупы", "Упущ. ₽", "Регионы"],
    missingTable,
  );
  missingSheet.getRange(`A3:J${missingTable.length + 3}`).format.wrapText = true;

  setWidths(methodSheet, [180, 900]);
  setValues(methodSheet, 1, 1, [["Методика и ограничения"]]);
  styleTitle(methodSheet.getRange("A1"), 18);
  const bullets = mdText
    .split("\n")
    .filter((line) => line.startsWith("- ") || line.startsWith("Главная задача"))
    .slice(0, 18)
    .map((line) => ["", line.replace(/^- /, "")]);
  writeTable(methodSheet, 3, 1, ["Блок", "Описание"], [
    ["Источник", "Остатки и ИЛ за 28.04.2026-28.05.2026; экономика за 01.04.2026-24.05.2026."],
    ["Цель", "Показать, какие SKU и регионы нужно дозагрузить, чтобы вернуть спрос, поднять локальность и снизить ИРП."],
    ["Важно", "Упущенные выкупы WB очищены от дублей на уровне SKU; складские строки используются для приоритизации, а не для механического суммирования."],
    ...bullets,
  ]);
  methodSheet.getRange("A3:B24").format.wrapText = true;

  // Basic number formats.
  for (const sheet of [skuSheet, regionSheet, regionSummarySheet, missingSheet]) {
    sheet.getRange("A:Z").format.font.size = 10;
  }
  skuSheet.getRange(`K4:O${skuTable.length + 3}`).numberFormat = [["#,##0"]];
  regionSheet.getRange(`J4:J${regionTable.length + 3}`).numberFormat = [["0%"]];
  regionSheet.getRange(`O4:Q${regionTable.length + 3}`).numberFormat = [["#,##0"]];
  regionSummarySheet.getRange(`B4:G${regionSummaryTable.length + 3}`).numberFormat = [["#,##0"]];

  // Conditional highlighting for missing costs and low localization.
  for (let i = 0; i < regionRows.length; i += 1) {
    const row = i + 4;
    if (regionRows[i].localPct < 0.2) {
      regionSheet.getRange(`J${row}`).format.fill.color = COLORS.paleRed;
      regionSheet.getRange(`J${row}`).format.font.color = COLORS.red;
    } else if (regionRows[i].localPct < 0.6) {
      regionSheet.getRange(`J${row}`).format.fill.color = COLORS.paleAmber;
      regionSheet.getRange(`J${row}`).format.font.color = COLORS.amber;
    }
    if (regionRows[i].missingCost) {
      regionSheet.getRange(`S${row}`).format.fill.color = COLORS.paleRed;
    }
  }
  for (let i = 0; i < skuRows.length; i += 1) {
    const row = i + 4;
    if (skuRows[i].missingCost) {
      skuSheet.getRange(`P${row}`).format.fill.color = COLORS.paleRed;
    }
  }

  // Freeze panes / filters are intentionally omitted: the workbook is built for static review and source tables stay editable.
  const inspect = await workbook.inspect({
    kind: "table",
    range: "Дашборд!A1:I10",
    include: "values,formulas",
    tableMaxRows: 12,
    tableMaxCols: 10,
  });
  console.log(inspect.ndjson);
  const errors = await workbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
    options: { useRegex: true, maxResults: 50 },
    summary: "formula errors",
  });
  console.log(errors.ndjson);

  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(OUTPUT_XLSX);
  console.log(OUTPUT_XLSX);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
