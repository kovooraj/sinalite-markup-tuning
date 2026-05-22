import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import { PriceEngineData } from "./parsePriceEngine";
import { computeNewPrice, ScenarioDef, SCENARIO_BY_ID } from "./markupEngine";
import { bandOf } from "./qtyBands";

export interface BuildRepricedOpts {
  pe: PriceEngineData;
  scenarioId: string; // e.g. "A_Current_Locked"
  usdRate: number; // CAD * usdRate = USD (e.g. 0.70)
}

const FMT_CURRENCY = '"$"#,##0.00;("$"#,##0.00);"-"';
const FMT_PCT = "0.0%";
const FMT_INT = "#,##0";

const HEADER_FILL: ExcelJS.FillPattern = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "FF1F3864" },
};
const HIGHLIGHT_FILL: ExcelJS.FillPattern = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "FFFFF2CC" },
};
const HEADER_FONT: Partial<ExcelJS.Font> = {
  name: "Arial",
  bold: true,
  color: { argb: "FFFFFFFF" },
  size: 11,
};
const BODY_FONT: Partial<ExcelJS.Font> = { name: "Arial", size: 10 };

export async function buildRepricedXlsx(opts: BuildRepricedOpts): Promise<Blob> {
  const scenario: ScenarioDef | undefined = SCENARIO_BY_ID[opts.scenarioId];
  if (!scenario) throw new Error(`Unknown scenario: ${opts.scenarioId}`);

  const wb = new ExcelJS.Workbook();
  wb.creator = "SinaLite Markup Tuning";
  wb.created = new Date();
  const ws = wb.addWorksheet("Repriced", {
    views: [{ state: "frozen", xSplit: 3, ySplit: 1 }],
  });

  const usdHeader = `New Price USD (@${opts.usdRate.toFixed(2)})`;

  ws.columns = [
    { header: "Product", key: "product", width: 36 },
    { header: "qty", key: "qty", width: 9 },
    { header: "Turnaround", key: "turnaround", width: 22 },
    { header: "size", key: "size", width: 10 },
    { header: "Stock", key: "stock", width: 32 },
    { header: "Coating", key: "coating", width: 14 },
    { header: "Bundling", key: "bundling", width: 22 },
    { header: "Scoring", key: "scoring", width: 14 },
    { header: "sale price", key: "salePrice", width: 12 },
    { header: "PE3 cost no markup", key: "pe3Cost", width: 16 },
    { header: "Consolidated Markup", key: "consolidatedMarkup", width: 16 },
    { header: "Base Cost (CAD)", key: "baseCost", width: 14 },
    { header: "Finishing Cost (CAD)", key: "finCost", width: 16 },
    { header: "Base Markup %", key: "baseMarkup", width: 12 },
    { header: "Finishing Markup %", key: "finMarkup", width: 14 },
    { header: "Marked-up Base (CAD)", key: "markedBase", width: 16 },
    { header: "Marked-up Finishing (CAD)", key: "markedFin", width: 18 },
    { header: "New Price CAD", key: "newCad", width: 14 },
    { header: usdHeader, key: "newUsd", width: 18 },
    { header: "Delta vs sale price (CAD)", key: "delta", width: 18 },
  ];

  // Apply header styling
  const headerRow = ws.getRow(1);
  headerRow.eachCell((cell) => {
    cell.fill = HEADER_FILL;
    cell.font = HEADER_FONT;
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  });
  headerRow.height = 30;

  // Body rows
  for (const r of opts.pe.rows) {
    const band = bandOf(r.qty);
    const baseMkup = scenario.grad.base[band];
    const finMkup = scenario.grad.fin[band];
    const res = computeNewPrice(
      {
        qty: r.qty,
        baseCost: r.baseCost,
        finCost: r.finCost,
        isRush: r.turnaround === "Rush (NBD)",
        currentSalePrice: 0, // 0 disables cap rule — repricer shows uncapped model price
      },
      scenario.grad
    );
    const newCad = res.uncappedPrice;
    const newUsd = newCad * opts.usdRate;
    const delta = newCad - r.currentSalePrice;
    ws.addRow({
      product: r.productName,
      qty: r.qty,
      turnaround: r.rawTurnaround,
      size: r.size,
      stock: r.stock,
      coating: r.coating,
      bundling: r.bundling,
      scoring: r.scoring,
      salePrice: r.currentSalePrice,
      pe3Cost: r.pe3CostTotal,
      consolidatedMarkup: r.consolidatedMarkup,
      baseCost: round(r.baseCost, 4),
      finCost: round(r.finCost, 4),
      baseMarkup: baseMkup,
      finMarkup: finMkup,
      markedBase: round(res.basePrice, 4),
      markedFin: round(res.finPrice, 4),
      newCad: round(newCad, 2),
      newUsd: round(newUsd, 2),
      delta: round(delta, 2),
    });
  }

  // Body styling — formats, fonts, highlights
  const lastRow = ws.rowCount;
  const fmtMap: Record<string, string> = {
    qty: FMT_INT,
    salePrice: FMT_CURRENCY,
    pe3Cost: FMT_CURRENCY,
    consolidatedMarkup: FMT_PCT,
    baseCost: FMT_CURRENCY,
    finCost: FMT_CURRENCY,
    baseMarkup: FMT_PCT,
    finMarkup: FMT_PCT,
    markedBase: FMT_CURRENCY,
    markedFin: FMT_CURRENCY,
    newCad: FMT_CURRENCY,
    newUsd: FMT_CURRENCY,
    delta: FMT_CURRENCY,
  };
  for (let r = 2; r <= lastRow; r++) {
    const row = ws.getRow(r);
    row.eachCell((cell) => {
      cell.font = BODY_FONT;
    });
    for (const [key, fmt] of Object.entries(fmtMap)) {
      const colNum = ws.getColumn(key).number;
      row.getCell(colNum).numFmt = fmt;
    }
    // Yellow highlight on New Price CAD + USD (R + S)
    row.getCell(ws.getColumn("newCad").number).fill = HIGHLIGHT_FILL;
    row.getCell(ws.getColumn("newUsd").number).fill = HIGHLIGHT_FILL;
  }

  const buf = await wb.xlsx.writeBuffer();
  return new Blob([buf as ArrayBuffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

function round(n: number, digits: number): number {
  const m = Math.pow(10, digits);
  return Math.round(n * m) / m;
}

export async function downloadRepricedXlsx(opts: BuildRepricedOpts): Promise<void> {
  const blob = await buildRepricedXlsx(opts);
  const scenarioSuffix = opts.scenarioId === "A_Current_Locked" ? "" : `_${opts.scenarioId}`;
  saveAs(blob, `${opts.pe.productSlug}_Repriced${scenarioSuffix}.xlsx`);
}
