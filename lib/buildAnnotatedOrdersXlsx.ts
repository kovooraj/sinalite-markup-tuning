import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import { OrderReplayData, orderRowKey } from "./parseOrderReplay";
import {
  PriceEngineData,
  indexPriceEngine,
  pickByPriceProximity,
} from "./parsePriceEngine";
import { commonDimensions, useStockInKey } from "./computeScenarios";
import { computeNewPrice, SCENARIO_BY_ID, ScenarioDef } from "./markupEngine";
import { bandOf, nbdBandOf } from "./qtyBands";

export interface BuildAnnotatedOrdersOpts {
  pe: PriceEngineData;
  order: OrderReplayData;
  scenarioId: string;
  usdRate: number;
  productSlug: string;
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
const TOTAL_FILL: ExcelJS.FillPattern = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "FFEFEFEF" },
};
const HEADER_FONT: Partial<ExcelJS.Font> = {
  name: "Arial",
  bold: true,
  color: { argb: "FFFFFFFF" },
  size: 11,
};
const BODY_FONT: Partial<ExcelJS.Font> = { name: "Arial", size: 10 };

const DIM_DISPLAY: Record<string, string> = {
  coating: "Coating",
  bundling: "Bundling",
  scoring: "Scoring",
  cover: "Cover",
  binding: "Binding",
  pages: "Pages",
  sides: "Sides",
  finishing: "Finishing",
  lamination: "Lamination",
  foil: "Foil",
  embossing: "Embossing",
  diecutting: "Die Cutting",
  corner: "Corner",
  perforation: "Perforation",
  drilling: "Drilling",
};

function round(n: number, digits: number): number {
  const m = Math.pow(10, digits);
  return Math.round(n * m) / m;
}

export async function buildAnnotatedOrdersXlsx(
  opts: BuildAnnotatedOrdersOpts
): Promise<Blob> {
  const scenario: ScenarioDef | undefined = SCENARIO_BY_ID[opts.scenarioId];
  if (!scenario) throw new Error(`Unknown scenario: ${opts.scenarioId}`);

  const common = commonDimensions(opts.pe, opts.order);
  const useStock = useStockInKey(opts.pe, opts.order);
  const { buckets } = indexPriceEngine(opts.pe, common, useStock);

  const wb = new ExcelJS.Workbook();
  wb.creator = "SinaLite Markup Tuning";
  wb.created = new Date();
  const ws = wb.addWorksheet(`Annotated_${opts.scenarioId}`, {
    views: [{ state: "frozen", xSplit: 4, ySplit: 1 }],
  });

  const usdHeader = `New Price USD (@${opts.usdRate.toFixed(2)})`;

  // Dynamic columns: original-order identifying fields + product dims +
  // computed markup chain + delta math
  type Col = { header: string; key: string; width: number; fmt?: string };
  const cols: Col[] = [
    { header: "Description", key: "description", width: 36 },
    { header: "Qty", key: "qty", width: 9, fmt: FMT_INT },
    { header: "Stock", key: "stock", width: 28 },
    { header: "Size", key: "size", width: 12 },
    { header: "Turnaround", key: "turnaround", width: 18 },
  ];
  // Order-side dimensions are typically a subset of pe.dimensions; show
  // whichever are present in either file so the row is fully identified.
  const allDims = Array.from(
    new Set([...opts.pe.dimensions, ...opts.order.dimensions])
  ).sort();
  for (const dim of allDims) {
    cols.push({
      header: DIM_DISPLAY[dim] ?? dim,
      key: `dim_${dim}`,
      width: 16,
    });
  }
  cols.push(
    { header: "Orders", key: "orders", width: 8, fmt: FMT_INT },
    { header: "Avg Paid CAD", key: "avgPaid", width: 12, fmt: FMT_CURRENCY },
    { header: "PE3 Base Cost", key: "baseCost", width: 13, fmt: FMT_CURRENCY },
    { header: "PE3 Fin Cost", key: "finCost", width: 12, fmt: FMT_CURRENCY },
    {
      header: "Current Sale Price",
      key: "currentSalePrice",
      width: 16,
      fmt: FMT_CURRENCY,
    },
    { header: "Base Mkup %", key: "baseMkup", width: 11, fmt: FMT_PCT },
    { header: "Fin Mkup %", key: "finMkup", width: 11, fmt: FMT_PCT },
    { header: "NBD Mkup %", key: "nbdMkup", width: 11, fmt: FMT_PCT },
    {
      header: "Marked-up Base",
      key: "markedBase",
      width: 14,
      fmt: FMT_CURRENCY,
    },
    {
      header: "Marked-up Fin",
      key: "markedFin",
      width: 14,
      fmt: FMT_CURRENCY,
    },
    {
      header: "Subtotal Pre-NBD",
      key: "subtotalPreNbd",
      width: 16,
      fmt: FMT_CURRENCY,
    },
    {
      header: "Uncapped New Price",
      key: "uncappedNew",
      width: 18,
      fmt: FMT_CURRENCY,
    },
    { header: "Capped?", key: "capped", width: 9 },
    { header: "New Price CAD", key: "newCad", width: 14, fmt: FMT_CURRENCY },
    { header: usdHeader, key: "newUsd", width: 18, fmt: FMT_CURRENCY },
    {
      header: "Delta per Order",
      key: "deltaPerOrder",
      width: 14,
      fmt: FMT_CURRENCY,
    },
    {
      header: "Total Delta (CAD)",
      key: "totalDelta",
      width: 16,
      fmt: FMT_CURRENCY,
    },
    {
      header: "Annualized Delta",
      key: "annualizedDelta",
      width: 16,
      fmt: FMT_CURRENCY,
    }
  );

  ws.columns = cols.map(({ header, key, width }) => ({ header, key, width }));

  const headerRow = ws.getRow(1);
  headerRow.eachCell((cell) => {
    cell.fill = HEADER_FILL;
    cell.font = HEADER_FONT;
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  });
  headerRow.height = 30;

  let aggregateDelta = 0;
  let aggregatePaid = 0;
  let aggregateNew = 0;
  let matchedRowCount = 0;
  let unmatchedRowCount = 0;

  for (const r of opts.order.rows) {
    const matchKey = orderRowKey(r, common, useStock);
    const pe = pickByPriceProximity(buckets.get(matchKey), r.avgPaid);

    const baseRow: Record<string, string | number> = {
      description: r.description,
      qty: r.qty,
      stock: r.stock || "—",
      size: r.size,
      turnaround: r.turnaround,
      orders: r.orders,
      avgPaid: round(r.avgPaid, 2),
    };
    for (const dim of allDims) {
      baseRow[`dim_${dim}`] = r.dims[dim] ?? "";
    }

    if (!pe) {
      // Unmatched — emit row with empty markup math so the user can see
      // which orders couldn't be matched against the catalog
      ws.addRow({
        ...baseRow,
        baseCost: 0,
        finCost: 0,
        currentSalePrice: 0,
        baseMkup: 0,
        finMkup: 0,
        nbdMkup: 0,
        markedBase: 0,
        markedFin: 0,
        subtotalPreNbd: 0,
        uncappedNew: 0,
        capped: "no match",
        newCad: 0,
        newUsd: 0,
        deltaPerOrder: 0,
        totalDelta: 0,
        annualizedDelta: 0,
      });
      unmatchedRowCount += 1;
      continue;
    }

    const band = bandOf(r.qty);
    const baseMkup = scenario.grad.base[band];
    const finMkup = scenario.grad.fin[band];
    const isRush = r.turnaround === "Rush (NBD)";
    const nbdMkup = isRush ? scenario.grad.nbd[nbdBandOf(r.qty)] : 0;
    const res = computeNewPrice(
      {
        qty: r.qty,
        baseCost: pe.baseCost,
        finCost: pe.finCost,
        isRush,
        currentSalePrice: pe.currentSalePrice,
      },
      scenario.grad
    );
    const subtotalPreNbd = res.basePrice + res.finPrice;
    const deltaPerOrder = res.finalPrice - r.avgPaid;
    const totalDelta = deltaPerOrder * r.orders;
    const annualized = totalDelta * 4;

    aggregateDelta += totalDelta;
    aggregatePaid += r.avgPaid * r.orders;
    aggregateNew += res.finalPrice * r.orders;
    matchedRowCount += 1;

    ws.addRow({
      ...baseRow,
      baseCost: round(pe.baseCost, 4),
      finCost: round(pe.finCost, 4),
      currentSalePrice: round(pe.currentSalePrice, 2),
      baseMkup,
      finMkup,
      nbdMkup,
      markedBase: round(res.basePrice, 4),
      markedFin: round(res.finPrice, 4),
      subtotalPreNbd: round(subtotalPreNbd, 2),
      uncappedNew: round(res.uncappedPrice, 2),
      capped: res.capped ? "yes" : "no",
      newCad: round(res.finalPrice, 2),
      newUsd: round(res.finalPrice * opts.usdRate, 2),
      deltaPerOrder: round(deltaPerOrder, 2),
      totalDelta: round(totalDelta, 2),
      annualizedDelta: round(annualized, 2),
    });
  }

  // Apply body styling + highlight the New Price + Annualized Delta columns
  const lastBodyRow = ws.rowCount;
  for (let r = 2; r <= lastBodyRow; r++) {
    const row = ws.getRow(r);
    row.eachCell((cell) => {
      cell.font = BODY_FONT;
    });
    for (const c of cols) {
      if (!c.fmt) continue;
      const colNum = ws.getColumn(c.key).number;
      row.getCell(colNum).numFmt = c.fmt;
    }
    row.getCell(ws.getColumn("newCad").number).fill = HIGHLIGHT_FILL;
    row.getCell(ws.getColumn("annualizedDelta").number).fill = HIGHLIGHT_FILL;
  }

  // TOTAL row
  const totalRow: Record<string, string | number> = {
    description: `TOTAL (${matchedRowCount} matched · ${unmatchedRowCount} unmatched)`,
    qty: "",
    stock: "",
    size: "",
    turnaround: "",
    orders: opts.order.totals.orders,
    avgPaid: round(aggregatePaid / Math.max(opts.order.totals.orders, 1), 2),
    baseCost: 0,
    finCost: 0,
    currentSalePrice: round(aggregatePaid, 2),
    baseMkup: 0,
    finMkup: 0,
    nbdMkup: 0,
    markedBase: 0,
    markedFin: 0,
    subtotalPreNbd: 0,
    uncappedNew: 0,
    capped: "",
    newCad: round(aggregateNew, 2),
    newUsd: round(aggregateNew * opts.usdRate, 2),
    deltaPerOrder: 0,
    totalDelta: round(aggregateDelta, 2),
    annualizedDelta: round(aggregateDelta * 4, 2),
  };
  for (const dim of allDims) totalRow[`dim_${dim}`] = "";
  ws.addRow(totalRow);
  const totalR = ws.rowCount;
  const trRow = ws.getRow(totalR);
  trRow.eachCell((cell) => {
    cell.font = { ...BODY_FONT, bold: true };
    cell.fill = TOTAL_FILL;
  });
  for (const c of cols) {
    if (!c.fmt) continue;
    trRow.getCell(ws.getColumn(c.key).number).numFmt = c.fmt;
  }

  // Header note row above the data — small explainer
  ws.insertRow(1, []);
  ws.getRow(1).height = 18;
  ws.mergeCells(1, 1, 1, cols.length);
  const titleCell = ws.getCell(1, 1);
  titleCell.value = `Scenario ${opts.scenarioId} · Base ${scenario.baseFormatted} · Finishing ${scenario.finFormatted} · NBD ${scenario.nbdFormatted} · Annualized Delta = sum(Total Delta) × 4`;
  titleCell.font = { name: "Arial", italic: true, color: { argb: "FF555555" }, size: 10 };
  titleCell.alignment = { vertical: "middle" };

  const buf = await wb.xlsx.writeBuffer();
  return new Blob([buf as ArrayBuffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

export async function downloadAnnotatedOrdersXlsx(
  opts: BuildAnnotatedOrdersOpts
): Promise<void> {
  const blob = await buildAnnotatedOrdersXlsx(opts);
  saveAs(
    blob,
    `${opts.productSlug}_Orders_Annotated_${opts.scenarioId}.xlsx`
  );
}
