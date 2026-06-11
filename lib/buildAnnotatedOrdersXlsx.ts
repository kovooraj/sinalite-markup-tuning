import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import { OrderReplayData } from "./parseOrderReplay";
import { PriceEngineData } from "./parsePriceEngine";
import { computeAllScenarios } from "./computeScenarios";
import { computeNewPrice, resolveScenario, ScenarioDef } from "./markupEngine";
import { bandOf, nbdBandOf } from "./qtyBands";

export interface BuildAnnotatedOrdersOpts {
  pe: PriceEngineData;
  order: OrderReplayData;
  scenarioId: string;
  usdRate: number;
  productSlug: string;
  /** When false the cap rule is NOT applied — final price floats free of
   * the PE3 list price and delta vs list can go positive. Defaults to true. */
  applyCapRule?: boolean;
  /** Extra scenario definitions beyond the built-in A–H table — the user's
   * custom scenario and any saved scenarios. When scenarioId matches one of
   * these, that definition is used. */
  extraScenarios?: readonly ScenarioDef[];
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
  const scenario: ScenarioDef | undefined = resolveScenario(
    opts.scenarioId,
    opts.extraScenarios
  );
  if (!scenario) throw new Error(`Unknown scenario: ${opts.scenarioId}`);

  const applyCapRule = opts.applyCapRule ?? true;
  // Use the same matching pass the live UI uses — gets us the resolved
  // tuples with match quality (exact / snapped-qty / no-match) and the
  // closest-qty fallback applied.
  const scResult = computeAllScenarios(opts.order, opts.pe, { applyCapRule });
  const resolved = scResult.resolved;

  const wb = new ExcelJS.Workbook();
  wb.creator = "SinaLite Markup Tuning";
  wb.created = new Date();
  // Excel caps worksheet names at 31 chars — saved-scenario ids can be long
  const ws = wb.addWorksheet(`Annotated_${opts.scenarioId}`.slice(0, 31), {
    views: [{ state: "frozen", xSplit: 4, ySplit: 1 }],
  });

  const usdHeader = `New Price USD (@${opts.usdRate.toFixed(2)})`;

  // Dynamic columns: identifier columns (UID, Order ID, etc.) →
  // SKU descriptors → product dims → computed markup chain → delta math
  type Col = { header: string; key: string; width: number; fmt?: string };
  const cols: Col[] = [];
  // Identifier columns first so SL_ numbers / Order IDs are immediately
  // scannable at the left edge of the sheet
  for (const idCol of opts.order.identifierColumns) {
    const isDate = idCol.key === "orderDate";
    const isLongId = idCol.key === "uid" || idCol.key === "orderId";
    cols.push({
      header: idCol.label,
      key: `id_${idCol.key}`,
      width: isDate ? 20 : isLongId ? 16 : 12,
    });
  }
  cols.push(
    { header: "Description", key: "description", width: 36 },
    { header: "Qty", key: "qty", width: 9, fmt: FMT_INT },
    { header: "Stock", key: "stock", width: 28 },
    { header: "Size", key: "size", width: 12 },
    { header: "Turnaround", key: "turnaround", width: 18 }
  );
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
    {
      header: "Actual Paid (CAD)",
      key: "avgPaid",
      width: 14,
      fmt: FMT_CURRENCY,
    },
    { header: "PE3 Base Cost", key: "baseCost", width: 13, fmt: FMT_CURRENCY },
    { header: "PE3 Fin Cost", key: "finCost", width: 12, fmt: FMT_CURRENCY },
    {
      header: "PE3 List Price",
      key: "currentSalePrice",
      width: 14,
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
    { header: "Match Quality", key: "matchQuality", width: 26 },
    { header: "New Price CAD", key: "newCad", width: 14, fmt: FMT_CURRENCY },
    { header: usdHeader, key: "newUsd", width: 18, fmt: FMT_CURRENCY },
    {
      header: "Δ vs PE3 List (per order)",
      key: "deltaPerOrder",
      width: 18,
      fmt: FMT_CURRENCY,
    },
    {
      header: "Total Δ vs PE3 List (CAD)",
      key: "totalDelta",
      width: 20,
      fmt: FMT_CURRENCY,
    },
    {
      header: "Current Margin",
      key: "currentMargin",
      width: 14,
      fmt: FMT_CURRENCY,
    },
    {
      header: "New Margin",
      key: "newMargin",
      width: 14,
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
  let aggregatePaid = 0; // sum of avgPaid × orders (context, not the delta baseline)
  let aggregateList = 0; // sum of PE3 list × orders (delta baseline)
  let aggregateNew = 0;
  let aggregateBaseCost = 0;
  let aggregateFinCost = 0;
  let aggregateMarkedBase = 0;
  let aggregateMarkedFin = 0;
  let aggregateSubtotal = 0;
  let aggregateUncapped = 0;
  let aggregateCurrentMargin = 0;
  let aggregateNewMargin = 0;
  let matchedRowCount = 0;
  let unmatchedRowCount = 0;

  for (const { row: r, pe, matchQuality, snappedFromQty } of resolved) {
    const baseRow: Record<string, string | number> = {
      description: r.description,
      qty: r.qty,
      stock: r.stock || "—",
      size: r.size,
      turnaround: r.turnaround,
      orders: r.orders,
      avgPaid: round(r.avgPaid, 2),
    };
    for (const idCol of opts.order.identifierColumns) {
      baseRow[`id_${idCol.key}`] = r.identifiers[idCol.key] ?? "";
    }
    for (const dim of allDims) {
      baseRow[`dim_${dim}`] = r.dims[dim] ?? "";
    }

    if (!pe || matchQuality === "no-match") {
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
        matchQuality: "no match",
        newCad: 0,
        newUsd: 0,
        deltaPerOrder: 0,
        totalDelta: 0,
        currentMargin: 0,
        newMargin: 0,
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
      scenario.grad,
      applyCapRule
    );
    const subtotalPreNbd = res.basePrice + res.finPrice;
    // Delta is the catalog-repricing measure: New Price (capped) − PE3 List Price.
    // When cap rule applied: capped orders → delta = 0; uncapped → delta ≤ 0.
    // When cap rule disabled: delta can go positive.
    const deltaPerOrder = res.finalPrice - pe.currentSalePrice;
    const totalDelta = deltaPerOrder * r.orders;
    const variantCost = pe.baseCost + pe.finCost;
    const currentMargin = r.avgPaid - variantCost;
    const newMargin = res.finalPrice - variantCost;

    aggregateDelta += totalDelta;
    aggregatePaid += r.avgPaid * r.orders;
    aggregateList += pe.currentSalePrice * r.orders;
    aggregateNew += res.finalPrice * r.orders;
    aggregateBaseCost += pe.baseCost * r.orders;
    aggregateFinCost += pe.finCost * r.orders;
    aggregateMarkedBase += res.basePrice * r.orders;
    aggregateMarkedFin += res.finPrice * r.orders;
    aggregateSubtotal += subtotalPreNbd * r.orders;
    aggregateUncapped += res.uncappedPrice * r.orders;
    aggregateCurrentMargin += currentMargin * r.orders;
    aggregateNewMargin += newMargin * r.orders;
    matchedRowCount += 1;

    const matchQualityLabel =
      matchQuality === "snapped-qty" && snappedFromQty !== null
        ? `snapped to qty ${snappedFromQty.toLocaleString()}`
        : "exact";

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
      matchQuality: matchQualityLabel,
      newCad: round(res.finalPrice, 2),
      newUsd: round(res.finalPrice * opts.usdRate, 2),
      deltaPerOrder: round(deltaPerOrder, 2),
      totalDelta: round(totalDelta, 2),
      currentMargin: round(currentMargin, 2),
      newMargin: round(newMargin, 2),
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
    row.getCell(ws.getColumn("totalDelta").number).fill = HIGHLIGHT_FILL;
  }

  // TOTAL row
  const annualizedSum = aggregateDelta * 4;
  // Total order count summed across data rows — for per-order detail (each
  // row = 1 order) this equals matchedRowCount + unmatchedRowCount; for
  // per-sku aggregated it equals sum of the Orders column.
  const totalOrders = matchedRowCount + unmatchedRowCount > 0
    ? opts.order.rows.reduce((sum, r) => sum + r.orders, 0)
    : 0;
  const totalRow: Record<string, string | number> = {
    description: `TOTAL (${matchedRowCount} matched · ${unmatchedRowCount} unmatched) · Annualized × 4 = $${annualizedSum.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`,
    qty: "",
    stock: "",
    size: "",
    turnaround: "",
    orders: totalOrders,
    avgPaid: round(aggregatePaid, 2),
    baseCost: round(aggregateBaseCost, 2),
    finCost: round(aggregateFinCost, 2),
    currentSalePrice: round(aggregateList, 2),
    baseMkup: 0,
    finMkup: 0,
    nbdMkup: 0,
    markedBase: round(aggregateMarkedBase, 2),
    markedFin: round(aggregateMarkedFin, 2),
    subtotalPreNbd: round(aggregateSubtotal, 2),
    uncappedNew: round(aggregateUncapped, 2),
    capped: "",
    matchQuality: "",
    newCad: round(aggregateNew, 2),
    newUsd: round(aggregateNew * opts.usdRate, 2),
    deltaPerOrder: round(aggregateDelta, 2),
    totalDelta: round(aggregateDelta, 2),
    currentMargin: round(aggregateCurrentMargin, 2),
    newMargin: round(aggregateNewMargin, 2),
  };
  for (const dim of allDims) totalRow[`dim_${dim}`] = "";
  for (const idCol of opts.order.identifierColumns)
    totalRow[`id_${idCol.key}`] = "";
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
  const capLabel = applyCapRule ? "Cap rule: APPLIED" : "Cap rule: DISABLED";
  titleCell.value = `Scenario ${opts.scenarioId} · Base ${scenario.baseFormatted} · Finishing ${scenario.finFormatted} · NBD ${scenario.nbdFormatted}  ||  ${capLabel} · Δ vs PE3 List = New Price ${applyCapRule ? "(capped)" : "(uncapped)"} − PE3 List Price · Annualized = sum(Total Δ) × 4`;
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
