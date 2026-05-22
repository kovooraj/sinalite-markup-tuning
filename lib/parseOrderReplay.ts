import * as XLSX from "xlsx";
import { normalizeSize, normalizeStock, normalizeTurnaround } from "./normalize";
import { bandLabelOf } from "./qtyBands";
import {
  detectDimensions,
  normalizeDimValue,
  RowDimensions,
  buildKey,
} from "./dimensions";

export interface OrderReplayRow {
  description: string;
  qty: number;
  qtyBandLabel: string;
  turnaround: "Standard" | "Rush (NBD)";
  stock: string;
  size: string;
  /** All optional dimensional fields detected on this row. */
  dims: RowDimensions;
  orders: number;
  actualPaid: number;
  avgPaid: number;
  /** Pre-computed Scenario A price per order (only available in legacy format).
   * For per-order detail input this stays 0 and self-check is skipped. */
  newPricePerOrder: number;
  newRevenue: number;
  delta: number;
  pctChange: number;
  baseCost: number;
  baseSP: number;
  variantSP: number;
}

export type OrderReplayFormat = "per_sku_aggregated" | "per_order_detail";

export interface OrderReplayData {
  format: OrderReplayFormat;
  /** Sorted list of dimensional field keys detected. */
  dimensions: string[];
  /** True when at least one row carries a non-empty Stock value. */
  hasStock: boolean;
  rows: OrderReplayRow[];
  totals: {
    orders: number;
    actualPaid: number;
    newRevenue: number;
    deltaAtScenarioA: number;
  };
  unmatched: { count: number; revenue: number; reasons: string[] };
  warnings: string[];
  hasPrecomputedScenarioA: boolean;
}

function toNumber(v: unknown): number {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return v;
  const n = Number(v);
  return isFinite(n) ? n : 0;
}

function toStr(v: unknown): string {
  if (v === null || v === undefined) return "";
  return String(v).trim();
}

function findColIndex(
  header: (string | number | null)[],
  candidates: string[]
): number {
  for (const cand of candidates) {
    const i = header.findIndex(
      (h) =>
        typeof h === "string" && h.trim().toLowerCase() === cand.toLowerCase()
    );
    if (i >= 0) return i;
  }
  return -1;
}

const COL_ALIASES = {
  description: ["Description", "Product Name", "Product"],
  qty: ["Qty", "Print Qty", "Actual Print Qty", "PE3 Qty Used"],
  qtyBand: ["Qty Band"],
  turnaround: ["Turnaround", "PE3 Turnaround", "Turnaround (Raw)"],
  stock: ["Stock", "PE3 Stock"], // intentionally NOT "Product Name" — they differ semantically
  size: ["Size", "PE3 Size", "Size (Ordered)", "Ordered Size", "size"],
  orders: ["Orders", "# Orders", "Line Items"],
  actualPaid: ["Actual Paid $", "Total Paid", "Total Revenue", "Net Revenue"],
  avgPaid: ["Avg Paid $", "Avg Paid", "Actual Sale Price", "Sale Price"],
  currency: ["Currency"],
  newPricePerOrder: ["New Price/Order", "New Price"],
  newRevenue: ["New Revenue $", "New Revenue"],
  delta: ["Delta $", "Delta"],
  pctChange: ["% Change", "Pct Change"],
  baseCost: ["Base Cost"],
  baseSP: ["Base SP"],
  variantSP: ["Variant SP", "PE3 Sale Price"],
};

interface ColMap {
  description: number;
  qty: number;
  qtyBand: number;
  turnaround: number;
  stock: number;
  size: number;
  orders: number;
  actualPaid: number;
  avgPaid: number;
  currency: number;
  newPricePerOrder: number;
  newRevenue: number;
  delta: number;
  pctChange: number;
  baseCost: number;
  baseSP: number;
  variantSP: number;
}

function buildColMap(header: (string | number | null)[]): ColMap {
  const m = {} as ColMap;
  for (const [k, list] of Object.entries(COL_ALIASES) as [
    keyof ColMap,
    string[],
  ][]) {
    m[k] = findColIndex(header, list);
  }
  return m;
}

/** Sheets we never want to pick as the order source, even if they happen to
 * have some columns we recognize. */
const SHEET_IGNORE_PATTERNS = [/^summary$/i, /^index$/i, /^readme$/i];

interface SheetCandidate {
  name: string;
  format: OrderReplayFormat;
  score: number;
  rowCount: number;
}

/** Score a single sheet for "how good is this as the order-replay source?".
 * Higher = better. Returns null if the sheet looks useless. */
function scoreSheet(
  wb: XLSX.WorkBook,
  name: string
): SheetCandidate | null {
  if (SHEET_IGNORE_PATTERNS.some((re) => re.test(name))) return null;
  const ws = wb.Sheets[name];
  const aoa = XLSX.utils.sheet_to_json<(string | number | null)[]>(ws, {
    header: 1,
    defval: null,
  });
  if (aoa.length < 2) return null;
  const m = buildColMap(aoa[0]);
  if (m.qty < 0 || m.size < 0 || m.avgPaid < 0) return null;

  // Aggregated format wins big if it has the pre-computed New Price column
  if (m.orders >= 0 && m.newPricePerOrder >= 0) {
    return {
      name,
      format: "per_sku_aggregated",
      score: 100 + (aoa.length - 1),
      rowCount: aoa.length - 1,
    };
  }
  // Per-order detail: score = column richness + row count
  let s = 10;
  if (m.stock >= 0) s += 5;
  if (m.turnaround >= 0) s += 3;
  if (m.currency >= 0) s += 2;
  if (m.description >= 0) s += 1;
  s += aoa.length - 1; // more rows = better signal
  return {
    name,
    format: "per_order_detail",
    score: s,
    rowCount: aoa.length - 1,
  };
}

function chooseSheet(
  wb: XLSX.WorkBook
): { sheetName: string; format: OrderReplayFormat } | null {
  // First, prefer the canonical Sinalite sheet names if present
  const canonical: { name: string; format: OrderReplayFormat }[] = [
    { name: "Per-SKU Detail", format: "per_sku_aggregated" },
    { name: "Matched Orders (Price Analysis)", format: "per_order_detail" },
    { name: "All Orders + PE3 Match", format: "per_order_detail" },
    { name: "All Orders", format: "per_order_detail" },
  ];
  for (const c of canonical) {
    if (wb.SheetNames.includes(c.name)) {
      // Confirm the canonical sheet has plausible columns, otherwise fall
      // through to scoring.
      const score = scoreSheet(wb, c.name);
      if (score) {
        return { sheetName: c.name, format: c.format };
      }
    }
  }
  // Score every sheet; pick the best.
  const candidates = wb.SheetNames.map((n) => scoreSheet(wb, n)).filter(
    (c): c is SheetCandidate => c !== null
  );
  candidates.sort((a, b) => b.score - a.score);
  if (candidates.length === 0) return null;
  return { sheetName: candidates[0].name, format: candidates[0].format };
}

function readDims(
  row: (string | number | null)[],
  dimIdx: Record<string, number>
): RowDimensions {
  const dims: RowDimensions = {};
  for (const [dim, colIdx] of Object.entries(dimIdx)) {
    const v = normalizeDimValue(row[colIdx]);
    if (v) dims[dim] = v;
  }
  return dims;
}

function rowsForPerSku(
  aoa: (string | number | null)[][],
  m: ColMap,
  dimIdx: Record<string, number>
): {
  rows: OrderReplayRow[];
  totals: { orders: number; actualPaid: number; newRevenue: number };
} {
  const rows: OrderReplayRow[] = [];
  let totalOrders = 0;
  let totalPaid = 0;
  let totalNewRev = 0;
  for (let r = 1; r < aoa.length; r++) {
    const row = aoa[r];
    if (!row || row.length === 0) continue;
    const qty = toNumber(row[m.qty]);
    if (!qty) continue;
    const turnaround = normalizeTurnaround(toStr(row[m.turnaround]));
    const stock = m.stock >= 0 ? normalizeStock(toStr(row[m.stock])) : "";
    const size = normalizeSize(toStr(row[m.size]));
    const orders = toNumber(row[m.orders]);
    const actualPaid = toNumber(row[m.actualPaid]);
    const avgPaid = toNumber(row[m.avgPaid]);
    const newPricePerOrder = toNumber(row[m.newPricePerOrder]);
    const newRevenue = toNumber(row[m.newRevenue]);
    rows.push({
      description: toStr(row[m.description]),
      qty,
      qtyBandLabel: toStr(row[m.qtyBand]) || bandLabelOf(qty),
      turnaround,
      stock,
      size,
      dims: readDims(row, dimIdx),
      orders,
      actualPaid,
      avgPaid,
      newPricePerOrder,
      newRevenue,
      delta: toNumber(row[m.delta]),
      pctChange: toNumber(row[m.pctChange]),
      baseCost: toNumber(row[m.baseCost]),
      baseSP: toNumber(row[m.baseSP]),
      variantSP: toNumber(row[m.variantSP]),
    });
    totalOrders += orders;
    totalPaid += actualPaid;
    totalNewRev += newRevenue;
  }
  return {
    rows,
    totals: { orders: totalOrders, actualPaid: totalPaid, newRevenue: totalNewRev },
  };
}

function rowsForPerOrder(
  aoa: (string | number | null)[][],
  m: ColMap,
  dimIdx: Record<string, number>,
  cadFromUsd: number
): {
  rows: OrderReplayRow[];
  totals: { orders: number; actualPaid: number; newRevenue: number };
  usdRowsConverted: number;
} {
  const rows: OrderReplayRow[] = [];
  let totalOrders = 0;
  let totalPaid = 0;
  let usdConverted = 0;
  for (let r = 1; r < aoa.length; r++) {
    const row = aoa[r];
    if (!row || row.length === 0) continue;
    const qty = toNumber(row[m.qty]);
    if (!qty) continue;
    const stock = m.stock >= 0 ? normalizeStock(toStr(row[m.stock])) : "";
    const size = normalizeSize(toStr(row[m.size]));
    if (!size) continue;
    const turnaround = normalizeTurnaround(toStr(row[m.turnaround]));
    let salePrice = toNumber(row[m.avgPaid]);
    if (!salePrice) continue;
    const currency = m.currency >= 0 ? toStr(row[m.currency]).toUpperCase() : "CAD";
    if (currency === "USD") {
      salePrice = salePrice * cadFromUsd;
      usdConverted += 1;
    }
    rows.push({
      description: toStr(row[m.description]) || `${stock} / ${size} / ${qty}`,
      qty,
      qtyBandLabel: bandLabelOf(qty),
      turnaround,
      stock,
      size,
      dims: readDims(row, dimIdx),
      orders: 1,
      actualPaid: salePrice,
      avgPaid: salePrice,
      newPricePerOrder: 0,
      newRevenue: 0,
      delta: 0,
      pctChange: 0,
      baseCost: 0,
      baseSP: 0,
      variantSP: m.variantSP >= 0 ? toNumber(row[m.variantSP]) : 0,
    });
    totalOrders += 1;
    totalPaid += salePrice;
  }
  return {
    rows,
    totals: { orders: totalOrders, actualPaid: totalPaid, newRevenue: 0 },
    usdRowsConverted: usdConverted,
  };
}

function countUnmatchedFromLegacySheet(wb: XLSX.WorkBook): {
  count: number;
  revenue: number;
  reasons: Set<string>;
} {
  const reasons = new Set<string>();
  let count = 0;
  let revenue = 0;
  if (!wb.SheetNames.includes("Unmatched")) {
    return { count, revenue, reasons };
  }
  const ws = wb.Sheets["Unmatched"];
  const aoa = XLSX.utils.sheet_to_json<(string | number | null)[]>(ws, {
    header: 1,
    defval: null,
  });
  let headerRow = -1;
  for (let i = 0; i < Math.min(aoa.length, 5); i++) {
    const r = aoa[i];
    if (
      r &&
      r.some((c) => typeof c === "string" && c.toLowerCase().includes("orders")) &&
      r.some((c) => typeof c === "string" && c.toLowerCase().includes("revenue"))
    ) {
      headerRow = i;
      break;
    }
  }
  if (headerRow < 0) return { count, revenue, reasons };
  const h = aoa[headerRow];
  const orderIdx = h.findIndex(
    (c) => typeof c === "string" && c.toLowerCase().includes("orders")
  );
  const revIdx = h.findIndex(
    (c) => typeof c === "string" && c.toLowerCase().includes("revenue")
  );
  const reasonIdx = h.findIndex(
    (c) => typeof c === "string" && c.toLowerCase().includes("reason")
  );
  for (let i = headerRow + 1; i < aoa.length; i++) {
    const r = aoa[i];
    if (!r) continue;
    const o = toNumber(r[orderIdx]);
    const v = toNumber(r[revIdx]);
    if (!o && !v) continue;
    count += o;
    revenue += v;
    if (reasonIdx >= 0) {
      const reason = toStr(r[reasonIdx]);
      if (reason) reasons.add(reason);
    }
  }
  return { count, revenue, reasons };
}

function countUnmatchedFromExtraSheets(
  wb: XLSX.WorkBook,
  cadFromUsd: number
): { count: number; revenue: number; reasons: Set<string> } {
  const reasons = new Set<string>();
  let count = 0;
  let revenue = 0;
  // Only count sheets that are unambiguously unmatched. Generic "Rush Orders"
  // (without the "(No PE3 Match)" suffix) is often a duplicate listing of the
  // main sheet rather than an unmatched bucket, so we don't include it here.
  const candidates = [
    "Custom Sizes (No PE3 Match)",
    "Rush Orders (No PE3 Match)",
    "Other Unmatched",
    "Custom Sizes",
    "Unmatched Orders",
    "Unmatched",
  ];
  for (const name of candidates) {
    if (!wb.SheetNames.includes(name)) continue;
    const ws = wb.Sheets[name];
    const aoa = XLSX.utils.sheet_to_json<(string | number | null)[]>(ws, {
      header: 1,
      defval: null,
    });
    if (aoa.length < 2) continue;
    const header = aoa[0];
    const m = buildColMap(header);
    if (m.avgPaid < 0) continue;
    for (let r = 1; r < aoa.length; r++) {
      const row = aoa[r];
      if (!row || row.length === 0) continue;
      let v = toNumber(row[m.avgPaid]);
      if (!v) continue;
      const currency =
        m.currency >= 0 ? toStr(row[m.currency]).toUpperCase() : "CAD";
      if (currency === "USD") v = v * cadFromUsd;
      revenue += v;
      count += 1;
    }
    reasons.add(name);
  }
  return { count, revenue, reasons };
}

export async function parseOrderReplay(
  file: File,
  opts: { usdRate?: number } = {}
): Promise<OrderReplayData> {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });
  const choice = chooseSheet(wb);
  if (!choice) {
    throw new Error(
      `Order replay couldn't find a usable sheet. Sheets found: ${wb.SheetNames.join(", ") || "(none)"}.`
    );
  }
  const { sheetName, format } = choice;
  const ws = wb.Sheets[sheetName];
  const aoa = XLSX.utils.sheet_to_json<(string | number | null)[]>(ws, {
    header: 1,
    raw: true,
    defval: null,
  });
  if (aoa.length < 2) throw new Error(`Sheet "${sheetName}" has no data rows.`);
  const m = buildColMap(aoa[0]);
  const dimIdx = detectDimensions(aoa[0]);
  const dimensions = Object.keys(dimIdx).sort();

  // Required-column guard per format. Stock is OPTIONAL — many product files
  // omit it because the file is scoped to one product variant.
  if (format === "per_sku_aggregated") {
    const required: [keyof ColMap, string][] = [
      ["qty", "Qty"],
      ["size", "Size"],
      ["orders", "Orders"],
      ["avgPaid", "Avg Paid $"],
    ];
    for (const [k, label] of required) {
      if (m[k] < 0) {
        throw new Error(`Aggregated order replay missing required column "${label}".`);
      }
    }
  } else {
    const required: [keyof ColMap, string][] = [
      ["qty", "Qty / Print Qty"],
      ["size", "Size / PE3 Size"],
      ["avgPaid", "Actual Sale Price / Avg Paid / Sale Price"],
    ];
    for (const [k, label] of required) {
      if (m[k] < 0) {
        throw new Error(
          `Per-order replay missing required column "${label}". Sheet: ${sheetName}.`
        );
      }
    }
  }

  const usdRate = opts.usdRate && opts.usdRate > 0 ? opts.usdRate : 0.7;
  const cadFromUsd = 1 / usdRate;

  const warnings: string[] = [];
  let result: {
    rows: OrderReplayRow[];
    totals: { orders: number; actualPaid: number; newRevenue: number };
  };
  let usdConverted = 0;

  if (format === "per_sku_aggregated") {
    result = rowsForPerSku(aoa, m, dimIdx);
  } else {
    const r2 = rowsForPerOrder(aoa, m, dimIdx, cadFromUsd);
    result = { rows: r2.rows, totals: r2.totals };
    usdConverted = r2.usdRowsConverted;
    if (usdConverted > 0) {
      warnings.push(
        `${usdConverted} USD orders converted to CAD using ×${cadFromUsd.toFixed(4)} (inverse of ${usdRate.toFixed(2)} USD rate).`
      );
    }
  }

  const unmatched =
    format === "per_sku_aggregated"
      ? countUnmatchedFromLegacySheet(wb)
      : countUnmatchedFromExtraSheets(wb, cadFromUsd);

  return {
    format,
    dimensions,
    hasStock: result.rows.some((r) => !!r.stock),
    rows: result.rows,
    totals: {
      orders: result.totals.orders,
      actualPaid: result.totals.actualPaid,
      newRevenue: result.totals.newRevenue,
      deltaAtScenarioA: result.totals.newRevenue - result.totals.actualPaid,
    },
    unmatched: {
      count: unmatched.count,
      revenue: unmatched.revenue,
      reasons: Array.from(unmatched.reasons),
    },
    warnings,
    hasPrecomputedScenarioA: format === "per_sku_aggregated",
  };
}

/** Convenience: build a lookup key for an order-replay row using the given
 * dimension subset. If `useStock` is false the stock field is collapsed to
 * "" so PE and order keys still match when one side lacks a Stock column. */
export function orderRowKey(
  r: OrderReplayRow,
  includedDims: string[],
  useStock = true
): string {
  return buildKey(
    {
      stock: useStock ? r.stock : "",
      size: r.size,
      qty: r.qty,
      turnaround: r.turnaround,
    },
    r.dims,
    includedDims
  );
}
