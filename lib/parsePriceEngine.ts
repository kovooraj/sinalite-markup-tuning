import * as XLSX from "xlsx";
import {
  detectDimensions,
  normalizeDimValue,
  buildKey,
  RowDimensions,
} from "./dimensions";
import { normalizeStock, normalizeSize, normalizeTurnaround } from "./normalize";

export interface PriceEngineRow {
  productName: string;
  qty: number;
  stock: string;
  size: string;
  turnaround: "Standard" | "Rush (NBD)";
  rawTurnaround: string;
  /** All optional dimensional fields detected on this row (e.g. coating,
   * bundling, scoring for postcards; pages, cover, binding for booklets). */
  dims: RowDimensions;
  currentSalePrice: number;
  pe3CostTotal: number;
  consolidatedMarkup: number;
  baseCost: number;
  finCost: number;
}

export interface PriceEngineData {
  productName: string;
  productSlug: string;
  /** Sorted list of dimensional field keys detected (e.g. ["bundling",
   * "coating", "scoring"] for postcards, ["binding", "cover", "pages"] for
   * booklets). */
  dimensions: string[];
  rows: PriceEngineRow[];
  /** Indexes the rows by (stock|size|qty|turnaround|...all dimensions...).
   * For matching to an order replay with a different schema, the caller
   * should rebuild a sub-index using the dimensional intersection of the
   * two files. */
  warnings: string[];
}

const FINISHING_MACHINE_PATTERNS = [
  /rosback/i,
  /longford/i,
  /shrink\s*wrap/i,
  /bundler/i,
  /score/i,
  /stahlfold/i,
  /horizon\s+bookmaker/i,
  /stitchmaster/i,
];

function isFinishingMachine(name: string): boolean {
  if (!name) return false;
  return FINISHING_MACHINE_PATTERNS.some((re) => re.test(name));
}

function extractCostBreakdown(cells: (string | number | null)[]): {
  baseCost: number;
  finCost: number;
} {
  let baseCost = 0;
  let finCost = 0;
  let i = 0;
  while (i < cells.length) {
    const key = cells[i];
    if (key === null || key === undefined || key === "") {
      i += 1;
      continue;
    }
    if (typeof key === "string" && key.trim().toLowerCase() === "name") {
      const machine = String(cells[i + 1] ?? "").trim();
      let j = i + 2;
      let totalCost: number | null = null;
      while (j < cells.length) {
        const k = cells[j];
        if (typeof k === "string" && k.trim().toLowerCase() === "name") break;
        if (typeof k === "string" && k.trim() === "totalCost") {
          const v = cells[j + 1];
          if (typeof v === "number") totalCost = v;
          else if (typeof v === "string" && v.trim() !== "" && !isNaN(Number(v)))
            totalCost = Number(v);
          break;
        }
        j += 1;
      }
      if (totalCost !== null) {
        if (isFinishingMachine(machine)) finCost += totalCost;
        else baseCost += totalCost;
      }
      i = j;
    } else {
      i += 1;
    }
  }
  return { baseCost, finCost };
}

function toNumber(v: unknown): number {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return v;
  const n = Number(v);
  return isFinite(n) ? n : 0;
}

function slugify(name: string): string {
  return name
    .replace(/[^A-Za-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "")
    .slice(0, 64);
}

function findColIndex(
  header: (string | number | null)[],
  candidates: string[]
): number {
  for (const cand of candidates) {
    const i = header.findIndex(
      (h) => typeof h === "string" && h.trim().toLowerCase() === cand.toLowerCase()
    );
    if (i >= 0) return i;
  }
  return -1;
}

export async function parsePriceEngine(file: File): Promise<PriceEngineData> {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });
  const sheetName = wb.SheetNames[0];
  if (!sheetName) throw new Error("Price engine file has no sheets.");
  const ws = wb.Sheets[sheetName];
  const aoa = XLSX.utils.sheet_to_json<(string | number | null)[]>(ws, {
    header: 1,
    raw: true,
    defval: null,
  });
  if (aoa.length < 2) throw new Error("Price engine file has no data rows.");

  const header = aoa[0];

  // Required columns
  const idxQty = findColIndex(header, ["qty", "Qty"]);
  const idxStock = findColIndex(header, ["Stock", "PE3 Stock"]);
  const idxSize = findColIndex(header, ["size", "Size", "PE3 Size"]);
  const idxTurnaround = findColIndex(header, [
    "Turnaround",
    "PE3 Turnaround",
    "Turnaround (Raw)",
  ]);
  const idxSale = findColIndex(header, ["sale price", "Sale Price", "PE3 Sale Price"]);
  const idxPe3 = findColIndex(header, [
    "PE3 cost no markup",
    "PE3 Cost No Markup",
    "PE3 cost",
    "PE3 Cost",
  ]);

  const required: [number, string][] = [
    [idxQty, "qty"],
    [idxStock, "Stock"],
    [idxSize, "size"],
    [idxTurnaround, "Turnaround"],
    [idxSale, "sale price"],
    [idxPe3, "PE3 cost no markup (or PE3 cost)"],
  ];
  for (const [i, label] of required) {
    if (i < 0) {
      throw new Error(
        `Price engine missing required column "${label}". Required: qty, Stock, size, Turnaround, sale price, PE3 cost. ` +
          `Dimensional columns (Coating, Bundling, Scoring, Pages, Cover, Binding, Sides, Finishing, etc.) are auto-detected — any subset is fine.`
      );
    }
  }

  // Optional dimensions
  const dimIdx = detectDimensions(header);
  const dimensions = Object.keys(dimIdx).sort();

  // Optional auxiliary columns
  const idxConsolidatedMarkup = findColIndex(header, [
    "Consolidated Markup",
    "Consolidated_Markup",
  ]);
  let breakdownStart = header.findIndex(
    (h) => typeof h === "string" && h.trim().toLowerCase() === "breakdown"
  );
  if (breakdownStart < 0) {
    // Best guess: just after PE3 cost columns
    breakdownStart = Math.max(idxPe3, idxConsolidatedMarkup) + 1;
  }

  const productNameFromHeader = (header[0] as string) || sheetName;

  const rows: PriceEngineRow[] = [];
  const warnings: string[] = [];

  for (let r = 1; r < aoa.length; r++) {
    const row = aoa[r];
    if (!row || row.length === 0) continue;
    const qty = toNumber(row[idxQty]);
    if (!qty) continue;

    const productName = (row[0] as string) || productNameFromHeader;
    const stock = normalizeStock(row[idxStock] as string);
    const size = normalizeSize(row[idxSize] as string);
    const rawTurnaround = String(row[idxTurnaround] ?? "");
    const turnaround = normalizeTurnaround(rawTurnaround);
    const currentSalePrice = toNumber(row[idxSale]);
    const pe3CostTotal = toNumber(row[idxPe3]);
    const consolidatedMarkup =
      idxConsolidatedMarkup >= 0
        ? toNumber(row[idxConsolidatedMarkup])
        : pe3CostTotal > 0
          ? currentSalePrice / pe3CostTotal
          : 0;

    // Collect dimensional values for this row
    const dims: RowDimensions = {};
    for (const [dim, colIdx] of Object.entries(dimIdx)) {
      const v = normalizeDimValue(row[colIdx]);
      if (v) dims[dim] = v;
    }

    const breakdownCells = row.slice(breakdownStart);
    const { baseCost, finCost } = extractCostBreakdown(breakdownCells);

    rows.push({
      productName: String(productName).trim(),
      qty,
      stock,
      size,
      turnaround,
      rawTurnaround,
      dims,
      currentSalePrice,
      pe3CostTotal,
      consolidatedMarkup,
      baseCost,
      finCost,
    });
  }

  if (rows.length === 0) {
    throw new Error("Price engine parsed but contained no data rows.");
  }

  if (dimensions.length === 0) {
    warnings.push(
      "No optional dimensional columns detected (e.g. Coating, Bundling, Pages). Lookups will rely on Stock/Size/Qty/Turnaround alone, which may cause ambiguity if multiple variants share these."
    );
  }

  const productName = rows[0]?.productName || productNameFromHeader || sheetName;
  return {
    productName,
    productSlug: slugify(productName),
    dimensions,
    rows,
    warnings,
  };
}

/** Build a fast lookup index keyed by (required + included dims). The caller
 * provides which dimensions to include — typically the intersection of the
 * price-engine's and order-replay's dimensions. */
export function indexPriceEngine(
  pe: PriceEngineData,
  includedDims: string[]
): Map<string, PriceEngineRow> {
  const map = new Map<string, PriceEngineRow>();
  for (const r of pe.rows) {
    const key = buildKey(
      { stock: r.stock, size: r.size, qty: r.qty, turnaround: r.turnaround },
      r.dims,
      includedDims
    );
    map.set(key, r);
  }
  return map;
}

/** Identify the base catalog rows for this product. "Base" means: no
 * finishing add-ons selected. A row is "base" when every dimension whose
 * value typically signals add-on presence (bundling, scoring, finishing,
 * lamination, etc.) is empty or carries a "no" / "none" value. */
const ADDON_DIMS = new Set([
  "bundling",
  "scoring",
  "finishing",
  "lamination",
  "foil",
  "embossing",
  "diecutting",
  "corner",
  "perforation",
  "drilling",
]);

export function isBaseRow(r: PriceEngineRow, dimensions: string[]): boolean {
  for (const d of dimensions) {
    if (!ADDON_DIMS.has(d)) continue;
    const v = r.dims[d];
    if (!v) continue; // empty = no add-on
    const lc = v.toLowerCase();
    if (lc.includes("no bundling") || lc === "none" || lc === "free" || lc === "no")
      continue;
    return false;
  }
  return true;
}
