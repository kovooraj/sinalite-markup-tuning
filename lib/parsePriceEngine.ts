import * as XLSX from "xlsx";
import {
  detectDimensions,
  normalizeDimValue,
  buildKey,
  RowDimensions,
} from "./dimensions";
import { normalizeStock, normalizeSize, normalizeTurnaround } from "./normalize";
import { isCsvFile, parseCsvToAoa, CsvCell } from "./parseCsv";

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
  /** True when at least one row has a non-empty Stock value. When false the
   * file likely has no Stock column (booklets-style); the matching layer
   * will drop stock from the lookup key on both sides. */
  hasStock: boolean;
  rows: PriceEngineRow[];
  warnings: string[];
  /** Cost centers (machine / stage names) encountered while decomposing the
   * breakdown column, classified by how this tool bucketed them. */
  costCenters: { base: string[]; finishing: string[] };
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

function extractCostBreakdown(
  cells: (string | number | null)[],
  baseSeen: Set<string>,
  finSeen: Set<string>
): {
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
      if (totalCost !== null && machine) {
        if (isFinishingMachine(machine)) {
          finCost += totalCost;
          finSeen.add(machine);
        } else {
          baseCost += totalCost;
          baseSeen.add(machine);
        }
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
  let aoa: CsvCell[][];
  let sheetNameForLabel: string;

  if (isCsvFile(file)) {
    // Direct CSV path — bypasses SheetJS so very large files (Brochures
    // PE3 is 110 MB / 8.6M cells) don't blow up the cell-object model
    // with "Too many properties to enumerate".
    const buf = await file.arrayBuffer();
    const text = new TextDecoder("utf-8").decode(new Uint8Array(buf));
    aoa = parseCsvToAoa(text);
    sheetNameForLabel = file.name.replace(/\.[^.]+$/, "");
  } else {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const sheetName = wb.SheetNames[0];
    if (!sheetName) throw new Error("Price engine file has no sheets.");
    const ws = wb.Sheets[sheetName];
    aoa = XLSX.utils.sheet_to_json<CsvCell[]>(ws, {
      header: 1,
      raw: true,
      defval: null,
    });
    sheetNameForLabel = sheetName;
  }
  if (aoa.length < 2) throw new Error("Price engine file has no data rows.");

  const header = aoa[0];

  // Required columns
  const idxQty = findColIndex(header, ["qty", "Qty", "Print Qty"]);
  const idxStock = findColIndex(header, ["Stock", "PE3 Stock"]); // optional
  // For "size", prefer the unambiguous "Dimensions" column when it exists
  // (booklets PE3 puts physical size in "Dimensions" and uses the "size"
  // column for the page count instead). Fall back to "size" otherwise.
  const idxDimensionsCol = findColIndex(header, ["Dimensions", "PE3 Dimensions"]);
  const idxSizeCol = findColIndex(header, ["size", "Size", "PE3 Size"]);
  const idxSize = idxDimensionsCol >= 0 ? idxDimensionsCol : idxSizeCol;
  // If both exist, treat the "size" column as Pages — its values look like
  // "8pg", "16pg", etc. and represent the page count dimension.
  const idxPagesFromSize =
    idxDimensionsCol >= 0 && idxSizeCol >= 0 && idxSizeCol !== idxDimensionsCol
      ? idxSizeCol
      : -1;
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
    [idxSize, "size"],
    [idxTurnaround, "Turnaround"],
    [idxSale, "sale price"],
    [idxPe3, "PE3 cost no markup (or PE3 cost)"],
  ];
  for (const [i, label] of required) {
    if (i < 0) {
      throw new Error(
        `Price engine missing required column "${label}". Required: qty, size, Turnaround, sale price, PE3 cost. ` +
          `Stock + dimensional columns (Coating, Bundling, Scoring, Pages, Cover, Binding, Sides, Finishing, etc.) are auto-detected — any subset is fine.`
      );
    }
  }

  // Optional dimensions — auto-detected from header aliases
  const dimIdx = detectDimensions(header);
  // If we re-purposed the "size" column as Pages (booklets), register it
  // here so dim discovery sees it.
  if (idxPagesFromSize >= 0 && !("pages" in dimIdx)) {
    dimIdx["pages"] = idxPagesFromSize;
  }
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

  const productNameFromHeader = (header[0] as string) || sheetNameForLabel;

  const rows: PriceEngineRow[] = [];
  const warnings: string[] = [];
  const baseSeen = new Set<string>();
  const finSeen = new Set<string>();

  for (let r = 1; r < aoa.length; r++) {
    const row = aoa[r];
    if (!row || row.length === 0) continue;
    const qty = toNumber(row[idxQty]);
    if (!qty) continue;

    const productName = (row[0] as string) || productNameFromHeader;
    const stock =
      idxStock >= 0 ? normalizeStock(row[idxStock] as string) : "";
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
    const { baseCost, finCost } = extractCostBreakdown(
      breakdownCells,
      baseSeen,
      finSeen
    );

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

  const productName =
    rows[0]?.productName || productNameFromHeader || sheetNameForLabel;
  const hasStock = rows.some((r) => !!r.stock);
  return {
    productName,
    productSlug: slugify(productName),
    dimensions,
    hasStock,
    rows,
    warnings,
    costCenters: {
      base: Array.from(baseSeen).sort(),
      finishing: Array.from(finSeen).sort(),
    },
  };
}

/** Build a multi-valued index keyed by (required + included dims). Each
 * bucket holds every PE3 row that matches the key — used by callers to pick
 * the best variant for a given order (e.g. by price proximity to what the
 * customer paid). */
export function indexPriceEngine(
  pe: PriceEngineData,
  includedDims: string[],
  useStock = true
): { buckets: Map<string, PriceEngineRow[]>; collisions: number } {
  const buckets = new Map<string, PriceEngineRow[]>();
  let collisions = 0;
  for (const r of pe.rows) {
    const key = buildKey(
      {
        stock: useStock ? r.stock : "",
        size: r.size,
        qty: r.qty,
        turnaround: r.turnaround,
      },
      r.dims,
      includedDims
    );
    let arr = buckets.get(key);
    if (!arr) {
      arr = [];
      buckets.set(key, arr);
    } else {
      collisions += 1;
    }
    arr.push(r);
  }
  return { buckets, collisions };
}

/** Build a no-qty index keyed by (stock|size|turnaround|...dims). Used by
 * the closest-qty fallback when an order's exact qty isn't in the PE3
 * catalog. */
export function indexPriceEngineNoQty(
  pe: PriceEngineData,
  includedDims: string[],
  useStock = true
): Map<string, PriceEngineRow[]> {
  const buckets = new Map<string, PriceEngineRow[]>();
  for (const r of pe.rows) {
    const key = buildKey(
      {
        stock: useStock ? r.stock : "",
        size: r.size,
        qty: 0, // strip qty
        turnaround: r.turnaround,
      },
      r.dims,
      includedDims
    );
    let arr = buckets.get(key);
    if (!arr) {
      arr = [];
      buckets.set(key, arr);
    }
    arr.push(r);
  }
  return buckets;
}

/** Build a no-qty key for an order row using the same convention as the
 * no-qty index. */
export function orderRowKeyNoQty(
  row: { stock: string; size: string; turnaround: string; dims: RowDimensions },
  includedDims: string[],
  useStock = true
): string {
  return buildKey(
    {
      stock: useStock ? row.stock : "",
      size: row.size,
      qty: 0,
      turnaround: row.turnaround,
    },
    row.dims,
    includedDims
  );
}

/** Pick the PE3 row whose qty is closest to the target order qty. Used as
 * the closest-qty fallback. */
export function pickByClosestQty(
  candidates: PriceEngineRow[] | undefined,
  targetQty: number,
  paidPrice: number
): { pe: PriceEngineRow | null; snappedFromQty: number | null } {
  if (!candidates || candidates.length === 0)
    return { pe: null, snappedFromQty: null };
  let best = candidates[0];
  let bestDiff = Math.abs(best.qty - targetQty);
  for (let i = 1; i < candidates.length; i++) {
    const c = candidates[i];
    const d = Math.abs(c.qty - targetQty);
    if (d < bestDiff) {
      best = c;
      bestDiff = d;
    } else if (d === bestDiff && paidPrice > 0) {
      // tie-break by price proximity
      if (
        Math.abs(c.currentSalePrice - paidPrice) <
        Math.abs(best.currentSalePrice - paidPrice)
      ) {
        best = c;
      }
    }
  }
  return { pe: best, snappedFromQty: best.qty };
}

/** Pick the PE3 row that best matches an order based on price proximity.
 * Among all PE3 rows matching the order's key, choose the one whose
 * currentSalePrice is closest to what the customer actually paid. This
 * implicitly identifies which catalog variant the order corresponds to,
 * even when the order replay lacks the dimensional columns needed to
 * distinguish variants directly. */
export function pickByPriceProximity(
  candidates: PriceEngineRow[] | undefined,
  paidPrice: number
): PriceEngineRow | null {
  if (!candidates || candidates.length === 0) return null;
  if (candidates.length === 1) return candidates[0];
  if (paidPrice <= 0) return candidates[0];
  let best = candidates[0];
  let bestDiff = Math.abs(best.currentSalePrice - paidPrice);
  for (let i = 1; i < candidates.length; i++) {
    const d = Math.abs(candidates[i].currentSalePrice - paidPrice);
    if (d < bestDiff) {
      best = candidates[i];
      bestDiff = d;
    }
  }
  return best;
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
