import * as XLSX from "xlsx";
import {
  lookupKey,
  normalizeBundling,
  normalizeCoating,
  normalizeScoring,
  normalizeSize,
  normalizeStock,
  normalizeTurnaround,
} from "./normalize";

export interface PriceEngineRow {
  productName: string;
  qty: number;
  stock: string;
  coating: string;
  size: string;
  turnaround: "Standard" | "Rush (NBD)";
  rawTurnaround: string;
  bundling: string;
  scoring: string;
  currentSalePrice: number;
  pe3CostTotal: number;
  consolidatedMarkup: number;
  baseCost: number;
  finCost: number;
  key: string;
}

export interface PriceEngineData {
  productName: string;
  productSlug: string;
  rows: PriceEngineRow[];
  byKey: Map<string, PriceEngineRow>;
  /** Same as byKey but with coating stripped to "" — used to match order-replay
   * rows which don't carry a coating column. Assumes the price-engine file is
   * scoped to a single coating; if multiple coatings appear, last value wins. */
  byKeyNoCoating: Map<string, PriceEngineRow>;
  coatings: string[];
  warnings: string[];
}

const FINISHING_MACHINE_PATTERNS = [
  /rosback/i,
  /longford/i,
  /shrink\s*wrap/i,
  /bundler/i,
  /score\b/i,
];

function isFinishingMachine(name: string): boolean {
  if (!name) return false;
  return FINISHING_MACHINE_PATTERNS.some((re) => re.test(name));
}

// Walk the flattened breakdown columns starting at colStart.
// Pattern: cells of (key, value) pairs. Whenever key === "name", the next
// cell is the machine name; the matching key === "totalCost" in the same
// block gives that machine's total cost.
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
      // Walk forward in the same row until the next "name" or end, looking for
      // this machine's `totalCost` value.
      let j = i + 2;
      let totalCost: number | null = null;
      while (j < cells.length) {
        const k = cells[j];
        if (typeof k === "string" && k.trim().toLowerCase() === "name") break;
        if (typeof k === "string" && k.trim() === "totalCost") {
          const v = cells[j + 1];
          if (typeof v === "number") totalCost = v;
          else if (typeof v === "string" && v.trim() !== "" && !isNaN(Number(v))) totalCost = Number(v);
          // first totalCost in this block is the machine's total — stop scanning
          break;
        }
        j += 1;
      }
      if (totalCost !== null) {
        if (isFinishingMachine(machine)) finCost += totalCost;
        else baseCost += totalCost;
      }
      // jump to next block
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

export async function parsePriceEngine(file: File): Promise<PriceEngineData> {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });
  const sheetName = wb.SheetNames[0];
  if (!sheetName) throw new Error("Price engine xlsx has no sheets.");
  const ws = wb.Sheets[sheetName];
  const aoa = XLSX.utils.sheet_to_json<(string | number | null)[]>(ws, {
    header: 1,
    raw: true,
    defval: null,
  });
  if (aoa.length < 2) throw new Error("Price engine xlsx has no data rows.");

  const header = aoa[0];
  // Validate expected leading columns
  const expectedHeaders = [
    "qty",
    "Stock",
    "Coating",
    "size",
    "Turnaround",
    "Bundling",
    "Scoring",
    "sale price",
    "PE3 cost no markup",
    "Consolidated Markup",
  ];
  const idx: Record<string, number> = {};
  for (const name of expectedHeaders) {
    const i = header.findIndex(
      (h) => typeof h === "string" && h.trim().toLowerCase() === name.toLowerCase()
    );
    if (i < 0) {
      throw new Error(
        `Price engine missing column "${name}". Expected columns include: ${expectedHeaders.join(", ")}.`
      );
    }
    idx[name] = i;
  }
  // breakdown columns start at "breakdown" header if present, else col 11 (index 11)
  let breakdownStart = header.findIndex(
    (h) => typeof h === "string" && h.trim().toLowerCase() === "breakdown"
  );
  if (breakdownStart < 0) breakdownStart = 11;
  // Product name is in column A on each row (or use sheet name fallback)
  const productNameFromHeader = (header[0] as string) || sheetName;

  const rows: PriceEngineRow[] = [];
  const byKey = new Map<string, PriceEngineRow>();
  const byKeyNoCoating = new Map<string, PriceEngineRow>();
  const coatingsSet = new Set<string>();
  const warnings: string[] = [];
  let dupCount = 0;

  for (let r = 1; r < aoa.length; r++) {
    const row = aoa[r];
    if (!row || row.length === 0) continue;
    const qty = toNumber(row[idx["qty"]]);
    if (!qty) continue; // skip blanks

    const productName = (row[0] as string) || productNameFromHeader;
    const stock = normalizeStock(row[idx["Stock"]] as string);
    const coating = normalizeCoating(row[idx["Coating"]] as string);
    const size = normalizeSize(row[idx["size"]] as string);
    const rawTurnaround = String(row[idx["Turnaround"]] ?? "");
    const turnaround = normalizeTurnaround(rawTurnaround);
    const bundling = normalizeBundling(row[idx["Bundling"]] as string);
    const scoring = normalizeScoring(row[idx["Scoring"]] as string);
    const currentSalePrice = toNumber(row[idx["sale price"]]);
    const pe3CostTotal = toNumber(row[idx["PE3 cost no markup"]]);
    const consolidatedMarkup = toNumber(row[idx["Consolidated Markup"]]);

    const breakdownCells = row.slice(breakdownStart);
    const { baseCost, finCost } = extractCostBreakdown(breakdownCells);

    const key = lookupKey({ stock, coating, size, qty, turnaround, bundling, scoring });
    const keyNC = lookupKey({ stock, coating: "", size, qty, turnaround, bundling, scoring });
    const peRow: PriceEngineRow = {
      productName: String(productName).trim(),
      qty,
      stock,
      coating,
      size,
      turnaround,
      rawTurnaround,
      bundling,
      scoring,
      currentSalePrice,
      pe3CostTotal,
      consolidatedMarkup,
      baseCost,
      finCost,
      key,
    };
    if (byKey.has(key)) dupCount += 1;
    byKey.set(key, peRow);
    byKeyNoCoating.set(keyNC, peRow);
    if (coating) coatingsSet.add(coating);
    rows.push(peRow);
  }

  if (dupCount > 0) {
    warnings.push(`${dupCount} duplicate SKU keys in price engine — last value wins.`);
  }
  if (rows.length === 0) {
    throw new Error("Price engine parsed but contained no data rows.");
  }

  const productName = rows[0]?.productName || productNameFromHeader || sheetName;
  const coatings = Array.from(coatingsSet).sort();
  if (coatings.length > 1) {
    warnings.push(
      `Price engine contains multiple coatings (${coatings.join(", ")}). Order-replay lookups assume one coating per file; last-row coating wins per SKU.`
    );
  }
  return {
    productName,
    productSlug: slugify(productName),
    rows,
    byKey,
    byKeyNoCoating,
    coatings,
    warnings,
  };
}
