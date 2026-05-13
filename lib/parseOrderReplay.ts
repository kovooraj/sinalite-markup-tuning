import * as XLSX from "xlsx";
import {
  lookupKey,
  normalizeBundling,
  normalizeScoring,
  normalizeSize,
  normalizeStock,
  normalizeTurnaround,
} from "./normalize";

export interface OrderReplayRow {
  description: string;
  qty: number;
  qtyBandLabel: string;
  turnaround: "Standard" | "Rush (NBD)";
  stock: string;
  size: string;
  bundling: string;
  scoring: string;
  orders: number;
  actualPaid: number;
  avgPaid: number;
  newPricePerOrder: number;
  newRevenue: number;
  delta: number;
  pctChange: number;
  baseCost: number;
  baseSP: number;
  variantSP: number;
  /** Lookup key matching the price engine. Coating is unknown from this sheet
   * so it's empty here and the price-engine match will need a coating-agnostic
   * fallback. */
  keyNoCoating: string;
}

export interface OrderReplayData {
  rows: OrderReplayRow[];
  totals: {
    orders: number;
    actualPaid: number;
    newRevenue: number;
    deltaAtScenarioA: number;
  };
  unmatched: { count: number; revenue: number; reasons: string[] };
  warnings: string[];
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

const EXPECTED_COLS = [
  "Description",
  "Qty",
  "Qty Band",
  "Turnaround",
  "Stock",
  "Size",
  "Bundling",
  "Scoring",
  "Orders",
  "Actual Paid $",
  "Avg Paid $",
  "New Price/Order",
  "New Revenue $",
  "Delta $",
  "% Change",
  "Base Cost",
  "Base SP",
  "Variant SP",
];

export async function parseOrderReplay(file: File): Promise<OrderReplayData> {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });
  if (!wb.SheetNames.includes("Per-SKU Detail")) {
    throw new Error(
      `Order replay missing "Per-SKU Detail" sheet. Sheets found: ${wb.SheetNames.join(", ")}`
    );
  }
  const ws = wb.Sheets["Per-SKU Detail"];
  const aoa = XLSX.utils.sheet_to_json<(string | number | null)[]>(ws, {
    header: 1,
    raw: true,
    defval: null,
  });
  if (aoa.length < 2) throw new Error('"Per-SKU Detail" sheet has no data rows.');
  const header = aoa[0];
  const idx: Record<string, number> = {};
  for (const name of EXPECTED_COLS) {
    const i = header.findIndex(
      (h) => typeof h === "string" && h.trim().toLowerCase() === name.toLowerCase()
    );
    if (i < 0) throw new Error(`Order replay "Per-SKU Detail" missing column "${name}".`);
    idx[name] = i;
  }

  const rows: OrderReplayRow[] = [];
  let totalOrders = 0;
  let totalPaid = 0;
  let totalNewRev = 0;

  for (let r = 1; r < aoa.length; r++) {
    const row = aoa[r];
    if (!row || row.length === 0) continue;
    const qty = toNumber(row[idx["Qty"]]);
    if (!qty) continue;
    const description = toStr(row[idx["Description"]]);
    const qtyBandLabel = toStr(row[idx["Qty Band"]]);
    const turnaround = normalizeTurnaround(toStr(row[idx["Turnaround"]]));
    const stock = normalizeStock(toStr(row[idx["Stock"]]));
    const size = normalizeSize(toStr(row[idx["Size"]]));
    const bundling = normalizeBundling(toStr(row[idx["Bundling"]]));
    const scoring = normalizeScoring(toStr(row[idx["Scoring"]]));
    const orders = toNumber(row[idx["Orders"]]);
    const actualPaid = toNumber(row[idx["Actual Paid $"]]);
    const avgPaid = toNumber(row[idx["Avg Paid $"]]);
    const newPricePerOrder = toNumber(row[idx["New Price/Order"]]);
    const newRevenue = toNumber(row[idx["New Revenue $"]]);
    const delta = toNumber(row[idx["Delta $"]]);
    const pctChange = toNumber(row[idx["% Change"]]);
    const baseCost = toNumber(row[idx["Base Cost"]]);
    const baseSP = toNumber(row[idx["Base SP"]]);
    const variantSP = toNumber(row[idx["Variant SP"]]);

    const keyNoCoating = lookupKey({
      stock,
      coating: "",
      size,
      qty,
      turnaround,
      bundling,
      scoring,
    });

    rows.push({
      description,
      qty,
      qtyBandLabel,
      turnaround,
      stock,
      size,
      bundling,
      scoring,
      orders,
      actualPaid,
      avgPaid,
      newPricePerOrder,
      newRevenue,
      delta,
      pctChange,
      baseCost,
      baseSP,
      variantSP,
      keyNoCoating,
    });

    totalOrders += orders;
    totalPaid += actualPaid;
    totalNewRev += newRevenue;
  }

  let unmatchedCount = 0;
  let unmatchedRev = 0;
  const unmatchedReasonsSet = new Set<string>();
  if (wb.SheetNames.includes("Unmatched")) {
    const wsU = wb.Sheets["Unmatched"];
    const aoaU = XLSX.utils.sheet_to_json<(string | number | null)[]>(wsU, {
      header: 1,
      defval: null,
    });
    // header at row 2 typically; iterate looking for plausible rows
    let headerRowU = -1;
    for (let i = 0; i < Math.min(aoaU.length, 5); i++) {
      const r = aoaU[i];
      if (
        r &&
        r.some((c) => typeof c === "string" && c.toLowerCase().includes("orders")) &&
        r.some((c) => typeof c === "string" && c.toLowerCase().includes("revenue"))
      ) {
        headerRowU = i;
        break;
      }
    }
    if (headerRowU >= 0) {
      const h = aoaU[headerRowU];
      const orderIdx = h.findIndex(
        (c) => typeof c === "string" && c.toLowerCase().includes("orders")
      );
      const revIdx = h.findIndex(
        (c) => typeof c === "string" && c.toLowerCase().includes("revenue")
      );
      const reasonIdx = h.findIndex(
        (c) => typeof c === "string" && c.toLowerCase().includes("reason")
      );
      for (let i = headerRowU + 1; i < aoaU.length; i++) {
        const r = aoaU[i];
        if (!r) continue;
        const o = toNumber(r[orderIdx]);
        const v = toNumber(r[revIdx]);
        if (!o && !v) continue;
        unmatchedCount += o;
        unmatchedRev += v;
        if (reasonIdx >= 0) {
          const reason = toStr(r[reasonIdx]);
          if (reason) unmatchedReasonsSet.add(reason);
        }
      }
    }
  }

  return {
    rows,
    totals: {
      orders: totalOrders,
      actualPaid: totalPaid,
      newRevenue: totalNewRev,
      deltaAtScenarioA: totalNewRev - totalPaid,
    },
    unmatched: {
      count: unmatchedCount,
      revenue: unmatchedRev,
      reasons: Array.from(unmatchedReasonsSet),
    },
    warnings: [],
  };
}
