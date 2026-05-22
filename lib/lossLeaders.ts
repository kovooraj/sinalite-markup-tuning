import { OrderReplayData, OrderReplayRow, orderRowKey } from "./parseOrderReplay";
import {
  PriceEngineData,
  indexPriceEngine,
  isBaseRow,
  pickByPriceProximity,
} from "./parsePriceEngine";
import { commonDimensions, useStockInKey } from "./computeScenarios";

export type Verdict = "JUSTIFIED" | "MARGINAL" | "BORDERLINE" | "NOT JUSTIFIED";

export interface LossLeaderRow {
  sizeQty: string;
  size: string;
  qty: number;
  baseMarginUsd: number;
  addonMarginUsd: number;
  netUsd: number;
  baseOrders: number;
  verdict: Verdict;
  action: string;
}

export interface LossLeadersOutput {
  rows: LossLeaderRow[];
  totals: {
    baseMarginUsd: number;
    addonMarginUsd: number;
    netUsd: number;
  };
  callout: {
    sizeFamily: string;
    qtys: number[];
    totalLoss: number;
  } | null;
}

function verdictOf(net: number): { verdict: Verdict; action: string } {
  if (net >= 500) return { verdict: "JUSTIFIED", action: "Keep" };
  if (net >= 0) return { verdict: "MARGINAL", action: "Monitor" };
  if (net >= -200) return { verdict: "BORDERLINE", action: "Review Q3" };
  return { verdict: "NOT JUSTIFIED", action: "Raise / kill" };
}

function fmtQtySimple(q: number): string {
  return q.toLocaleString("en-US");
}

/** Decide whether an order row is "base" (no add-on dimensions set) using the
 * order's own dims and the price engine's add-on dimension list. */
function isOrderBaseRow(r: OrderReplayRow): boolean {
  // Use the same is-base logic across products: any add-on-style dimension
  // (bundling, scoring, finishing, etc.) is empty OR carries a "no/none" sentinel.
  const addonDims = [
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
  ];
  for (const d of addonDims) {
    const v = r.dims[d];
    if (!v) continue;
    const lc = v.toLowerCase();
    if (lc.includes("no bundling") || lc === "none" || lc === "free" || lc === "no")
      continue;
    return false;
  }
  return true;
}

export function computeLossLeaders(
  order: OrderReplayData,
  pe: PriceEngineData
): LossLeadersOutput {
  const common = commonDimensions(pe, order);
  const useStock = useStockInKey(pe, order);
  const { buckets } = indexPriceEngine(pe, common, useStock);

  type Agg = {
    size: string;
    qty: number;
    baseOrders: number;
    baseMargin: number;
    addonOrders: number;
    addonMargin: number;
  };
  const map = new Map<string, Agg>();
  function get(size: string, qty: number): Agg {
    const k = `${size}|${qty}`;
    let a = map.get(k);
    if (!a) {
      a = { size, qty, baseOrders: 0, baseMargin: 0, addonOrders: 0, addonMargin: 0 };
      map.set(k, a);
    }
    return a;
  }

  for (const r of order.rows) {
    if (r.turnaround !== "Standard") continue;
    const isBase = isOrderBaseRow(r);
    const agg = get(r.size, r.qty);
    const peRow = pickByPriceProximity(
      buckets.get(orderRowKey(r, common, useStock)),
      r.avgPaid
    );
    const variantCost = peRow ? peRow.baseCost + peRow.finCost : r.baseCost;
    const marginPerOrder = r.avgPaid - variantCost;
    if (isBase) {
      agg.baseOrders += r.orders;
      agg.baseMargin += marginPerOrder * r.orders;
    } else {
      agg.addonOrders += r.orders;
      agg.addonMargin += marginPerOrder * r.orders;
    }
  }

  const all = Array.from(map.values()).filter((a) => a.baseOrders > 0);
  all.sort((a, b) => b.baseOrders - a.baseOrders);
  const top = all.slice(0, 20);
  top.sort((a, b) => b.baseMargin + b.addonMargin - (a.baseMargin + a.addonMargin));

  const rows: LossLeaderRow[] = top.map((a) => {
    const net = a.baseMargin + a.addonMargin;
    const { verdict, action } = verdictOf(net);
    return {
      sizeQty: `${a.size} / ${fmtQtySimple(a.qty)}`,
      size: a.size,
      qty: a.qty,
      baseMarginUsd: a.baseMargin,
      addonMarginUsd: a.addonMargin,
      netUsd: net,
      baseOrders: a.baseOrders,
      verdict,
      action,
    };
  });

  const totals = rows.reduce(
    (acc, r) => {
      acc.baseMarginUsd += r.baseMarginUsd;
      acc.addonMarginUsd += r.addonMarginUsd;
      acc.netUsd += r.netUsd;
      return acc;
    },
    { baseMarginUsd: 0, addonMarginUsd: 0, netUsd: 0 }
  );

  const lossesBySize = new Map<string, { totalLoss: number; qtys: Set<number> }>();
  for (const r of rows) {
    if (r.verdict !== "NOT JUSTIFIED") continue;
    let agg = lossesBySize.get(r.size);
    if (!agg) {
      agg = { totalLoss: 0, qtys: new Set() };
      lossesBySize.set(r.size, agg);
    }
    agg.totalLoss += r.netUsd;
    agg.qtys.add(r.qty);
  }
  let callout: LossLeadersOutput["callout"] = null;
  let worstFamily: string | null = null;
  let worstLoss = 0;
  for (const [size, info] of lossesBySize) {
    if (info.qtys.size >= 2 && info.totalLoss < worstLoss) {
      worstLoss = info.totalLoss;
      worstFamily = size;
    }
  }
  if (worstFamily) {
    const info = lossesBySize.get(worstFamily)!;
    callout = {
      sizeFamily: worstFamily,
      qtys: Array.from(info.qtys).sort((a, b) => a - b),
      totalLoss: info.totalLoss,
    };
  }

  return { rows, totals, callout };
}

// re-export to avoid unused-import warnings on isBaseRow if future consumers
// reach for it from here
export { isBaseRow };
