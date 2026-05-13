import { OrderReplayData } from "./parseOrderReplay";
import { PriceEngineData } from "./parsePriceEngine";
import { lookupKey } from "./normalize";

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
  // The largest concentration of NOT JUSTIFIED losses, for the callout line
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

function fmtQty(q: number): string {
  if (q >= 1000) return `${(q / 1000).toFixed(q % 1000 ? 1 : 0).replace(/\.0$/, "")},000`.replace("k,000", "k").replace("000,000", "0,000");
  return q.toString();
}

function fmtQtySimple(q: number): string {
  return q.toLocaleString("en-US");
}

export function computeLossLeaders(
  order: OrderReplayData,
  pe: PriceEngineData
): LossLeadersOutput {
  // Aggregate by (size, qty) — separately for base vs variant.
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
    // We only consider Standard turnaround orders for loss-leader analysis;
    // rush orders carry different economics and are evaluated separately in
    // the NBD lift narrative.
    if (r.turnaround !== "Standard") continue;
    const isBase = r.bundling === "No bundling - FREE" && r.scoring === "None";
    const agg = get(r.size, r.qty);
    // Margin per order = (avgPaid - variantCost) where variantCost comes from
    // the price engine. The replay's "Base Cost" column is base-only; for the
    // variant we look up baseCost + finCost.
    const peRow = pe.byKeyNoCoating.get(r.keyNoCoating);
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

  // Sort all aggregates by base order volume descending; take top 20 that have
  // at least 1 base order (to be visible to reps).
  const all = Array.from(map.values()).filter((a) => a.baseOrders > 0);
  all.sort((a, b) => b.baseOrders - a.baseOrders);
  const top = all.slice(0, 20);
  // Now sort the chosen 20 by net margin descending (matches reference doc).
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

  // Callout: find the size family carrying the biggest concentration of
  // NOT JUSTIFIED losses across multiple qty breaks.
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
