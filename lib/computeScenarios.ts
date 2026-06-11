import { computeNewPrice, SCENARIOS, ScenarioDef } from "./markupEngine";
import { OrderReplayRow, OrderReplayData, orderRowKey } from "./parseOrderReplay";
import {
  PriceEngineData,
  PriceEngineRow,
  indexPriceEngine,
  pickByPriceProximity,
  indexPriceEngineNoQty,
  orderRowKeyNoQty,
  pickByClosestQty,
} from "./parsePriceEngine";
import { bandOf } from "./qtyBands";

export interface ScenarioResult {
  id: string;
  label: string;
  baseFormatted: string;
  finFormatted: string;
  nbdFormatted: string;
  deltaUsd: number;
  pctDelta: number;
  annualizedUsd: number;
  dist: {
    noChange: number;
    decrease1to5: number;
    decreaseGt5: number;
    increase1to5: number;
    increaseGt5: number;
  };
  deltaByBand: [number, number, number, number];
  deltaByTurnaround: { standard: number; rush: number };
  ordersByTurnaround: { standard: number; rush: number };
  paidByTurnaround: { standard: number; rush: number };
  newRevByTurnaround: { standard: number; rush: number };
  totalOrders: number;
  totalPaid: number;
  totalNewRev: number;
  matchedRows: number;
  unmatchedRows: number;
}

export interface ScenariosOutput {
  scenarios: ScenarioResult[];
  matchedOrderRows: number;
  unmatchedOrderRows: number;
  unmatchedSamples: string[];
  /** Dimensions used for matching — intersection of the two files, possibly
   * trimmed by the progressive key-relaxation pass when the strict key
   * produced zero matches. */
  commonDimensions: string[];
  /** Number of PE3 rows that collided on the same matching key. Non-zero
   * means the index averaged costs across multiple variants — surface to
   * the user so they know matches are approximate. */
  peCollisions: number;
  useStockInKey: boolean;
  /** Dimensions the relaxation pass dropped (in the order they were dropped)
   * because the strict key produced no matches. Empty when the full
   * intersection matched cleanly. */
  droppedDimensions: string[];
  /** True if Stock was dropped from the matching key by the relaxation
   * pass — it was set initially but produced 0 matches at first. */
  droppedStock: boolean;
  /** Count of order rows resolved by exact key match. */
  matchedExact: number;
  /** Count of order rows resolved by closest-qty fallback. */
  matchedSnapped: number;
  /** Per-row resolved tuples — used by the Annotated Orders writer to
   * surface match quality per row. */
  resolved: ResolvedRow[];
}

function classifyPct(p: number): keyof ScenarioResult["dist"] {
  if (Math.abs(p) <= 0.01) return "noChange";
  if (p < 0 && p >= -0.05) return "decrease1to5";
  if (p < -0.05) return "decreaseGt5";
  if (p > 0.01 && p <= 0.05) return "increase1to5";
  return "increaseGt5";
}

export function commonDimensions(
  pe: PriceEngineData,
  order: OrderReplayData
): string[] {
  const peSet = new Set(pe.dimensions);
  return order.dimensions.filter((d) => peSet.has(d));
}

/** Both files must have Stock for it to be used in the matching key. If
 * either side lacks Stock, we collapse it to "" on both sides so they still
 * match on the remaining fields. */
export function useStockInKey(
  pe: PriceEngineData,
  order: OrderReplayData
): boolean {
  return pe.hasStock && order.hasStock;
}

/** Match quality for a single resolved order row. */
export type MatchQuality = "exact" | "snapped-qty" | "no-match";

export interface ResolvedRow {
  row: OrderReplayRow;
  pe: PriceEngineRow | null;
  matchQuality: MatchQuality;
  /** When matchQuality === "snapped-qty", the PE3 qty we snapped to. */
  snappedFromQty: number | null;
}

/** Try to match each order against the price engine, progressively relaxing
 * the matching key until a non-trivial fraction of orders match. After the
 * key is chosen, any remaining unmatched rows fall back to closest-qty
 * lookup (same stock/size/turnaround/dims, nearest qty). Returns the
 * resolved tuples along with metadata describing which dimensions (and/or
 * stock) were dropped. */
function progressivelyMatch(
  order: OrderReplayData,
  pe: PriceEngineData
): {
  resolved: ResolvedRow[];
  matched: number;
  matchedExact: number;
  matchedSnapped: number;
  unmatched: number;
  unmatchedSamples: string[];
  commonDims: string[];
  droppedDimensions: string[];
  droppedStock: boolean;
  collisions: number;
} {
  const initialDims = commonDimensions(pe, order);
  const initialUseStock = useStockInKey(pe, order);
  // Target match rate above which we accept the current key
  const MIN_MATCH_RATE = 0.1;

  // Attempt sequence: start strict, drop one dimension at a time, then drop
  // stock if still 0 matches.
  // Order of dropping: from highest-cardinality / least-likely-to-overlap.
  // Heuristic: try the original "common" set first, then drop dims in reverse
  // alphabetical order so deterministic (and longer-string dims tend to be
  // more product-specific like "binding", "pages").
  const dropOrder: { drop: string[]; stock: boolean }[] = [];
  // Try the full key first
  dropOrder.push({ drop: [], stock: initialUseStock });
  // Then drop dims one at a time, accumulating
  const dimsToTry = [...initialDims].sort();
  for (let i = 1; i <= dimsToTry.length; i++) {
    dropOrder.push({ drop: dimsToTry.slice(0, i), stock: initialUseStock });
  }
  // Finally try without stock (and without dims one at a time again)
  if (initialUseStock) {
    dropOrder.push({ drop: [], stock: false });
    for (let i = 1; i <= dimsToTry.length; i++) {
      dropOrder.push({ drop: dimsToTry.slice(0, i), stock: false });
    }
  }

  let bestAttempt: {
    matched: number;
    unmatched: number;
    unmatchedSamples: string[];
    commonDims: string[];
    droppedDimensions: string[];
    droppedStock: boolean;
    resolved: Array<{ row: OrderReplayRow; pe: PriceEngineRow | null }>;
    collisions: number;
    useStock: boolean;
  } | null = null;

  for (const attempt of dropOrder) {
    const dimsForThisRun = dimsToTry.filter((d) => !attempt.drop.includes(d));
    const { buckets, collisions } = indexPriceEngine(
      pe,
      dimsForThisRun,
      attempt.stock
    );
    const resolved: Array<{ row: OrderReplayRow; pe: PriceEngineRow | null }> =
      [];
    let matched = 0;
    let unmatched = 0;
    const unmatchedSamples: string[] = [];
    for (const r of order.rows) {
      const key = orderRowKey(r, dimsForThisRun, attempt.stock);
      const peRow = pickByPriceProximity(buckets.get(key), r.avgPaid);
      resolved.push({ row: r, pe: peRow });
      if (peRow) matched += 1;
      else {
        unmatched += 1;
        if (unmatchedSamples.length < 10) unmatchedSamples.push(r.description);
      }
    }
    const matchRate = order.rows.length > 0 ? matched / order.rows.length : 0;
    const candidate = {
      matched,
      unmatched,
      unmatchedSamples,
      commonDims: dimsForThisRun,
      droppedDimensions: attempt.drop,
      droppedStock: initialUseStock && !attempt.stock,
      resolved,
      collisions,
      useStock: attempt.stock,
    };
    if (!bestAttempt || matched > bestAttempt.matched) {
      bestAttempt = candidate;
    }
    if (matchRate >= MIN_MATCH_RATE) {
      bestAttempt = candidate;
      break;
    }
  }
  const chosen =
    bestAttempt ?? {
      matched: 0,
      unmatched: order.rows.length,
      unmatchedSamples: order.rows.slice(0, 10).map((r) => r.description),
      commonDims: initialDims,
      droppedDimensions: [],
      droppedStock: false,
      resolved: order.rows.map((row) => ({ row, pe: null })),
      collisions: 0,
      useStock: initialUseStock,
    };

  // Closest-qty fallback: for rows that still didn't match, look up the
  // PE3 catalog by everything-except-qty and snap to the nearest qty in
  // the same stock/size/turnaround/dims bucket. The PE3 catalog typically
  // has standard breaks (50/100/250/500/1000/2500/...); orders at custom
  // qtys (200, 600, 1500, 3500...) need this fallback to get a sensible
  // cost basis rather than being dropped.
  const noQtyBuckets = indexPriceEngineNoQty(
    pe,
    chosen.commonDims,
    chosen.useStock
  );
  const finalResolved: ResolvedRow[] = [];
  let matchedExact = 0;
  let matchedSnapped = 0;
  let stillUnmatched = 0;
  const stillUnmatchedSamples: string[] = [];

  for (const { row, pe: exact } of chosen.resolved) {
    if (exact) {
      finalResolved.push({
        row,
        pe: exact,
        matchQuality: "exact",
        snappedFromQty: null,
      });
      matchedExact += 1;
      continue;
    }
    const candidates = noQtyBuckets.get(
      orderRowKeyNoQty(row, chosen.commonDims, chosen.useStock)
    );
    const { pe: snapped, snappedFromQty } = pickByClosestQty(
      candidates,
      row.qty,
      row.avgPaid
    );
    if (snapped) {
      finalResolved.push({
        row,
        pe: snapped,
        matchQuality: "snapped-qty",
        snappedFromQty,
      });
      matchedSnapped += 1;
    } else {
      finalResolved.push({
        row,
        pe: null,
        matchQuality: "no-match",
        snappedFromQty: null,
      });
      stillUnmatched += 1;
      if (stillUnmatchedSamples.length < 10)
        stillUnmatchedSamples.push(row.description);
    }
  }

  return {
    resolved: finalResolved,
    matched: matchedExact + matchedSnapped,
    matchedExact,
    matchedSnapped,
    unmatched: stillUnmatched,
    unmatchedSamples:
      stillUnmatched > 0 ? stillUnmatchedSamples : chosen.unmatchedSamples,
    commonDims: chosen.commonDims,
    droppedDimensions: chosen.droppedDimensions,
    droppedStock: chosen.droppedStock,
    collisions: chosen.collisions,
  };
}

export function computeAllScenarios(
  order: OrderReplayData,
  pe: PriceEngineData,
  opts: { applyCapRule?: boolean; extraScenarios?: ScenarioDef[] } = {}
): ScenariosOutput {
  const applyCapRule = opts.applyCapRule ?? true;
  const result = progressivelyMatch(order, pe);

  const results: ScenarioResult[] = [];
  for (const s of [...SCENARIOS, ...(opts.extraScenarios ?? [])]) {
    results.push(runScenario(s, result.resolved, applyCapRule));
  }

  return {
    scenarios: results,
    matchedOrderRows: result.matched,
    unmatchedOrderRows: result.unmatched,
    unmatchedSamples: result.unmatchedSamples,
    commonDimensions: result.commonDims,
    peCollisions: result.collisions,
    useStockInKey: !result.droppedStock && useStockInKey(pe, order),
    droppedDimensions: result.droppedDimensions,
    droppedStock: result.droppedStock,
    matchedExact: result.matchedExact,
    matchedSnapped: result.matchedSnapped,
    resolved: result.resolved,
  };
}

function runScenario(
  s: ScenarioDef,
  resolved: ResolvedRow[],
  applyCapRule: boolean
): ScenarioResult {
  // Delta everywhere is now computed as (New Price capped) − PE3 List Price.
  // This is the catalog-repricing view: capped rows produce $0 delta (new = list),
  // uncapped rows produce a negative delta (new < list). The previous
  // realized-revenue framing (vs avgPaid) has been replaced everywhere.
  let totalDelta = 0;
  let totalList = 0; // baseline = sum of PE3 list × orders
  let totalPaid = 0; // kept for context / paid-by-turnaround stats
  let totalNewRev = 0;
  let totalOrders = 0;
  const deltaByBand: [number, number, number, number] = [0, 0, 0, 0];
  const dist = { noChange: 0, decrease1to5: 0, decreaseGt5: 0, increase1to5: 0, increaseGt5: 0 };
  const deltaByTurnaround = { standard: 0, rush: 0 };
  const ordersByTurnaround = { standard: 0, rush: 0 };
  const paidByTurnaround = { standard: 0, rush: 0 };
  const newRevByTurnaround = { standard: 0, rush: 0 };
  let matched = 0;
  let unmatched = 0;

  for (const { row, pe } of resolved) {
    if (!pe) {
      unmatched += 1;
      continue;
    }
    matched += 1;
    const r = computeNewPrice(
      {
        qty: row.qty,
        baseCost: pe.baseCost,
        finCost: pe.finCost,
        isRush: row.turnaround === "Rush (NBD)",
        currentSalePrice: pe.currentSalePrice,
      },
      s.grad,
      applyCapRule
    );
    const newRev = r.finalPrice * row.orders;
    const paidRev = row.avgPaid * row.orders;
    const listRev = pe.currentSalePrice * row.orders;
    const delta = newRev - listRev; // catalog-repricing baseline
    totalDelta += delta;
    totalNewRev += newRev;
    totalPaid += paidRev;
    totalList += listRev;
    totalOrders += row.orders;
    deltaByBand[bandOf(row.qty)] += delta;
    if (row.turnaround === "Rush (NBD)") {
      deltaByTurnaround.rush += delta;
      ordersByTurnaround.rush += row.orders;
      paidByTurnaround.rush += paidRev;
      newRevByTurnaround.rush += newRev;
    } else {
      deltaByTurnaround.standard += delta;
      ordersByTurnaround.standard += row.orders;
      paidByTurnaround.standard += paidRev;
      newRevByTurnaround.standard += newRev;
    }
    if (pe.currentSalePrice > 0) {
      const pct = (r.finalPrice - pe.currentSalePrice) / pe.currentSalePrice;
      dist[classifyPct(pct)] += row.orders;
    } else {
      dist.noChange += row.orders;
    }
  }

  const pctDelta = totalList > 0 ? totalDelta / totalList : 0;
  const distFractions = {
    noChange: totalOrders > 0 ? dist.noChange / totalOrders : 0,
    decrease1to5: totalOrders > 0 ? dist.decrease1to5 / totalOrders : 0,
    decreaseGt5: totalOrders > 0 ? dist.decreaseGt5 / totalOrders : 0,
    increase1to5: totalOrders > 0 ? dist.increase1to5 / totalOrders : 0,
    increaseGt5: totalOrders > 0 ? dist.increaseGt5 / totalOrders : 0,
  };

  return {
    id: s.id,
    label: s.label,
    baseFormatted: s.baseFormatted,
    finFormatted: s.finFormatted,
    nbdFormatted: s.nbdFormatted,
    deltaUsd: totalDelta,
    pctDelta,
    annualizedUsd: totalDelta * 4,
    dist: distFractions,
    deltaByBand,
    deltaByTurnaround,
    ordersByTurnaround,
    paidByTurnaround,
    newRevByTurnaround,
    totalOrders,
    totalPaid,
    totalNewRev,
    matchedRows: matched,
    unmatchedRows: unmatched,
  };
}
