import { computeNewPrice, SCENARIOS, ScenarioDef } from "./markupEngine";
import { OrderReplayRow, OrderReplayData, orderRowKey } from "./parseOrderReplay";
import {
  PriceEngineData,
  PriceEngineRow,
  indexPriceEngine,
  pickByPriceProximity,
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
  /** Dimensions used for matching — intersection of the two files. */
  commonDimensions: string[];
  /** Number of PE3 rows that collided on the same matching key. Non-zero
   * means the index averaged costs across multiple variants — surface to
   * the user so they know matches are approximate. */
  peCollisions: number;
  useStockInKey: boolean;
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

export function computeAllScenarios(
  order: OrderReplayData,
  pe: PriceEngineData
): ScenariosOutput {
  const common = commonDimensions(pe, order);
  const useStock = useStockInKey(pe, order);
  const { buckets, collisions } = indexPriceEngine(pe, common, useStock);

  const resolved: Array<{ row: OrderReplayRow; pe: PriceEngineRow | null }> = [];
  let matched = 0;
  let unmatched = 0;
  const unmatchedSamples: string[] = [];
  for (const r of order.rows) {
    const key = orderRowKey(r, common, useStock);
    const peRow = pickByPriceProximity(buckets.get(key), r.avgPaid);
    resolved.push({ row: r, pe: peRow });
    if (peRow) matched += 1;
    else {
      unmatched += 1;
      if (unmatchedSamples.length < 10) unmatchedSamples.push(r.description);
    }
  }

  const results: ScenarioResult[] = [];
  for (const s of SCENARIOS) {
    results.push(runScenario(s, resolved));
  }

  return {
    scenarios: results,
    matchedOrderRows: matched,
    unmatchedOrderRows: unmatched,
    unmatchedSamples,
    commonDimensions: common,
    peCollisions: collisions,
    useStockInKey: useStock,
  };
}

function runScenario(
  s: ScenarioDef,
  resolved: Array<{ row: OrderReplayRow; pe: PriceEngineRow | null }>
): ScenarioResult {
  let totalDelta = 0;
  let totalPaid = 0;
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
      s.grad
    );
    const newRev = r.finalPrice * row.orders;
    const paidRev = row.avgPaid * row.orders;
    const delta = newRev - paidRev;
    totalDelta += delta;
    totalNewRev += newRev;
    totalPaid += paidRev;
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
    if (row.avgPaid > 0) {
      const pct = (r.finalPrice - row.avgPaid) / row.avgPaid;
      dist[classifyPct(pct)] += row.orders;
    } else {
      dist.noChange += row.orders;
    }
  }

  const pctDelta = totalPaid > 0 ? totalDelta / totalPaid : 0;
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
