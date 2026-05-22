import { computeNewPrice, SCENARIO_BY_ID } from "./markupEngine";
import { OrderReplayData, orderRowKey } from "./parseOrderReplay";
import { PriceEngineData, indexPriceEngine } from "./parsePriceEngine";
import { commonDimensions } from "./computeScenarios";

export interface SelfCheckResult {
  scenarioAComputedDelta: number;
  scenarioAReplayDelta: number;
  aggregateDiff: number;
  maxRowDiff: number;
  rowsCompared: number;
  rowsMissed: number;
  ok: boolean;
  message: string;
}

const AGG_TOLERANCE = 50;
const ROW_TOLERANCE = 1.0;

export function runSelfCheck(
  order: OrderReplayData,
  pe: PriceEngineData
): SelfCheckResult {
  if (!order.hasPrecomputedScenarioA) {
    return {
      scenarioAComputedDelta: 0,
      scenarioAReplayDelta: 0,
      aggregateDiff: 0,
      maxRowDiff: 0,
      rowsCompared: 0,
      rowsMissed: 0,
      ok: true,
      message:
        "Self-check skipped — per-order detail input has no pre-computed Scenario A column to diff against.",
    };
  }
  const scenarioA = SCENARIO_BY_ID["A_Current_Locked"];
  const common = commonDimensions(pe, order);
  const peIndex = indexPriceEngine(pe, common);

  let totalDelta = 0;
  let maxRowDiff = 0;
  let rowsCompared = 0;
  let rowsMissed = 0;

  for (const r of order.rows) {
    const match = peIndex.get(orderRowKey(r, common));
    if (!match) {
      rowsMissed += 1;
      continue;
    }
    const res = computeNewPrice(
      {
        qty: r.qty,
        baseCost: match.baseCost,
        finCost: match.finCost,
        isRush: r.turnaround === "Rush (NBD)",
        currentSalePrice: match.currentSalePrice,
      },
      scenarioA.grad
    );
    const newRev = res.finalPrice * r.orders;
    totalDelta += newRev - r.avgPaid * r.orders;
    if (r.newPricePerOrder > 0) {
      const diff = Math.abs(res.finalPrice - r.newPricePerOrder);
      if (diff > maxRowDiff) maxRowDiff = diff;
    }
    rowsCompared += 1;
  }

  const replayDelta = order.totals.deltaAtScenarioA;
  const aggDiff = totalDelta - replayDelta;
  const ok = Math.abs(aggDiff) <= AGG_TOLERANCE && maxRowDiff <= ROW_TOLERANCE;

  let message: string;
  if (ok) {
    message = `Scenario A matches replay within ±$${ROW_TOLERANCE.toFixed(2)} per row.`;
  } else if (rowsMissed > rowsCompared) {
    message =
      "Most rows could not be matched between price engine and replay. Likely product/coating mismatch.";
  } else {
    message = `Scenario A drift detected: aggregate diff ${aggDiff.toFixed(2)} (max row diff ${maxRowDiff.toFixed(2)}). Likely a label normalization issue.`;
  }

  return {
    scenarioAComputedDelta: totalDelta,
    scenarioAReplayDelta: replayDelta,
    aggregateDiff: aggDiff,
    maxRowDiff,
    rowsCompared,
    rowsMissed,
    ok,
    message,
  };
}
