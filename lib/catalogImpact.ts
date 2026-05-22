import { computeNewPrice, SCENARIO_BY_ID } from "./markupEngine";
import { PriceEngineData, indexPriceEngine, isBaseRow } from "./parsePriceEngine";
import { ScenarioResult } from "./computeScenarios";
import { bandOf, BASE_BAND_LABELS } from "./qtyBands";
import { buildKey } from "./dimensions";

export interface BaseCatalogImpact {
  impactUsd: number;
  pct: number;
  totalCurrentSP: number;
  rowsScanned: number;
  deltaByBand: [number, number, number, number];
  concentrationLabel: string;
}

export interface FinishingCatalogImpact {
  impactUsd: number;
  pct: number;
  cellsUplifted: number;
  cellsReduced: number;
  rowsScanned: number;
}

export interface NbdLift {
  annualizedUsd: number;
  lowUsd: number;
  highUsd: number;
  orders3mo: number;
  baseRevenue3mo: number;
}

export function computeBaseCatalogImpact(pe: PriceEngineData): BaseCatalogImpact {
  const A = SCENARIO_BY_ID["A_Current_Locked"];
  const baseRows = pe.rows.filter(
    (r) => r.turnaround === "Standard" && isBaseRow(r, pe.dimensions)
  );
  let impact = 0;
  let totalCurrent = 0;
  const deltaByBand: [number, number, number, number] = [0, 0, 0, 0];
  for (const r of baseRows) {
    if (r.currentSalePrice <= 0) continue;
    const res = computeNewPrice(
      {
        qty: r.qty,
        baseCost: r.baseCost,
        finCost: 0,
        isRush: false,
        currentSalePrice: r.currentSalePrice,
      },
      A.grad
    );
    const delta = res.finalPrice - r.currentSalePrice;
    impact += delta;
    totalCurrent += r.currentSalePrice;
    deltaByBand[bandOf(r.qty)] += delta;
  }

  let maxLossIdx = 0;
  let maxLoss = Infinity;
  for (let i = 0; i < 4; i++) {
    if (deltaByBand[i] < maxLoss) {
      maxLoss = deltaByBand[i];
      maxLossIdx = i;
    }
  }
  let concentrationLabel = "";
  if (maxLoss < 0 && Math.abs(maxLoss) >= 0.5 * Math.abs(impact || 1)) {
    const labels = ["100–1k", "1k–5k", "5k–25k", "25k+"];
    concentrationLabel = `Concentrated in ${labels[maxLossIdx]} band.`;
  }

  const pct = totalCurrent > 0 ? impact / totalCurrent : 0;

  return {
    impactUsd: impact,
    pct,
    totalCurrentSP: totalCurrent,
    rowsScanned: baseRows.length,
    deltaByBand,
    concentrationLabel,
  };
}

export function computeFinishingCatalogImpact(
  pe: PriceEngineData
): FinishingCatalogImpact {
  const A = SCENARIO_BY_ID["A_Current_Locked"];
  // Need to look up the base counterpart of each variant row. Build an index
  // of base rows keyed on (stock, size, qty, turnaround). Use ALL of the price
  // engine's dimensions for the variant key (since both sides are PE rows),
  // but use NONE for the base-counterpart key — i.e. strip dimensions to find
  // the no-add-on row at the same stock/size/qty/turnaround.
  const baseIndex = new Map<string, typeof pe.rows[number]>();
  for (const r of pe.rows) {
    if (r.turnaround !== "Standard") continue;
    if (!isBaseRow(r, pe.dimensions)) continue;
    const k = buildKey(
      { stock: r.stock, size: r.size, qty: r.qty, turnaround: r.turnaround },
      {},
      []
    );
    baseIndex.set(k, r);
  }

  let impact = 0;
  let totalCurrentPremium = 0;
  let uplifted = 0;
  let reduced = 0;
  let scanned = 0;
  for (const r of pe.rows) {
    if (r.turnaround !== "Standard") continue;
    if (isBaseRow(r, pe.dimensions)) continue;
    if (r.currentSalePrice <= 0) continue;
    const baseK = buildKey(
      { stock: r.stock, size: r.size, qty: r.qty, turnaround: r.turnaround },
      {},
      []
    );
    const baseRow = baseIndex.get(baseK);
    if (!baseRow || baseRow.currentSalePrice <= 0) continue;
    const currentPremium = r.currentSalePrice - baseRow.currentSalePrice;
    if (currentPremium <= 0) continue;
    const newPremium = r.finCost * (1 + A.grad.fin[bandOf(r.qty)]);
    const delta = newPremium - currentPremium;
    impact += delta;
    totalCurrentPremium += currentPremium;
    scanned += 1;
    if (delta > 0) uplifted += 1;
    else if (delta < 0) reduced += 1;
  }
  const pct = totalCurrentPremium > 0 ? impact / totalCurrentPremium : 0;
  return {
    impactUsd: impact,
    pct,
    cellsUplifted: uplifted,
    cellsReduced: reduced,
    rowsScanned: scanned,
  };
}

export function computeNbdLift(scenarioA: ScenarioResult): NbdLift {
  const rushDelta3mo = scenarioA.deltaByTurnaround.rush;
  const annualized = rushDelta3mo * 4;
  return {
    annualizedUsd: annualized,
    lowUsd: annualized * 0.85,
    highUsd: annualized * 1.15,
    orders3mo: scenarioA.ordersByTurnaround.rush,
    baseRevenue3mo: scenarioA.paidByTurnaround.rush,
  };
}

// avoid unused warning on indexPriceEngine if not used elsewhere
void indexPriceEngine;
export const BAND_LABELS_OUT = BASE_BAND_LABELS;
