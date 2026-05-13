import { computeNewPrice, SCENARIO_BY_ID } from "./markupEngine";
import { PriceEngineData, PriceEngineRow } from "./parsePriceEngine";
import { ScenarioResult } from "./computeScenarios";
import { bandOf, BASE_BAND_LABELS } from "./qtyBands";
import { lookupKey } from "./normalize";

export interface BaseCatalogImpact {
  impactUsd: number;
  pct: number;
  totalCurrentSP: number;
  rowsScanned: number;
  // Loss share per band, used to build the concentration narrative
  deltaByBand: [number, number, number, number];
  concentrationLabel: string; // e.g. "Concentrated in 25k+ band."
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
  baseRevenue3mo: number; // standard-revenue context
}

export function computeBaseCatalogImpact(pe: PriceEngineData): BaseCatalogImpact {
  const A = SCENARIO_BY_ID["A_Current_Locked"];
  const baseRows = pe.rows.filter(
    (r) =>
      r.bundling === "No bundling - FREE" &&
      r.scoring === "None" &&
      r.turnaround === "Standard"
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

  // Concentration: find the band with the largest absolute negative delta
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
  // For each variant row (bundling or scoring), find the matching base row
  // (same stock/coating/size/qty/turnaround=Standard, bundling=No, scoring=None).
  let impact = 0;
  let totalCurrentPremium = 0;
  let uplifted = 0;
  let reduced = 0;
  let scanned = 0;
  for (const r of pe.rows) {
    if (r.turnaround !== "Standard") continue;
    const hasFinishing = r.bundling !== "No bundling - FREE" || r.scoring !== "None";
    if (!hasFinishing) continue;
    if (r.currentSalePrice <= 0) continue;
    const baseKey = lookupKey({
      stock: r.stock,
      coating: r.coating,
      size: r.size,
      qty: r.qty,
      turnaround: "Standard",
      bundling: "No bundling - FREE",
      scoring: "None",
    });
    const baseRow = pe.byKey.get(baseKey);
    if (!baseRow || baseRow.currentSalePrice <= 0) continue;
    const currentPremium = r.currentSalePrice - baseRow.currentSalePrice;
    if (currentPremium <= 0) continue; // skip nonsense
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

export const BAND_LABELS_OUT = BASE_BAND_LABELS;
