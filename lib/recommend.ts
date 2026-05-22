import { ScenarioResult } from "./computeScenarios";

export interface RecommendationOpts {
  targetMinPct: number; // e.g. -0.01 = -1%
  targetMaxPct: number; // e.g. 0.005 = +0.5%
}

export interface Recommendation {
  inBand: ScenarioResult[];
  outOfBand: ScenarioResult[];
  recommended: ScenarioResult | null;
  reason: string;
}

export function recommend(
  scenarios: ScenarioResult[],
  opts: RecommendationOpts
): Recommendation {
  const inBand = scenarios.filter(
    (s) => s.pctDelta >= opts.targetMinPct && s.pctDelta <= opts.targetMaxPct
  );
  const outOfBand = scenarios.filter(
    (s) => !(s.pctDelta >= opts.targetMinPct && s.pctDelta <= opts.targetMaxPct)
  );

  let recommended: ScenarioResult | null = null;
  let reason: string;

  // Implements the sinalite-pricing-model SOP recommendation heuristic
  // (references/markup-tables.md, "Recommendation Heuristic"):
  //   1. If finance accepts the band and A is in band → ship A (locked
  //      decision, predictable customer impact).
  //   2. If A is out of band but the user wants revenue-neutral-or-positive
  //      → prefer F_Aggressive, then E_Combined_Mod, then any other in-band.
  //   3. Otherwise, fall back to the closest scenario to the band.
  const a = scenarios.find((s) => s.id === "A_Current_Locked");
  const aInBand =
    !!a && a.pctDelta >= opts.targetMinPct && a.pctDelta <= opts.targetMaxPct;
  if (aInBand) {
    recommended = a!;
    reason = `A_Current_Locked is in target band (${(a!.pctDelta * 100).toFixed(2)}%) — SOP default per sinalite-pricing-model: locked decision, customer impact most predictable.`;
  } else if (inBand.length > 0) {
    const priority = ["F_Aggressive", "E_Combined_Mod"];
    const preferred = priority
      .map((id) => inBand.find((s) => s.id === id))
      .find((s): s is ScenarioResult => !!s);
    recommended = preferred ?? inBand[0];
    const pctStr = `${(recommended.pctDelta * 100).toFixed(2)}%`;
    if (recommended.id === "F_Aggressive") {
      reason = `A_Current_Locked is outside target. F_Aggressive (${pctStr}) hits the band — SOP pick when finance wants revenue-positive.`;
    } else if (recommended.id === "E_Combined_Mod") {
      reason = `A_Current_Locked is outside target. E_Combined_Mod (${pctStr}) hits the band with minimal customer-impact change — SOP pick when finance wants revenue-positive with smaller customer impact shift.`;
    } else {
      reason = `${recommended.id} (${pctStr}) is in target band; A is outside.`;
    }
  } else {
    // Nothing in band — pick the scenario closest to the band.
    let best = scenarios[0];
    let bestDist = bandDistance(best.pctDelta, opts);
    for (const s of scenarios) {
      const d = bandDistance(s.pctDelta, opts);
      if (d < bestDist) {
        best = s;
        bestDist = d;
      }
    }
    recommended = best;
    reason = `No scenario fully within target — closest is ${best.id} (${(best.pctDelta * 100).toFixed(2)}%).`;
  }

  return { inBand, outOfBand, recommended, reason };
}

function bandDistance(pct: number, opts: RecommendationOpts): number {
  if (pct < opts.targetMinPct) return opts.targetMinPct - pct;
  if (pct > opts.targetMaxPct) return pct - opts.targetMaxPct;
  return 0;
}
