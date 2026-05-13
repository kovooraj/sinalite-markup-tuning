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

  if (inBand.length === 0) {
    // Pick the scenario closest to the band (smallest distance from the band)
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
  } else {
    // Prefer the one with smallest absolute distribution drift vs scenario A
    const a = scenarios.find((s) => s.id === "A_Current_Locked");
    if (!a) {
      recommended = inBand[0];
      reason = `${recommended.id} is in target band.`;
    } else {
      let best = inBand[0];
      let bestDrift = distDrift(best, a);
      for (const s of inBand) {
        const d = distDrift(s, a);
        if (d < bestDrift) {
          best = s;
          bestDrift = d;
        }
      }
      recommended = best;
      reason = `${best.id} is in target band with the smallest customer-impact shift vs A.`;
    }
  }

  return { inBand, outOfBand, recommended, reason };
}

function bandDistance(pct: number, opts: RecommendationOpts): number {
  if (pct < opts.targetMinPct) return opts.targetMinPct - pct;
  if (pct > opts.targetMaxPct) return pct - opts.targetMaxPct;
  return 0;
}

function distDrift(s: ScenarioResult, a: ScenarioResult): number {
  const keys: Array<keyof ScenarioResult["dist"]> = [
    "noChange",
    "decrease1to5",
    "decreaseGt5",
    "increase1to5",
    "increaseGt5",
  ];
  let sum = 0;
  for (const k of keys) sum += Math.abs(s.dist[k] - a.dist[k]);
  return sum;
}
