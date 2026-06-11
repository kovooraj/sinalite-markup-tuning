import { bandOf, nbdBandOf } from "./qtyBands";

export interface Gradient {
  base: readonly [number, number, number, number]; // 100-1k, 1k-5k, 5k-25k, 25k-100k
  fin: readonly [number, number, number, number];
  nbd: readonly [number, number, number]; // 100-5k, 5k-25k, 25k-100k
}

export interface ScenarioDef {
  id: string;
  label: string;
  baseFormatted: string;
  finFormatted: string;
  nbdFormatted: string;
  grad: Gradient;
  /** Plain-English description of what this scenario does and when to pick it. */
  description: string;
}

const fmt4 = (g: readonly [number, number, number, number]) =>
  g.map((v) => `${Math.round(v * 100)}%`).join("/");
const fmt3 = (g: readonly [number, number, number]) =>
  g.map((v) => `${Math.round(v * 100)}%`).join("/");

function makeScenario(
  id: string,
  label: string,
  base: Gradient["base"],
  fin: Gradient["fin"],
  nbd: Gradient["nbd"],
  description: string
): ScenarioDef {
  return {
    id,
    label,
    baseFormatted: fmt4(base),
    finFormatted: fmt4(fin),
    nbdFormatted: fmt3(nbd),
    grad: { base, fin, nbd },
    description,
  };
}

export const SCENARIOS: readonly ScenarioDef[] = [
  makeScenario(
    "A_Current_Locked",
    "Current locked (35/30/25/20 base, 60/50/40/35 fin, 30/25/20 NBD)",
    [0.35, 0.3, 0.25, 0.2],
    [0.6, 0.5, 0.4, 0.35],
    [0.3, 0.25, 0.2],
    "The May 2, 2026 locked decision. Volume-discount gradient on base (35/30/25/20%), steeper gradient on finishing add-ons (60/50/40/35%), modest rush surcharge (30/25/20%). Operational default — customer-impact distribution is already measured and approved. Use unless finance has an explicit reason to deviate."
  ),
  makeScenario(
    "B_Lift_25kBase",
    "Lift 25k+ base to 25% (35/30/25/25)",
    [0.35, 0.3, 0.25, 0.25],
    [0.6, 0.5, 0.4, 0.35],
    [0.3, 0.25, 0.2],
    "Same as A everywhere except the 25k+ qty band's base markup is lifted from 20% to 25%. Recovers high-volume revenue without touching smaller orders. Use only if competitive concerns at high volume don't dominate."
  ),
  makeScenario(
    "C_Base_Plus5",
    "Base +5pt across board (40/35/30/25)",
    [0.4, 0.35, 0.3, 0.25],
    [0.6, 0.5, 0.4, 0.35],
    [0.3, 0.25, 0.2],
    "Lifts all four base markup bands by +5pt (40/35/30/25%). Across-the-board base price increase; finishing and NBD untouched. Use when pilot showed too much revenue loss but you only want to move the base lever."
  ),
  makeScenario(
    "D_NBD_Plus5",
    "NBD +5pt across board (35/30/25)",
    [0.35, 0.3, 0.25, 0.2],
    [0.6, 0.5, 0.4, 0.35],
    [0.35, 0.3, 0.25],
    "Same base and finishing as A, but NBD rush surcharge bumped to 35/30/25% (was 30/25/20%). Targets rush-order revenue specifically. Co-preferred with A when competitive concerns dominate high-volume pricing."
  ),
  makeScenario(
    "E_Combined_Mod",
    "Combined: Lift 25k+ base + NBD +5pt",
    [0.35, 0.3, 0.25, 0.25],
    [0.6, 0.5, 0.4, 0.35],
    [0.35, 0.3, 0.25],
    "Combines B (lift 25k+ base) with D (NBD +5pt). Modest revenue improvement, customer-impact distribution barely changes. SOP's pick when finance wants revenue-positive with minimal customer-impact shift."
  ),
  makeScenario(
    "F_Aggressive",
    "Aggressive: Base +5pt + NBD +5pt + Fin +5pt",
    [0.4, 0.35, 0.3, 0.25],
    [0.65, 0.55, 0.45, 0.4],
    [0.35, 0.3, 0.25],
    "Lifts base, finishing, AND rush each by +5pt across all bands. The most aggressive scenario — consistently the only revenue-positive option in pilot data. SOP's pick when finance explicitly wants revenue-positive."
  ),
  makeScenario(
    "G_Conservative",
    "Conservative: only lift 25k+ base to 25%",
    [0.35, 0.3, 0.25, 0.25],
    [0.6, 0.5, 0.4, 0.35],
    [0.3, 0.25, 0.2],
    "Identical gradients to B (only the 25k+ base band lifts to 25%). Listed separately to surface it as the most conservative single-band deviation from A."
  ),
  makeScenario(
    "H_BaseGrad_5pt",
    "Base 5pt steeper gradient (40/35/30/20)",
    [0.4, 0.35, 0.3, 0.2],
    [0.6, 0.5, 0.4, 0.35],
    [0.3, 0.25, 0.2],
    "Steeper base gradient — the first 3 bands lift +5pt but the 25k+ band stays at 20%. Protects high-volume buyers while extracting more from small/mid orders. Use when high-volume competitiveness matters more than gross revenue."
  ),
] as const;

export const SCENARIO_BY_ID: Record<string, ScenarioDef> = Object.fromEntries(
  SCENARIOS.map((s) => [s.id, s])
);

export const CUSTOM_SCENARIO_ID = "X_Custom";

/** Build a user-defined scenario from raw gradient fractions (e.g. 0.35 = 35%).
 * Used by the "Custom scenario" checkbox in the UI — the resulting ScenarioDef
 * flows through the same compute / report / download pipeline as A–H. */
export function makeCustomScenario(
  base: Gradient["base"],
  fin: Gradient["fin"],
  nbd: Gradient["nbd"]
): ScenarioDef {
  return makeScenario(
    CUSTOM_SCENARIO_ID,
    `Custom (${fmt4(base)} base, ${fmt4(fin)} fin, ${fmt3(nbd)} NBD)`,
    base,
    fin,
    nbd,
    "User-defined custom scenario entered via the Custom scenario checkbox. Base, finishing, and turnaround (NBD) markups are exactly the percentages you typed — the report, 1-pager, and 3-month tuning all use these values when this scenario is selected."
  );
}

/** Resolve a scenario id against the built-in table, falling back to a
 * caller-supplied custom scenario when the id matches it. */
export function resolveScenario(
  id: string,
  custom?: ScenarioDef | null
): ScenarioDef | undefined {
  if (custom && custom.id === id) return custom;
  return SCENARIO_BY_ID[id];
}

export interface PriceComputeInput {
  qty: number;
  baseCost: number;
  finCost: number;
  isRush: boolean;
  currentSalePrice: number;
}

export interface PriceComputeResult {
  basePrice: number;
  finPrice: number;
  subtotal: number;
  uncappedPrice: number;
  finalPrice: number;
  capped: boolean;
}

export function computeNewPrice(
  inp: PriceComputeInput,
  grad: Gradient,
  applyCap = true
): PriceComputeResult {
  const band = bandOf(inp.qty);
  const basePrice = inp.baseCost * (1 + grad.base[band]);
  const finPrice = inp.finCost * (1 + grad.fin[band]);
  let subtotal = basePrice + finPrice;
  if (inp.isRush) {
    const nb = nbdBandOf(inp.qty);
    subtotal = subtotal * (1 + grad.nbd[nb]);
  }
  const uncapped = subtotal;
  // Cap rule: only applied when applyCap is true AND a positive list price
  // is available. When disabled, the new price floats free of the catalog
  // list and may exceed it (delta can go positive).
  const finalPrice =
    applyCap && inp.currentSalePrice > 0 && uncapped > inp.currentSalePrice
      ? inp.currentSalePrice
      : uncapped;
  return {
    basePrice,
    finPrice,
    subtotal,
    uncappedPrice: uncapped,
    finalPrice,
    capped: finalPrice < uncapped,
  };
}
