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
  nbd: Gradient["nbd"]
): ScenarioDef {
  return {
    id,
    label,
    baseFormatted: fmt4(base),
    finFormatted: fmt4(fin),
    nbdFormatted: fmt3(nbd),
    grad: { base, fin, nbd },
  };
}

export const SCENARIOS: readonly ScenarioDef[] = [
  makeScenario(
    "A_Current_Locked",
    "Current locked (35/30/25/20 base, 60/50/40/35 fin, 30/25/20 NBD)",
    [0.35, 0.3, 0.25, 0.2],
    [0.6, 0.5, 0.4, 0.35],
    [0.3, 0.25, 0.2]
  ),
  makeScenario(
    "B_Lift_25kBase",
    "Lift 25k+ base to 25% (35/30/25/25)",
    [0.35, 0.3, 0.25, 0.25],
    [0.6, 0.5, 0.4, 0.35],
    [0.3, 0.25, 0.2]
  ),
  makeScenario(
    "C_Base_Plus5",
    "Base +5pt across board (40/35/30/25)",
    [0.4, 0.35, 0.3, 0.25],
    [0.6, 0.5, 0.4, 0.35],
    [0.3, 0.25, 0.2]
  ),
  makeScenario(
    "D_NBD_Plus5",
    "NBD +5pt across board (35/30/25)",
    [0.35, 0.3, 0.25, 0.2],
    [0.6, 0.5, 0.4, 0.35],
    [0.35, 0.3, 0.25]
  ),
  makeScenario(
    "E_Combined_Mod",
    "Combined: Lift 25k+ base + NBD +5pt",
    [0.35, 0.3, 0.25, 0.25],
    [0.6, 0.5, 0.4, 0.35],
    [0.35, 0.3, 0.25]
  ),
  makeScenario(
    "F_Aggressive",
    "Aggressive: Base +5pt + NBD +5pt + Fin +5pt",
    [0.4, 0.35, 0.3, 0.25],
    [0.65, 0.55, 0.45, 0.4],
    [0.35, 0.3, 0.25]
  ),
  makeScenario(
    "G_Conservative",
    "Conservative: only lift 25k+ base to 25%",
    [0.35, 0.3, 0.25, 0.25],
    [0.6, 0.5, 0.4, 0.35],
    [0.3, 0.25, 0.2]
  ),
  makeScenario(
    "H_BaseGrad_5pt",
    "Base 5pt steeper gradient (40/35/30/20)",
    [0.4, 0.35, 0.3, 0.2],
    [0.6, 0.5, 0.4, 0.35],
    [0.3, 0.25, 0.2]
  ),
] as const;

export const SCENARIO_BY_ID: Record<string, ScenarioDef> = Object.fromEntries(
  SCENARIOS.map((s) => [s.id, s])
);

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
  grad: Gradient
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
  const finalPrice =
    inp.currentSalePrice > 0 && uncapped > inp.currentSalePrice
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
