export type BaseBandIndex = 0 | 1 | 2 | 3;
export type NbdBandIndex = 0 | 1 | 2;

export const BASE_BAND_LABELS = [
  "100 – 1,000",
  "1,001 – 5,000",
  "5,001 – 25,000",
  "25,001 – 100,000",
] as const;

export const NBD_BAND_LABELS = [
  "100 – 5,000",
  "5,001 – 25,000",
  "25,001 – 100,000",
] as const;

export function bandOf(qty: number): BaseBandIndex {
  if (qty <= 1000) return 0;
  if (qty <= 5000) return 1;
  if (qty <= 25000) return 2;
  return 3;
}

export function nbdBandOf(qty: number): NbdBandIndex {
  if (qty <= 5000) return 0;
  if (qty <= 25000) return 1;
  return 2;
}

export function bandLabelOf(qty: number): string {
  return BASE_BAND_LABELS[bandOf(qty)];
}
