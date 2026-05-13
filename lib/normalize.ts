export function normalizeTurnaround(raw: string | null | undefined): "Standard" | "Rush (NBD)" {
  if (!raw) return "Standard";
  const s = String(raw).toLowerCase().trim();
  if (s.includes("next business day") || s.includes("nbd") || s.includes("rush")) {
    return "Rush (NBD)";
  }
  return "Standard";
}

export function normalizeBundling(raw: string | null | undefined): string {
  if (!raw) return "No bundling - FREE";
  const s = String(raw).trim();
  if (!s) return "No bundling - FREE";
  return s;
}

export function normalizeScoring(raw: string | null | undefined): string {
  if (!raw) return "None";
  const s = String(raw).trim();
  if (!s || s.toLowerCase() === "none") return "None";
  return s;
}

export function normalizeStock(raw: string | null | undefined): string {
  if (!raw) return "";
  return String(raw).trim();
}

export function normalizeSize(raw: string | null | undefined): string {
  if (!raw) return "";
  return String(raw).trim();
}

export function normalizeCoating(raw: string | null | undefined): string {
  if (!raw) return "";
  return String(raw).trim();
}

export function lookupKey(parts: {
  stock: string;
  coating?: string;
  size: string;
  qty: number;
  turnaround: "Standard" | "Rush (NBD)";
  bundling: string;
  scoring: string;
}): string {
  return [
    normalizeStock(parts.stock),
    normalizeCoating(parts.coating ?? ""),
    normalizeSize(parts.size),
    parts.qty,
    parts.turnaround,
    normalizeBundling(parts.bundling),
    normalizeScoring(parts.scoring),
  ].join("|");
}
