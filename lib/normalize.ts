export function normalizeTurnaround(
  raw: string | null | undefined
): "Standard" | "Rush (NBD)" {
  if (!raw) return "Standard";
  const s = String(raw).toLowerCase().trim();
  if (s.includes("next business day") || s.includes("nbd") || s.includes("rush")) {
    return "Rush (NBD)";
  }
  return "Standard";
}

export function normalizeStock(raw: string | null | undefined): string {
  if (!raw) return "";
  return String(raw).trim();
}

export function normalizeSize(raw: string | null | undefined): string {
  if (!raw) return "";
  return String(raw).trim();
}
