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
  let s = String(raw).trim();
  // Strip trailing ".0" from numeric components so "8.5 x 11.0" matches
  // "8.5 x 11". Handles both " x " and "x" separators with any whitespace.
  s = s.replace(/(\d+)\.0+\b/g, "$1");
  return s;
}
