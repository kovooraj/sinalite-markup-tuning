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
  // "8.5 x 11". Handles any whitespace around the separator.
  s = s.replace(/(\d+)\.0+\b/g, "$1");
  // Canonical dimension form: "<smaller> x <larger>". This collapses
  // "5.5 x 8.5" and "8.5 x 5.5" to the same key so the order replay (which
  // may list size as Height × Width) and the price engine (Width × Height)
  // still match. Only applies when the value is exactly two numeric
  // components separated by "x" — leaves anything else (e.g. "8pg",
  // "Custom") untouched.
  const m = s.match(/^(\d+(?:\.\d+)?)\s*[xX]\s*(\d+(?:\.\d+)?)$/);
  if (m) {
    const a = parseFloat(m[1]);
    const b = parseFloat(m[2]);
    if (isFinite(a) && isFinite(b)) {
      const [lo, hi] = a <= b ? [m[1], m[2]] : [m[2], m[1]];
      s = `${lo} x ${hi}`;
    }
  }
  return s;
}
