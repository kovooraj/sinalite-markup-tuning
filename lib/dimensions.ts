/**
 * Generic dimensional-field system. Different products carry different
 * "variant" columns — postcards have Coating/Bundling/Scoring, booklets have
 * Cover/Binding/Pages, business cards have Sides/Finishing, etc. Rather than
 * hard-code one product's schema, every dimensional field is optional and
 * discovered at parse time.
 *
 * Required fields (always assumed present in any product):
 *   - qty            (number)
 *   - stock          (string)
 *   - size           (string)
 *   - turnaround     (Standard | Rush (NBD))
 *   - sale price     (CAD number)
 *   - PE3 cost       (number)
 *
 * Everything below is a recognized optional dimension. If a column matching
 * one of these aliases appears in the input, that dimension becomes part of
 * the lookup key for that file.
 */

export interface DimensionDef {
  /** Canonical field name used internally and in lookup keys. */
  key: string;
  /** Aliases — case-insensitive header strings that map to this field. */
  aliases: string[];
}

export const DIMENSIONS: readonly DimensionDef[] = [
  { key: "coating", aliases: ["Coating", "PE3 Coating"] },
  { key: "bundling", aliases: ["Bundling", "PE3 Bundling", "Packaging", "Packaging Option"] },
  { key: "scoring", aliases: ["Scoring", "PE3 Scoring"] },
  { key: "cover", aliases: ["Cover", "PE3 Cover"] },
  { key: "binding", aliases: ["Binding", "PE3 Binding"] },
  { key: "pages", aliases: ["Pages", "PE3 Pages", "Page Count"] },
  { key: "sides", aliases: ["Sides", "PE3 Sides"] },
  { key: "finishing", aliases: ["Finishing", "PE3 Finishing"] },
  { key: "lamination", aliases: ["Lamination", "PE3 Lamination"] },
  { key: "foil", aliases: ["Foil", "PE3 Foil"] },
  { key: "embossing", aliases: ["Embossing", "PE3 Embossing"] },
  { key: "diecutting", aliases: ["Die Cutting", "Diecutting", "PE3 Die Cutting"] },
  { key: "corner", aliases: ["Corner", "Rounded Corner", "PE3 Corner"] },
  { key: "perforation", aliases: ["Perforation", "PE3 Perforation"] },
  { key: "drilling", aliases: ["Drilling", "PE3 Drilling"] },
];

/** Build a Map from header-name (lowercased) → canonical dimension key. */
export function buildHeaderToDimensionMap(): Map<string, string> {
  const m = new Map<string, string>();
  for (const d of DIMENSIONS) {
    for (const a of d.aliases) m.set(a.toLowerCase(), d.key);
  }
  return m;
}

export type RowDimensions = Record<string, string>;

/** Build a lookup key for a row. The order of dimensional fields included
 * is fixed by sorting the field keys alphabetically, so the same key is
 * produced regardless of header order. */
export function buildKey(
  base: { stock: string; size: string; qty: number; turnaround: string },
  dims: RowDimensions,
  includedDimFields: string[]
): string {
  const sorted = [...includedDimFields].sort();
  const parts: string[] = [
    base.stock.toLowerCase().trim(),
    base.size.toLowerCase().trim(),
    String(base.qty),
    base.turnaround,
  ];
  for (const f of sorted) {
    const v = dims[f] ?? "";
    parts.push(v.toLowerCase().trim());
  }
  return parts.join("|");
}

/** Detect which canonical dimensions exist in the given header row. Returns
 * a Record<dimensionKey, columnIndex>. */
export function detectDimensions(
  header: (string | number | null)[]
): Record<string, number> {
  const hToD = buildHeaderToDimensionMap();
  const found: Record<string, number> = {};
  for (let i = 0; i < header.length; i++) {
    const cell = header[i];
    if (typeof cell !== "string") continue;
    const norm = cell.trim().toLowerCase();
    const dim = hToD.get(norm);
    if (dim && !(dim in found)) found[dim] = i;
  }
  return found;
}

/** Normalize a raw dimension value: trim, fall back to empty string. */
export function normalizeDimValue(raw: unknown): string {
  if (raw === null || raw === undefined) return "";
  const s = String(raw).trim();
  if (!s) return "";
  // Treat sentinel "None" as empty so absence and "None" collapse.
  if (s.toLowerCase() === "none") return "";
  return s;
}
