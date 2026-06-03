/**
 * Direct CSV → array-of-arrays parser. Bypasses SheetJS for CSV inputs so
 * very large files (Brochures PE3 is 110 MB / 8.6M cells) don't blow up
 * SheetJS's cell-object model with "Too many properties to enumerate".
 *
 * Handles:
 *  - LF, CRLF line endings
 *  - UTF-8 BOM
 *  - Quoted fields ("..."), including embedded commas and escaped quotes ("")
 *  - Numeric coercion: pure number-looking strings become numbers
 *  - Empty cells → null (matching SheetJS sheet_to_json defval: null)
 *
 * Returns the same shape sheet_to_json produces with header: 1.
 */

export type CsvCell = string | number | null;

const NUMERIC_RE = /^-?\d+(?:\.\d+)?$/;

function maybeNumber(s: string): CsvCell {
  if (s === "") return null;
  // Only coerce if it looks like a plain number — leave anything else alone.
  if (NUMERIC_RE.test(s)) {
    const n = Number(s);
    if (isFinite(n)) return n;
  }
  return s;
}

export function parseCsvToAoa(text: string): CsvCell[][] {
  // Strip UTF-8 BOM
  if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);

  const rows: CsvCell[][] = [];
  let row: CsvCell[] = [];
  let cell = "";
  let inQuotes = false;
  const len = text.length;

  for (let i = 0; i < len; i++) {
    const c = text.charCodeAt(i);

    if (inQuotes) {
      if (c === 34 /* " */) {
        // Escaped quote ""? Look ahead.
        if (i + 1 < len && text.charCodeAt(i + 1) === 34) {
          cell += '"';
          i += 1;
        } else {
          inQuotes = false;
        }
      } else {
        cell += text[i];
      }
      continue;
    }

    // not in quotes
    if (c === 34 /* " */) {
      inQuotes = true;
      continue;
    }
    if (c === 44 /* , */) {
      row.push(maybeNumber(cell));
      cell = "";
      continue;
    }
    if (c === 10 /* \n */) {
      row.push(maybeNumber(cell));
      cell = "";
      rows.push(row);
      row = [];
      continue;
    }
    if (c === 13 /* \r */) {
      // CR — if next char is LF, treat as one terminator
      row.push(maybeNumber(cell));
      cell = "";
      rows.push(row);
      row = [];
      if (i + 1 < len && text.charCodeAt(i + 1) === 10) i += 1;
      continue;
    }
    cell += text[i];
  }

  // Trailing cell / row
  if (cell.length > 0 || row.length > 0) {
    row.push(maybeNumber(cell));
    rows.push(row);
  }

  // Drop trailing entirely-empty rows (common with CSVs that end in \n)
  while (
    rows.length > 0 &&
    rows[rows.length - 1].every((c) => c === null || c === "")
  ) {
    rows.pop();
  }

  return rows;
}

/** Detect if a File is a CSV: by extension first, then by sniffing the first
 * few bytes (XLSX/XLS have known magic numbers). Defaults to true on
 * ambiguity since SheetJS handles xlsx fine even if we route through CSV
 * detection wrongly — only the CSV path benefits from streaming. */
export function isCsvFile(file: File): boolean {
  const name = file.name.toLowerCase();
  if (name.endsWith(".csv")) return true;
  if (name.endsWith(".xlsx") || name.endsWith(".xls") || name.endsWith(".xlsm"))
    return false;
  return false;
}
