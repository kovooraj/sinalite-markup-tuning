// Quick CSV smoke test — verifies parsePriceEngine handles a CSV input.
// Requires a CSV at the path passed (or the default).

import { readFile } from "node:fs/promises";
import { parsePriceEngine } from "../lib/parsePriceEngine";

const PRICE_CSV =
  process.argv[2] ?? "C:\\Users\\kovoo\\OneDrive\\Desktop\\Postcards 14PT + Matte Finish Price .xlsx.csv";

class FakeFile {
  _buf: Buffer;
  name: string;
  size: number;
  constructor(buf: Buffer, name: string) {
    this._buf = buf;
    this.name = name;
    this.size = buf.length;
  }
  async arrayBuffer(): Promise<ArrayBuffer> {
    return this._buf.buffer.slice(
      this._buf.byteOffset,
      this._buf.byteOffset + this._buf.byteLength
    ) as ArrayBuffer;
  }
}

async function main() {
  const buf = await readFile(PRICE_CSV);
  const f = new FakeFile(buf, "price.csv") as unknown as File;
  console.log(`Reading CSV: ${PRICE_CSV} (${(buf.length / 1024).toFixed(1)} KB)`);
  const pe = await parsePriceEngine(f);
  console.log(`Product: ${pe.productName}`);
  console.log(`Rows parsed: ${pe.rows.length}`);
  console.log(`Coatings: ${pe.coatings.join(", ") || "—"}`);
  console.log(`Warnings: ${pe.warnings.length}`);
  for (const w of pe.warnings) console.log(`  - ${w}`);

  // Spot-check first 3 rows
  console.log(`\nFirst 3 rows:`);
  for (const r of pe.rows.slice(0, 3)) {
    console.log(
      `  qty=${r.qty} stock="${r.stock.slice(0, 40)}" size=${r.size} ` +
        `bundling="${r.bundling.slice(0, 30)}" base=$${r.baseCost.toFixed(2)} fin=$${r.finCost.toFixed(2)} sale=$${r.currentSalePrice.toFixed(2)}`
    );
  }
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
