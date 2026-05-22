// Smoke test against the per-order replay format the user uploaded.
import { readFile } from "node:fs/promises";
import { parsePriceEngine } from "../lib/parsePriceEngine";
import { parseOrderReplay } from "../lib/parseOrderReplay";
import { computeAllScenarios } from "../lib/computeScenarios";

const PRICE =
  process.argv[2] ??
  "C:\\Users\\kovoo\\OneDrive\\Desktop\\Postcards 14PT + Matte Finish Price .xlsx.csv";
const ORDER =
  process.argv[3] ??
  "C:\\Users\\kovoo\\OneDrive\\Desktop\\Postcards_14PT_Matte_3Month_Orders.xlsx";

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
  const pricBuf = await readFile(PRICE);
  const ordBuf = await readFile(ORDER);
  const pe = await parsePriceEngine(new FakeFile(pricBuf, "p.csv") as unknown as File);
  console.log(`Price engine: ${pe.productName} · ${pe.rows.length} rows · dims ${pe.dimensions.join(",")}`);

  const order = await parseOrderReplay(
    new FakeFile(ordBuf, "o.xlsx") as unknown as File,
    { usdRate: 0.7 }
  );
  console.log(`Order replay format: ${order.format}`);
  console.log(`  rows: ${order.rows.length}, orders: ${order.totals.orders}, paid CAD: $${order.totals.actualPaid.toFixed(2)}`);
  console.log(`  unmatched: ${order.unmatched.count} orders / $${order.unmatched.revenue.toFixed(2)}`);
  console.log(`  warnings: ${order.warnings.length}`);
  for (const w of order.warnings) console.log(`   - ${w}`);
  console.log(`  hasPrecomputedScenarioA: ${order.hasPrecomputedScenarioA}`);

  const sc = computeAllScenarios(order, pe);
  console.log(`\nScenario matches: ${sc.matchedOrderRows} / unmatched: ${sc.unmatchedOrderRows}`);
  console.log(`\nScenarios:`);
  for (const s of sc.scenarios) {
    console.log(`  ${s.id.padEnd(20)} Δ$${s.deltaUsd.toFixed(0).padStart(10)} (${(s.pctDelta * 100).toFixed(2)}%)`);
  }
}

main().catch((e) => { console.error(e); process.exit(1); });
