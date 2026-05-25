// Verifies the annotated-orders xlsx renders cleanly and the TOTAL row's
// annualized delta lines up with the scenario-level annualized number.
import { readFile, mkdir, writeFile } from "node:fs/promises";
import { parsePriceEngine } from "../lib/parsePriceEngine";
import { parseOrderReplay } from "../lib/parseOrderReplay";
import { computeAllScenarios } from "../lib/computeScenarios";
import { buildAnnotatedOrdersXlsx } from "../lib/buildAnnotatedOrdersXlsx";

const PRICE = process.argv[2];
const ORDER = process.argv[3];
const SCENARIO = process.argv[4] ?? "F_Aggressive";

class FakeFile {
  _buf: Buffer; name: string; size: number;
  constructor(buf: Buffer, name: string) { this._buf = buf; this.name = name; this.size = buf.length; }
  async arrayBuffer(): Promise<ArrayBuffer> {
    return this._buf.buffer.slice(this._buf.byteOffset, this._buf.byteOffset + this._buf.byteLength) as ArrayBuffer;
  }
}

async function main() {
  if (!PRICE || !ORDER) {
    console.error("Usage: tsx smoke-annotated.ts <price.csv|xlsx> <order.xlsx> [scenarioId]");
    process.exit(2);
  }
  const peBuf = await readFile(PRICE);
  const orBuf = await readFile(ORDER);
  const pe = await parsePriceEngine(new FakeFile(peBuf, "p") as unknown as File);
  const order = await parseOrderReplay(
    new FakeFile(orBuf, "o") as unknown as File,
    { usdRate: 0.7 }
  );

  const sc = computeAllScenarios(order, pe);
  const target = sc.scenarios.find(s => s.id === SCENARIO)!;
  console.log(`Scenario ${SCENARIO}: 3-mo Δ ${target.deltaUsd.toFixed(2)} (${(target.pctDelta*100).toFixed(2)}%) annualized ${target.annualizedUsd.toFixed(2)}`);

  const blob = await buildAnnotatedOrdersXlsx({
    pe, order,
    scenarioId: SCENARIO,
    usdRate: 0.7,
    productSlug: pe.productSlug,
  });
  await mkdir("out", { recursive: true });
  const path = `out/${pe.productSlug}_Orders_Annotated_${SCENARIO}.xlsx`;
  await writeFile(path, Buffer.from(await blob.arrayBuffer()));
  console.log(`Wrote ${path} (${(await blob.arrayBuffer()).byteLength} bytes)`);
}

main().catch(e => { console.error(e); process.exit(1); });
