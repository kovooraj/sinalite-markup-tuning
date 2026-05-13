// End-to-end smoke test using the same lib modules the browser will use.
// Run via: npx tsx scripts/smoke.ts [pricePath] [orderPath]

import { readFile, mkdir, writeFile } from "node:fs/promises";
import { parsePriceEngine } from "../lib/parsePriceEngine";
import { parseOrderReplay } from "../lib/parseOrderReplay";
import { computeAllScenarios } from "../lib/computeScenarios";
import { runSelfCheck } from "../lib/selfCheck";
import {
  computeBaseCatalogImpact,
  computeFinishingCatalogImpact,
  computeNbdLift,
} from "../lib/catalogImpact";
import { computeLossLeaders } from "../lib/lossLeaders";
import { buildScenariosXlsx } from "../lib/buildScenariosXlsx";
import { buildOnePagerDocx } from "../lib/buildOnePagerDocx";

const PRICE_FILE =
  process.argv[2] ??
  "C:\\Users\\kovoo\\OneDrive\\Desktop\\Postcards 14PT + Matte Finish price data.xlsx";
const ORDER_FILE =
  process.argv[3] ??
  "C:\\Users\\kovoo\\OneDrive\\Desktop\\Postcards_14PT_AQ_3Month_Order_Replay.xlsx";

// Mock File with the minimal shape our parsers use
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
const priceBuf = await readFile(PRICE_FILE);
const orderBuf = await readFile(ORDER_FILE);
const priceFile = new FakeFile(priceBuf, "price.xlsx") as unknown as File;
const orderFile = new FakeFile(orderBuf, "order.xlsx") as unknown as File;

console.log(`Reading price engine: ${PRICE_FILE}`);
const pe = await parsePriceEngine(priceFile);
console.log(`  product: ${pe.productName}`);
console.log(`  rows: ${pe.rows.length}, coatings: ${pe.coatings.join(", ") || "—"}`);
if (pe.warnings.length) for (const w of pe.warnings) console.log(`  WARN: ${w}`);

console.log(`Reading order replay: ${ORDER_FILE}`);
const order = await parseOrderReplay(orderFile);
console.log(
  `  rows: ${order.rows.length}, orders: ${order.totals.orders}, paid: $${order.totals.actualPaid.toFixed(2)}`
);
console.log(
  `  unmatched in sheet: ${order.unmatched.count} orders / $${order.unmatched.revenue.toFixed(2)}`
);

const scenarios = computeAllScenarios(order, pe);
console.log(
  `\nScenario lookup: matched ${scenarios.matchedOrderRows}, unmatched ${scenarios.unmatchedOrderRows}`
);
if (scenarios.unmatchedSamples.length) {
  console.log(`  unmatched samples (first 5):`);
  scenarios.unmatchedSamples.slice(0, 5).forEach((s) =>
    console.log(`    - ${s.slice(0, 100)}`)
  );
}

console.log("\nScenarios (3-mo Δ | % Δ | annualized):");
for (const s of scenarios.scenarios) {
  console.log(
    `  ${s.id.padEnd(20)} ${("$" + s.deltaUsd.toFixed(2)).padStart(12)}  ${(
      s.pctDelta * 100
    )
      .toFixed(2)
      .padStart(7)}%   $${s.annualizedUsd.toFixed(0)}`
  );
}

const sc = runSelfCheck(order, pe);
console.log(
  `\nSelf-check: ok=${sc.ok}, rowsCompared=${sc.rowsCompared}, missed=${sc.rowsMissed}, maxRowDiff=${sc.maxRowDiff.toFixed(4)}, aggDiff=${sc.aggregateDiff.toFixed(2)}`
);
console.log(
  `  scenarioA computed=${sc.scenarioAComputedDelta.toFixed(2)}, replay=${sc.scenarioAReplayDelta.toFixed(2)}`
);

const baseImpact = computeBaseCatalogImpact(pe);
console.log(
  `\nBase catalog impact: $${baseImpact.impactUsd.toFixed(0)} (${(
    baseImpact.pct * 100
  ).toFixed(1)}%) on ${baseImpact.rowsScanned} base rows`
);
console.log(`  concentration: ${baseImpact.concentrationLabel || "—"}`);
console.log(
  `  band deltas: [${baseImpact.deltaByBand.map((v) => v.toFixed(0)).join(", ")}]`
);

const finImpact = computeFinishingCatalogImpact(pe);
console.log(
  `Finishing impact: $${finImpact.impactUsd.toFixed(0)} (${(
    finImpact.pct * 100
  ).toFixed(1)}%); ${finImpact.cellsUplifted} up / ${finImpact.cellsReduced} down on ${finImpact.rowsScanned} variant rows`
);

const a = scenarios.scenarios.find((s) => s.id === "A_Current_Locked")!;
const nbd = computeNbdLift(a);
console.log(
  `NBD lift: annualized $${nbd.annualizedUsd.toFixed(0)} (low ${nbd.lowUsd.toFixed(
    0
  )} / high ${nbd.highUsd.toFixed(0)}); 3-mo orders=${nbd.orders3mo}, base rev=$${nbd.baseRevenue3mo.toFixed(0)}`
);

const ll = computeLossLeaders(order, pe);
console.log(`\nLoss leaders (top ${ll.rows.length}):`);
for (const r of ll.rows) {
  console.log(
    `  ${r.sizeQty.padEnd(20)}  base=${r.baseMarginUsd.toFixed(0).padStart(8)}  addon=${r.addonMarginUsd
      .toFixed(0)
      .padStart(8)}  net=${r.netUsd.toFixed(0).padStart(8)}  ${r.verdict.padEnd(15)}  ${r.action}`
  );
}
console.log(
  `  TOTAL  base=${ll.totals.baseMarginUsd.toFixed(0)}  addon=${ll.totals.addonMarginUsd.toFixed(0)}  net=${ll.totals.netUsd.toFixed(0)}`
);
if (ll.callout) {
  console.log(
    `  callout: ${ll.callout.sizeFamily} at qtys ${ll.callout.qtys.join(",")} totaling $${ll.callout.totalLoss.toFixed(0)}`
  );
} else {
  console.log("  callout: —");
}

await mkdir("out", { recursive: true });

const scXlsx = await buildScenariosXlsx({
  productName: pe.productName,
  productSlug: pe.productSlug,
  results: scenarios.scenarios,
});
await writeFile(
  `out/${pe.productSlug}_Markup_Tuning_Scenarios.xlsx`,
  Buffer.from(await scXlsx.arrayBuffer())
);

const docxBlob = await buildOnePagerDocx({
  productName: pe.productName,
  productSlug: pe.productSlug,
  lockedDate: new Date().toISOString().slice(0, 10),
  owner: "Mike",
  archivePath: "\\Costing\\Postcards_Pricing\\",
  baseImpact,
  finImpact,
  nbdLift: nbd,
  lossLeaders: ll,
});
await writeFile(
  `out/${pe.productSlug}_Markup_Reference_OnePager.docx`,
  Buffer.from(await docxBlob.arrayBuffer())
);

console.log(`\nWrote out/${pe.productSlug}_Markup_Tuning_Scenarios.xlsx`);
console.log(`Wrote out/${pe.productSlug}_Markup_Reference_OnePager.docx`);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
