// Quick check: Stahlfold B30 #1 buckets as BASE for Brochure products and
// stays FINISHING for everything else. Run via: npx tsx scripts/check-brochure-override.ts
import { parsePriceEngine } from "../lib/parsePriceEngine";

class FakeFile {
  _buf: Buffer;
  name: string;
  constructor(content: string, name: string) {
    this._buf = Buffer.from(content, "utf-8");
    this.name = name;
  }
  async arrayBuffer(): Promise<ArrayBuffer> {
    return this._buf.buffer.slice(
      this._buf.byteOffset,
      this._buf.byteOffset + this._buf.byteLength
    ) as ArrayBuffer;
  }
}

function makeCsv(product: string): string {
  const header = `${product},qty,size,Turnaround,sale price,PE3 cost no markup,breakdown,,,,,,,,`;
  const row = `${product},500,8.5x11,Standard,100,40,name,Komori G40,totalCost,30,name,"Stahlfold B30 #1",totalCost,10`;
  return `${header}\n${row}\n`;
}

async function run(product: string) {
  const pe = await parsePriceEngine(
    new FakeFile(makeCsv(product), `${product}.csv`) as unknown as File
  );
  const r = pe.rows[0];
  console.log(
    `${product.padEnd(28)} base=$${r.baseCost} fin=$${r.finCost}  baseCenters=[${pe.costCenters.base}]  finCenters=[${pe.costCenters.finishing}]`
  );
  return r;
}

async function main() {
  const brochure = await run("Brochures 100lb Gloss");
  const postcard = await run("Postcards 14PT AQ");

  const ok =
    brochure.baseCost === 40 &&
    brochure.finCost === 0 &&
    postcard.baseCost === 30 &&
    postcard.finCost === 10;
  console.log(ok ? "PASS" : "FAIL");
  if (!ok) process.exit(1);
}

main();
