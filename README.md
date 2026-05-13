# SinaLite Markup Tuning

A browser-only Next.js tool that turns two spreadsheets into two locked-pricing reports.

**Inputs:**
1. Price-engine xlsx — full SKU catalog with PE3 cost decomposition (e.g. `Postcards 14PT + Matte Finish price data.xlsx`)
2. 3-month order replay xlsx — must contain a `Per-SKU Detail` sheet (output of `scripts/replay_orders.py` in the pricing model SOP)

**Outputs:**
1. `<Product>_Markup_Reference_OnePager.docx` — the 1-pager with the three locked markup tables, formulas, catalog impact stats, and the top-20 loss-leader verdict table
2. `<Product>_Markup_Tuning_Scenarios.xlsx` — the 8-scenario comparison sheet (A_Current_Locked … H_BaseGrad_5pt) with 3-mo Δ, %Δ, annualized, customer impact distribution, and Δ by qty band

Everything runs in the browser via [SheetJS](https://github.com/SheetJS/sheetjs), [ExcelJS](https://github.com/exceljs/exceljs), and [docx](https://docx.js.org/). No upload, no server, no backend.

## Methodology

This tool encodes the **Base + Add-On pricing model** validated on Postcards 14PT + AQ in May 2026:

- **Base markup gradient:** 35% / 30% / 25% / 20% (by qty band: 100–1k, 1k–5k, 5k–25k, 25k–100k)
- **Finishing markup gradient:** 60% / 50% / 40% / 35%
- **NBD rush markup:** 30% / 25% / 20% (3 bands: 100–5k, 5k–25k, 25k–100k)
- **Cap rule (non-negotiable):** `final price = MIN(model price, current published price)` — no customer ever sees a list-price increase

See the canonical SOP in the `sinalite-pricing-model` skill for the full 9-phase methodology.

## Local dev

```bash
npm install
npm run dev
# open http://localhost:3000
```

Type-check + production build:

```bash
npx tsc --noEmit
npm run build
```

## Smoke test against the reference files

```bash
npx tsx scripts/smoke.ts \
  "C:\path\to\Postcards 14PT + Matte Finish price data.xlsx" \
  "C:\path\to\Postcards_14PT_AQ_3Month_Order_Replay.xlsx"
```

Outputs land in `./out/`. The script prints per-scenario deltas, the self-check result (Scenario A computed-vs-replay), and the loss-leader verdict. For a clean diff against the reference, the two input files must be for the **same coating** (e.g. AQ + AQ).

## Deploy to Vercel

1. Push to GitHub.
2. Import the repo at <https://vercel.com/new>.
3. Framework preset: Next.js (auto-detected).
4. No env vars required. No build customization.
5. Deploy — first deploy completes in ~1 minute.

Subsequent pushes to `main` redeploy automatically.

## Architecture

```
app/page.tsx                 ← single page UI
components/
  FileDrop.tsx               ← drag-and-drop zone (react-dropzone)
  MetaForm.tsx               ← locked date / owner / archive path / target band
  ResultsPanel.tsx           ← scenario table + recommendation + download buttons
lib/
  qtyBands.ts                ← bandOf(qty) / nbdBandOf(qty)
  normalize.ts               ← canonicalize stock/bundling/scoring/turnaround strings
  parsePriceEngine.ts        ← xlsx → SKU records w/ denylist-based base + fin cost decomposition
  parseOrderReplay.ts        ← xlsx → per-SKU order rows
  markupEngine.ts            ← 8 scenarios + computeNewPrice(row, gradient)
  computeScenarios.ts        ← aggregate Δ$, % Δ, distribution, qty bands, by-turnaround
  selfCheck.ts               ← cross-validate Scenario A vs replay's pre-computed columns
  catalogImpact.ts           ← OnePager Section 1/2/3 narrative stats
  lossLeaders.ts             ← top-20 SKU verdict
  recommend.ts               ← pick the winning scenario given a target revenue impact band
  buildScenariosXlsx.ts      ← ExcelJS writer matching reference layout
  buildOnePagerDocx.ts       ← docx writer matching reference layout
scripts/smoke.ts             ← end-to-end Node smoke test
```

## Cost decomposition

The price engine xlsx flattens the cost breakdown across ~80 columns as repeating `name | <machine> | totalCost | <value>` triples. `parsePriceEngine.ts` walks each row and bins every machine's `totalCost` into one of two buckets:

- **Finishing marginal cost:** machines matching `Rosback`, `Longford`, `Shrink Wrap`, `Bundler`, `Score`
- **Base cost:** everything else (Prepress, HP Indigo / Heidelberg / any press, Polar / any cutter, Packing, BDOSetup)

The denylist approach keeps the parser product-agnostic — same code handles Postcards (HP Indigo), Business Cards, or anything else.

## Self-check

The order replay xlsx already contains `New Price/Order` and `New Revenue $` columns computed at the locked gradient (Scenario A). On every run we re-compute Scenario A from scratch and diff row-by-row. If aggregate drift > $50 or max-row diff > $1, the UI surfaces a warning suggesting a key normalization issue (e.g. uploading the Matte price engine against an AQ replay).

## Out of scope (v1)

These deliverables from the skill SOP are not produced by this tool — they're follow-ups that can become additional download buttons:

- `00_<Product>_Pricing_OnePager_Dev.docx` (dev implementation spec)
- `00_CS_Enablement_Pricing_Change_FAQ.docx`
- `<Product>_SKU_Revenue_Split.xlsx`, `<Product>_NBD_Rush_Impact.xlsx`, `<Product>_PE3_Cost_Grids.xlsx`

## License

Internal Sinalite tooling.
