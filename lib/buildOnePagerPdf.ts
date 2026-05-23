import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";
import { saveAs } from "file-saver";
import { BaseCatalogImpact, FinishingCatalogImpact, NbdLift } from "./catalogImpact";
import { LossLeadersOutput } from "./lossLeaders";
import { SCENARIO_BY_ID } from "./markupEngine";
import { Recommendation } from "./recommend";

export interface OnePagerPdfOpts {
  productName: string;
  productSlug: string;
  lockedDate: string;
  owner: string;
  baseImpact: BaseCatalogImpact;
  finImpact: FinishingCatalogImpact;
  nbdLift: NbdLift;
  lossLeaders: LossLeadersOutput;
  recommendation: Recommendation;
  targetMinPct: number;
  targetMaxPct: number;
}

// --- formatters ---------------------------------------------------------
// All values rendered into the PDF must be ASCII. jsPDF's default Helvetica
// doesn't carry the Unicode minus (U+2212), em dash, multiplication, warning,
// or star glyphs, and falls back in ways that break kerning and cause line
// overflow. asciiSafe() and asciiClean() normalize everything to plain ASCII.

function asciiSafe(s: string): string {
  return s
    .replace(/[−–—]/g, "-") // minus, en dash, em dash
    .replace(/[×]/g, "x") // multiplication sign
    .replace(/[→]/g, "->") // right arrow
    .replace(/[⚠⚡]/g, "!") // warning
    .replace(/[★☆]/g, "*") // black/white star
    .replace(/[✓✔]/g, "+") // checkmarks
    .replace(/[ ]/g, " ") // nbsp
    .replace(/[‘’]/g, "'") // smart single quotes
    .replace(/[“”]/g, '"') // smart double quotes
    .replace(/[Δ∆]/g, "Delta") // greek delta + math increment
    // Catch-all: any remaining non-ASCII char becomes "?" so a stray
    // Unicode glyph never silently breaks line measurement again.
    .replace(/[^\x20-\x7E\n]/g, "?");
}

function fmtUsd(v: number, opts: { signed?: boolean } = {}): string {
  const abs = Math.abs(v);
  const sign = opts.signed ? (v >= 0 ? "+" : "-") : v < 0 ? "-" : "";
  return `${sign}$${abs.toLocaleString("en-US", {
    minimumFractionDigits: 0,
    maximumFractionDigits: 0,
  })}`;
}
function fmtUsdK(v: number): string {
  const abs = Math.abs(v);
  const sign = v < 0 ? "-" : "+";
  return `${sign}$${(abs / 1000).toFixed(0)}K`;
}
function fmtPct(v: number): string {
  const sign = v < 0 ? "-" : "+";
  return `${sign}${Math.abs(v * 100).toFixed(1)}%`;
}

// --- layout helpers -----------------------------------------------------

const PAGE_W = 612; // letter, pt
const MARGIN = 28; // ~0.39"
const CONTENT_W = PAGE_W - MARGIN * 2;

interface Cursor {
  y: number;
}

function ensureSpace(doc: jsPDF, c: Cursor, needed: number) {
  const pageH = doc.internal.pageSize.getHeight();
  if (c.y + needed > pageH - MARGIN) {
    doc.addPage();
    c.y = MARGIN;
  }
}

// jsPDF positions text by its baseline, not its top. To keep "draw at c.y"
// mean "text TOP aligns with c.y" (so it doesn't overlap content above), we
// advance the cursor by the font's ascender BEFORE calling doc.text, then
// advance again past remaining lines + a small bottom gap.
const ASCENDER_RATIO = 0.85; // helvetica ascender ≈ 85% of nominal size
const LINE_GAP_RATIO = 1.25; // ~25% line spacing
const BOTTOM_GAP = 4;

function heading(doc: jsPDF, c: Cursor, text: string) {
  const size = 11;
  const ascender = size * ASCENDER_RATIO;
  const topGap = 8;
  ensureSpace(doc, c, topGap + ascender + 4);
  c.y += topGap + ascender;
  doc.setFont("helvetica", "bold");
  doc.setFontSize(size);
  doc.setTextColor(20, 20, 20);
  doc.text(asciiSafe(text), MARGIN, c.y);
  c.y += 4;
}

function para(
  doc: jsPDF,
  c: Cursor,
  text: string,
  opts: { bold?: boolean; italic?: boolean; size?: number; color?: [number, number, number] } = {}
) {
  const size = opts.size ?? 9;
  const ascender = size * ASCENDER_RATIO;
  const lineHeight = size * LINE_GAP_RATIO;
  const style = opts.bold
    ? opts.italic
      ? "bolditalic"
      : "bold"
    : opts.italic
      ? "italic"
      : "normal";
  doc.setFont("helvetica", style);
  doc.setFontSize(size);
  doc.setTextColor(...(opts.color ?? [60, 60, 60]));
  const clean = asciiSafe(text);
  const lines = doc.splitTextToSize(clean, CONTENT_W);
  const totalHeight =
    ascender + (lines.length - 1) * lineHeight + BOTTOM_GAP;
  ensureSpace(doc, c, totalHeight);
  c.y += ascender;
  // jspdf can render the whole array in one call; it auto-steps by font size
  // but our lineHeight is a touch larger, so iterate to control spacing.
  for (let i = 0; i < lines.length; i++) {
    if (i > 0) c.y += lineHeight;
    doc.text(lines[i], MARGIN, c.y);
  }
  c.y += BOTTOM_GAP;
}

function cleanRows(rows: string[][]): string[][] {
  return rows.map((r) => r.map(asciiSafe));
}

function markupTable(
  doc: jsPDF,
  c: Cursor,
  head: string[],
  rows: string[][]
) {
  ensureSpace(doc, c, 12 + rows.length * 14);
  autoTable(doc, {
    startY: c.y,
    head: [head.map(asciiSafe)],
    body: cleanRows(rows),
    margin: { left: MARGIN, right: MARGIN },
    styles: { font: "helvetica", fontSize: 8.5, cellPadding: 3, lineColor: [180, 180, 180], lineWidth: 0.3, overflow: "linebreak" },
    headStyles: { fillColor: [31, 56, 100], textColor: [255, 255, 255], fontStyle: "bold", halign: "left" },
    tableWidth: CONTENT_W,
  });
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  c.y = (doc as any).lastAutoTable.finalY + 4;
}

// --- main builder -------------------------------------------------------

export function buildOnePagerPdf(opts: OnePagerPdfOpts): Blob {
  const doc = new jsPDF({ unit: "pt", format: "letter" });
  const c: Cursor = { y: MARGIN };
  const A = SCENARIO_BY_ID["A_Current_Locked"];

  // Title
  doc.setFont("helvetica", "bold");
  doc.setFontSize(15);
  doc.setTextColor(20, 20, 20);
  const titleAscender = 15 * ASCENDER_RATIO;
  c.y += titleAscender;
  doc.text(
    asciiSafe(`${opts.productName} - Locked Markup Reference`),
    MARGIN,
    c.y
  );
  c.y += 8;

  // Subtitle
  para(
    doc,
    c,
    `Decisions locked ${opts.lockedDate}  |  Cap rule applies to ALL: final price = MIN(model price, current published price)`,
    { italic: true, size: 8.5 }
  );

  // Section 1 — BASE
  heading(doc, c, "1. BASE Product Markup");
  para(doc, c, "Applied to PE3 base cost. Volume-discount gradient — markup decreases as qty rises.");
  para(doc, c, "Formula:  Base Price = MIN(PE3 Base Cost × (1 + Markup %), Current Sale Price)", { bold: true });
  markupTable(doc, c,
    ["Qty Band", "Markup %", "Multiplier", "Example: $30 cost ->"],
    [
      ["100 - 1,000", `${Math.round(A.grad.base[0] * 100)}%`, `${(1 + A.grad.base[0]).toFixed(2)}x`, `$${(30 * (1 + A.grad.base[0])).toFixed(2)}`],
      ["1,001 - 5,000", `${Math.round(A.grad.base[1] * 100)}%`, `${(1 + A.grad.base[1]).toFixed(2)}x`, `$${(30 * (1 + A.grad.base[1])).toFixed(2)}`],
      ["5,001 - 25,000", `${Math.round(A.grad.base[2] * 100)}%`, `${(1 + A.grad.base[2]).toFixed(2)}x`, `$${(30 * (1 + A.grad.base[2])).toFixed(2)}`],
      ["25,001 - 100,000", `${Math.round(A.grad.base[3] * 100)}%`, `${(1 + A.grad.base[3]).toFixed(2)}x`, `$${(30 * (1 + A.grad.base[3])).toFixed(2)}`],
    ]
  );
  const concLine = opts.baseImpact.concentrationLabel ? ` ${opts.baseImpact.concentrationLabel}` : "";
  para(doc, c,
    `Catalog impact: ${fmtUsd(opts.baseImpact.impactUsd, { signed: true })} (${fmtPct(opts.baseImpact.pct)}) on base SKUs.${concLine} Cap rule means no customer ever sees a price increase.`,
    { italic: true }
  );

  // Section 2 — UPCHARGES
  heading(doc, c, "2. UPCHARGES — Finishing Add-Ons");
  para(doc, c, "Applies to all 8 bundling types (Single band 100s/50s/25s, Double band 100s/50s/25s, Shrink Wrap 50s/25s) + Score in Half.");
  para(doc, c, "Formula:  Add-on Price = PE3 Marginal Cost × (1 + Markup %)", { bold: true });
  markupTable(doc, c,
    ["Qty Band", "Markup %", "Multiplier", "Example: $10 marginal cost ->"],
    [
      ["100 - 1,000", `${Math.round(A.grad.fin[0] * 100)}%`, `${(1 + A.grad.fin[0]).toFixed(2)}x`, `$${(10 * (1 + A.grad.fin[0])).toFixed(2)} add-on`],
      ["1,001 - 5,000", `${Math.round(A.grad.fin[1] * 100)}%`, `${(1 + A.grad.fin[1]).toFixed(2)}x`, `$${(10 * (1 + A.grad.fin[1])).toFixed(2)} add-on`],
      ["5,001 - 25,000", `${Math.round(A.grad.fin[2] * 100)}%`, `${(1 + A.grad.fin[2]).toFixed(2)}x`, `$${(10 * (1 + A.grad.fin[2])).toFixed(2)} add-on`],
      ["25,001 - 100,000", `${Math.round(A.grad.fin[3] * 100)}%`, `${(1 + A.grad.fin[3]).toFixed(2)}x`, `$${(10 * (1 + A.grad.fin[3])).toFixed(2)} add-on`],
    ]
  );
  para(doc, c,
    `Catalog impact: ${fmtUsd(opts.finImpact.impactUsd, { signed: true })} (${fmtPct(opts.finImpact.pct)}) vs current implied premiums. ${opts.finImpact.cellsUplifted} cells uplifted, ${opts.finImpact.cellsReduced} reduced.`,
    { italic: true }
  );

  // Section 3 — TURNAROUND / NBD
  heading(doc, c, "3. TURNAROUND — Rush / Next Business Day (NBD)");
  para(doc, c, "Multiplier applied to TOTAL order subtotal (base + bundling + scoring) when customer selects rush.");
  para(doc, c, "Formula:  Final Price = (Base + Add-ons) × (1 + NBD Markup),  capped at current published NBD price", { bold: true });
  markupTable(doc, c,
    ["Qty Band", "Markup %", "Multiplier", "Example: $50 subtotal ->"],
    [
      ["100 - 5,000", `${Math.round(A.grad.nbd[0] * 100)}%`, `${(1 + A.grad.nbd[0]).toFixed(2)}x`, `$${(50 * (1 + A.grad.nbd[0])).toFixed(2)} final`],
      ["5,001 - 25,000", `${Math.round(A.grad.nbd[1] * 100)}%`, `${(1 + A.grad.nbd[1]).toFixed(2)}x`, `$${(50 * (1 + A.grad.nbd[1])).toFixed(2)} final`],
      ["25,001 - 100,000", `${Math.round(A.grad.nbd[2] * 100)}%`, `${(1 + A.grad.nbd[2]).toFixed(2)}x`, `$${(50 * (1 + A.grad.nbd[2])).toFixed(2)} final`],
    ]
  );
  para(doc, c,
    `Estimated 12-mo realized lift: ${fmtUsdK(opts.nbdLift.lowUsd)}–${fmtUsdK(opts.nbdLift.highUsd)} (${opts.nbdLift.orders3mo} NBD orders, $${(opts.nbdLift.baseRevenue3mo / 1000).toFixed(0)}K base revenue, depending on retention).`,
    { italic: true }
  );

  // Section 4 — LOSS LEADERS
  heading(doc, c, "4. LOSS LEADERS — SKUs Sold Below Cost");
  para(doc, c, "Top SKUs ranked by 90-day base order volume. Verdict reflects net 90-day margin (base + add-on).");
  const llRows = opts.lossLeaders.rows.map((r) => [
    r.sizeQty,
    fmtUsd(r.baseMarginUsd, { signed: true }),
    fmtUsd(r.addonMarginUsd, { signed: true }),
    fmtUsd(r.netUsd, { signed: true }),
    r.verdict,
    r.action,
  ]);
  const llTotalLabel = opts.lossLeaders.totals.netUsd >= 0 ? "Net + but thin" : "Net negative — flag";
  const totalRow = [
    `TOTAL (top ${opts.lossLeaders.rows.length})`,
    fmtUsd(opts.lossLeaders.totals.baseMarginUsd, { signed: true }),
    fmtUsd(opts.lossLeaders.totals.addonMarginUsd, { signed: true }),
    fmtUsd(opts.lossLeaders.totals.netUsd, { signed: true }),
    llTotalLabel,
    "—",
  ];
  ensureSpace(doc, c, 60);
  autoTable(doc, {
    startY: c.y,
    head: [["Size x Qty", "Base $", "Add-on $", "Net 90-day $", "Verdict", "Action"]],
    body: cleanRows([...llRows, totalRow]),
    margin: { left: MARGIN, right: MARGIN },
    styles: { font: "helvetica", fontSize: 8, cellPadding: 2.5, lineColor: [180, 180, 180], lineWidth: 0.3, overflow: "linebreak" },
    headStyles: { fillColor: [31, 56, 100], textColor: [255, 255, 255], fontStyle: "bold" },
    didParseCell: (data) => {
      if (data.section === "body" && data.row.index === llRows.length) {
        data.cell.styles.fillColor = [244, 244, 244];
        data.cell.styles.fontStyle = "bold";
      }
    },
    tableWidth: CONTENT_W,
  });
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  c.y = (doc as any).lastAutoTable.finalY + 4;
  if (opts.lossLeaders.callout) {
    const qtys = opts.lossLeaders.callout.qtys
      .map((q) => (q >= 1000 ? `${q / 1000}k` : String(q)))
      .join(" / ");
    para(
      doc,
      c,
      `Note: The ${opts.lossLeaders.callout.sizeFamily} family at ${qtys} qtys drives ${fmtUsd(opts.lossLeaders.callout.totalLoss, { signed: true })} of the loss alone. Structurally unprofitable even with add-ons. Flagged for Q3 pricing review.`,
      { italic: true }
    );
  }

  // Section 5 — RECOMMENDED SCENARIO
  heading(doc, c, "5. RECOMMENDED SCENARIO");
  const rec = opts.recommendation;
  const recRow = rec.recommended;
  const targetBandStr = `[${(opts.targetMinPct * 100).toFixed(1)}%, ${(opts.targetMaxPct * 100).toFixed(1)}%]`;
  para(doc, c, `Target band: ${targetBandStr}. ${rec.inBand.length} of 8 scenarios fall inside this band.`);
  if (recRow) {
    para(doc, c,
      `Recommended: ${recRow.id} — Δ ${fmtUsd(recRow.deltaUsd, { signed: true })} (${(recRow.pctDelta * 100).toFixed(2)}%) over 3 months, annualized ${fmtUsd(recRow.annualizedUsd, { signed: true })}.`,
      { bold: true, size: 9.5 }
    );
  }

  let sopRule: string;
  const aInBand = rec.inBand.some((s) => s.id === "A_Current_Locked");
  if (aInBand && recRow?.id === "A_Current_Locked") {
    sopRule =
      "SOP Rule 1: When A_Current_Locked is in the target band, ship A. Customer impact distribution is already measured and operationally approved.";
  } else if (recRow?.id === "F_Aggressive") {
    sopRule =
      "SOP Rule 2: A_Current_Locked is outside target. F_Aggressive is the SOP pick when finance wants revenue-positive — base/finishing/NBD all lift +5pt.";
  } else if (recRow?.id === "E_Combined_Mod") {
    sopRule =
      "SOP Rule 3: A_Current_Locked is outside target. E_Combined_Mod is the SOP pick when finance wants revenue-positive with minimal customer-impact shift.";
  } else if (rec.inBand.length === 0) {
    sopRule =
      "No scenario fully within target band. Pick is the closest to the band; consider widening the band or revisiting cost assumptions.";
  } else {
    sopRule = `${recRow?.id ?? "—"} is in target band; A is outside and the SOP's revenue-positive picks (F, E) are not available.`;
  }
  para(doc, c, sopRule, { italic: true });

  const altScenarios = [...rec.inBand].sort((a, b) => a.pctDelta - b.pctDelta);
  if (altScenarios.length > 0) {
    para(doc, c, "In-band alternatives (sorted by % Delta):", { size: 8.5 });
    const altBody = altScenarios.map((s) => [
      s.id,
      s.baseFormatted,
      s.finFormatted,
      s.nbdFormatted,
      fmtUsd(s.deltaUsd, { signed: true }),
      `${(s.pctDelta * 100).toFixed(2)}%`,
      s.id === recRow?.id ? "* Recommended" : "+ In band",
    ]);
    ensureSpace(doc, c, 12 + altBody.length * 11);
    const recIdx = altScenarios.findIndex((s) => s.id === recRow?.id);
    autoTable(doc, {
      startY: c.y,
      head: [["Scenario", "Base", "Finishing", "NBD", "3-mo Delta", "% Delta", "Status"]],
      body: cleanRows(altBody),
      margin: { left: MARGIN, right: MARGIN },
      styles: { font: "helvetica", fontSize: 7.5, cellPadding: 2, lineColor: [180, 180, 180], lineWidth: 0.3, overflow: "linebreak" },
      headStyles: { fillColor: [31, 56, 100], textColor: [255, 255, 255], fontStyle: "bold" },
      didParseCell: (data) => {
        if (data.section === "body" && data.row.index === recIdx) {
          data.cell.styles.fontStyle = "bold";
          data.cell.styles.fillColor = [255, 242, 204];
        }
      },
      tableWidth: CONTENT_W,
    });
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    c.y = (doc as any).lastAutoTable.finalY + 4;
  }

  // Reference footer
  c.y += 8;
  doc.setFont("helvetica", "bold");
  doc.setFontSize(10);
  doc.setTextColor(20, 20, 20);
  const refAscender = 10 * ASCENDER_RATIO;
  c.y += refAscender;
  doc.text("Reference", MARGIN, c.y);
  c.y += 4;
  para(doc, c,
    `Notion: Pricing Model Comparison + 7 sub-pages  |  Owner: ${opts.owner || "—"}  |  Locked: ${opts.lockedDate}`,
    { size: 8 }
  );

  return doc.output("blob");
}

export function downloadOnePagerPdf(opts: OnePagerPdfOpts): void {
  const blob = buildOnePagerPdf(opts);
  saveAs(blob, `${opts.productSlug}_Markup_Reference_OnePager.pdf`);
}
