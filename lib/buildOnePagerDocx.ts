import {
  AlignmentType,
  BorderStyle,
  Document,
  HeadingLevel,
  Packer,
  PageOrientation,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
  convertInchesToTwip,
} from "docx";
import { saveAs } from "file-saver";
import { BaseCatalogImpact, FinishingCatalogImpact, NbdLift } from "./catalogImpact";
import { LossLeadersOutput } from "./lossLeaders";
import { SCENARIO_BY_ID } from "./markupEngine";
import { Recommendation } from "./recommend";

export interface OnePagerOpts {
  productName: string;
  productSlug: string;
  lockedDate: string; // YYYY-MM-DD or display string
  owner: string;
  baseImpact: BaseCatalogImpact;
  finImpact: FinishingCatalogImpact;
  nbdLift: NbdLift;
  lossLeaders: LossLeadersOutput;
  recommendation: Recommendation;
  targetMinPct: number;
  targetMaxPct: number;
}

function fmtUsd(v: number, opts: { signed?: boolean; abs?: boolean } = {}): string {
  const abs = Math.abs(v);
  const sign = opts.signed ? (v >= 0 ? "+" : "−") : v < 0 ? "−" : "";
  return `${sign}$${abs.toLocaleString("en-US", {
    minimumFractionDigits: 0,
    maximumFractionDigits: 0,
  })}`;
}

function fmtUsdK(v: number): string {
  const abs = Math.abs(v);
  const sign = v < 0 ? "−" : "+";
  return `${sign}$${(abs / 1000).toFixed(0)}K`;
}

function fmtPct(v: number): string {
  const sign = v < 0 ? "−" : "+";
  return `${sign}${Math.abs(v * 100).toFixed(1)}%`;
}

function fmtPctNoSign(v: number): string {
  return `${(v * 100).toFixed(1)}%`;
}

function paragraph(text: string, opts: { bold?: boolean; italic?: boolean; size?: number; spacing?: number } = {}): Paragraph {
  return new Paragraph({
    spacing: { after: opts.spacing ?? 40 },
    children: [
      new TextRun({
        text,
        bold: opts.bold,
        italics: opts.italic,
        size: opts.size,
      }),
    ],
  });
}

function sectionHeading(text: string): Paragraph {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 140, after: 60 },
    children: [new TextRun({ text, bold: true, size: 22 })],
  });
}

const BORDERS_ALL = {
  top: { style: BorderStyle.SINGLE, size: 4, color: "808080" },
  bottom: { style: BorderStyle.SINGLE, size: 4, color: "808080" },
  left: { style: BorderStyle.SINGLE, size: 4, color: "808080" },
  right: { style: BorderStyle.SINGLE, size: 4, color: "808080" },
};

function cell(text: string, opts: { bold?: boolean; shade?: string; align?: (typeof AlignmentType)[keyof typeof AlignmentType] } = {}): TableCell {
  return new TableCell({
    borders: BORDERS_ALL,
    shading: opts.shade ? { type: ShadingType.CLEAR, fill: opts.shade, color: "auto" } : undefined,
    children: [
      new Paragraph({
        alignment: opts.align ?? AlignmentType.LEFT,
        children: [new TextRun({ text, bold: opts.bold, size: 18 })],
      }),
    ],
  });
}

function headerRow(headers: string[]): TableRow {
  return new TableRow({
    children: headers.map((h) => cell(h, { bold: true, shade: "EFEFEF" })),
    tableHeader: true,
  });
}

function dataRow(vals: string[]): TableRow {
  return new TableRow({ children: vals.map((v) => cell(v)) });
}

function totalRow(vals: string[]): TableRow {
  return new TableRow({
    children: vals.map((v) => cell(v, { bold: true, shade: "F4F4F4" })),
  });
}

function markupTable(
  headers: string[],
  rows: string[][]
): Table {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [headerRow(headers), ...rows.map(dataRow)],
  });
}

export async function buildOnePagerDocx(opts: OnePagerOpts): Promise<Blob> {
  const A = SCENARIO_BY_ID["A_Current_Locked"];
  // Tables: Base, Finishing, NBD
  const baseTable = markupTable(
    ["Qty Band", "Markup %", "Multiplier", "Example: $30 cost  →"],
    [
      ["100 – 1,000", `${Math.round(A.grad.base[0] * 100)}%`, `${(1 + A.grad.base[0]).toFixed(2)}×`, `$${(30 * (1 + A.grad.base[0])).toFixed(2)}`],
      ["1,001 – 5,000", `${Math.round(A.grad.base[1] * 100)}%`, `${(1 + A.grad.base[1]).toFixed(2)}×`, `$${(30 * (1 + A.grad.base[1])).toFixed(2)}`],
      ["5,001 – 25,000", `${Math.round(A.grad.base[2] * 100)}%`, `${(1 + A.grad.base[2]).toFixed(2)}×`, `$${(30 * (1 + A.grad.base[2])).toFixed(2)}`],
      ["25,001 – 100,000", `${Math.round(A.grad.base[3] * 100)}%`, `${(1 + A.grad.base[3]).toFixed(2)}×`, `$${(30 * (1 + A.grad.base[3])).toFixed(2)}`],
    ]
  );

  const finTable = markupTable(
    ["Qty Band", "Markup %", "Multiplier", "Example: $10 marginal cost  →"],
    [
      ["100 – 1,000", `${Math.round(A.grad.fin[0] * 100)}%`, `${(1 + A.grad.fin[0]).toFixed(2)}×`, `$${(10 * (1 + A.grad.fin[0])).toFixed(2)} add-on`],
      ["1,001 – 5,000", `${Math.round(A.grad.fin[1] * 100)}%`, `${(1 + A.grad.fin[1]).toFixed(2)}×`, `$${(10 * (1 + A.grad.fin[1])).toFixed(2)} add-on`],
      ["5,001 – 25,000", `${Math.round(A.grad.fin[2] * 100)}%`, `${(1 + A.grad.fin[2]).toFixed(2)}×`, `$${(10 * (1 + A.grad.fin[2])).toFixed(2)} add-on`],
      ["25,001 – 100,000", `${Math.round(A.grad.fin[3] * 100)}%`, `${(1 + A.grad.fin[3]).toFixed(2)}×`, `$${(10 * (1 + A.grad.fin[3])).toFixed(2)} add-on`],
    ]
  );

  const nbdTable = markupTable(
    ["Qty Band", "Markup %", "Multiplier", "Example: $50 subtotal  →"],
    [
      ["100 – 5,000", `${Math.round(A.grad.nbd[0] * 100)}%`, `${(1 + A.grad.nbd[0]).toFixed(2)}×`, `$${(50 * (1 + A.grad.nbd[0])).toFixed(2)} final`],
      ["5,001 – 25,000", `${Math.round(A.grad.nbd[1] * 100)}%`, `${(1 + A.grad.nbd[1]).toFixed(2)}×`, `$${(50 * (1 + A.grad.nbd[1])).toFixed(2)} final`],
      ["25,001 – 100,000", `${Math.round(A.grad.nbd[2] * 100)}%`, `${(1 + A.grad.nbd[2]).toFixed(2)}×`, `$${(50 * (1 + A.grad.nbd[2])).toFixed(2)} final`],
    ]
  );

  // Loss leader table — 6 columns
  const llHeader = headerRow(["Size × Qty", "Base $", "Add-on $", "Net 90-day $", "Verdict", "Action"]);
  const llRows = opts.lossLeaders.rows.map((r) =>
    new TableRow({
      children: [
        cell(r.sizeQty),
        cell(fmtUsd(r.baseMarginUsd, { signed: true })),
        cell(fmtUsd(r.addonMarginUsd, { signed: true })),
        cell(fmtUsd(r.netUsd, { signed: true })),
        cell(r.verdict),
        cell(r.action),
      ],
    })
  );
  const llTotalLabel =
    opts.lossLeaders.totals.netUsd >= 0 ? "Net + but thin" : "Net negative — flag";
  const llTotal = totalRow([
    `TOTAL (top ${opts.lossLeaders.rows.length})`,
    fmtUsd(opts.lossLeaders.totals.baseMarginUsd, { signed: true }),
    fmtUsd(opts.lossLeaders.totals.addonMarginUsd, { signed: true }),
    fmtUsd(opts.lossLeaders.totals.netUsd, { signed: true }),
    llTotalLabel,
    "—",
  ]);
  const llTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [llHeader, ...llRows, llTotal],
  });

  const concLine = opts.baseImpact.concentrationLabel
    ? ` ${opts.baseImpact.concentrationLabel}`
    : "";

  const calloutText = opts.lossLeaders.callout
    ? `⚠ The ${opts.lossLeaders.callout.sizeFamily} family at ${opts.lossLeaders.callout.qtys
        .map((q) => (q >= 1000 ? `${q / 1000}k` : String(q)))
        .join(" / ")} qtys drives ${fmtUsd(opts.lossLeaders.callout.totalLoss, {
        signed: true,
      })} of the loss alone. Structurally unprofitable even with add-ons. Flagged for Q3 pricing review.`
    : "";

  // Build Recommendation table — recommended scenario + in-band alternatives
  const rec = opts.recommendation;
  const recRow = rec.recommended;
  const recPctStr = recRow ? `${(recRow.pctDelta * 100).toFixed(2)}%` : "—";
  const recDeltaStr = recRow ? fmtUsd(recRow.deltaUsd, { signed: true }) : "—";
  const recAnnualStr = recRow ? fmtUsd(recRow.annualizedUsd, { signed: true }) : "—";
  const targetBandStr = `[${(opts.targetMinPct * 100).toFixed(1)}%, ${(opts.targetMaxPct * 100).toFixed(1)}%]`;

  // Detect which SOP rule fired
  const aInBand = rec.inBand.some((s) => s.id === "A_Current_Locked");
  let sopRule = "";
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

  // Alternatives table — show every in-band scenario sorted by %Δ ascending
  const altScenarios = [...rec.inBand].sort((a, b) => a.pctDelta - b.pctDelta);
  const altHeader = headerRow([
    "Scenario",
    "Base",
    "Finishing",
    "NBD",
    "3-mo Δ",
    "% Δ",
    "Status",
  ]);
  const altRows = altScenarios.map((s) => {
    const isPick = s.id === recRow?.id;
    return new TableRow({
      children: [
        cell(s.id, { bold: isPick }),
        cell(s.baseFormatted),
        cell(s.finFormatted),
        cell(s.nbdFormatted),
        cell(fmtUsd(s.deltaUsd, { signed: true })),
        cell(`${(s.pctDelta * 100).toFixed(2)}%`),
        cell(isPick ? "★ Recommended" : "✓ In band", { bold: isPick }),
      ],
    });
  });
  const altTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [altHeader, ...altRows],
  });

  const doc = new Document({
    creator: "SinaLite Markup Tuning",
    title: `${opts.productName} — Markup Reference`,
    styles: {
      default: {
        document: { run: { font: "Calibri", size: 18 } },
      },
    },
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: convertInchesToTwip(0.27),
              bottom: convertInchesToTwip(0.27),
              left: convertInchesToTwip(0.4),
              right: convertInchesToTwip(0.4),
            },
            size: { orientation: PageOrientation.PORTRAIT },
          },
        },
        children: [
          new Paragraph({
            spacing: { after: 80 },
            children: [
              new TextRun({
                text: `${opts.productName} — Locked Markup Reference`,
                bold: true,
                size: 28,
              }),
            ],
          }),
          paragraph(
            `Decisions locked ${opts.lockedDate}  |  Cap rule applies to ALL: final price = MIN(model price, current published price)`,
            { italic: true }
          ),

          sectionHeading("1. BASE Product Markup"),
          paragraph(
            "Applied to PE3 base cost. Volume-discount gradient — markup decreases as qty rises."
          ),
          paragraph(
            "Formula:  Base Price = MIN(PE3 Base Cost × (1 + Markup %), Current Sale Price)",
            { bold: true }
          ),
          baseTable,
          paragraph(
            `Catalog impact: ${fmtUsd(opts.baseImpact.impactUsd, { signed: true })} (${fmtPct(opts.baseImpact.pct)}) on base SKUs.${concLine} Cap rule means no customer ever sees a price increase.`,
            { italic: true, spacing: 60 }
          ),

          sectionHeading("2. UPCHARGES — Finishing Add-Ons"),
          paragraph(
            "Applies to all 8 bundling types (Single band 100s/50s/25s, Double band 100s/50s/25s, Shrink Wrap 50s/25s) + Score in Half."
          ),
          paragraph("Formula:  Add-on Price = PE3 Marginal Cost × (1 + Markup %)", {
            bold: true,
          }),
          finTable,
          paragraph(
            `Catalog impact: ${fmtUsd(opts.finImpact.impactUsd, { signed: true })} (${fmtPct(opts.finImpact.pct)}) vs current implied premiums. ${opts.finImpact.cellsUplifted} cells uplifted, ${opts.finImpact.cellsReduced} reduced.`,
            { italic: true, spacing: 60 }
          ),

          sectionHeading("3. TURNAROUND — Rush / Next Business Day (NBD)"),
          paragraph(
            "Multiplier applied to TOTAL order subtotal (base + bundling + scoring) when customer selects rush."
          ),
          paragraph(
            "Formula:  Final Price = (Base + Add-ons) × (1 + NBD Markup),  capped at current published NBD price",
            { bold: true }
          ),
          nbdTable,
          paragraph(
            `Estimated 12-mo realized lift: ${fmtUsdK(opts.nbdLift.lowUsd)}–${fmtUsdK(opts.nbdLift.highUsd)} (${opts.nbdLift.orders3mo} NBD orders, $${(opts.nbdLift.baseRevenue3mo / 1000).toFixed(0)}K base revenue, depending on retention).`,
            { italic: true, spacing: 60 }
          ),

          sectionHeading("4. LOSS LEADERS — SKUs Sold Below Cost"),
          paragraph(
            "Top SKUs ranked by 90-day base order volume. Verdict reflects net 90-day margin (base + add-on)."
          ),
          llTable,
          ...(calloutText
            ? [paragraph(calloutText, { italic: true, spacing: 80 })]
            : []),

          sectionHeading("5. RECOMMENDED SCENARIO"),
          paragraph(
            `Target band: ${targetBandStr}. ${rec.inBand.length} of 8 scenarios fall inside this band.`
          ),
          paragraph(
            `Recommended: ${recRow?.id ?? "—"} — Δ ${recDeltaStr} (${recPctStr}) over 3 months, annualized ${recAnnualStr}.`,
            { bold: true }
          ),
          paragraph(sopRule, { italic: true, spacing: 80 }),
          ...(altScenarios.length > 0
            ? [
                paragraph("In-band alternatives (sorted by % Δ):", { spacing: 40 }),
                altTable,
              ]
            : []),

          new Paragraph({
            spacing: { before: 120, after: 40 },
            children: [new TextRun({ text: "Reference", bold: true, size: 20 })],
          }),
          paragraph(
            `Notion: Pricing Model Comparison + 7 sub-pages  |  Owner: ${opts.owner || "—"}  |  Locked: ${opts.lockedDate}`
          ),
        ],
      },
    ],
  });

  const blob = await Packer.toBlob(doc);
  return blob;
}

export async function downloadOnePagerDocx(opts: OnePagerOpts): Promise<void> {
  const blob = await buildOnePagerDocx(opts);
  saveAs(blob, `${opts.productSlug}_Markup_Reference_OnePager.docx`);
}

// Suppress unused import warnings used for re-export hints
void fmtPctNoSign;
