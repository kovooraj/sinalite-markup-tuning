import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import { ScenarioResult } from "./computeScenarios";
import { BASE_BAND_LABELS } from "./qtyBands";
import { Recommendation } from "./recommend";

export interface BuildScenariosOpts {
  productName: string;
  productSlug: string;
  results: ScenarioResult[];
  recommendation: Recommendation;
  targetMinPct: number;
  targetMaxPct: number;
  applyCapRule?: boolean;
}

const CURRENCY_FMT = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)';
const PCT_FMT = "0.00%";

export async function buildScenariosXlsx(opts: BuildScenariosOpts): Promise<Blob> {
  const wb = new ExcelJS.Workbook();
  wb.creator = "SinaLite Markup Tuning";
  wb.created = new Date();
  const ws = wb.addWorksheet("Markup Tuning Scenarios", {
    views: [{ state: "frozen", ySplit: 3 }],
  });

  // Column widths
  ws.columns = [
    { width: 22 },
    { width: 70 },
    { width: 18 },
    { width: 18 },
    { width: 18 },
    { width: 14 },
    { width: 10 },
    { width: 14 },
  ];

  // R1: Title (merged A1:H1)
  ws.mergeCells("A1:H1");
  const titleCell = ws.getCell("A1");
  titleCell.value = "Markup Tuning Scenarios — 3-Month Order Replay";
  titleCell.font = { bold: true, size: 14 };
  titleCell.alignment = { vertical: "middle" };

  // R2: Subtitle
  const capLabel = (opts.applyCapRule ?? true) ? "Cap rule: APPLIED" : "Cap rule: DISABLED";
  ws.mergeCells("A2:H2");
  ws.getCell("A2").value =
    `${capLabel} · Δ = New Price ${(opts.applyCapRule ?? true) ? "(capped)" : "(uncapped)"} − PE3 List Price, summed over 3 months.`;
  ws.getCell("A2").font = { italic: true, color: { argb: "FF555555" } };

  // R4: Scenario header — delta is now "New Price (capped) − PE3 List Price"
  const header1 = [
    "Scenario",
    "Description",
    "Base Markup",
    "Finishing Markup",
    "NBD Markup",
    "3-mo Δ vs PE3 List",
    "% Δ",
    "Annualized",
  ];
  ws.getRow(4).values = header1;
  ws.getRow(4).font = { bold: true };
  ws.getRow(4).alignment = { horizontal: "left" };

  // R5..R12: per-scenario rows — highlight the recommended scenario in yellow
  const recommendedId = opts.recommendation.recommended?.id;
  opts.results.forEach((s, i) => {
    const r = 5 + i;
    ws.getRow(r).values = [
      s.id,
      s.label,
      s.baseFormatted,
      s.finFormatted,
      s.nbdFormatted,
      s.deltaUsd,
      s.pctDelta,
      s.annualizedUsd,
    ];
    ws.getCell(`F${r}`).numFmt = CURRENCY_FMT;
    ws.getCell(`G${r}`).numFmt = PCT_FMT;
    ws.getCell(`H${r}`).numFmt = CURRENCY_FMT;
    if (s.id === recommendedId) {
      const yellow: ExcelJS.FillPattern = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFFF2CC" },
      };
      for (const col of ["A", "B", "C", "D", "E", "F", "G", "H"]) {
        const cell = ws.getCell(`${col}${r}`);
        cell.fill = yellow;
        cell.font = { ...(cell.font ?? {}), bold: true };
      }
    }
  });

  // Section offsets are dynamic — saved scenarios can push the count past
  // the original 8, so each section starts 3 rows after the previous one ends
  // (matching the original fixed layout's spacing at n = 8).
  const n = opts.results.length;
  const distTitleRow = 4 + n + 3; // 15 when n = 8

  // Customer Impact Distribution title
  ws.mergeCells(`A${distTitleRow}:H${distTitleRow}`);
  ws.getCell(`A${distTitleRow}`).value =
    "Customer Impact Distribution (% of orders)";
  ws.getCell(`A${distTitleRow}`).font = { bold: true };

  // Header
  const header2 = [
    "Scenario",
    "No Change",
    "Decrease 1-5%",
    "Decrease >5%",
    "Increase 1-5%",
    "Increase >5%",
  ];
  ws.getRow(distTitleRow + 1).values = header2;
  ws.getRow(distTitleRow + 1).font = { bold: true };

  // Per scenario
  opts.results.forEach((s, i) => {
    const r = distTitleRow + 2 + i;
    ws.getRow(r).values = [
      s.id,
      s.dist.noChange,
      s.dist.decrease1to5,
      s.dist.decreaseGt5,
      s.dist.increase1to5,
      s.dist.increaseGt5,
    ];
    for (const col of ["B", "C", "D", "E", "F"]) {
      ws.getCell(`${col}${r}`).numFmt = PCT_FMT;
    }
  });

  // 3-Month $ Delta by Qty Band title
  const bandTitleRow = distTitleRow + 1 + n + 3; // 27 when n = 8
  ws.mergeCells(`A${bandTitleRow}:H${bandTitleRow}`);
  ws.getCell(`A${bandTitleRow}`).value = "3-Month $ Delta by Qty Band";
  ws.getCell(`A${bandTitleRow}`).font = { bold: true };

  // Header
  const header3 = ["Scenario", ...BASE_BAND_LABELS];
  ws.getRow(bandTitleRow + 1).values = header3;
  ws.getRow(bandTitleRow + 1).font = { bold: true };

  // Per scenario
  opts.results.forEach((s, i) => {
    const r = bandTitleRow + 2 + i;
    ws.getRow(r).values = [s.id, ...s.deltaByBand];
    for (const col of ["B", "C", "D", "E"]) {
      ws.getCell(`${col}${r}`).numFmt = CURRENCY_FMT;
    }
  });

  // Recommendation block — pick + SOP-aligned reasoning
  const recTitleRow = bandTitleRow + 1 + n + 3; // 39 when n = 8
  const rec = opts.recommendation;
  const recRow = rec.recommended;
  const targetBandStr = `[${(opts.targetMinPct * 100).toFixed(1)}%, ${(opts.targetMaxPct * 100).toFixed(1)}%]`;
  ws.mergeCells(`A${recTitleRow}:H${recTitleRow}`);
  ws.getCell(`A${recTitleRow}`).value =
    "Recommendation (per sinalite-pricing-model SOP)";
  ws.getCell(`A${recTitleRow}`).font = { bold: true, size: 12 };

  ws.mergeCells(`A${recTitleRow + 1}:H${recTitleRow + 1}`);
  ws.getCell(`A${recTitleRow + 1}`).value =
    `Target band: ${targetBandStr} · ${rec.inBand.length} of ${rec.inBand.length + rec.outOfBand.length} scenarios in band`;
  ws.getCell(`A${recTitleRow + 1}`).font = {
    italic: true,
    color: { argb: "FF555555" },
  };

  if (recRow) {
    ws.mergeCells(`A${recTitleRow + 2}:H${recTitleRow + 2}`);
    ws.getCell(`A${recTitleRow + 2}`).value = `Pick: ${recRow.id} — Δ ${recRow.deltaUsd >= 0 ? "+" : "−"}$${Math.abs(recRow.deltaUsd).toLocaleString("en-US", { maximumFractionDigits: 0 })} (${(recRow.pctDelta * 100).toFixed(2)}%) over 3 months, annualized ${recRow.annualizedUsd >= 0 ? "+" : "−"}$${Math.abs(recRow.annualizedUsd).toLocaleString("en-US", { maximumFractionDigits: 0 })}`;
    ws.getCell(`A${recTitleRow + 2}`).font = { bold: true };
  }

  ws.mergeCells(`A${recTitleRow + 3}:H${recTitleRow + 5}`);
  ws.getCell(`A${recTitleRow + 3}`).value = rec.reason;
  ws.getCell(`A${recTitleRow + 3}`).alignment = {
    wrapText: true,
    vertical: "top",
  };
  ws.getRow(recTitleRow + 3).height = 18;
  ws.getRow(recTitleRow + 4).height = 18;
  ws.getRow(recTitleRow + 5).height = 18;

  const buffer = await wb.xlsx.writeBuffer();
  return new Blob([buffer as ArrayBuffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

export async function downloadScenariosXlsx(opts: BuildScenariosOpts): Promise<void> {
  const blob = await buildScenariosXlsx(opts);
  saveAs(blob, `${opts.productSlug}_Markup_Tuning_Scenarios.xlsx`);
}
