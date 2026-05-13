import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import { ScenarioResult } from "./computeScenarios";
import { BASE_BAND_LABELS } from "./qtyBands";

export interface BuildScenariosOpts {
  productName: string;
  productSlug: string;
  results: ScenarioResult[];
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
  ws.mergeCells("A2:H2");
  ws.getCell("A2").value =
    "Goal: find markup combinations that are revenue-neutral or positive while keeping customer impact contained.";
  ws.getCell("A2").font = { italic: true, color: { argb: "FF555555" } };

  // R4: Scenario header
  const header1 = [
    "Scenario",
    "Description",
    "Base Markup",
    "Finishing Markup",
    "NBD Markup",
    "3-mo Δ",
    "% Δ",
    "Annualized",
  ];
  ws.getRow(4).values = header1;
  ws.getRow(4).font = { bold: true };
  ws.getRow(4).alignment = { horizontal: "left" };

  // R5..R12: per-scenario rows
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
  });

  // R15: Customer Impact Distribution title
  ws.mergeCells("A15:H15");
  ws.getCell("A15").value = "Customer Impact Distribution (% of orders)";
  ws.getCell("A15").font = { bold: true };

  // R16: Header
  const header2 = [
    "Scenario",
    "No Change",
    "Decrease 1-5%",
    "Decrease >5%",
    "Increase 1-5%",
    "Increase >5%",
  ];
  ws.getRow(16).values = header2;
  ws.getRow(16).font = { bold: true };

  // R17..R24: per scenario
  opts.results.forEach((s, i) => {
    const r = 17 + i;
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

  // R27: 3-Month $ Delta by Qty Band title
  ws.mergeCells("A27:H27");
  ws.getCell("A27").value = "3-Month $ Delta by Qty Band";
  ws.getCell("A27").font = { bold: true };

  // R28: header
  const header3 = ["Scenario", ...BASE_BAND_LABELS];
  ws.getRow(28).values = header3;
  ws.getRow(28).font = { bold: true };

  // R29..R36: per scenario
  opts.results.forEach((s, i) => {
    const r = 29 + i;
    ws.getRow(r).values = [s.id, ...s.deltaByBand];
    for (const col of ["B", "C", "D", "E"]) {
      ws.getCell(`${col}${r}`).numFmt = CURRENCY_FMT;
    }
  });

  const buffer = await wb.xlsx.writeBuffer();
  return new Blob([buffer as ArrayBuffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

export async function downloadScenariosXlsx(opts: BuildScenariosOpts): Promise<void> {
  const blob = await buildScenariosXlsx(opts);
  saveAs(blob, `${opts.productSlug}_Markup_Tuning_Scenarios.xlsx`);
}
