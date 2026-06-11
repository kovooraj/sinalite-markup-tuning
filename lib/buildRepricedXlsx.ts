import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import { PriceEngineData } from "./parsePriceEngine";
import { computeNewPrice, ScenarioDef, resolveScenario } from "./markupEngine";
import { bandOf, nbdBandOf } from "./qtyBands";

export interface BuildRepricedOpts {
  pe: PriceEngineData;
  scenarioId: string;
  usdRate: number;
  /** When true (default) the cap rule applies: new price never exceeds the
   * PE3 list price. When false the new price floats free and delta vs list
   * can go positive. */
  applyCapRule?: boolean;
  /** User-defined custom scenario (from the Custom scenario checkbox). When
   * scenarioId matches its id, this definition is used instead of the
   * built-in A–H table. */
  customScenario?: ScenarioDef | null;
}

const FMT_CURRENCY = '"$"#,##0.00;("$"#,##0.00);"-"';
const FMT_PCT = "0.0%";
const FMT_INT = "#,##0";

const HEADER_FILL: ExcelJS.FillPattern = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "FF1F3864" },
};
const HIGHLIGHT_FILL: ExcelJS.FillPattern = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "FFFFF2CC" },
};
const HEADER_FONT: Partial<ExcelJS.Font> = {
  name: "Arial",
  bold: true,
  color: { argb: "FFFFFFFF" },
  size: 11,
};
const BODY_FONT: Partial<ExcelJS.Font> = { name: "Arial", size: 10 };

// Map dimension key → display column header for output
const DIM_DISPLAY: Record<string, string> = {
  coating: "Coating",
  bundling: "Bundling",
  scoring: "Scoring",
  cover: "Cover",
  binding: "Binding",
  pages: "Pages",
  sides: "Sides",
  finishing: "Finishing",
  lamination: "Lamination",
  foil: "Foil",
  embossing: "Embossing",
  diecutting: "Die Cutting",
  corner: "Corner",
  perforation: "Perforation",
  drilling: "Drilling",
};

export async function buildRepricedXlsx(opts: BuildRepricedOpts): Promise<Blob> {
  const scenario: ScenarioDef | undefined = resolveScenario(
    opts.scenarioId,
    opts.customScenario
  );
  if (!scenario) throw new Error(`Unknown scenario: ${opts.scenarioId}`);

  const wb = new ExcelJS.Workbook();
  wb.creator = "SinaLite Markup Tuning";
  wb.created = new Date();
  const ws = wb.addWorksheet("Repriced", {
    views: [{ state: "frozen", xSplit: 3, ySplit: 1 }],
  });

  const usdHeader = `New Price USD (@${opts.usdRate.toFixed(2)})`;

  // Dynamic columns: standard fields + product-specific dimensions + computed
  type Col = { header: string; key: string; width: number; fmt?: string };
  const cols: Col[] = [
    { header: "Product", key: "product", width: 36 },
    { header: "qty", key: "qty", width: 9, fmt: FMT_INT },
    { header: "Turnaround", key: "turnaround", width: 22 },
    { header: "size", key: "size", width: 10 },
    { header: "Stock", key: "stock", width: 32 },
  ];
  for (const dim of opts.pe.dimensions) {
    cols.push({
      header: DIM_DISPLAY[dim] ?? dim,
      key: `dim_${dim}`,
      width: 18,
    });
  }
  cols.push(
    { header: "sale price", key: "salePrice", width: 12, fmt: FMT_CURRENCY },
    { header: "PE3 cost no markup", key: "pe3Cost", width: 16, fmt: FMT_CURRENCY },
    { header: "Consolidated Markup", key: "consolidatedMarkup", width: 16, fmt: FMT_PCT },
    { header: "Base Cost (CAD)", key: "baseCost", width: 14, fmt: FMT_CURRENCY },
    { header: "Finishing Cost (CAD)", key: "finCost", width: 16, fmt: FMT_CURRENCY },
    { header: "Base Markup %", key: "baseMarkup", width: 12, fmt: FMT_PCT },
    { header: "Finishing Markup %", key: "finMarkup", width: 14, fmt: FMT_PCT },
    { header: "NBD Markup %", key: "nbdMarkup", width: 12, fmt: FMT_PCT },
    { header: "Marked-up Base (CAD)", key: "markedBase", width: 16, fmt: FMT_CURRENCY },
    { header: "Marked-up Finishing (CAD)", key: "markedFin", width: 18, fmt: FMT_CURRENCY },
    { header: "Subtotal Before NBD (CAD)", key: "subtotalPreNbd", width: 18, fmt: FMT_CURRENCY },
    { header: "New Price CAD", key: "newCad", width: 14, fmt: FMT_CURRENCY },
    { header: usdHeader, key: "newUsd", width: 18, fmt: FMT_CURRENCY },
    { header: "Delta vs sale price (CAD)", key: "delta", width: 18, fmt: FMT_CURRENCY }
  );

  ws.columns = cols.map(({ header, key, width }) => ({ header, key, width }));

  const headerRow = ws.getRow(1);
  headerRow.eachCell((cell) => {
    cell.fill = HEADER_FILL;
    cell.font = HEADER_FONT;
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  });
  headerRow.height = 30;

  for (const r of opts.pe.rows) {
    const band = bandOf(r.qty);
    const baseMkup = scenario.grad.base[band];
    const finMkup = scenario.grad.fin[band];
    const isRush = r.turnaround === "Rush (NBD)";
    const nbdMkup = isRush ? scenario.grad.nbd[nbdBandOf(r.qty)] : 0;
    const res = computeNewPrice(
      {
        qty: r.qty,
        baseCost: r.baseCost,
        finCost: r.finCost,
        isRush,
        currentSalePrice: r.currentSalePrice,
      },
      scenario.grad,
      opts.applyCapRule ?? true
    );
    const subtotalPreNbd = res.basePrice + res.finPrice;
    const newCad = res.finalPrice;
    const newUsd = newCad * opts.usdRate;
    const delta = newCad - r.currentSalePrice;
    const rowData: Record<string, string | number> = {
      product: r.productName,
      qty: r.qty,
      turnaround: r.rawTurnaround,
      size: r.size,
      stock: r.stock,
      salePrice: r.currentSalePrice,
      pe3Cost: r.pe3CostTotal,
      consolidatedMarkup: r.consolidatedMarkup,
      baseCost: round(r.baseCost, 4),
      finCost: round(r.finCost, 4),
      baseMarkup: baseMkup,
      finMarkup: finMkup,
      nbdMarkup: nbdMkup,
      markedBase: round(res.basePrice, 4),
      markedFin: round(res.finPrice, 4),
      subtotalPreNbd: round(subtotalPreNbd, 2),
      newCad: round(newCad, 2),
      newUsd: round(newUsd, 2),
      delta: round(delta, 2),
    };
    for (const dim of opts.pe.dimensions) {
      rowData[`dim_${dim}`] = r.dims[dim] ?? "";
    }
    ws.addRow(rowData);
  }

  const lastRow = ws.rowCount;
  for (let r = 2; r <= lastRow; r++) {
    const row = ws.getRow(r);
    row.eachCell((cell) => {
      cell.font = BODY_FONT;
    });
    for (const c of cols) {
      if (!c.fmt) continue;
      const colNum = ws.getColumn(c.key).number;
      row.getCell(colNum).numFmt = c.fmt;
    }
    row.getCell(ws.getColumn("newCad").number).fill = HIGHLIGHT_FILL;
    row.getCell(ws.getColumn("newUsd").number).fill = HIGHLIGHT_FILL;
  }

  const buf = await wb.xlsx.writeBuffer();
  return new Blob([buf as ArrayBuffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

function round(n: number, digits: number): number {
  const m = Math.pow(10, digits);
  return Math.round(n * m) / m;
}

export async function downloadRepricedXlsx(opts: BuildRepricedOpts): Promise<void> {
  const blob = await buildRepricedXlsx(opts);
  const scenarioSuffix =
    opts.scenarioId === "A_Current_Locked" ? "" : `_${opts.scenarioId}`;
  saveAs(blob, `${opts.pe.productSlug}_Repriced${scenarioSuffix}.xlsx`);
}
