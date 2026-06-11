"use client";

import { useState } from "react";
import { Gradient, SCENARIO_BY_ID } from "@/lib/markupEngine";
import { BASE_BAND_LABELS, NBD_BAND_LABELS } from "@/lib/qtyBands";

/** Editable gradient values, stored as fractions (0.35 = 35%). */
export interface CustomGradientValues {
  base: [number, number, number, number];
  fin: [number, number, number, number];
  nbd: [number, number, number];
}

export function defaultCustomGradient(): CustomGradientValues {
  const a: Gradient = SCENARIO_BY_ID["A_Current_Locked"].grad;
  return {
    base: [...a.base] as CustomGradientValues["base"],
    fin: [...a.fin] as CustomGradientValues["fin"],
    nbd: [...a.nbd] as CustomGradientValues["nbd"],
  };
}

interface Props {
  open: boolean;
  values: CustomGradientValues;
  onApply: (next: CustomGradientValues) => void;
  onCancel: () => void;
}

/** Percent strings for editing — one per band, e.g. "35". */
interface DraftStrings {
  base: [string, string, string, string];
  fin: [string, string, string, string];
  nbd: [string, string, string];
}

function toDraft(v: CustomGradientValues): DraftStrings {
  const f = (x: number) => String(Math.round(x * 10000) / 100);
  return {
    base: v.base.map(f) as DraftStrings["base"],
    fin: v.fin.map(f) as DraftStrings["fin"],
    nbd: v.nbd.map(f) as DraftStrings["nbd"],
  };
}

function parseDraft(d: DraftStrings): CustomGradientValues | null {
  const p = (s: string): number | null => {
    const n = Number(s);
    if (!isFinite(n) || s.trim() === "") return null;
    if (n < 0 || n > 500) return null;
    return n / 100;
  };
  const base = d.base.map(p);
  const fin = d.fin.map(p);
  const nbd = d.nbd.map(p);
  if ([...base, ...fin, ...nbd].some((v) => v === null)) return null;
  return {
    base: base as CustomGradientValues["base"],
    fin: fin as CustomGradientValues["fin"],
    nbd: nbd as CustomGradientValues["nbd"],
  };
}

function SectionTable({
  title,
  subtitle,
  labels,
  values,
  onChange,
}: {
  title: string;
  subtitle: string;
  labels: readonly string[];
  values: string[];
  onChange: (idx: number, v: string) => void;
}) {
  return (
    <div>
      <div className="text-sm font-semibold text-zinc-800">{title}</div>
      <div className="mb-2 text-xs text-zinc-500">{subtitle}</div>
      <table className="w-full border-collapse text-xs">
        <thead>
          <tr className="bg-zinc-100 text-left">
            <th className="border border-zinc-200 px-2 py-1">Qty Band</th>
            <th className="border border-zinc-200 px-2 py-1 w-32">Markup %</th>
            <th className="border border-zinc-200 px-2 py-1 w-24">Multiplier</th>
          </tr>
        </thead>
        <tbody>
          {labels.map((label, i) => {
            const n = Number(values[i]);
            const ok = isFinite(n) && values[i].trim() !== "" && n >= 0 && n <= 500;
            return (
              <tr key={label}>
                <td className="border border-zinc-200 px-2 py-1">{label}</td>
                <td className="border border-zinc-200 px-2 py-1">
                  <div className="flex items-center gap-1">
                    <input
                      type="number"
                      step="1"
                      min="0"
                      max="500"
                      value={values[i]}
                      onChange={(e) => onChange(i, e.target.value)}
                      className={[
                        "w-20 rounded border px-2 py-1 text-right text-xs",
                        ok ? "border-zinc-300" : "border-red-400 bg-red-50",
                      ].join(" ")}
                    />
                    <span className="text-zinc-500">%</span>
                  </div>
                </td>
                <td className="border border-zinc-200 px-2 py-1 text-zinc-600">
                  {ok ? `${(1 + n / 100).toFixed(2)}×` : "—"}
                </td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

export function CustomScenarioModal({ open, values, onApply, onCancel }: Props) {
  // The body is mounted only while open, so its draft state re-seeds from the
  // saved values on every open — Cancel → reopen never shows abandoned edits.
  if (!open) return null;
  return <ModalBody values={values} onApply={onApply} onCancel={onCancel} />;
}

function ModalBody({ values, onApply, onCancel }: Omit<Props, "open">) {
  const [draft, setDraft] = useState<DraftStrings>(() => toDraft(values));

  const parsed = parseDraft(draft);

  function setBand(section: keyof DraftStrings, idx: number, v: string) {
    setDraft((d) => {
      const next = { ...d, [section]: [...d[section]] } as DraftStrings;
      (next[section] as string[])[idx] = v;
      return next;
    });
  }

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4"
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) onCancel();
      }}
    >
      <div className="max-h-[90vh] w-full max-w-lg overflow-y-auto rounded-lg bg-white p-5 shadow-xl">
        <h2 className="text-lg font-semibold text-zinc-900">
          Custom scenario — markup breakdown
        </h2>
        <p className="mt-1 text-xs text-zinc-600">
          Edit the markup % per qty band. Defaults are Scenario A (current
          locked). When applied, the custom scenario is added to the 3-month
          tuning table and the 1-pager, scenarios xlsx, repriced catalog, and
          annotated orders are all based on it.
        </p>

        <div className="mt-4 space-y-5">
          <SectionTable
            title="1. BASE product markup"
            subtitle="Applied to PE3 base cost. Locked default: 35/30/25/20%."
            labels={BASE_BAND_LABELS}
            values={draft.base}
            onChange={(i, v) => setBand("base", i, v)}
          />
          <SectionTable
            title="2. FINISHING add-on markup"
            subtitle="Applied to PE3 marginal (finishing) cost. Locked default: 60/50/40/35%."
            labels={BASE_BAND_LABELS}
            values={draft.fin}
            onChange={(i, v) => setBand("fin", i, v)}
          />
          <SectionTable
            title="3. TURNAROUND — Rush / NBD markup"
            subtitle="Multiplier on the full subtotal (base + add-ons) for rush orders. Locked default: 30/25/20%."
            labels={NBD_BAND_LABELS}
            values={draft.nbd}
            onChange={(i, v) => setBand("nbd", i, v)}
          />
        </div>

        {!parsed && (
          <p className="mt-3 text-xs text-red-600">
            Fix the highlighted fields — each markup must be a number between 0
            and 500.
          </p>
        )}

        <div className="mt-5 flex justify-end gap-3">
          <button
            type="button"
            onClick={onCancel}
            className="rounded border border-zinc-300 px-4 py-2 text-sm font-medium text-zinc-700 hover:bg-zinc-50"
          >
            Cancel
          </button>
          <button
            type="button"
            disabled={!parsed}
            onClick={() => parsed && onApply(parsed)}
            className="rounded bg-blue-600 px-4 py-2 text-sm font-medium text-white hover:bg-blue-700 disabled:opacity-50"
          >
            Apply custom scenario
          </button>
        </div>
      </div>
    </div>
  );
}
