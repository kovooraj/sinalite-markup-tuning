"use client";

import { InfoIcon } from "./InfoIcon";

export interface MetaFormValues {
  lockedDate: string;
  owner: string;
  targetMinPct: number;
  targetMaxPct: number;
}

interface Props {
  values: MetaFormValues;
  onChange: (next: MetaFormValues) => void;
}

const TIPS = {
  lockedDate:
    "Goes into the 1-pager subtitle (“Decisions locked YYYY-MM-DD”) and the Reference footer at the bottom of the docx. Defaults to today.",
  owner:
    "Person accountable for the locked decisions. Appears in the Reference footer of the 1-pager — e.g. “Owner: Mike”. Leave blank to show “—”.",
  targetMin:
    "Lower bound of the acceptable revenue-impact band (per finance — typically −1%). Scenarios with 3-mo %Δ below this are tagged “Outside”. Does not change the numbers in the output files; only drives the recommendation badge in the results panel.",
  targetMax:
    "Upper bound of the acceptable revenue-impact band (typically +0.5%). Scenarios within [min, max] are tagged “In band”, and the one with the smallest customer-impact drift vs Scenario A gets the “★ Recommended” star. Output files are unchanged.",
};

export function MetaForm({ values, onChange }: Props) {
  function set<K extends keyof MetaFormValues>(k: K, v: MetaFormValues[K]) {
    onChange({ ...values, [k]: v });
  }
  return (
    <div className="grid grid-cols-1 gap-3 md:grid-cols-2">
      <label className="text-sm">
        <div className="mb-1 flex items-center text-xs font-semibold text-zinc-700">
          <span>Locked date</span>
          <InfoIcon text={TIPS.lockedDate} />
        </div>
        <input
          type="date"
          value={values.lockedDate}
          onChange={(e) => set("lockedDate", e.target.value)}
          className="w-full rounded border border-zinc-300 px-2 py-1.5 text-sm"
        />
      </label>
      <label className="text-sm">
        <div className="mb-1 flex items-center text-xs font-semibold text-zinc-700">
          <span>Owner</span>
          <InfoIcon text={TIPS.owner} />
        </div>
        <input
          type="text"
          value={values.owner}
          onChange={(e) => set("owner", e.target.value)}
          placeholder="e.g. Mike"
          className="w-full rounded border border-zinc-300 px-2 py-1.5 text-sm"
        />
      </label>
      <label className="text-sm">
        <div className="mb-1 flex items-center text-xs font-semibold text-zinc-700">
          <span>Target band — min %</span>
          <InfoIcon text={TIPS.targetMin} />
        </div>
        <input
          type="number"
          step="0.1"
          value={values.targetMinPct * 100}
          onChange={(e) => set("targetMinPct", Number(e.target.value) / 100)}
          className="w-full rounded border border-zinc-300 px-2 py-1.5 text-sm"
        />
      </label>
      <label className="text-sm">
        <div className="mb-1 flex items-center text-xs font-semibold text-zinc-700">
          <span>Target band — max %</span>
          <InfoIcon text={TIPS.targetMax} />
        </div>
        <input
          type="number"
          step="0.1"
          value={values.targetMaxPct * 100}
          onChange={(e) => set("targetMaxPct", Number(e.target.value) / 100)}
          className="w-full rounded border border-zinc-300 px-2 py-1.5 text-sm"
        />
      </label>
    </div>
  );
}
