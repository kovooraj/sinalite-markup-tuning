"use client";

export interface MetaFormValues {
  lockedDate: string;
  owner: string;
  archivePath: string;
  targetMinPct: number;
  targetMaxPct: number;
}

interface Props {
  values: MetaFormValues;
  onChange: (next: MetaFormValues) => void;
}

export function MetaForm({ values, onChange }: Props) {
  function set<K extends keyof MetaFormValues>(k: K, v: MetaFormValues[K]) {
    onChange({ ...values, [k]: v });
  }
  return (
    <div className="grid grid-cols-1 gap-3 md:grid-cols-2">
      <label className="text-sm">
        <div className="mb-1 text-xs font-semibold text-zinc-700">Locked date</div>
        <input
          type="date"
          value={values.lockedDate}
          onChange={(e) => set("lockedDate", e.target.value)}
          className="w-full rounded border border-zinc-300 px-2 py-1.5 text-sm"
        />
      </label>
      <label className="text-sm">
        <div className="mb-1 text-xs font-semibold text-zinc-700">Owner</div>
        <input
          type="text"
          value={values.owner}
          onChange={(e) => set("owner", e.target.value)}
          placeholder="e.g. Mike"
          className="w-full rounded border border-zinc-300 px-2 py-1.5 text-sm"
        />
      </label>
      <label className="text-sm md:col-span-2">
        <div className="mb-1 text-xs font-semibold text-zinc-700">
          Archive path (optional)
        </div>
        <input
          type="text"
          value={values.archivePath}
          onChange={(e) => set("archivePath", e.target.value)}
          placeholder="\Costing\Postcards_14PT_AQ_Pricing\"
          className="w-full rounded border border-zinc-300 px-2 py-1.5 text-sm"
        />
      </label>
      <label className="text-sm">
        <div className="mb-1 text-xs font-semibold text-zinc-700">
          Target band — min %
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
        <div className="mb-1 text-xs font-semibold text-zinc-700">
          Target band — max %
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
