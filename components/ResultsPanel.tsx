"use client";

import { ScenariosOutput } from "@/lib/computeScenarios";
import { SelfCheckResult } from "@/lib/selfCheck";
import { Recommendation } from "@/lib/recommend";

interface Props {
  productName: string;
  orderRows: number;
  totalOrders: number;
  unmatched: number;
  selfCheck: SelfCheckResult;
  scenarios: ScenariosOutput;
  recommendation: Recommendation;
  sanity: string[];
  generating: boolean;
  onDownloadOnePager: () => void;
  onDownloadScenarios: () => void;
  onDownloadRepriced: () => void;
}

function fmtUsd(v: number): string {
  const abs = Math.abs(v);
  const sign = v >= 0 ? "+" : "−";
  return `${sign}$${abs.toLocaleString("en-US", { maximumFractionDigits: 0 })}`;
}

function fmtPct(v: number): string {
  const sign = v >= 0 ? "+" : "−";
  return `${sign}${Math.abs(v * 100).toFixed(1)}%`;
}

export function ResultsPanel(props: Props) {
  const { recommendation, scenarios, selfCheck } = props;
  return (
    <div className="rounded-lg border border-zinc-200 bg-white p-5 shadow-sm">
      <h2 className="mb-3 text-lg font-semibold">Results — {props.productName}</h2>

      <ul className="space-y-1 text-sm">
        <li>
          {selfCheck.ok ? "✅" : "⚠️"}{" "}
          <span className="text-zinc-700">{selfCheck.message}</span>
        </li>
        <li>
          📊 {props.orderRows} SKU rows · {props.totalOrders.toLocaleString()} orders ·{" "}
          {props.unmatched} unmatched (replay-side){" "}
          {scenarios.unmatchedOrderRows > 0
            ? `· ${scenarios.unmatchedOrderRows} replay rows had no price-engine match`
            : ""}
        </li>
        {recommendation.recommended && (
          <li>
            📌 Recommended: <span className="font-semibold">{recommendation.recommended.id}</span>{" "}
            ({fmtUsd(recommendation.recommended.deltaUsd)} ·{" "}
            {fmtPct(recommendation.recommended.pctDelta)}) — {recommendation.reason}
          </li>
        )}
      </ul>

      {props.sanity.length > 0 && (
        <div className="mt-3 rounded border border-amber-200 bg-amber-50 p-3 text-xs text-amber-900">
          <div className="font-semibold mb-1">Sanity warnings</div>
          <ul className="list-disc pl-4 space-y-0.5">
            {props.sanity.map((s, i) => (
              <li key={i}>{s}</li>
            ))}
          </ul>
        </div>
      )}

      <div className="mt-4 overflow-x-auto">
        <table className="min-w-full border-collapse text-xs">
          <thead>
            <tr className="bg-zinc-100 text-left">
              <th className="border border-zinc-200 px-2 py-1">Scenario</th>
              <th className="border border-zinc-200 px-2 py-1">Base</th>
              <th className="border border-zinc-200 px-2 py-1">Finishing</th>
              <th className="border border-zinc-200 px-2 py-1">NBD</th>
              <th className="border border-zinc-200 px-2 py-1 text-right">3-mo Δ</th>
              <th className="border border-zinc-200 px-2 py-1 text-right">% Δ</th>
              <th className="border border-zinc-200 px-2 py-1 text-right">Annualized</th>
              <th className="border border-zinc-200 px-2 py-1">Status</th>
            </tr>
          </thead>
          <tbody>
            {scenarios.scenarios.map((s) => {
              const inBand = recommendation.inBand.some((x) => x.id === s.id);
              const isRecommended = recommendation.recommended?.id === s.id;
              return (
                <tr
                  key={s.id}
                  className={isRecommended ? "bg-green-50" : inBand ? "bg-zinc-50" : ""}
                >
                  <td className="border border-zinc-200 px-2 py-1 font-mono">{s.id}</td>
                  <td className="border border-zinc-200 px-2 py-1">{s.baseFormatted}</td>
                  <td className="border border-zinc-200 px-2 py-1">{s.finFormatted}</td>
                  <td className="border border-zinc-200 px-2 py-1">{s.nbdFormatted}</td>
                  <td className="border border-zinc-200 px-2 py-1 text-right">
                    {fmtUsd(s.deltaUsd)}
                  </td>
                  <td className="border border-zinc-200 px-2 py-1 text-right">
                    {fmtPct(s.pctDelta)}
                  </td>
                  <td className="border border-zinc-200 px-2 py-1 text-right">
                    {fmtUsd(s.annualizedUsd)}
                  </td>
                  <td className="border border-zinc-200 px-2 py-1">
                    {isRecommended ? "★ Recommended" : inBand ? "✓ In band" : "— Outside"}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      <div className="mt-5 flex flex-wrap gap-3">
        <button
          type="button"
          onClick={props.onDownloadOnePager}
          disabled={props.generating}
          className="rounded bg-zinc-900 px-4 py-2 text-sm font-medium text-white hover:bg-zinc-800 disabled:opacity-50"
        >
          ⬇ Download 1-Pager (.docx)
        </button>
        <button
          type="button"
          onClick={props.onDownloadScenarios}
          disabled={props.generating}
          className="rounded bg-zinc-900 px-4 py-2 text-sm font-medium text-white hover:bg-zinc-800 disabled:opacity-50"
        >
          ⬇ Download Scenarios (.xlsx)
        </button>
        <button
          type="button"
          onClick={props.onDownloadRepriced}
          disabled={props.generating}
          className="rounded bg-zinc-900 px-4 py-2 text-sm font-medium text-white hover:bg-zinc-800 disabled:opacity-50"
        >
          ⬇ Download Repriced Catalog (.xlsx)
        </button>
      </div>
    </div>
  );
}
