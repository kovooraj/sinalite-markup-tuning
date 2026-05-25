"use client";

import { ScenariosOutput } from "@/lib/computeScenarios";
import { SelfCheckResult } from "@/lib/selfCheck";
import { Recommendation } from "@/lib/recommend";
import { SCENARIO_BY_ID } from "@/lib/markupEngine";
import { InfoIcon } from "./InfoIcon";

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
  onDownloadRepriced: (scenarioId: string) => void;
  onDownloadAnnotatedOrders: (scenarioId: string) => void;
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
        <li className="text-xs text-zinc-500">
          🧩 Matched on {scenarios.commonDimensions.length} dimension
          {scenarios.commonDimensions.length === 1 ? "" : "s"}:{" "}
          {scenarios.commonDimensions.length > 0
            ? scenarios.commonDimensions.join(", ")
            : "stock, size, qty, turnaround only (no shared variants)"}
        </li>
      </ul>

      {recommendation.recommended && (
        <div className="mt-4 rounded border border-green-300 bg-green-50 p-3">
          <div className="flex items-baseline gap-2">
            <span className="text-sm font-semibold text-green-900">
              📌 Recommended:
            </span>
            <span className="font-mono text-base font-bold text-green-900">
              {recommendation.recommended.id}
            </span>
            <span className="text-sm text-green-900">
              Δ {fmtUsd(recommendation.recommended.deltaUsd)} ·{" "}
              {fmtPct(recommendation.recommended.pctDelta)} · annualized{" "}
              {fmtUsd(recommendation.recommended.annualizedUsd)}
            </span>
          </div>
          <p className="mt-2 text-xs text-green-900 leading-relaxed">
            <span className="font-semibold">Why: </span>
            {recommendation.reason}
          </p>
          <p className="mt-1 text-[11px] text-green-800/70 italic">
            Source: sinalite-pricing-model SOP · Recommendation Heuristic ·{" "}
            {recommendation.inBand.length} of 8 scenarios in your target band.
          </p>
        </div>
      )}

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
              <th className="border border-zinc-200 px-2 py-1">Repriced Catalog</th>
              <th className="border border-zinc-200 px-2 py-1">Annotated Orders</th>
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
                  <td className="border border-zinc-200 px-2 py-1 font-mono">
                    <span className="inline-flex items-center">
                      <span>{s.id}</span>
                      {SCENARIO_BY_ID[s.id]?.description && (
                        <InfoIcon text={SCENARIO_BY_ID[s.id].description} />
                      )}
                    </span>
                  </td>
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
                  <td className="border border-zinc-200 px-2 py-1 text-center">
                    <button
                      type="button"
                      onClick={() => props.onDownloadRepriced(s.id)}
                      disabled={props.generating}
                      className={[
                        "rounded px-2 py-1 text-xs font-medium text-white disabled:opacity-50",
                        isRecommended
                          ? "bg-green-600 hover:bg-green-700"
                          : "bg-zinc-700 hover:bg-zinc-800",
                      ].join(" ")}
                      title={`Download repriced PE3 catalog using ${s.id}`}
                    >
                      {isRecommended ? "⬇ ★ Download" : "⬇ Download"}
                    </button>
                  </td>
                  <td className="border border-zinc-200 px-2 py-1 text-center">
                    <button
                      type="button"
                      onClick={() => props.onDownloadAnnotatedOrders(s.id)}
                      disabled={props.generating}
                      className={[
                        "rounded px-2 py-1 text-xs font-medium text-white disabled:opacity-50",
                        isRecommended
                          ? "bg-green-600 hover:bg-green-700"
                          : "bg-zinc-700 hover:bg-zinc-800",
                      ].join(" ")}
                      title={`Download annotated 3-month orders with ${s.id}'s markup chain applied per order`}
                    >
                      {isRecommended ? "⬇ ★ Download" : "⬇ Download"}
                    </button>
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
          ⬇ Download 1-Pager (.pdf)
        </button>
        <button
          type="button"
          onClick={props.onDownloadScenarios}
          disabled={props.generating}
          className="rounded bg-zinc-900 px-4 py-2 text-sm font-medium text-white hover:bg-zinc-800 disabled:opacity-50"
        >
          ⬇ Download Scenarios (.xlsx)
        </button>
      </div>
      <p className="mt-2 text-xs text-zinc-500">
        Per-scenario downloads (any row above):{" "}
        <span className="font-medium">Repriced Catalog</span> — the full PE3
        catalog with new prices applied;{" "}
        <span className="font-medium">Annotated Orders</span> — your 3-month
        order data with each order&apos;s base + finishing + NBD markup chain
        shown end-to-end, totals + annualized delta at the bottom. The ★ green
        row is the system&apos;s pick.
      </p>
    </div>
  );
}
