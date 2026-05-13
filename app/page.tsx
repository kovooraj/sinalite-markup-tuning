"use client";

import { useMemo, useState } from "react";
import { FileDrop } from "@/components/FileDrop";
import { MetaForm, MetaFormValues } from "@/components/MetaForm";
import { ResultsPanel } from "@/components/ResultsPanel";
import { parsePriceEngine, PriceEngineData } from "@/lib/parsePriceEngine";
import { parseOrderReplay, OrderReplayData } from "@/lib/parseOrderReplay";
import { computeAllScenarios, ScenariosOutput } from "@/lib/computeScenarios";
import { runSelfCheck, SelfCheckResult } from "@/lib/selfCheck";
import { recommend, Recommendation } from "@/lib/recommend";
import {
  computeBaseCatalogImpact,
  computeFinishingCatalogImpact,
  computeNbdLift,
} from "@/lib/catalogImpact";
import { computeLossLeaders } from "@/lib/lossLeaders";
import { downloadScenariosXlsx } from "@/lib/buildScenariosXlsx";
import { downloadOnePagerDocx } from "@/lib/buildOnePagerDocx";

interface ComputedState {
  pe: PriceEngineData;
  order: OrderReplayData;
  scenarios: ScenariosOutput;
  selfCheck: SelfCheckResult;
  recommendation: Recommendation;
  sanity: string[];
  lossLeaders: ReturnType<typeof computeLossLeaders>;
  baseImpact: ReturnType<typeof computeBaseCatalogImpact>;
  finImpact: ReturnType<typeof computeFinishingCatalogImpact>;
  nbdLift: ReturnType<typeof computeNbdLift>;
}

function todayIsoDate(): string {
  return new Date().toISOString().slice(0, 10);
}

function computeSanity(
  pe: PriceEngineData,
  order: OrderReplayData,
  scenarios: ScenariosOutput
): string[] {
  const warnings: string[] = [];
  const cm = pe.rows
    .map((r) => r.consolidatedMarkup)
    .filter((v) => v > 0 && isFinite(v))
    .sort((a, b) => a - b);
  if (cm.length > 0) {
    const median = cm[Math.floor(cm.length / 2)];
    if (median < 0.7 || median > 1.5) {
      warnings.push(
        `Median realized markup (Consolidated Markup): ${median.toFixed(2)}× — outside the pilot's [0.7, 1.5] sanity band.`
      );
    }
  }
  const belowCost = pe.rows.filter(
    (r) =>
      r.currentSalePrice > 0 &&
      r.pe3CostTotal > 0 &&
      r.currentSalePrice < r.pe3CostTotal
  ).length;
  const belowPct = pe.rows.length > 0 ? belowCost / pe.rows.length : 0;
  if (belowPct > 0.1) {
    warnings.push(
      `${(belowPct * 100).toFixed(1)}% of catalog rows priced below cost — exceeds pilot baseline (~5.5%).`
    );
  }
  const pureBaseOrders = order.rows
    .filter((r) => r.bundling === "No bundling - FREE" && r.scoring === "None")
    .reduce((sum, r) => sum + r.orders, 0);
  const totalOrders = order.totals.orders;
  const pureBaseShare = totalOrders > 0 ? pureBaseOrders / totalOrders : 0;
  if (pureBaseShare < 0.2 || pureBaseShare > 0.7) {
    warnings.push(
      `Pure-base order share: ${(pureBaseShare * 100).toFixed(1)}% — outside the pilot's [20%, 70%] band.`
    );
  }
  const a = scenarios.scenarios.find((s) => s.id === "A_Current_Locked");
  if (a && Math.abs(a.pctDelta) > 0.03) {
    warnings.push(
      `Scenario A revenue impact: ${(a.pctDelta * 100).toFixed(2)}% — outside the pilot ±3% comfort band.`
    );
  }
  if (a && a.dist.noChange < 0.3) {
    warnings.push(
      `Customer "no change" share: ${(a.dist.noChange * 100).toFixed(1)}% — cap rule may not be applying correctly.`
    );
  }
  return warnings;
}

export default function Home() {
  const [priceFile, setPriceFile] = useState<File | null>(null);
  const [orderFile, setOrderFile] = useState<File | null>(null);
  const [priceErr, setPriceErr] = useState<string>("");
  const [orderErr, setOrderErr] = useState<string>("");
  const [generating, setGenerating] = useState(false);
  const [computed, setComputed] = useState<ComputedState | null>(null);
  const [meta, setMeta] = useState<MetaFormValues>({
    lockedDate: todayIsoDate(),
    owner: "",
    archivePath: "",
    targetMinPct: -0.01,
    targetMaxPct: 0.005,
  });

  const canGenerate = !!priceFile && !!orderFile && !generating;

  async function handleGenerate() {
    if (!priceFile || !orderFile) return;
    setGenerating(true);
    setPriceErr("");
    setOrderErr("");
    setComputed(null);
    try {
      let pe: PriceEngineData;
      try {
        pe = await parsePriceEngine(priceFile);
      } catch (err) {
        setPriceErr((err as Error).message);
        throw err;
      }
      let order: OrderReplayData;
      try {
        order = await parseOrderReplay(orderFile);
      } catch (err) {
        setOrderErr((err as Error).message);
        throw err;
      }
      const scenarios = computeAllScenarios(order, pe);
      if (scenarios.matchedOrderRows === 0) {
        setOrderErr(
          "Zero SKU rows from the order replay matched the price engine. Likely a different product or coating."
        );
        throw new Error("zero matches");
      }
      const selfCheck = runSelfCheck(order, pe);
      const recommendation = recommend(scenarios.scenarios, {
        targetMinPct: meta.targetMinPct,
        targetMaxPct: meta.targetMaxPct,
      });
      const sanity = computeSanity(pe, order, scenarios);
      const baseImpact = computeBaseCatalogImpact(pe);
      const finImpact = computeFinishingCatalogImpact(pe);
      const a = scenarios.scenarios.find((s) => s.id === "A_Current_Locked")!;
      const nbdLift = computeNbdLift(a);
      const lossLeaders = computeLossLeaders(order, pe);

      setComputed({
        pe,
        order,
        scenarios,
        selfCheck,
        recommendation,
        sanity,
        baseImpact,
        finImpact,
        nbdLift,
        lossLeaders,
      });
    } catch {
      // errors surfaced inline above
    } finally {
      setGenerating(false);
    }
  }

  async function handleDownloadScenarios() {
    if (!computed) return;
    setGenerating(true);
    try {
      await downloadScenariosXlsx({
        productName: computed.pe.productName,
        productSlug: computed.pe.productSlug,
        results: computed.scenarios.scenarios,
      });
    } finally {
      setGenerating(false);
    }
  }

  async function handleDownloadOnePager() {
    if (!computed) return;
    setGenerating(true);
    try {
      await downloadOnePagerDocx({
        productName: computed.pe.productName,
        productSlug: computed.pe.productSlug,
        lockedDate: meta.lockedDate || todayIsoDate(),
        owner: meta.owner,
        archivePath: meta.archivePath,
        baseImpact: computed.baseImpact,
        finImpact: computed.finImpact,
        nbdLift: computed.nbdLift,
        lossLeaders: computed.lossLeaders,
      });
    } finally {
      setGenerating(false);
    }
  }

  const pricePreview = useMemo(() => {
    if (!computed || !priceFile) return undefined;
    return `${computed.pe.rows.length.toLocaleString()} rows · ${computed.pe.productName}`;
  }, [computed, priceFile]);

  const orderPreview = useMemo(() => {
    if (!computed || !orderFile) return undefined;
    return `${computed.order.rows.length.toLocaleString()} SKU rows · ${computed.order.totals.orders.toLocaleString()} orders`;
  }, [computed, orderFile]);

  return (
    <div className="mx-auto max-w-4xl px-4 py-10 sm:py-16">
      <header className="mb-8">
        <h1 className="text-2xl font-bold tracking-tight text-zinc-900">
          SinaLite Markup Tuning
        </h1>
        <p className="mt-1 text-sm text-zinc-600">
          Upload a price-engine export and a 3-month order replay → get the locked-markup
          1-pager and tuning scenarios. All processing happens in your browser; files
          never leave your machine.
        </p>
      </header>

      <section className="space-y-5">
        <div className="grid grid-cols-1 gap-4 md:grid-cols-2">
          <FileDrop
            label="Price Engine .xlsx"
            hint="Sample SKU catalog with PE3 cost decomposition"
            file={priceFile}
            onSelect={(f) => {
              setPriceFile(f);
              setPriceErr("");
            }}
            preview={pricePreview}
            error={priceErr}
          />
          <FileDrop
            label="3-Month Order Replay .xlsx"
            hint={`Must contain "Per-SKU Detail" sheet from the replay script`}
            file={orderFile}
            onSelect={(f) => {
              setOrderFile(f);
              setOrderErr("");
            }}
            preview={orderPreview}
            error={orderErr}
          />
        </div>

        <div className="rounded-lg border border-zinc-200 bg-white p-4">
          <MetaForm values={meta} onChange={setMeta} />
        </div>

        <div className="flex items-center gap-3">
          <button
            type="button"
            onClick={handleGenerate}
            disabled={!canGenerate}
            className="rounded bg-blue-600 px-5 py-2.5 text-sm font-medium text-white hover:bg-blue-700 disabled:opacity-50"
          >
            {generating ? "Generating…" : "Generate Reports"}
          </button>
          {!canGenerate && !generating && (
            <span className="text-xs text-zinc-500">
              Upload both files to enable.
            </span>
          )}
        </div>

        {computed && (
          <ResultsPanel
            productName={computed.pe.productName}
            orderRows={computed.order.rows.length}
            totalOrders={computed.order.totals.orders}
            unmatched={computed.order.unmatched.count}
            selfCheck={computed.selfCheck}
            scenarios={computed.scenarios}
            recommendation={computed.recommendation}
            sanity={computed.sanity}
            generating={generating}
            onDownloadOnePager={handleDownloadOnePager}
            onDownloadScenarios={handleDownloadScenarios}
          />
        )}
      </section>

      <footer className="mt-12 border-t border-zinc-200 pt-4 text-xs text-zinc-500">
        Locked methodology: Base + Add-On pricing, validated May 2026. Cap rule is
        non-negotiable — no customer ever sees a list-price increase.
      </footer>
    </div>
  );
}
