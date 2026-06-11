import { Gradient, makeSavedScenario, ScenarioDef } from "./markupEngine";

/** A user-saved scenario as persisted locally and in Supabase. Gradients are
 * stored as fractions (0.35 = 35%). */
export interface SavedScenarioRecord {
  name: string;
  base: Gradient["base"];
  fin: Gradient["fin"];
  nbd: Gradient["nbd"];
}

const STORAGE_KEY = "sinalite-markup-tuning:saved-scenarios";

// Shared store: every visitor of the app reads/writes the same scenario
// list in the SinaLite "AI Projects" Supabase project. The publishable key
// is public by design (RLS scopes it to this one table). localStorage acts
// as an instant-load cache and offline fallback.
const SUPABASE_URL = "https://mzlnnxpnfxbjmywsxcfc.supabase.co";
const SUPABASE_KEY = "sb_publishable_J-xtl-nmayo_kZi4MttNxA_85GmQGjl";
const REST_URL = `${SUPABASE_URL}/rest/v1/markup_saved_scenarios`;
const REST_HEADERS: Record<string, string> = {
  apikey: SUPABASE_KEY,
  Authorization: `Bearer ${SUPABASE_KEY}`,
  "Content-Type": "application/json",
};

function isValidTuple(v: unknown, len: number): v is number[] {
  return (
    Array.isArray(v) &&
    v.length === len &&
    v.every((x) => typeof x === "number" && isFinite(x) && x >= 0 && x <= 5)
  );
}

function sanitize(parsed: unknown): SavedScenarioRecord[] {
  if (!Array.isArray(parsed)) return [];
  return parsed.filter(
    (r): r is SavedScenarioRecord =>
      !!r &&
      typeof (r as SavedScenarioRecord).name === "string" &&
      (r as SavedScenarioRecord).name.trim() !== "" &&
      isValidTuple((r as SavedScenarioRecord).base, 4) &&
      isValidTuple((r as SavedScenarioRecord).fin, 4) &&
      isValidTuple((r as SavedScenarioRecord).nbd, 3)
  );
}

/** Load saved scenarios from localStorage. Safe to call anywhere — returns
 * [] on the server, on parse errors, or when nothing is saved. */
export function loadSavedScenarios(): SavedScenarioRecord[] {
  if (typeof window === "undefined") return [];
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    if (!raw) return [];
    return sanitize(JSON.parse(raw));
  } catch {
    return [];
  }
}

export function savedRecordToScenario(rec: SavedScenarioRecord): ScenarioDef {
  return makeSavedScenario(rec.name, rec.base, rec.fin, rec.nbd);
}

// --- external-store wrapper (for useSyncExternalStore) -------------------
// localStorage is read once into a cached snapshot; writes update the cache,
// persist, and notify subscribers. The cached reference stays stable between
// writes, which is what useSyncExternalStore requires.

const EMPTY: SavedScenarioRecord[] = [];
let cache: SavedScenarioRecord[] | null = null;
const listeners = new Set<() => void>();

function notify() {
  for (const cb of listeners) cb();
}

function writeLocal(list: SavedScenarioRecord[]) {
  if (typeof window === "undefined") return;
  try {
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(list));
  } catch {
    // localStorage full / blocked — local cache just won't survive reload
  }
}

export function subscribeSavedScenarios(cb: () => void): () => void {
  listeners.add(cb);
  return () => listeners.delete(cb);
}

export function getSavedScenariosSnapshot(): SavedScenarioRecord[] {
  if (cache === null) cache = loadSavedScenarios();
  return cache;
}

/** Server snapshot — always empty; the client snapshot replaces it after
 * hydration. */
export function getSavedScenariosServerSnapshot(): SavedScenarioRecord[] {
  return EMPTY;
}

function setLocalState(list: SavedScenarioRecord[]): void {
  cache = list;
  writeLocal(list);
  notify();
}

// --- remote sync ----------------------------------------------------------

/** Pull the shared scenario list from Supabase and replace the local cache.
 * Silently keeps the local cache when offline or on any error. */
export async function refreshSavedScenariosFromRemote(): Promise<void> {
  try {
    const res = await fetch(`${REST_URL}?select=name,base,fin,nbd&order=name`, {
      headers: REST_HEADERS,
    });
    if (!res.ok) return;
    const rows = sanitize(await res.json());
    setLocalState(rows);
  } catch {
    // offline / blocked — keep whatever localStorage had
  }
}

function pushUpsert(rec: SavedScenarioRecord): void {
  void fetch(`${REST_URL}?on_conflict=name`, {
    method: "POST",
    headers: {
      ...REST_HEADERS,
      Prefer: "resolution=merge-duplicates,return=minimal",
    },
    body: JSON.stringify([
      {
        name: rec.name,
        base: rec.base,
        fin: rec.fin,
        nbd: rec.nbd,
        updated_at: new Date().toISOString(),
      },
    ]),
  }).catch(() => {});
}

function pushDelete(name: string): void {
  void fetch(`${REST_URL}?name=eq.${encodeURIComponent(name)}`, {
    method: "DELETE",
    headers: { ...REST_HEADERS, Prefer: "return=minimal" },
  }).catch(() => {});
}

/** Add or replace a saved scenario — updates the local cache immediately and
 * pushes to the shared Supabase table in the background. Any record whose
 * derived scenario id collides with the new one is replaced (locally and
 * remotely) so ids stay unique. */
export function upsertSavedScenario(rec: SavedScenarioRecord): void {
  const current = getSavedScenariosSnapshot();
  const newId = savedRecordToScenario(rec).id;
  const displaced = current.filter(
    (r) => savedRecordToScenario(r).id === newId && r.name !== rec.name
  );
  const next = [
    ...current.filter((r) => savedRecordToScenario(r).id !== newId),
    rec,
  ];
  setLocalState(next);
  for (const d of displaced) pushDelete(d.name);
  pushUpsert(rec);
}

/** Delete a saved scenario by name — locally and from the shared table. */
export function removeSavedScenario(name: string): void {
  const current = getSavedScenariosSnapshot();
  setLocalState(current.filter((r) => r.name !== name));
  pushDelete(name);
}
