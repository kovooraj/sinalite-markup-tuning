import { Gradient, makeSavedScenario, ScenarioDef } from "./markupEngine";

/** A user-saved scenario as persisted in localStorage. Gradients are stored
 * as fractions (0.35 = 35%). */
export interface SavedScenarioRecord {
  name: string;
  base: Gradient["base"];
  fin: Gradient["fin"];
  nbd: Gradient["nbd"];
}

const STORAGE_KEY = "sinalite-markup-tuning:saved-scenarios";

function isValidTuple(v: unknown, len: number): v is number[] {
  return (
    Array.isArray(v) &&
    v.length === len &&
    v.every((x) => typeof x === "number" && isFinite(x) && x >= 0 && x <= 5)
  );
}

/** Load saved scenarios from localStorage. Safe to call anywhere — returns
 * [] on the server, on parse errors, or when nothing is saved. */
export function loadSavedScenarios(): SavedScenarioRecord[] {
  if (typeof window === "undefined") return [];
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    if (!raw) return [];
    const parsed: unknown = JSON.parse(raw);
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

export function setSavedScenarios(list: SavedScenarioRecord[]): void {
  cache = list;
  if (typeof window !== "undefined") {
    try {
      window.localStorage.setItem(STORAGE_KEY, JSON.stringify(list));
    } catch {
      // localStorage full / blocked — saved scenarios won't survive reload
    }
  }
  for (const cb of listeners) cb();
}
