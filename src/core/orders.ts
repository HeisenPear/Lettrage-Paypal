import Papa from "papaparse";
import { keyVariants, type OrderSource } from "./activity";

/** Un export de commandes (CSV) : entetes + lignes de valeurs brutes. */
export interface OrderFile {
  name: string;
  headers: string[];
  rows: string[][];
}

export function parseOrderFile(text: string, name: string): OrderFile {
  const parsed = Papa.parse<string[]>(text, { skipEmptyLines: "greedy" });
  const data = parsed.data.map((row) => row.map((cell) => String(cell ?? "").trim()));
  const [headers = [], ...rows] = data;
  return { name, headers, rows };
}

/** Nombre de lignes de l'export dont la colonne `col` correspond a l'une des cles recherchees. */
export function countMatches(file: OrderFile, col: number, wanted: Set<string>): number {
  if (col < 0) return 0;
  let n = 0;
  for (const row of file.rows) {
    const cell = row[col] ?? "";
    if (cell && keyVariants(cell).some((k) => wanted.has(k))) n += 1;
  }
  return n;
}

/**
 * Repere la colonne du numero de commande sans connaitre le format de l'export :
 * c'est celle dont les valeurs recoupent le plus de paiements PayPal.
 */
export function detectKeyColumn(
  file: OrderFile,
  wanted: Set<string>
): { index: number; matches: number } {
  let best = { index: -1, matches: 0 };
  for (let c = 0; c < file.headers.length; c++) {
    const matches = countMatches(file, c, wanted);
    if (matches > best.matches) best = { index: c, matches };
  }
  return best;
}

/** Valeurs distinctes d'une colonne (au plus `limit`, `truncated` si il y en a davantage). */
export function distinctValues(
  file: OrderFile,
  col: number,
  limit = 40
): { values: string[]; truncated: boolean } {
  const seen = new Set<string>();
  for (const row of file.rows) {
    const v = row[col] ?? "";
    if (v) seen.add(v);
    if (seen.size > limit) return { values: [...seen].slice(0, limit), truncated: true };
  }
  return { values: [...seen], truncated: false };
}

export interface OrderSettings {
  keyColumn: number;
  /** -1 : toutes les commandes du fichier ont la meme activite (`fixedActivity`). */
  activityColumn: number;
  /** Valeur de la colonne d'activite -> libelle d'activite ("" = ignorer cette valeur). */
  valueMap: Record<string, string>;
  fixedActivity: string;
}

/** Activite associee a une valeur de la colonne d'activite ("" si elle est ignoree ou inconnue). */
function mappedActivity(valueMap: Record<string, string>, value: string): string {
  return Object.prototype.hasOwnProperty.call(valueMap, value) ? valueMap[value] : "";
}

/** Transforme un export de commandes en table cle de rapprochement -> activite. */
export function buildOrderSource(file: OrderFile, s: OrderSettings): OrderSource {
  const keys = new Map<string, string>();
  for (const row of file.rows) {
    const activity =
      s.activityColumn >= 0 ? mappedActivity(s.valueMap, row[s.activityColumn] ?? "") : s.fixedActivity;
    if (!activity) continue;
    for (const k of keyVariants(row[s.keyColumn] ?? "")) if (!keys.has(k)) keys.set(k, activity);
  }
  return { name: file.name, keys };
}
