import { applyActivities, emptyActivityConfig, type ActivityConfig } from "./activity";
import { parseCsvText } from "./csv";
import { buildDashboard } from "./dashboard";
import { applyLettrage } from "./lettrage";
import { normalizeRows } from "./normalize";
import { detectProfile, getProfile, type Profile } from "./profiles";
import type { PipelineResult, Transaction } from "./types";

export interface PipelineOptions {
  /** "auto" (defaut) ou l'id d'un profil precis. */
  profileId?: string;
  /** Force un separateur de colonnes, sinon auto-detecte. */
  delimiter?: string;
  /** Taux de TVA pour l'estimation de la synthese (0.20 par defaut). */
  vatRate?: number;
  /** Regles d'affectation des operations aux activites (sinon : tout reste a affecter). */
  activity?: ActivityConfig;
}

function chooseProfile(headers: string[], profileId?: string): Profile {
  if (profileId && profileId !== "auto") {
    const explicit = getProfile(profileId);
    if (explicit) return explicit;
  }
  return detectProfile(headers).profile;
}

function compareForOutput(a: Transaction, b: Transaction): number {
  const ta = a.date ? a.date.getTime() : Number.MAX_SAFE_INTEGER;
  const tb = b.date ? b.date.getTime() : Number.MAX_SAFE_INTEGER;
  if (ta !== tb) return ta - tb;
  return a.index - b.index;
}

/** Lance toute la chaine de traitement a partir du texte CSV. */
export function runPipeline(text: string, options: PipelineOptions = {}): PipelineResult {
  const vatRate = options.vatRate ?? 0.2;
  const { headers, rows } = parseCsvText(text, options.delimiter);
  const profile = chooseProfile(headers, options.profileId);

  const { transactions, mapping, missingFields } = normalizeRows(rows, headers, profile);

  // Le lettrage s'appuie sur les positions d'origine : on l'applique avant
  // de trier pour l'affichage.
  const lettrage = applyLettrage(transactions);
  // Les groupes d'operations liees (remboursement d'un paiement...) servent a propager l'activite.
  const activityStats = applyActivities(transactions, options.activity ?? emptyActivityConfig());
  const ordered = [...transactions].sort(compareForOutput);

  const dashboard = buildDashboard(ordered);

  const currencies = [...new Set(ordered.map((t) => t.currency).filter(Boolean))].sort();
  const dated = ordered.map((t) => t.date).filter((d): d is Date => d != null);
  const periodStart = dated.length ? new Date(Math.min(...dated.map((d) => d.getTime()))) : null;
  const periodEnd = dated.length ? new Date(Math.max(...dated.map((d) => d.getTime()))) : null;

  return {
    transactions: ordered,
    mapping,
    missingFields,
    dashboard,
    activityStats,
    stats: {
      rowCount: rows.length,
      txCount: ordered.length,
      groupCount: lettrage.groupCount,
      lettredCount: lettrage.lettredCount,
      currencies,
      periodStart,
      periodEnd,
    },
    profileId: profile.id,
    profileLabel: profile.label,
    vatRate,
  };
}
