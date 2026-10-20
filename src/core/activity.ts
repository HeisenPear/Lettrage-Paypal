import { isRetrait } from "./classify";
import { norm } from "./profiles/profile";
import type { ActivityStats, Transaction } from "./types";

// ---------------------------------------------------------------------------
// Activites
// ---------------------------------------------------------------------------
export interface Activity {
  id: string;
  label: string;
  /** Fragments (sans accents, en minuscules) qui designent l'activite dans un texte libre. */
  aliases: string[];
}

/** Activites gerees. Pour en ajouter une, ajouter une entree ici. */
export const ACTIVITIES: Activity[] = [
  { id: "jocondienne", label: "La Jocondienne", aliases: ["jocondienne"] },
  { id: "duhalle", label: "Duhallé", aliases: ["duhalle"] },
];

/** Operation dont l'activite n'a pas pu etre determinee. */
export const UNASSIGNED = "À affecter";
/** Mouvement du compte (virement vers la banque...) : commun a toutes les activites. */
export const COMMON = "Compte commun";

/** Retrouve l'activite designee par un texte libre ("duhalle-boutique.fr", "La Jocondienne"...). */
export function resolveActivity(text: string): string | null {
  const n = norm(text);
  if (!n) return null;
  const hit = ACTIVITIES.find(
    (a) => n === norm(a.label) || a.aliases.some((alias) => n.includes(alias))
  );
  return hit ? hit.label : null;
}

/** Origines possibles d'une affectation (affichees dans la colonne "Origine affectation"). */
export const SOURCE = {
  manual: "Affectation manuelle",
  rule: "Règle",
  orders: "Export commandes",
  signature: "Signature boutique",
  linked: "Opération liée",
  fallback: "Valeur par défaut",
  account: "Mouvement du compte",
} as const;

// ---------------------------------------------------------------------------
// Regles de reconnaissance (texte)
// ---------------------------------------------------------------------------
export type RuleField =
  | "article"
  | "invoiceNumber"
  | "customerRef"
  | "note"
  | "name"
  | "email"
  | "type";
export type RuleOp = "contains" | "startsWith" | "regex";

export interface TextRule {
  field: RuleField;
  op: RuleOp;
  value: string;
  /** Libelle d'activite (ou texte libre que resolveActivity sait reconnaitre). */
  activity: string;
}

export const RULE_FIELDS: Record<RuleField, string> = {
  article: "Titre de l'objet / référence",
  invoiceNumber: "N° de facture",
  customerRef: "Numéro de client (signature boutique)",
  note: "Remarque",
  name: "Nom du client",
  email: "Email du client",
  type: "Type de transaction",
};

export const RULE_OPS: Record<RuleOp, string> = {
  contains: "contient",
  startsWith: "commence par",
  regex: "correspond à (regex)",
};

function ruleMatches(rule: TextRule, t: Transaction): boolean {
  if (!rule.value) return false;
  const raw = String(t[rule.field] ?? "");
  if (rule.op === "regex") {
    try {
      return new RegExp(rule.value, "i").test(raw);
    } catch {
      return false; // regex invalide : la regle est ignoree
    }
  }
  const haystack = norm(raw);
  const needle = norm(rule.value);
  return rule.op === "startsWith" ? haystack.startsWith(needle) : haystack.includes(needle);
}

function describeRule(rule: TextRule): string {
  return `${RULE_FIELDS[rule.field]} ${RULE_OPS[rule.op]} « ${rule.value} »`;
}

// ---------------------------------------------------------------------------
// Cles de rapprochement (n° de commande, n° de transaction PayPal)
// ---------------------------------------------------------------------------
/** Formes sous lesquelles une valeur peut etre comparee ("CMD-50748526" -> "50748526"). */
export function keyVariants(raw: string): string[] {
  const v = (raw ?? "").trim().toUpperCase();
  if (v.length < 4) return [];
  const digits = v.replace(/\D/g, "");
  return digits.length >= 6 && digits !== v ? [v, digits] : [v];
}

/**
 * Cles sous lesquelles une operation peut etre retrouvee dans un fichier externe,
 * de la plus precise a la plus generale : l'operation elle-meme, sa commande,
 * puis le paiement d'origine. La premiere cle connue l'emporte.
 */
export function transactionKeys(t: Transaction): string[] {
  const keys = new Set<string>();
  const add = (raw: string): void => keyVariants(raw).forEach((k) => keys.add(k));
  add(t.transactionId);
  add(t.invoiceNumber);
  const order = t.article.match(/^\d+ (\d+) \d+$/); // reference "<boutique> <commande> <suffixe>"
  if (order) add(order[1]);
  add(t.referenceId);
  return [...keys];
}

/** Ensemble des cles des operations (hors mouvements du compte) : sert a reperer la bonne colonne d'un export. */
export function transactionKeySet(txs: Transaction[]): Set<string> {
  const set = new Set<string>();
  for (const t of txs) if (!isRetrait(t)) transactionKeys(t).forEach((k) => set.add(k));
  return set;
}

/**
 * Lit des affectations manuelles : une par ligne, "<n° de commande ou de transaction> ; <activité>".
 * Les lignes vides et celles qui commencent par # sont ignorees.
 */
export function parseManual(text: string): {
  map: Map<string, string>;
  accepted: number;
  ignored: number;
} {
  const map = new Map<string, string>();
  let accepted = 0;
  let ignored = 0;
  for (const line of text.split(/\r?\n/)) {
    const l = line.trim();
    if (!l || l.startsWith("#")) continue;
    const m = l.match(/^([^;\t=,]+)[;\t=,]\s*(.+)$/);
    const activity = m ? resolveActivity(m[2]) : null;
    const keys = m ? keyVariants(m[1]) : [];
    if (!m || !activity || keys.length === 0) {
      ignored += 1;
      continue;
    }
    keys.forEach((k) => map.set(k, activity));
    accepted += 1;
  }
  return { map, accepted, ignored };
}

// ---------------------------------------------------------------------------
// Configuration et moteur
// ---------------------------------------------------------------------------
/** Commandes d'une boutique (issues d'un export) avec l'activite de chacune. */
export interface OrderSource {
  name: string;
  keys: Map<string, string>; // cle de rapprochement -> libelle d'activite
}

export interface ActivityConfig {
  rules: TextRule[];
  manual: Map<string, string>;
  sources: OrderSource[];
  /** Activite donnee aux operations non identifiees (UNASSIGNED = les laisser a affecter). */
  fallback: string;
}

export function emptyActivityConfig(): ActivityConfig {
  return { rules: [], manual: new Map(), sources: [], fallback: UNASSIGNED };
}

/**
 * Signature laissee par certaines boutiques (module PayPal de PrestaShop) dans
 * "Numero de client" : "..._Cart ID: 27572_Shop name: La Jocondienne".
 */
export function signatureActivity(t: Transaction): string | null {
  const m = t.customerRef.match(/shop name:\s*(.+)$/i);
  return m ? resolveActivity(m[1]) : null;
}

function lookup(keys: string[], map: Map<string, string>): string | undefined {
  for (const k of keys) {
    const hit = map.get(k);
    if (hit) return hit;
  }
  return undefined;
}

/**
 * Affecte chaque operation a une activite. Ordre de priorite :
 * 1. affectation manuelle, 2. regle de reconnaissance, 3. export de commandes,
 * 4. signature de la boutique, 5. operation liee deja affectee (remboursement,
 * suspension... d'un meme paiement), 6. valeur par defaut.
 * Les mouvements du compte (virements vers la banque) restent communs.
 * A appeler une fois le lettrage applique (il fournit les groupes d'operations liees).
 */
export function applyActivities(txs: Transaction[], cfg: ActivityConfig): ActivityStats {
  const assign = (t: Transaction, activity: string, source: string): void => {
    t.activity = activity;
    t.activitySource = source;
  };

  for (const t of txs) {
    t.activity = "";
    t.activitySource = "";

    if (isRetrait(t) || t.category === "RECHARGE_BANQUE") {
      assign(t, COMMON, SOURCE.account);
      continue;
    }

    const keys = transactionKeys(t);

    const manual = lookup(keys, cfg.manual);
    if (manual) {
      assign(t, manual, SOURCE.manual);
      continue;
    }

    const rule = cfg.rules.find((r) => ruleMatches(r, t));
    if (rule) {
      assign(t, resolveActivity(rule.activity) ?? rule.activity, `${SOURCE.rule} · ${describeRule(rule)}`);
      continue;
    }

    for (const source of cfg.sources) {
      const hit = lookup(keys, source.keys);
      if (hit) {
        assign(t, hit, `${SOURCE.orders} · ${source.name}`);
        break;
      }
    }
    if (t.activity) continue;

    const signature = signatureActivity(t);
    if (signature) assign(t, signature, SOURCE.signature);
  }

  // Une operation liee (remboursement, suspension...) suit l'activite de son paiement.
  const groups = new Map<number, Transaction[]>();
  for (const t of txs) {
    if (t.group <= 0) continue;
    const members = groups.get(t.group);
    if (members) members.push(t);
    else groups.set(t.group, [t]);
  }
  let conflicts = 0;
  for (const members of groups.values()) {
    const known = new Set(members.map((m) => m.activity).filter((a) => a && a !== COMMON));
    if (known.size > 1) conflicts += 1;
    if (known.size !== 1) continue;
    const [only] = known;
    for (const m of members) if (!m.activity) assign(m, only, SOURCE.linked);
  }

  for (const t of txs) {
    if (t.activity) continue;
    if (cfg.fallback !== UNASSIGNED) assign(t, cfg.fallback, SOURCE.fallback);
    else assign(t, UNASSIGNED, "");
  }

  const byActivity: Record<string, number> = {};
  const bySource: Record<string, number> = {};
  for (const t of txs) {
    byActivity[t.activity] = (byActivity[t.activity] ?? 0) + 1;
    if (t.activitySource) {
      const key = t.activitySource.split(" · ")[0];
      bySource[key] = (bySource[key] ?? 0) + 1;
    }
  }
  return { byActivity, bySource, conflicts };
}
