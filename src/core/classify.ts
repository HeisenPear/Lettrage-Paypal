import { norm } from "./profiles/profile";
import type { Transaction } from "./types";

// Definitions uniques de ce qui compte (ou non) comme vente, remboursement...
// Le classeur Excel et les indicateurs de l'application s'appuient sur ces
// memes fonctions, pour que les chiffres ne divergent jamais.

const FAILED = /refus|echou|annul|supprim|denied|failed|expir|rejet|reverse|echec/;

/** Operations de gestion de compte : jamais des ventes, meme si leur montant est positif. */
const NOT_A_SALE = /suspension|deblocage|annulation|retrait|remboursement|litige/;

/** Vrai si l'etat indique une operation echouee / annulee / rejetee. */
export function isFailed(status: string): boolean {
  return FAILED.test(norm(status));
}

export const isRetrait = (t: Transaction): boolean => norm(t.type).includes("retrait");

export const isRemboursement = (t: Transaction): boolean => norm(t.type).includes("remboursement");

export const isCredit = (t: Transaction): boolean =>
  norm(t.impact).startsWith("cred") || (!t.impact && t.net >= 0);

/**
 * Les lignes "Memo" (paiement annule, en attente...) ne modifient pas le solde
 * PayPal : la somme des Net hors Memo retombe exactement sur le solde affiche.
 */
export const impactsBalance = (t: Transaction): boolean => norm(t.impact) !== "memo";

/** Operation terminee : ni en attente, ni annulee ou rejetee. */
export const isSettled = (t: Transaction): boolean => !t.pending && !isFailed(t.status);

/** Vente comptabilisable dans le chiffre d'affaires. */
export function isSale(t: Transaction): boolean {
  if (NOT_A_SALE.test(norm(t.type))) return false;
  return t.net > 0 && isSettled(t) && impactsBalance(t);
}

/** Remboursement comptabilisable (vient en diminution du chiffre d'affaires). */
export const isRefund = (t: Transaction): boolean =>
  isRemboursement(t) && isSettled(t) && impactsBalance(t);
