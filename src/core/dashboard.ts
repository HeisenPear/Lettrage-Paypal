import { impactsBalance, isRefund, isRetrait, isSale } from "./classify";
import type { DashboardStats, Transaction } from "./types";

/**
 * Agrege les operations pour les indicateurs de l'application. Le CA, les
 * commissions et le net ne comptent que les ventes reellement encaissees
 * (ni remboursements, ni retraits, ni operations annulees ou en attente).
 */
export function buildDashboard(transactions: Transaction[]): DashboardStats {
  const d: DashboardStats = {
    nbPaiements: 0,
    caBrut: 0,
    commissions: 0,
    netEncaisse: 0,
    nbRemboursements: 0,
    totalRemboursements: 0,
    nbRetraits: 0,
    totalRetraits: 0,
    nbAttente: 0,
    totalAttente: 0,
    nbAutres: 0,
  };

  for (const t of transactions) {
    if (isSale(t)) {
      d.nbPaiements += 1;
      d.caBrut += t.gross;
      d.commissions += t.fee;
      d.netEncaisse += t.net;
    } else if (isRefund(t)) {
      d.nbRemboursements += 1;
      d.totalRemboursements += t.net || t.gross;
    } else if (isRetrait(t) && impactsBalance(t)) {
      d.nbRetraits += 1;
      d.totalRetraits += t.net || t.gross;
    } else if (t.pending) {
      d.nbAttente += 1;
      d.totalAttente += t.gross;
    } else {
      d.nbAutres += 1;
    }
  }

  const r2 = (n: number): number => Math.round(n * 100) / 100;
  d.caBrut = r2(d.caBrut);
  d.commissions = r2(d.commissions);
  d.netEncaisse = r2(d.netEncaisse);
  d.totalRemboursements = r2(d.totalRemboursements);
  d.totalRetraits = r2(d.totalRetraits);
  d.totalAttente = r2(d.totalAttente);
  return d;
}
