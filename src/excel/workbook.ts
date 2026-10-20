import ExcelJS from "exceljs";
import { ACTIVITIES, UNASSIGNED } from "../core/activity";
import {
  impactsBalance,
  isCredit,
  isRefund,
  isRemboursement,
  isRetrait,
  isSale,
} from "../core/classify";
import { norm } from "../core/profiles/profile";
import type { PipelineResult, Transaction } from "../core/types";
import { C, FONT, MONEY_FMT, MONEY_TOTAL, colLetter, fillForType, setFill } from "./styles";

// ---------------------------------------------------------------------------
// Helpers de mise en forme (police Arial partout, comme le gabarit)
// ---------------------------------------------------------------------------
interface CellOpts {
  fill?: string;
  bold?: boolean;
  italic?: boolean;
  size?: number;
  color?: string;
  align?: "left" | "center" | "right";
  wrap?: boolean;
  numFmt?: string;
}

function put(
  ws: ExcelJS.Worksheet,
  r: number,
  c: number,
  value: ExcelJS.CellValue,
  o: CellOpts = {}
): ExcelJS.Cell {
  const cell = ws.getCell(r, c);
  cell.value = value;
  cell.font = {
    name: FONT,
    size: o.size ?? 9,
    bold: !!o.bold,
    italic: !!o.italic,
    color: { argb: o.color ?? C.ink },
  };
  cell.alignment = { vertical: "middle", horizontal: o.align ?? "left", wrapText: !!o.wrap };
  if (o.fill) setFill(cell, o.fill);
  if (o.numFmt) cell.numFmt = o.numFmt;
  return cell;
}

/** Bandeau / section fusionne sur toute la largeur. */
function band(
  ws: ExcelJS.Worksheet,
  r: number,
  lastCol: number,
  text: string,
  o: CellOpts
): void {
  ws.mergeCells(r, 1, r, lastCol);
  put(ws, r, 1, text, o);
}

/** Ligne d'entete de colonnes (fond plein, blanc gras, centre). */
function headerRow(
  ws: ExcelJS.Worksheet,
  r: number,
  labels: string[],
  fill: string,
  height = 28,
  size = 10
): void {
  ws.getRow(r).height = height;
  labels.forEach((label, i) =>
    put(ws, r, i + 1, label, { fill, bold: true, color: C.white, align: "center", wrap: true, size })
  );
}

const eur = (n: number): string =>
  `${n.toLocaleString("fr-FR", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} €`;

/** Arrondi au centime : les totaux affiches s'additionnent exactement. */
const r2 = (n: number): number => Math.round(n * 100) / 100;

const isPayLater = (t: Transaction): boolean => norm(t.source).includes("pay later");

/** HT/TVA d'une ligne : ventes a partir du brut TTC, remboursements a partir du brut rembourse. */
function htTvaForLine(t: Transaction, rate: number): { ht: number; tva: number } | null {
  let ttc: number | null = null;
  if (isSale(t)) ttc = t.gross;
  else if (isRefund(t)) ttc = t.gross;
  if (ttc === null) return null;
  const ht = r2(ttc / (1 + rate));
  return { ht, tva: r2(ttc - ht) };
}

function timeMs(time: string): number {
  const m = (time || "").match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);
  if (!m) return 0;
  return Number(m[1]) * 3600000 + Number(m[2]) * 60000 + Number(m[3] || 0) * 1000;
}
/** Horodatage complet (date + heure) pour trier/decouper finement. */
function stamp(t: Transaction): number {
  return (t.date ? t.date.getTime() : 0) + timeMs(t.time);
}

function sortChrono(txs: Transaction[]): Transaction[] {
  return [...txs].sort((a, b) => stamp(a) - stamp(b) || a.index - b.index);
}

// ---------------------------------------------------------------------------
// 1. Transactions PayPal
// ---------------------------------------------------------------------------
const TX_HEADERS = [
  "Date", "Heure", "Nom du client", "Type de transaction", "État",
  "Montant brut (€)", "Commission (€)", "Net reçu (€)", "Impact", "Solde (€)",
  "N° Facture", "Référence commande", "N° Transaction PayPal", "N° Transaction origine",
  "Email client", "Source paiement", "Pays", "Lettrage",
];
const TX_WIDTHS = [12, 10, 28, 32, 12, 16, 15, 15, 10, 15, 20, 26, 22, 22, 34, 14, 7, 20];

function buildTransactions(wb: ExcelJS.Workbook, txs: Transaction[], rate: number): void {
  const ws = wb.addWorksheet("Transactions PayPal", { views: [{ state: "frozen", ySplit: 1 }] });
  const pct = Number((rate * 100).toFixed((rate * 100) % 1 ? 1 : 0));
  const headers = [
    ...TX_HEADERS,
    "Montant HT (€)",
    `TVA ${pct} % (€)`,
    "Activité",
    "Origine affectation",
  ];
  const widths = [...TX_WIDTHS, 15, 14, 18, 34];
  widths.forEach((w, i) => (ws.getColumn(i + 1).width = w));
  headerRow(ws, 1, headers, C.navy, 36, 10);

  let r = 2;
  for (const t of txs) {
    const fill = fillForType(t.type);
    const money = { align: "right" as const, numFmt: MONEY_FMT };
    put(ws, r, 1, t.dateRaw, { fill });
    put(ws, r, 2, t.time, { fill });
    put(ws, r, 3, t.name, { fill });
    put(ws, r, 4, t.type, { fill });
    put(ws, r, 5, t.status, { fill });
    put(ws, r, 6, t.gross, { fill, ...money });
    put(ws, r, 7, t.fee, { fill, ...money });
    put(ws, r, 8, t.net, { fill, ...money });
    put(ws, r, 9, t.impact, { fill });
    put(ws, r, 10, t.balance ?? "", { fill, ...money });
    put(ws, r, 11, t.invoiceNumber, { fill });
    put(ws, r, 12, t.article, { fill });
    put(ws, r, 13, t.transactionId, { fill });
    put(ws, r, 14, t.referenceId, { fill });
    put(ws, r, 15, t.email, { fill });
    put(ws, r, 16, t.source, { fill });
    put(ws, r, 17, t.country, { fill, align: "center" });
    put(ws, r, 18, t.lettrage, { fill, align: "center" });
    const vt = htTvaForLine(t, rate);
    put(ws, r, 19, vt ? vt.ht : "", { fill, ...money });
    put(ws, r, 20, vt ? vt.tva : "", { fill, ...money });
    const unassigned = t.activity === UNASSIGNED;
    put(ws, r, 21, t.activity, {
      fill,
      bold: true,
      color: unassigned ? C.orange : C.ink,
    });
    put(ws, r, 22, t.activitySource, { fill, italic: true, size: 8, color: C.grayTxt });
    ws.getRow(r).height = 16;
    r += 1;
  }
  ws.autoFilter = { from: "A1", to: `${colLetter(headers.length)}1` };
}

// ---------------------------------------------------------------------------
// Decoupage par virement (utilise par Virements + Rapprochement)
// ---------------------------------------------------------------------------
interface Periode {
  retrait: Transaction | null; // null = periode en cours (non viree)
  start: string;
  end: string;
  txs: Transaction[]; // transactions hors retrait
}

function decouperParVirement(txs: Transaction[]): Periode[] {
  const retraits = txs.filter(isRetrait);
  // Les lignes "Memo" (annulees / en attente) ne bougent pas le solde : elles n'entrent pas dans un virement.
  const others = txs.filter((t) => !isRetrait(t) && impactsBalance(t));
  const bounds = retraits.map(stamp);
  const firstDate = txs.length ? txs[0].dateRaw : "";

  const groups: Transaction[][] = retraits.map(() => []);
  const enCours: Transaction[] = [];
  for (const t of others) {
    const td = stamp(t);
    const idx = bounds.findIndex((b) => td <= b);
    if (idx === -1) enCours.push(t);
    else groups[idx].push(t);
  }

  const periodes: Periode[] = retraits.map((retrait, i) => ({
    retrait,
    start: i === 0 ? firstDate : retraits[i - 1].dateRaw,
    end: retrait.dateRaw,
    txs: groups[i],
  }));
  if (enCours.length && retraits.length) {
    periodes.push({
      retrait: null,
      start: retraits[retraits.length - 1].dateRaw,
      end: enCours[enCours.length - 1].dateRaw,
      txs: enCours,
    });
  } else if (enCours.length) {
    periodes.push({ retrait: null, start: firstDate, end: enCours[enCours.length - 1].dateRaw, txs: enCours });
  }
  return periodes;
}

/** Mention ajoutee aux feuilles de virements d'un classeur d'activite. */
const activityBankNote = (activity: string): string =>
  `ℹ️  Virements globaux du compte PayPal (toutes activités confondues). Les nombres de transactions, les crédits nets et les sous-totaux ne concernent que l'activité ${activity}.`;

// ---------------------------------------------------------------------------
// 2. Virements Banque
// ---------------------------------------------------------------------------
function buildVirements(wb: ExcelJS.Workbook, periodes: Periode[], activity?: string): void {
  const ws = wb.addWorksheet("Virements Banque");
  [6, 14, 14, 18, 28, 18, 18, 22, 22, 16].forEach((w, i) => (ws.getColumn(i + 1).width = w));

  band(ws, 1, 10, "VIREMENTS PAYPAL → BANQUE  —  Feuille de lettrage", {
    fill: C.navy, bold: true, color: C.white, size: 13, align: "center", wrap: true,
  });
  ws.getRow(1).height = 34;
  headerRow(ws, 2, [
    "#", "Date virement", "Heure", "Période couverte", "Identifiant bancaire",
    "Montant viré (€)", "Solde PayPal restant (€)", "N° transactions", "N° Crédits nets", "✓ Lettré",
  ], C.blue, 28, 10);

  let r = 3;
  let num = 1;
  let totalVire = 0;
  for (const p of periodes) {
    const fill = p.retrait ? "FFE3F2FD" : "FFFFFDE7";
    const creditsNets = r2(p.txs.filter(isCredit).reduce((s, t) => s + t.net, 0));
    const montant = p.retrait ? p.retrait.net : 0;
    totalVire += montant;
    put(ws, r, 1, num, { fill, align: "center", size: 10 });
    put(ws, r, 2, p.retrait ? p.retrait.dateRaw : "En cours", { fill, size: 10 });
    put(ws, r, 3, p.retrait ? p.retrait.time : "—", { fill, size: 10 });
    put(ws, r, 4, `${p.start} → ${p.end}${p.retrait ? "" : " (non viré)"}`, { fill, size: 10 });
    put(ws, r, 5, p.retrait ? p.retrait.bankId || "—" : "—", { fill, italic: true, size: 9, color: C.grayTxt });
    put(ws, r, 6, montant, { fill, bold: true, align: "right", numFmt: MONEY_FMT, size: 10 });
    put(ws, r, 7, p.retrait ? p.retrait.balance ?? "" : lastBalance(p.txs), { fill, align: "right", numFmt: MONEY_FMT, size: 10 });
    put(ws, r, 8, p.txs.length, { fill, align: "center", size: 10 });
    put(ws, r, 9, creditsNets, { fill, align: "right", numFmt: MONEY_FMT, size: 10 });
    put(ws, r, 10, p.retrait ? "" : "En cours", { fill: C.lettre, align: "center", size: 10 });
    ws.getRow(r).height = 20;
    r += 1;
    num += 1;
  }

  r += 1; // ligne vide
  const totalRow = r;
  ws.mergeCells(totalRow, 1, totalRow, 5);
  put(ws, totalRow, 1, "TOTAL VIREMENTS REÇUS EN BANQUE", { fill: C.dark, bold: true, color: C.white, align: "right", size: 10 });
  for (let c = 2; c <= 5; c++) setFill(ws.getCell(totalRow, c), C.dark);
  put(ws, totalRow, 6, r2(totalVire), { fill: C.dark, bold: true, color: C.white, align: "right", numFmt: MONEY_TOTAL, size: 11 });
  for (let c = 7; c <= 10; c++) setFill(ws.getCell(totalRow, c), C.dark);
  ws.getRow(totalRow).height = 22;

  const noteRow = totalRow + 2;
  band(ws, noteRow, 10,
    "ℹ️  Pour lettrer : Retrouvez chaque 'Identifiant bancaire' dans votre relevé bancaire. Le montant doit correspondre exactement à 'Montant viré'. Cochez la colonne ✓ Lettré.",
    { italic: true, size: 9, color: C.grayTxt, wrap: true });
  if (activity) {
    band(ws, noteRow + 1, 10, activityBankNote(activity), { italic: true, size: 9, color: C.orange, wrap: true });
    ws.getRow(noteRow + 1).height = 26;
  }
}

function lastBalance(txs: Transaction[]): number | string {
  for (let i = txs.length - 1; i >= 0; i--) if (txs[i].balance != null) return txs[i].balance as number;
  return "";
}

// ---------------------------------------------------------------------------
// 3. Rapprochement par virement
// ---------------------------------------------------------------------------
const RAP_SUB = ["#", "Date", "Heure", "Nom client", "Type transaction", "Brut (€)", "Commission (€)", "Net (€)", "Impact", "N° Facture", "N° Transaction"];

function buildRapprochement(wb: ExcelJS.Workbook, periodes: Periode[], activity?: string): void {
  const ws = wb.addWorksheet("Rapprochement");
  [6, 13, 10, 26, 30, 13, 14, 14, 14, 22, 16].forEach((w, i) => (ws.getColumn(i + 1).width = w));
  band(ws, 1, 11, "RAPPROCHEMENT PAR VIREMENT  —  Toutes transactions groupées par période de virement", {
    fill: C.navy, bold: true, color: C.white, size: 13, align: "center", wrap: true,
  });
  ws.getRow(1).height = 34;

  let r = 2;
  let vnum = 1;
  for (const p of periodes) {
    const titre = p.retrait
      ? `VIREMENT ${vnum}  ·  ${p.start} → ${p.end}  ·  Viré : ${eur(p.retrait.net)}`
      : `SOLDE EN COURS  ·  ${p.start} → ${p.end}  ·  Non encore viré`;
    band(ws, r, 11, titre, { fill: p.retrait ? C.blue : C.orange, bold: true, color: C.white, size: 11 });
    ws.getRow(r).height = 24;
    r += 1;

    headerRow(ws, r, RAP_SUB, C.slate, 20, 10);
    r += 1;

    let i = 1;
    let sb = 0, sc = 0, sn = 0, nc = 0, nd = 0;
    for (const t of p.txs) {
      const fill = fillForType(t.type);
      const credit = isCredit(t);
      const money = { align: "right" as const, numFmt: MONEY_FMT };
      put(ws, r, 1, i, { fill, align: "center" });
      put(ws, r, 2, t.dateRaw, { fill });
      put(ws, r, 3, t.time, { fill });
      put(ws, r, 4, t.name, { fill });
      put(ws, r, 5, t.type, { fill });
      put(ws, r, 6, t.gross, { fill, ...money });
      put(ws, r, 7, t.fee, { fill, ...money });
      put(ws, r, 8, t.net, { fill, ...money });
      put(ws, r, 9, t.impact, { fill, align: "center", bold: !credit, color: credit ? C.creditTxt : C.debitTxt });
      put(ws, r, 10, t.invoiceNumber, { fill });
      put(ws, r, 11, t.transactionId, { fill });
      ws.getRow(r).height = 15;
      sb += t.gross; sc += t.fee; sn += t.net;
      if (t.net >= 0) nc += 1; else nd += 1;
      i += 1;
      r += 1;
    }

    // Sous-total periode
    ws.mergeCells(r, 1, r, 5);
    put(ws, r, 1, `Sous-total période  (${nc} crédits / ${nd} débits)`, { fill: C.subtotal, bold: true });
    for (let c = 2; c <= 5; c++) setFill(ws.getCell(r, c), C.subtotal);
    put(ws, r, 6, r2(sb), { fill: C.subtotal, bold: true, align: "right", numFmt: MONEY_FMT });
    put(ws, r, 7, r2(sc), { fill: C.subtotal, bold: true, align: "right", numFmt: MONEY_FMT });
    put(ws, r, 8, r2(sn), { fill: C.subtotal, bold: true, align: "right", numFmt: MONEY_FMT });
    for (let c = 9; c <= 11; c++) setFill(ws.getCell(r, c), C.subtotal);
    r += 1;

    // Ligne virement banque
    if (p.retrait) {
      ws.mergeCells(r, 1, r, 5);
      put(ws, r, 1, `▶ VIREMENT BANQUE  ·  ID bancaire : ${p.retrait.bankId || "—"}`, { fill: "FFE3F2FD", bold: true });
      for (let c = 2; c <= 5; c++) setFill(ws.getCell(r, c), "FFE3F2FD");
      put(ws, r, 6, "", { fill: "FFE3F2FD" });
      put(ws, r, 7, "", { fill: "FFE3F2FD" });
      put(ws, r, 8, p.retrait.net, { fill: "FFE3F2FD", bold: true, align: "right", numFmt: MONEY_FMT });
      put(ws, r, 9, "VIREMENT", { fill: "FFE3F2FD", align: "center", bold: true });
      put(ws, r, 10, p.retrait.dateRaw, { fill: "FFE3F2FD" });
      put(ws, r, 11, `Solde PayPal restant : ${(p.retrait.balance ?? 0).toFixed(2)} €`, { fill: "FFE3F2FD" });
      r += 1;
    }
    r += 1; // ligne vide
    if (p.retrait) vnum += 1;
  }
  if (activity) {
    band(ws, r, 11, activityBankNote(activity), { italic: true, size: 9, color: C.orange, wrap: true });
    ws.getRow(r).height = 26;
  }
}

// ---------------------------------------------------------------------------
// 4. Remboursements
// ---------------------------------------------------------------------------
function buildRemboursements(wb: ExcelJS.Workbook, txs: Transaction[], byId: Map<string, Transaction>): void {
  const ws = wb.addWorksheet("Remboursements");
  [5, 13, 26, 16, 22, 22, 26, 18, 18, 22, 20].forEach((w, i) => (ws.getColumn(i + 1).width = w));
  band(ws, 1, 11, "REMBOURSEMENTS  —  Liste complète avec lien vers transaction d'origine", {
    fill: C.red, bold: true, color: C.white, size: 13, align: "center", wrap: true,
  });
  ws.getRow(1).height = 34;
  headerRow(ws, 2, [
    "#", "Date", "Nom client remboursé", "Net remboursé (€)", "N° Facture",
    "N° Transaction remboursement", "N° Transaction d'origine", "Date paiement original",
    "Montant original (€)", "Email client", "✓ Lettré",
  ], C.red, 28, 10);

  const fill = "FFFFEBEE";
  const rembs = txs.filter(isRemboursement);
  let r = 3;
  let num = 1;
  let total = 0;
  for (const t of rembs) {
    const orig = t.referenceId ? byId.get(t.referenceId) : undefined;
    put(ws, r, 1, num, { fill, align: "center" });
    put(ws, r, 2, t.dateRaw, { fill });
    put(ws, r, 3, t.name, { fill });
    put(ws, r, 4, t.net, { fill, align: "right", numFmt: MONEY_FMT });
    put(ws, r, 5, t.invoiceNumber, { fill });
    put(ws, r, 6, t.transactionId, { fill, italic: true, size: 8, color: C.grayTxt });
    put(ws, r, 7, t.referenceId, { fill, italic: true, size: 8, color: C.grayTxt });
    put(ws, r, 8, orig ? orig.dateRaw : "", { fill });
    put(ws, r, 9, orig ? orig.gross : "", { fill, align: "right", numFmt: MONEY_FMT });
    put(ws, r, 10, t.email, { fill });
    put(ws, r, 11, "", { fill: C.lettre, align: "center" });
    ws.getRow(r).height = 16;
    total += t.net;
    num += 1;
    r += 1;
  }
  ws.mergeCells(r, 1, r, 3);
  put(ws, r, 1, "TOTAL REMBOURSEMENTS", { fill: C.red, bold: true, color: C.white, align: "right", size: 10 });
  put(ws, r, 4, r2(total), { fill: C.red, bold: true, color: C.white, align: "right", numFmt: MONEY_TOTAL, size: 11 });
  for (let c = 5; c <= 11; c++) setFill(ws.getCell(r, c), C.red);
  ws.getRow(r).height = 22;

  band(ws, r + 2, 11,
    "ℹ️  Chaque remboursement est lié à sa transaction d'origine via le 'N° Transaction d'origine'. Vérifiez que le remboursement a bien été déduit du virement correspondant. Cochez ✓ Lettré une fois contrôlé.",
    { italic: true, size: 9, color: C.grayTxt, wrap: true });
}

// ---------------------------------------------------------------------------
// 5. Anomalies & Attentions
// ---------------------------------------------------------------------------
const ANOM_SUB = ["Catégorie", "Date", "Nom", "Type", "Net (€)", "Impact", "N° Facture", "N° Tx référence", "Source paiement", "Pays"];

function buildAnomalies(wb: ExcelJS.Workbook, txs: Transaction[]): void {
  const ws = wb.addWorksheet("Anomalies & Attentions");
  [14, 13, 26, 30, 14, 16, 22, 22, 18, 18].forEach((w, i) => (ws.getColumn(i + 1).width = w));
  band(ws, 1, 10, "ANOMALIES & POINTS D'ATTENTION  —  Transactions à vérifier manuellement", {
    fill: C.orange, bold: true, color: C.white, size: 13, align: "center", wrap: true,
  });
  ws.getRow(1).height = 34;

  const sections: { title: string; label: string; fill: string; head: string; rows: Transaction[] }[] = [
    { title: "▌ SUSPENSION DE PAIEMENT  —  Paiements temporairement retenus par PayPal", label: "Suspension de paiement", fill: "FFFFF3E0", head: C.anomHead, rows: txs.filter((t) => norm(t.type) === "suspension de paiement") },
    { title: "▌ DÉBLOCAGE DE PAIEMENT  —  Libération des suspensions — à faire correspondre", label: "Déblocage de paiement", fill: "FFF3E5F5", head: C.anomHead, rows: txs.filter((t) => norm(t.type) === "deblocage de paiement") },
    { title: "▌ SOLDE SUSPENDU POUR ENQUÊTE SUR UN LITIGE  —  Litige ouvert — fonds bloqués", label: "Solde suspendu pour enquête sur un litige", fill: "FFFFFDE7", head: C.anomHead, rows: txs.filter((t) => norm(t.type).includes("solde suspendu")) },
    { title: "▌ ANNULATION DE LA SUSPENSION POUR RÉSOLUTION DU LITIGE  —  Litige résolu — fonds libérés", label: "Annulation de la suspension pour résolution du litige", fill: "FFF1F8E9", head: C.anomHead, rows: txs.filter((t) => norm(t.type).includes("annulation de la suspension")) },
    { title: "▌ PAY LATER (Paiement en 4X)  —  Transactions financées par PayPal Credit", label: "Pay Later 4X", fill: "FFF3E5F5", head: C.purpleHead, rows: txs.filter(isPayLater) },
    { title: "▌ PAIEMENTS ANNULÉS OU EN ATTENTE  —  Lignes « Mémo » sans impact sur le solde — non comptées dans le CA", label: "Annulé / en attente", fill: "FFEEEEEE", head: C.anomHead, rows: txs.filter((t) => !impactsBalance(t)) },
  ];

  let r = 2;
  for (const s of sections) {
    if (!s.rows.length) continue;
    band(ws, r, 10, `${s.title}  (${s.rows.length} lignes)`, { fill: s.fill, bold: true, color: C.anomHead, size: 10 });
    ws.getRow(r).height = 22;
    r += 1;
    headerRow(ws, r, ANOM_SUB, s.head, 18, 10);
    r += 1;
    for (const t of s.rows) {
      put(ws, r, 1, s.label, { fill: s.fill });
      put(ws, r, 2, t.dateRaw, { fill: s.fill });
      put(ws, r, 3, t.name, { fill: s.fill });
      put(ws, r, 4, t.type, { fill: s.fill });
      put(ws, r, 5, t.net, { fill: s.fill, bold: true, align: "right", numFmt: MONEY_FMT, color: t.net < 0 ? C.debitTxt : C.creditTxt });
      put(ws, r, 6, t.impact, { fill: s.fill, align: "center" });
      put(ws, r, 7, t.invoiceNumber, { fill: s.fill });
      put(ws, r, 8, t.referenceId || t.transactionId, { fill: s.fill });
      put(ws, r, 9, t.source, { fill: s.fill });
      put(ws, r, 10, t.country, { fill: s.fill });
      ws.getRow(r).height = 15;
      r += 1;
    }
    ws.getRow(r).height = 8; // espace
    r += 1;
  }
  if (r === 2) band(ws, 2, 10, "Aucune anomalie détectée sur la période.", { italic: true, size: 10, color: C.grayTxt });
}

// ---------------------------------------------------------------------------
// 6. Résumé
// ---------------------------------------------------------------------------
const RESUME_ORDER = [
  "Paiement Express Checkout",
  "Suspension de paiement",
  "Remboursement de paiement",
  "Retrait initié par l'utilisateur",
  "Déblocage de paiement",
  "Solde suspendu pour enquête sur un litige",
  "Annulation de la suspension pour résolution du litige",
];

function buildResume(wb: ExcelJS.Workbook, txs: Transaction[], scope: string, control: boolean): void {
  const ws = wb.addWorksheet("Résumé");
  [36, 15, 15].forEach((w, i) => (ws.getColumn(i + 1).width = w));
  headerRow(ws, 1, [`Résumé des transactions${scope}`, "Nombre", "Montant net (€)"], C.navy, 28, 10);

  // Seules les lignes qui bougent le solde sont additionnees ; les "Memo" sont signalees a part.
  const groups = new Map<string, { count: number; net: number }>();
  for (const t of txs.filter(impactsBalance)) {
    const g = groups.get(t.type) ?? { count: 0, net: 0 };
    g.count += 1;
    g.net += t.net;
    groups.set(t.type, g);
  }
  const ordered = [
    ...RESUME_ORDER.filter((k) => groups.has(k)),
    ...[...groups.keys()].filter((k) => !RESUME_ORDER.includes(k)),
  ];

  let r = 2;
  let credits = 0;
  let debits = 0;
  for (const type of ordered) {
    const g = groups.get(type)!;
    const fill = fillForType(type);
    put(ws, r, 1, type, { fill, size: 10 });
    put(ws, r, 2, g.count, { fill, align: "center", size: 10 });
    put(ws, r, 3, r2(g.net), { fill, size: 10, numFmt: MONEY_FMT });
    if (g.net >= 0) credits += g.net; else debits += g.net;
    r += 1;
  }
  credits = r2(credits);
  debits = r2(debits);
  put(ws, r, 1, "TOTAL CRÉDITS", { fill: "FFE8F5E9", bold: true, size: 10 });
  put(ws, r, 2, "", { fill: "FFE8F5E9", size: 10 });
  const creditRow = r;
  put(ws, r, 3, credits, { fill: "FFE8F5E9", bold: true, size: 10, numFmt: MONEY_TOTAL });
  r += 1;
  put(ws, r, 1, "TOTAL DÉBITS", { fill: "FFFFEBEE", bold: true, size: 10 });
  put(ws, r, 2, "", { fill: "FFFFEBEE", size: 10 });
  const debitRow = r;
  put(ws, r, 3, debits, { fill: "FFFFEBEE", bold: true, size: 10, numFmt: MONEY_FMT });
  r += 1;
  const soldeNet = r2(credits + debits);
  put(ws, r, 1, "SOLDE NET", { fill: C.navy, bold: true, color: C.white, size: 10 });
  put(ws, r, 2, "", { fill: C.navy, size: 10 });
  const soldeRow = r;
  put(ws, r, 3, { formula: `C${creditRow}+C${debitRow}`, result: soldeNet }, { fill: C.navy, bold: true, color: C.white, align: "right", size: 10, numFmt: MONEY_TOTAL });
  r += 1;

  const memo = txs.filter((t) => !impactsBalance(t));
  const grey = "FFF1F1F1";
  if (memo.length) {
    r += 1;
    put(ws, r, 1, "Lignes « Mémo » (annulées / en attente)", { fill: grey, italic: true, size: 10 });
    put(ws, r, 2, memo.length, { fill: grey, align: "center", size: 10 });
    put(ws, r, 3, r2(memo.reduce((s, t) => s + t.net, 0)), { fill: grey, size: 10, numFmt: MONEY_FMT });
    r += 1;
    put(ws, r, 1, "Hors solde et hors chiffre d'affaires (aucun impact sur le solde PayPal)", { italic: true, size: 8, color: C.grayTxt });
    r += 1;
  }

  // Controle : le solde calcule doit retomber sur le solde affiche par PayPal.
  const withBalance = txs.filter((t) => t.balance != null);
  if (control && withBalance.length) {
    const first = withBalance[0];
    const last = withBalance[withBalance.length - 1];
    const opening = r2((first.balance as number) - (impactsBalance(first) ? first.net : 0));
    const closing = r2(opening + soldeNet);
    const shown = r2(last.balance as number);
    r += 1;
    put(ws, r, 1, "Solde d'ouverture (déduit de la 1re ligne)", { size: 10 });
    put(ws, r, 3, opening, { size: 10, numFmt: MONEY_FMT });
    const openingRow = r;
    r += 1;
    put(ws, r, 1, "Solde de clôture calculé", { size: 10 });
    put(ws, r, 3, { formula: `C${soldeRow}+C${openingRow}`, result: closing }, { size: 10, numFmt: MONEY_FMT });
    const closingRow = r;
    r += 1;
    put(ws, r, 1, "Solde de clôture affiché par PayPal", { size: 10 });
    put(ws, r, 3, shown, { size: 10, numFmt: MONEY_FMT });
    const shownRow = r;
    r += 1;
    const gap = r2(closing - shown);
    put(ws, r, 1, "ÉCART (doit être 0,00 €)", { bold: true, size: 10, fill: gap === 0 ? "FFE8F5E9" : "FFFFEBEE" });
    put(ws, r, 2, "", { fill: gap === 0 ? "FFE8F5E9" : "FFFFEBEE" });
    put(ws, r, 3, { formula: `ROUND(C${closingRow}-C${shownRow},2)`, result: gap }, { bold: true, size: 10, numFmt: MONEY_TOTAL, fill: gap === 0 ? "FFE8F5E9" : "FFFFEBEE" });
  }
}

// ---------------------------------------------------------------------------
// 7. Légende
// ---------------------------------------------------------------------------
const LEGENDE: [string, string][] = [
  ["Paiement Express Checkout  —  Crédit normal", "FFE8F5E9"],
  ["Suspension de paiement  —  Paiement en attente (débit temporaire)", "FFFFF3E0"],
  ["Déblocage de paiement  —  Libération de la suspension", "FFF3E5F5"],
  ["Remboursement de paiement  —  Remboursement client", "FFFFEBEE"],
  ["Retrait initié par l'utilisateur  —  Virement vers banque", "FFE3F2FD"],
  ["Solde suspendu pour enquête sur un litige  —  Litige en cours", "FFFFFDE7"],
  ["Annulation de la suspension pour résolution du litige  —  Litige résolu", "FFF1F8E9"],
];

function buildLegende(wb: ExcelJS.Workbook): void {
  const ws = wb.addWorksheet("Légende");
  ws.getColumn(1).width = 40;
  ws.getColumn(2).width = 20;
  headerRow(ws, 1, ["Type de transaction", "Couleur"], C.navy, 18, 10);
  let r = 2;
  for (const [label, fill] of LEGENDE) {
    put(ws, r, 1, label, { fill });
    put(ws, r, 2, "", { fill });
    r += 1;
  }
}

// ---------------------------------------------------------------------------
// 8. Chiffre d'affaires (TTC / TVA / HT)
// ---------------------------------------------------------------------------
const CA_MONTHS = [
  "Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
  "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre",
];

const caHeaders = (first: string, pct: number): string[] => [
  first, "Nb ventes", "CA brut TTC (€)", "Remboursements TTC (€)", "CA net TTC (€)",
  `TVA collectée ${pct} % (€)`, "CA net HT (€)", "Commissions PayPal (€)", "Net encaissé (€)",
];

interface CaLine { key: string; label: string; nb: number; brut: number; remb: number; comm: number }

/** Regroupe les ventes et remboursements comptabilisables selon `keyOf` (les cles de `order` d'abord). */
function caLines(
  txs: Transaction[],
  keyOf: (t: Transaction) => { key: string; label: string },
  order: string[] = []
): CaLine[] {
  const map = new Map<string, CaLine>();
  for (const t of txs) {
    const sale = isSale(t);
    if (!sale && !isRefund(t)) continue;
    const { key, label } = keyOf(t);
    const g = map.get(key) ?? { key, label, nb: 0, brut: 0, remb: 0, comm: 0 };
    if (sale) {
      g.nb += 1;
      g.brut += t.gross;
    } else {
      g.remb += t.gross; // le client recupere le TTC ; la commission eventuellement rendue va avec les commissions
    }
    g.comm += t.fee;
    map.set(key, g);
  }
  const rank = (k: string): number => {
    const i = order.indexOf(k);
    return i === -1 ? order.length : i;
  };
  return [...map.values()].sort((a, b) => rank(a.key) - rank(b.key) || a.key.localeCompare(b.key));
}

/** Ecrit un tableau de CA + sa ligne TOTAL (somme des lignes arrondies). Retourne la ligne suivante libre. */
function writeCaTable(
  ws: ExcelJS.Worksheet,
  start: number,
  lines: CaLine[],
  rate: number,
  fillOf: (l: CaLine) => string
): number {
  const money = { align: "right" as const, numFmt: MONEY_FMT };
  let r = start;
  let tNb = 0, tBrut = 0, tRemb = 0, tNet = 0, tComm = 0, tEnc = 0;
  for (const l of lines) {
    const brut = r2(l.brut);
    const remb = r2(l.remb);
    const comm = r2(l.comm);
    const netTtc = r2(brut + remb);
    const ht = r2(netTtc / (1 + rate));
    const tva = r2(netTtc - ht);
    const enc = r2(netTtc + comm);
    const fill = fillOf(l);
    put(ws, r, 1, l.label, { fill });
    put(ws, r, 2, l.nb, { fill, align: "center" });
    put(ws, r, 3, brut, { fill, ...money });
    put(ws, r, 4, remb, { fill, ...money });
    put(ws, r, 5, netTtc, { fill, ...money });
    put(ws, r, 6, tva, { fill, ...money });
    put(ws, r, 7, ht, { fill, ...money });
    put(ws, r, 8, comm, { fill, ...money });
    put(ws, r, 9, enc, { fill, ...money });
    tNb += l.nb; tBrut += brut; tRemb += remb; tNet += netTtc; tComm += comm; tEnc += enc;
    r += 1;
  }
  // Comme pour une declaration de TVA, HT et TVA du total sont calcules sur le total TTC (et non en
  // additionnant des lignes arrondies) : tous les tableaux d'une feuille affichent ainsi le meme total.
  const tHt = r2(tNet / (1 + rate));
  const tTva = r2(tNet - tHt);
  const tot = { align: "right" as const, numFmt: MONEY_TOTAL, bold: true, color: C.white };
  put(ws, r, 1, "TOTAL", { fill: C.dark, bold: true, color: C.white, align: "right" });
  put(ws, r, 2, tNb, { fill: C.dark, bold: true, color: C.white, align: "center" });
  put(ws, r, 3, r2(tBrut), { fill: C.dark, ...tot });
  put(ws, r, 4, r2(tRemb), { fill: C.dark, ...tot });
  put(ws, r, 5, r2(tNet), { fill: C.dark, ...tot });
  put(ws, r, 6, tTva, { fill: C.dark, ...tot });
  put(ws, r, 7, tHt, { fill: C.dark, ...tot });
  put(ws, r, 8, r2(tComm), { fill: C.dark, ...tot });
  put(ws, r, 9, r2(tEnc), { fill: C.dark, ...tot });
  ws.getRow(r).height = 22;
  return r + 1;
}

/** `all` : si fourni, ajoute la repartition du CA par activite (classeur du compte complet). */
function buildCA(
  wb: ExcelJS.Workbook,
  txs: Transaction[],
  rate: number,
  scope: string,
  all?: Transaction[]
): void {
  const ws = wb.addWorksheet("Chiffre d'affaires", { views: [{ state: "frozen", ySplit: 2 }] });
  [16, 10, 18, 20, 18, 17, 16, 18, 16].forEach((w, i) => (ws.getColumn(i + 1).width = w));
  const pct = Number((rate * 100).toFixed((rate * 100) % 1 ? 1 : 0));

  band(ws, 1, 9, `CHIFFRE D'AFFAIRES  —  TTC / TVA / HT  (TVA ${pct} %)${scope}`, {
    fill: C.navy, bold: true, color: C.white, size: 13, align: "center", wrap: true,
  });
  ws.getRow(1).height = 34;
  headerRow(ws, 2, caHeaders("Mois", pct), C.navy, 30, 10);

  const month = (t: Transaction): { key: string; label: string } => {
    if (!t.date) return { key: "0000-00", label: "Sans date" };
    const m = t.date.getMonth();
    const y = t.date.getFullYear();
    return { key: `${y}-${String(m + 1).padStart(2, "0")}`, label: `${CA_MONTHS[m]} ${y}` };
  };
  let r = writeCaTable(ws, 3, caLines(txs, month), rate, () => "FFE8F5E9");

  band(ws, r + 1, 9,
    `ℹ️  CA net TTC = ventes − remboursements.  TVA collectée = CA net TTC × ${pct}/${100 + pct}.  CA net HT = CA net TTC / ${(1 + rate).toFixed(2)}.  Les commissions PayPal sont des charges (déduites du net encaissé), sans TVA.  Les paiements annulés ou en attente ne sont pas comptés. HT et TVA du TOTAL sont calculés sur le total TTC : une ligne peut différer d'un centime.`,
    { italic: true, size: 9, color: C.grayTxt, wrap: true });
  ws.getRow(r + 1).height = 38;
  r += 3;

  if (all) {
    band(ws, r, 9, "RÉPARTITION PAR ACTIVITÉ  —  Même période, ventilée selon l'activité affectée à chaque opération", {
      fill: C.blue, bold: true, color: C.white, size: 11,
    });
    ws.getRow(r).height = 24;
    r += 1;
    headerRow(ws, r, caHeaders("Activité", pct), C.navy, 30, 10);
    r += 1;
    const activity = (t: Transaction): { key: string; label: string } => {
      const label = t.activity || UNASSIGNED;
      return { key: label, label };
    };
    const order = [...ACTIVITIES.map((a) => a.label), UNASSIGNED];
    writeCaTable(ws, r, caLines(all, activity, order), rate, (l) =>
      l.label === UNASSIGNED ? "FFFFF3E0" : "FFE8F5E9"
    );
  }
  ws.autoFilter = { from: "A2", to: "I2" };
}

// ---------------------------------------------------------------------------
// Assemblage
// ---------------------------------------------------------------------------
export interface WorkbookOptions {
  /**
   * Libelle d'une activite : le classeur ne contient alors que ses operations (memes feuilles,
   * memes colonnes). Sans cette option : le compte complet, toutes activites confondues.
   */
  activity?: string;
}

export function buildWorkbook(result: PipelineResult, opts: WorkbookOptions = {}): ExcelJS.Workbook {
  const wb = new ExcelJS.Workbook();
  wb.creator = "Lettrage Auto";
  wb.created = new Date();
  wb.calcProperties.fullCalcOnLoad = true; // Excel recalcule les formules (SOLDE NET, ecart) a l'ouverture

  const { activity } = opts;
  const all = sortChrono(result.transactions);
  // Classeur d'activite : ses operations seulement. Les virements, eux, restent ceux du compte.
  const txs = activity ? all.filter((t) => t.activity === activity) : all;
  const bankTxs = activity ? all.filter((t) => isRetrait(t) || t.activity === activity) : all;
  const byId = new Map<string, Transaction>();
  for (const t of all) if (t.transactionId) byId.set(t.transactionId, t);
  const periodes = decouperParVirement(bankTxs);
  const rate = result.vatRate ?? 0.2;
  const scope = activity ? `  —  ${activity}` : "";
  wb.title = `Lettrage PayPal${scope}`;

  buildTransactions(wb, txs, rate);
  buildVirements(wb, periodes, activity);
  buildRapprochement(wb, periodes, activity);
  buildRemboursements(wb, txs, byId);
  buildAnomalies(wb, txs);
  buildResume(wb, txs, scope, !activity);
  buildCA(wb, txs, rate, scope, activity ? undefined : all);
  buildLegende(wb);

  return wb;
}
