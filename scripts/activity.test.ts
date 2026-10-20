// Tests du moteur d'activites et du classeur. Lancer avec : npm test
import assert from "node:assert/strict";
import { test } from "node:test";
import ExcelJS from "exceljs";
import {
  COMMON,
  UNASSIGNED,
  emptyActivityConfig,
  parseManual,
  resolveActivity,
  transactionKeySet,
  type ActivityConfig,
} from "../src/core/activity";
import { buildOrderSource, countMatches, detectKeyColumn, distinctValues, parseOrderFile } from "../src/core/orders";
import { runPipeline } from "../src/core/pipeline";
import type { PipelineResult, Transaction } from "../src/core/types";
import { buildWorkbook } from "../src/excel/workbook";

const J = "La Jocondienne";
const D = "Duhallé";

// --- Jeu de donnees synthetique (format d'export PayPal francais) -----------------
const HEADERS = [
  "Date", "Heure", "Fuseau horaire", "Nom", "Type", "État", "Devise", "Avant commission",
  "Commission", "Net", "De l'adresse email", "À l'adresse email", "Numéro de transaction",
  "Titre de l'objet", "Numéro de la transaction de référence", "Numéro de facture",
  "Numéro de client", "Solde", "Impact sur le solde", "Identifiant bancaire",
  "Source de paiement", "Code du pays de l'acheteur pour cette transaction",
];
const PRESTA = "1.7.8.3_6.7.1_7.4.33_Cart ID: 27572_Shop name: La Jocondienne";
const id = (n: number): string => `PAYPALTX${String(n).padStart(9, "0")}`;

// [date, nom, type, etat, brut, commission, net, id, titre, ref, facture, clientRef, solde, impact, bankId]
type Row = [string, string, string, string, string, string, string, string, string, string, string, string, string, string, string];
const ROWS: Row[] = [
  ["01/05/2026", "Client A", "Paiement Express Checkout", "Terminé", "100,00", "-3,40", "96,60", id(1), "134613 50000001 256", "", "50000001", "", "96,60", "Crédit", ""],
  ["02/05/2026", "Client B", "Paiement Express Checkout", "Terminé", "50,00", "-1,90", "48,10", id(2), "134613 50000002 256", "", "50000002", "", "144,70", "Crédit", ""],
  ["03/05/2026", "Client C", "Paiement Express Checkout", "Terminé", "30,00", "-1,20", "28,80", id(3), "Sachet de 50 bouchons", "", "", PRESTA, "173,50", "Crédit", ""],
  ["04/05/2026", "Client A", "Remboursement de paiement", "Terminé", "-100,00", "0,00", "-100,00", id(4), "134613 50000001 256", id(1), "50000001", "", "73,50", "Débit", ""],
  ["05/05/2026", "Client D", "Paiement Express Checkout", "Annulé", "25,00", "0,00", "25,00", id(5), "134613 50000003 256", "", "50000003", "", "73,50", "Mémo", ""],
  ["06/05/2026", "", "Retrait initié par l'utilisateur", "En cours de traitement", "-73,50", "0,00", "-73,50", id(6), "", "", "", "", "0,00", "Débit", "2605060001"],
  ["07/05/2026", "Client E", "Paiement Express Checkout", "Terminé", "80,00", "-2,80", "77,20", id(7), "134613 50000004 256", "", "50000004", "", "77,20", "Crédit", ""],
  ["08/05/2026", "Client C", "Remboursement de paiement", "Terminé", "-30,00", "0,00", "-30,00", id(8), "", id(3), "", "", "47,20", "Débit", ""],
];

function csv(rows: Row[] = ROWS): string {
  const q = (v: string): string => `"${v.replace(/"/g, '""')}"`;
  const line = (r: Row): string =>
    [r[0], "10:00:00", "CEST", r[1], r[2], r[3], "EUR", r[4], r[5], r[6], "client@example.com",
      "vendeur@example.com", r[7], r[8], r[9], r[10], r[11], r[12], r[13], r[14], "PayPal", "FR"]
      .map(q).join(",");
  return [HEADERS.map(q).join(","), ...rows.map(line)].join("\n");
}

function run(activity?: ActivityConfig): PipelineResult {
  return runPipeline(csv(), { vatRate: 0.2, activity });
}
const byId = (r: PipelineResult, n: number): Transaction => {
  const t = r.transactions.find((x) => x.transactionId === id(n));
  assert.ok(t, `transaction ${n} introuvable`);
  return t;
};

/** Exports de commandes factices : un fichier par activite, avec une colonne "reference" qui ne correspond a rien. */
function sourcesFor(wanted: Set<string>): ActivityConfig["sources"] {
  const make = (text: string, name: string, activity: string): ActivityConfig["sources"][number] => {
    const file = parseOrderFile(text, name);
    const { index } = detectKeyColumn(file, wanted);
    return buildOrderSource(file, { keyColumn: index, activityColumn: -1, valueMap: {}, fixedActivity: activity });
  };
  return [
    make("Réf;N° commande;Client\nA1;50000001;Client A\nA2;50000004;Client E", "duhalle.csv", D),
    make("Réf;N° commande;Client\nB1;50000002;Client B", "jocondienne.csv", J),
  ];
}

// --- Activites ---------------------------------------------------------------------
test("resolveActivity reconnait les activites dans un texte libre", () => {
  assert.equal(resolveActivity("DUHALLÉ"), D);
  assert.equal(resolveActivity("duhalle-boutique.fr"), D);
  assert.equal(resolveActivity("la jocondienne"), J);
  assert.equal(resolveActivity("autre boutique"), null);
});

test("sans regle : la signature de boutique est reconnue, le reste est a affecter", () => {
  const r = run();
  assert.equal(byId(r, 3).activity, J);
  assert.equal(byId(r, 3).activitySource, "Signature boutique");
  assert.equal(byId(r, 1).activity, UNASSIGNED);
  assert.equal(byId(r, 7).activity, UNASSIGNED);
});

test("les virements vers la banque restent communs", () => {
  const w = byId(run(), 6);
  assert.equal(w.activity, COMMON);
  assert.equal(w.activitySource, "Mouvement du compte");
});

test("un remboursement sans n° de commande suit le paiement auquel il est lie", () => {
  const refund = byId(run(), 8);
  assert.equal(refund.activity, J);
  assert.equal(refund.activitySource, "Opération liée");
});

// --- Exports de commandes -----------------------------------------------------------
test("la colonne du n° de commande est trouvee par recoupement, sans connaitre les entetes", () => {
  const wanted = transactionKeySet(run().transactions);
  const file = parseOrderFile("Réf;N° commande;Client\nA1;50000001;Client A\nA2;50000004;Client E", "x.csv");
  const best = detectKeyColumn(file, wanted);
  assert.equal(best.index, 1);
  assert.equal(best.matches, 2);
  assert.equal(countMatches(file, 0, wanted), 0);
});

test("les exports de commandes affectent paiements ET remboursements du meme n° de commande", () => {
  const wanted = transactionKeySet(run().transactions);
  const r = run({ ...emptyActivityConfig(), sources: sourcesFor(wanted) });
  assert.equal(byId(r, 1).activity, D);
  assert.equal(byId(r, 4).activity, D, "le remboursement porte le n° de commande du paiement");
  assert.equal(byId(r, 4).activitySource, "Export commandes · duhalle.csv");
  assert.equal(byId(r, 2).activity, J);
  assert.equal(byId(r, 7).activity, D);
  assert.equal(byId(r, 5).activity, UNASSIGNED, "commande absente des exports");
});

test("un export dont une colonne donne l'activite : correspondance valeur -> activite", () => {
  const file = parseOrderFile("N° commande;Site\n50000001;duhalle-boutique.fr\n50000002;lajocondienne.fr\n50000004;autre", "mixte.csv");
  const { values, truncated } = distinctValues(file, 1);
  assert.deepEqual(values, ["duhalle-boutique.fr", "lajocondienne.fr", "autre"]);
  assert.equal(truncated, false);
  const valueMap = Object.fromEntries(values.map((v) => [v, resolveActivity(v) ?? ""]));
  const src = buildOrderSource(file, { keyColumn: 0, activityColumn: 1, valueMap, fixedActivity: "" });
  assert.equal(src.keys.get("50000001"), D);
  assert.equal(src.keys.get("50000002"), J);
  assert.equal(src.keys.has("50000004"), false, "valeur ignoree");
});

// --- Priorites -----------------------------------------------------------------------
test("valeur par defaut : les operations non identifiees prennent l'activite choisie", () => {
  const r = run({ ...emptyActivityConfig(), fallback: D });
  assert.equal(byId(r, 1).activity, D);
  assert.equal(byId(r, 1).activitySource, "Valeur par défaut");
  assert.equal(byId(r, 3).activity, J, "la signature reste prioritaire sur la valeur par defaut");
});

test("une regle de reconnaissance l'emporte sur la signature et sur l'export", () => {
  const wanted = transactionKeySet(run().transactions);
  const cfg: ActivityConfig = {
    ...emptyActivityConfig(),
    sources: sourcesFor(wanted),
    rules: [{ field: "article", op: "contains", value: "BOUCHONS", activity: D }],
  };
  const r = run(cfg);
  assert.equal(byId(r, 3).activity, D);
  assert.match(byId(r, 3).activitySource, /^Règle · /);
});

test("une regle sur le n° de facture (marqueur a la source) fonctionne avec 'commence par' et regex", () => {
  const starts = run({ ...emptyActivityConfig(), rules: [{ field: "invoiceNumber", op: "startsWith", value: "5000000", activity: J }] });
  assert.equal(byId(starts, 1).activity, J);
  const regex = run({ ...emptyActivityConfig(), rules: [{ field: "invoiceNumber", op: "regex", value: "^50000004$", activity: D }] });
  assert.equal(byId(regex, 7).activity, D);
  assert.equal(byId(regex, 1).activity, UNASSIGNED);
  const invalid = run({ ...emptyActivityConfig(), rules: [{ field: "article", op: "regex", value: "([", activity: D }] });
  assert.equal(byId(invalid, 1).activity, UNASSIGNED, "une regex invalide est ignoree sans planter");
});

test("l'affectation manuelle l'emporte sur tout, et un conflit entre operations liees est signale", () => {
  const wanted = transactionKeySet(run().transactions);
  const manual = parseManual("50000001 ; La Jocondienne\n# commentaire\nligne invalide\n50000004 = Duhallé");
  assert.equal(manual.accepted, 2);
  assert.equal(manual.ignored, 1);
  const r = run({ ...emptyActivityConfig(), sources: sourcesFor(wanted), manual: manual.map });
  assert.equal(byId(r, 1).activity, J, "le manuel bat l'export qui disait Duhallé");
  assert.equal(byId(r, 1).activitySource, "Affectation manuelle");
  assert.equal(r.activityStats.conflicts, 0);

  // Le remboursement (TX4) est force vers Duhallé alors que son paiement (TX1) va a La Jocondienne.
  const refundKey = id(4);
  const conflicting = run({ ...emptyActivityConfig(), manual: new Map([["50000001", J], [refundKey, D]]) });
  assert.equal(byId(conflicting, 4).activity, D);
  assert.equal(conflicting.activityStats.conflicts, 1);
});

test("le bilan d'affectation couvre toutes les operations", () => {
  const wanted = transactionKeySet(run().transactions);
  const r = run({ ...emptyActivityConfig(), sources: sourcesFor(wanted) });
  const total = Object.values(r.activityStats.byActivity).reduce((a, b) => a + b, 0);
  assert.equal(total, r.transactions.length);
  assert.equal(r.activityStats.byActivity[COMMON], 1);
});

// --- Chiffres : paiements annules exclus --------------------------------------------
test("les indicateurs ne comptent pas les paiements annules (impact Mémo)", () => {
  const d = run().dashboard;
  assert.equal(d.nbPaiements, 4);
  assert.equal(d.caBrut, 260);
});

// --- Classeur Excel -------------------------------------------------------------------
async function reload(wb: ExcelJS.Workbook): Promise<ExcelJS.Workbook> {
  const out = new ExcelJS.Workbook();
  await out.xlsx.load((await wb.xlsx.writeBuffer()) as ArrayBuffer);
  return out;
}
const text = (ws: ExcelJS.Worksheet, r: number, c: number): string => {
  const v = ws.getCell(r, c).value;
  return v == null ? "" : typeof v === "object" && "result" in v ? String(v.result) : String(v);
};
const num = (ws: ExcelJS.Worksheet, r: number, c: number): number => {
  const v = ws.getCell(r, c).value;
  return typeof v === "object" && v && "result" in v ? Number(v.result) : Number(v);
};
function rowOf(ws: ExcelJS.Worksheet, label: string): number {
  for (let r = 1; r <= ws.rowCount; r++) if (text(ws, r, 1) === label) return r;
  assert.fail(`ligne "${label}" introuvable dans ${ws.name}`);
}

test("la feuille CA exclut les paiements annules et ventile par activite", async () => {
  const wanted = transactionKeySet(run().transactions);
  const wb = await reload(buildWorkbook(run({ ...emptyActivityConfig(), sources: sourcesFor(wanted) })));
  const ca = wb.getWorksheet("Chiffre d'affaires")!;
  const total = rowOf(ca, "TOTAL");
  assert.equal(num(ca, total, 2), 4, "4 ventes (la commande annulee n'est pas comptee)");
  assert.equal(num(ca, total, 3), 260);
  assert.equal(num(ca, total, 4), -130);
  assert.equal(num(ca, total, 5), 130);
  assert.equal(num(ca, total, 6), 21.67, "TVA = 130 x 20/120");
  assert.equal(num(ca, total, 7), 108.33, "HT = 130 / 1,20");

  // Repartition par activite : Duhalle (TX1, TX4, TX7), Jocondienne (TX2, TX3, TX8), rien a affecter.
  assert.equal(num(ca, rowOf(ca, D), 5), 80, "Duhalle : 100 + 80 - 100");
  assert.equal(num(ca, rowOf(ca, J), 5), 50, "Jocondienne : 50 + 30 - 30");
});

test("un classeur par activite : leurs CA additionnes retombent sur celui du compte", async () => {
  const wanted = transactionKeySet(run().transactions);
  const result = run({ ...emptyActivityConfig(), sources: sourcesFor(wanted) });
  const net = async (activity?: string): Promise<number> => {
    const ws = (await reload(buildWorkbook(result, { activity }))).getWorksheet("Chiffre d'affaires")!;
    return num(ws, rowOf(ws, "TOTAL"), 5);
  };
  const [all, duhalle, jocondienne] = [await net(), await net(D), await net(J)];
  assert.equal(all, 130);
  assert.equal(duhalle + jocondienne, all);
});

test("le classeur d'activite ne contient que ses operations, mais les virements du compte", async () => {
  const wanted = transactionKeySet(run().transactions);
  const wb = await reload(buildWorkbook(run({ ...emptyActivityConfig(), sources: sourcesFor(wanted) }), { activity: D }));
  const tx = wb.getWorksheet("Transactions PayPal")!;
  const activities = new Set<string>();
  for (let r = 2; r <= tx.rowCount; r++) activities.add(text(tx, r, 21));
  assert.deepEqual([...activities], [D]);
  const vir = wb.getWorksheet("Virements Banque")!;
  assert.equal(num(vir, 3, 6), -73.5, "le virement reste celui du compte entier");
  assert.match(text(wb.getWorksheet("Résumé")!, 1, 1), /Duhallé$/);
});

test("le Résumé exclut les lignes Mémo et son controle de solde retombe sur zéro", async () => {
  const wb = await reload(buildWorkbook(run()));
  const res = wb.getWorksheet("Résumé")!;
  assert.equal(num(res, rowOf(res, "SOLDE NET"), 3), 47.2, "96,60+48,10+28,80-100-73,50+77,20-30 (hors paiement annule)");
  // Le solde calcule retombe sur celui que PayPal affiche : l'ecart est une formule ROUND(calcule - affiche).
  const calcule = num(res, rowOf(res, "Solde de clôture calculé"), 3);
  const affiche = num(res, rowOf(res, "Solde de clôture affiché par PayPal"), 3);
  assert.equal(calcule, 47.2);
  assert.equal(affiche, 47.2);
  const ecart = res.getCell(rowOf(res, "ÉCART (doit être 0,00 €)"), 3).value as { formula: string };
  assert.match(ecart.formula, /^ROUND\(C\d+-C\d+,2\)$/);
});

test("le rapprochement ne liste pas les lignes Mémo", async () => {
  const wb = await reload(buildWorkbook(run()));
  const rap = wb.getWorksheet("Rapprochement")!;
  for (let r = 1; r <= rap.rowCount; r++) assert.notEqual(text(rap, r, 11), id(5));
});
