// Demo Node : lit un CSV, execute le pipeline et ecrit un .xlsx.
// Usage : npm run demo [chemin/vers/source.csv] [chemin/vers/sortie.xlsx]
import { readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { decodeBytes } from "../src/core/parse";
import { runPipeline } from "../src/core/pipeline";
import { CATEGORY_LABELS } from "../src/core/types";
import { buildWorkbook } from "../src/excel/workbook";

const here = dirname(fileURLToPath(import.meta.url));
const csvPath = process.argv[2] ?? resolve(here, "../sample-data/paypal-exemple.csv");
const outPath = process.argv[3] ?? resolve(here, "../demo-output.xlsx");

const text = decodeBytes(new Uint8Array(readFileSync(csvPath)));
const result = runPipeline(text, { vatRate: 0.2 });

const eur = new Intl.NumberFormat("fr-FR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

console.log("=== Lettrage Auto — demo ===");
console.log(`Source            : ${csvPath}`);
console.log(`Plateforme        : ${result.profileLabel}`);
console.log(`Lignes lues       : ${result.stats.rowCount}`);
console.log(`Operations        : ${result.stats.txCount}`);
console.log(`Groupes lettres   : ${result.stats.groupCount}`);
console.log(`Devises           : ${result.stats.currencies.join(", ") || "—"}`);
console.log("");

console.log("Operations :");
for (const t of result.transactions) {
  const code = t.lettrage ? `[${t.lettrage}]` : "[ ]";
  console.log(
    `  ${code} ${t.dateRaw.padEnd(10)} ${CATEGORY_LABELS[t.category].padEnd(22)} ` +
      `brut ${eur.format(t.gross).padStart(10)}  net ${eur.format(t.net).padStart(10)}  ${t.transactionId}`
  );
}

const d = result.dashboard;
console.log("\nVentes encaissees :");
console.log(
  `  ${d.nbPaiements} paiements | CA brut ${eur.format(d.caBrut)} | ` +
    `commissions ${eur.format(d.commissions)} | net ${eur.format(d.netEncaisse)}`
);

console.log("\nActivites (aucune regle fournie : tout est a affecter sauf signatures reconnues) :");
for (const [name, n] of Object.entries(result.activityStats.byActivity)) {
  console.log(`  ${name.padEnd(16)} ${n}`);
}

const wb = buildWorkbook(result);
await wb.xlsx.writeFile(outPath);
console.log(`\nFichier Excel ecrit : ${outPath}`);
