import type { Profile } from "./profiles";
import { norm } from "./profiles/profile";
import {
  detectDateOrder,
  detectDecimal,
  parseAmount,
  parseDate,
  type Decimal,
  type DateOrder,
} from "./parse";
import type { Field, RawRow, ResolvedMapping, Transaction } from "./types";

/** Champs indispensables pour produire un lettrage exploitable. */
const ESSENTIAL: Field[] = ["date", "type"];

export interface NormalizeResult {
  transactions: Transaction[];
  mapping: ResolvedMapping;
  missingFields: Field[];
  decimal: Decimal;
  dateOrder: DateOrder;
}

/** Associe chaque champ normalise a l'entete reelle du fichier. */
export function resolveMapping(headers: string[], profile: Profile): ResolvedMapping {
  const byNorm = new Map<string, string>();
  for (const h of headers) byNorm.set(norm(h), h);

  const mapping: ResolvedMapping = {};
  for (const [field, candidates] of Object.entries(profile.columns) as [Field, string[]][]) {
    for (const cand of candidates) {
      const actual = byNorm.get(norm(cand));
      if (actual) {
        mapping[field] = actual;
        break;
      }
    }
  }
  return mapping;
}

export function normalizeRows(
  rows: RawRow[],
  headers: string[],
  profile: Profile
): NormalizeResult {
  const mapping = resolveMapping(headers, profile);
  const get = (row: RawRow, field: Field): string => {
    const header = mapping[field];
    return header ? (row[header] ?? "").trim() : "";
  };

  // Detection globale du format des montants et des dates.
  const amountSamples: string[] = [];
  const dateSamples: string[] = [];
  for (const row of rows) {
    for (const f of ["gross", "net", "fee", "balance", "debit", "credit"] as Field[]) {
      const v = get(row, f);
      if (v) amountSamples.push(v);
    }
    const d = get(row, "date");
    if (d) dateSamples.push(d);
  }
  const decimal: Decimal = profile.decimal === "auto" ? detectDecimal(amountSamples) : profile.decimal;
  const dateOrder: DateOrder =
    !profile.dateOrder || profile.dateOrder === "auto"
      ? detectDateOrder(dateSamples)
      : profile.dateOrder;

  const hasGross = !!mapping.gross;
  const hasDebitCredit = !!mapping.debit || !!mapping.credit;

  const transactions: Transaction[] = rows.map((row, i) => {
    let gross: number;
    if (hasGross) {
      gross = parseAmount(get(row, "gross"), decimal);
    } else if (hasDebitCredit) {
      const credit = parseAmount(get(row, "credit"), decimal);
      const debit = parseAmount(get(row, "debit"), decimal);
      // Le debit est exprime en valeur absolue dans les releves : on le soustrait.
      gross = credit - Math.abs(debit);
    } else {
      gross = parseAmount(get(row, "net"), decimal);
    }

    const fee = mapping.fee ? parseAmount(get(row, "fee"), decimal) : 0;
    let net: number;
    if (mapping.net) net = parseAmount(get(row, "net"), decimal);
    else net = gross + fee;

    const balance = mapping.balance ? parseAmount(get(row, "balance"), decimal) : null;
    const type = get(row, "type");
    const dateRaw = get(row, "date");
    const status = get(row, "status");
    const pending = /attente|pending|unclaimed|non reclame|on hold/.test(norm(status));

    return {
      index: i + 1,
      date: parseDate(dateRaw, dateOrder),
      dateRaw,
      time: get(row, "time"),
      type,
      category: profile.categorize(type, gross, net),
      name: get(row, "name"),
      email: get(row, "email"),
      currency: get(row, "currency").toUpperCase(),
      gross,
      fee,
      net,
      balance,
      article: get(row, "article"),
      invoiceNumber: get(row, "invoiceNumber"),
      impact: get(row, "impact"),
      source: get(row, "source"),
      country: get(row, "country"),
      bankId: get(row, "bankId"),
      transactionId: get(row, "transactionId"),
      referenceId: get(row, "referenceId"),
      customerRef: get(row, "customerRef"),
      note: get(row, "note"),
      status,
      pending,
      lettrage: "",
      group: 0,
      activity: "",
      activitySource: "",
    };
  });

  const missingFields = ESSENTIAL.filter((f) => !mapping[f]);
  if (!hasGross && !hasDebitCredit && !mapping.net) missingFields.push("gross");

  return { transactions, mapping, missingFields, decimal, dateOrder };
}
