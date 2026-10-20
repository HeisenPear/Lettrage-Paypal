// Types partages par tout le moteur de lettrage.

/**
 * Categorie comptable normalisee d'une operation.
 * Elle est independante de la plateforme source : chaque profil traduit ses
 * propres libelles vers ces categories.
 */
export type Category =
  | "VENTE" // encaissement / paiement recu
  | "REMBOURSEMENT" // remboursement emis vers un client
  | "FRAIS" // commission / frais isole (ligne dediee)
  | "RETRAIT_BANQUE" // virement de la plateforme vers le compte bancaire
  | "RECHARGE_BANQUE" // approvisionnement depuis la banque / carte
  | "LITIGE" // chargeback, litige, annulation, reversal
  | "CONVERSION" // conversion de devise
  | "TRANSFERT" // envoi d'argent / paiement sortant (achat)
  | "AUTRE"; // non classe

export const CATEGORY_LABELS: Record<Category, string> = {
  VENTE: "Vente / encaissement",
  REMBOURSEMENT: "Remboursement",
  FRAIS: "Frais / commission",
  RETRAIT_BANQUE: "Retrait vers banque",
  RECHARGE_BANQUE: "Approvisionnement",
  LITIGE: "Litige / chargeback",
  CONVERSION: "Conversion de devise",
  TRANSFERT: "Transfert / achat",
  AUTRE: "Autre",
};

/** Une ligne brute du CSV : entete -> valeur texte. */
export type RawRow = Record<string, string>;

/** Une operation normalisee, exploitable pour le lettrage. */
export interface Transaction {
  index: number; // position d'origine dans le fichier (1-based)
  date: Date | null;
  dateRaw: string;
  time: string;
  type: string; // libelle d'origine
  category: Category;
  name: string; // contrepartie / client
  email: string; // adresse email de la contrepartie
  currency: string;
  gross: number; // montant brut
  fee: number; // commission (negatif en general)
  net: number; // montant net
  balance: number | null; // solde apres operation
  article: string; // titre de l'objet / libelle article
  invoiceNumber: string; // numero de facture
  impact: string; // sens : Credit / Debit / Memo
  source: string; // source de paiement (PayPal, Pay Later...)
  country: string; // pays de l'acheteur
  bankId: string; // identifiant bancaire (pour les retraits)
  transactionId: string;
  referenceId: string; // numero de transaction associee
  customerRef: string; // "Numero de client" : signature de la boutique (ex. nom de boutique)
  note: string; // remarque libre
  status: string; // etat (Termine, En attente, Supprime...)
  pending: boolean; // operation en attente
  lettrage: string; // code de lettrage (rempli par le moteur)
  group: number; // identifiant interne du groupe lettre (0 = non lettre)
  activity: string; // activite affectee (rempli par le moteur d'activites)
  activitySource: string; // comment l'activite a ete determinee
}

/** Champs normalises que l'on cherche a mapper depuis les entetes du CSV. */
export type Field =
  | "date"
  | "time"
  | "type"
  | "name"
  | "email"
  | "currency"
  | "gross"
  | "fee"
  | "net"
  | "balance"
  | "debit"
  | "credit"
  | "article"
  | "invoiceNumber"
  | "impact"
  | "source"
  | "country"
  | "bankId"
  | "transactionId"
  | "referenceId"
  | "customerRef"
  | "note"
  | "status";

/** Resultat du mapping : pour chaque champ, l'entete reellement trouvee. */
export type ResolvedMapping = Partial<Record<Field, string>>;

/** Indicateurs du tableau de bord (sur la periode de l'export). */
export interface DashboardStats {
  nbPaiements: number;
  caBrut: number;
  commissions: number;
  netEncaisse: number;
  nbRemboursements: number;
  totalRemboursements: number;
  nbRetraits: number;
  totalRetraits: number;
  nbAttente: number;
  totalAttente: number;
  nbAutres: number;
}

export interface PipelineStats {
  rowCount: number;
  txCount: number;
  groupCount: number; // nombre de groupes effectivement lettres (>1 ligne)
  lettredCount: number; // nombre de lignes portant un code de lettrage
  currencies: string[];
  periodStart: Date | null;
  periodEnd: Date | null;
}

/** Bilan de l'affectation des operations aux activites. */
export interface ActivityStats {
  byActivity: Record<string, number>; // libelle d'activite -> nombre d'operations
  bySource: Record<string, number>; // origine de l'affectation -> nombre d'operations
  conflicts: number; // groupes lies dont les membres sont affectes a des activites differentes
}

export interface PipelineResult {
  transactions: Transaction[];
  mapping: ResolvedMapping;
  missingFields: Field[]; // champs essentiels introuvables
  dashboard: DashboardStats;
  activityStats: ActivityStats;
  stats: PipelineStats;
  profileId: string;
  profileLabel: string;
  vatRate: number;
}
