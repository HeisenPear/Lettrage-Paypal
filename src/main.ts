import "./style.css";
import { ACTIVITIES, transactionKeySet } from "./core/activity";
import { runPipeline } from "./core/pipeline";
import { PROFILES } from "./core/profiles";
import { decodeBytes } from "./core/parse";
import { type Field, type PipelineResult } from "./core/types";
import { fillForType } from "./excel/styles";
import { createActivityPanel } from "./ui/activityPanel";
import { chip } from "./ui/dom";
// ExcelJS est volumineux : on le charge a la demande (voir generate()).

// --- Acces aux elements ----------------------------------------------------
const $ = <T extends HTMLElement>(id: string): T => {
  const el = document.getElementById(id);
  if (!el) throw new Error(`Element introuvable: #${id}`);
  return el as T;
};

const dropEl = $("drop");
const fileInput = $<HTMLInputElement>("file");
const browseBtn = $<HTMLButtonElement>("browse");
const errorEl = $("error");
const configEl = $("config");
const resultsEl = $("results");
const fileNameEl = $<HTMLOutputElement>("fileName");
const profileSelect = $<HTMLSelectElement>("profile");
const delimiterSelect = $<HTMLSelectElement>("delimiter");
const vatInput = $<HTMLInputElement>("vat");
const statsEl = $("stats");
const mappingEl = $("mapping");
const previewTable = $<HTMLTableElement>("preview");
const previewCountEl = $("previewCount");
const activitiesEl = $("activities");
const downloadsEl = $("downloads");
const statusEl = $("status");

// --- Etat ------------------------------------------------------------------
interface State {
  fileName: string;
  text: string | null;
  profileId: string;
  delimiter: string;
  result: PipelineResult | null;
}
const state: State = {
  fileName: "",
  text: null,
  profileId: "auto",
  delimiter: "",
  result: null,
};

const nf = new Intl.NumberFormat("fr-FR", {
  minimumFractionDigits: 2,
  maximumFractionDigits: 2,
});

const FIELD_LABELS: Record<Field, string> = {
  date: "Date",
  time: "Heure",
  type: "Type / libellé",
  name: "Client / nom",
  email: "Email",
  currency: "Devise",
  gross: "Brut",
  fee: "Commission",
  net: "Net",
  balance: "Solde",
  debit: "Débit",
  credit: "Crédit",
  article: "Article",
  invoiceNumber: "N° facture",
  impact: "Sens",
  source: "Source",
  country: "Pays",
  bankId: "Identifiant bancaire",
  transactionId: "N° transaction",
  referenceId: "Réf. associée",
  customerRef: "Numéro de client",
  note: "Remarque",
  status: "État",
};

// Panneau des activites : affecte chaque operation a La Jocondienne / Duhalle.
const panel = createActivityPanel(activitiesEl, { onChange: run });

const argbToCss = (argb: string): string => `#${argb.slice(2)}`;
const fmtDate = (d: Date | null, raw = ""): string =>
  d ? d.toLocaleDateString("fr-FR") : raw;

// --- Initialisation des selecteurs -----------------------------------------
function populateProfiles(): void {
  profileSelect.innerHTML = "";
  const auto = document.createElement("option");
  auto.value = "auto";
  auto.textContent = "Auto-détection";
  profileSelect.appendChild(auto);
  for (const p of PROFILES) {
    const opt = document.createElement("option");
    opt.value = p.id;
    opt.textContent = p.label;
    profileSelect.appendChild(opt);
  }
  profileSelect.value = state.profileId;
}

// --- Erreurs ---------------------------------------------------------------
function showError(message: string): void {
  errorEl.textContent = message;
  errorEl.hidden = false;
}
function clearError(): void {
  errorEl.textContent = "";
  errorEl.hidden = true;
}

// --- Lecture du fichier -----------------------------------------------------
async function handleFile(file: File): Promise<void> {
  try {
    clearError();
    const bytes = new Uint8Array(await file.arrayBuffer());
    state.fileName = file.name;
    state.text = decodeBytes(bytes);
    fileNameEl.textContent = `${file.name} (${(file.size / 1024).toFixed(0)} Ko)`;
    populateProfiles();
    run();
  } catch (err) {
    showError(`Impossible de lire le fichier : ${(err as Error).message}`);
  }
}

// --- Traitement -------------------------------------------------------------
function run(): void {
  if (!state.text) return;
  try {
    clearError();
    const vatPct = Number.parseFloat(vatInput.value);
    const vatRate = Number.isFinite(vatPct) ? vatPct / 100 : 0.2;
    const result = runPipeline(state.text, {
      profileId: state.profileId,
      delimiter: state.delimiter || undefined,
      vatRate,
      activity: panel.buildConfig(),
    });
    state.result = result;

    if (result.stats.txCount === 0) {
      showError(
        "Aucune opération exploitable n'a été trouvée. Vérifiez le séparateur ou la plateforme sélectionnée."
      );
    }

    configEl.hidden = false;
    activitiesEl.hidden = false;
    resultsEl.hidden = false;
    panel.update(result.activityStats, transactionKeySet(result.transactions));
    renderStats(result);
    renderMapping(result);
    renderPreview(result);
    renderDownloads(result);
    statusEl.textContent = "";
  } catch (err) {
    showError(`Erreur de traitement : ${(err as Error).message}`);
  }
}

// --- Rendu : statistiques ---------------------------------------------------
function renderStats(result: PipelineResult): void {
  statsEl.innerHTML = "";
  const s = result.stats;
  const d = result.dashboard;
  const euro = (n: number) => `${nf.format(n)} €`;
  statsEl.append(
    chip("Plateforme", result.profileLabel),
    chip("Mouvements", String(s.txCount)),
    chip("Paiements", String(d.nbPaiements)),
    chip("CA brut", euro(d.caBrut)),
    chip("Commissions", euro(d.commissions)),
    chip("Net encaissé", euro(d.netEncaisse)),
    chip(
      "Période",
      s.periodStart ? `${fmtDate(s.periodStart)} → ${fmtDate(s.periodEnd)}` : "—"
    )
  );
}

// --- Rendu : mapping des colonnes ------------------------------------------
function renderMapping(result: PipelineResult): void {
  mappingEl.innerHTML = "";
  const entries = Object.entries(result.mapping) as [Field, string][];
  for (const [field, header] of entries) {
    const item = document.createElement("div");
    item.className = "map-item";
    const f = document.createElement("span");
    f.className = "map-field";
    f.textContent = FIELD_LABELS[field] ?? field;
    const h = document.createElement("span");
    h.className = "map-header";
    h.textContent = header;
    item.append(f, h);
    mappingEl.appendChild(item);
  }
  for (const field of result.missingFields) {
    const item = document.createElement("div");
    item.className = "map-item missing";
    item.textContent = `⚠ ${FIELD_LABELS[field] ?? field} introuvable`;
    mappingEl.appendChild(item);
  }
}

// --- Rendu : apercu ---------------------------------------------------------
function renderPreview(result: PipelineResult): void {
  previewTable.innerHTML = "";
  const cols = ["Date", "Client", "Type", "Brut", "Commission", "Net", "Impact"];
  const thead = document.createElement("thead");
  const htr = document.createElement("tr");
  for (const c of cols) {
    const th = document.createElement("th");
    th.textContent = c;
    htr.appendChild(th);
  }
  thead.appendChild(htr);
  previewTable.appendChild(thead);

  const tbody = document.createElement("tbody");
  const rows = result.transactions.slice(0, 25);
  for (const t of rows) {
    const tr = document.createElement("tr");
    const bg = argbToCss(fillForType(t.type));
    const cells: [string, boolean?][] = [
      [fmtDate(t.date, t.dateRaw)],
      [t.name],
      [t.type],
      [nf.format(t.gross), true],
      [nf.format(t.fee), true],
      [nf.format(t.net), true],
      [t.impact],
    ];
    cells.forEach(([text, num]) => {
      const td = document.createElement("td");
      td.textContent = text;
      td.style.background = bg;
      if (num) td.className = "num";
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  }
  previewTable.appendChild(tbody);
  previewCountEl.textContent = `(${rows.length} sur ${result.transactions.length})`;
}

// --- Telechargements --------------------------------------------------------
/** Un bouton pour le compte complet, puis un par activite (grise s'il n'y a rien a y mettre). */
function renderDownloads(result: PipelineResult): void {
  downloadsEl.replaceChildren();
  const add = (label: string, activity: string | undefined, disabled = false): void => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = activity ? "btn-secondary" : "btn-primary";
    button.textContent = label;
    button.disabled = disabled;
    button.addEventListener("click", () => void generate(activity));
    downloadsEl.append(button);
  };
  add("⬇ Compte complet", undefined);
  for (const { label } of ACTIVITIES) {
    const n = result.activityStats.byActivity[label] ?? 0;
    add(`⬇ ${label} (${n})`, label, n === 0);
  }
}

// --- Generation Excel -------------------------------------------------------
async function generate(activity?: string): Promise<void> {
  if (!state.result) return;
  try {
    downloadsEl.querySelectorAll("button").forEach((b) => (b.disabled = true));
    statusEl.textContent = "Génération en cours…";
    const [{ buildWorkbook }, { buildFilename, downloadWorkbook }] = await Promise.all([
      import("./excel/workbook"),
      import("./excel/download"),
    ]);
    const wb = buildWorkbook(state.result, { activity });
    await downloadWorkbook(wb, buildFilename(state.fileName, activity));
    statusEl.textContent = `✓ Fichier Excel téléchargé${activity ? ` (${activity})` : ""}`;
  } catch (err) {
    showError(`Échec de la génération : ${(err as Error).message}`);
    statusEl.textContent = "";
  } finally {
    if (state.result) renderDownloads(state.result); // retablit l'etat des boutons
  }
}

// --- Evenements -------------------------------------------------------------
browseBtn.addEventListener("click", () => fileInput.click());
fileInput.addEventListener("change", () => {
  const file = fileInput.files?.[0];
  if (file) void handleFile(file);
});

["dragenter", "dragover"].forEach((evt) =>
  dropEl.addEventListener(evt, (e) => {
    e.preventDefault();
    dropEl.classList.add("dragover");
  })
);
["dragleave", "drop"].forEach((evt) =>
  dropEl.addEventListener(evt, (e) => {
    e.preventDefault();
    if (evt === "dragleave" && e.target !== dropEl) return;
    dropEl.classList.remove("dragover");
  })
);
dropEl.addEventListener("drop", (e) => {
  const file = (e as DragEvent).dataTransfer?.files?.[0];
  if (file) void handleFile(file);
});
dropEl.addEventListener("click", (e) => {
  if (e.target === browseBtn) return;
  fileInput.click();
});

profileSelect.addEventListener("change", () => {
  state.profileId = profileSelect.value;
  run();
});
delimiterSelect.addEventListener("change", () => {
  state.delimiter = delimiterSelect.value;
  run();
});
vatInput.addEventListener("change", run);

populateProfiles();
