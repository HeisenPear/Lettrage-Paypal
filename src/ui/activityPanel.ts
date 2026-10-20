import {
  ACTIVITIES,
  COMMON,
  RULE_FIELDS,
  RULE_OPS,
  SOURCE,
  UNASSIGNED,
  parseManual,
  resolveActivity,
  type ActivityConfig,
  type RuleField,
  type RuleOp,
  type TextRule,
} from "../core/activity";
import {
  buildOrderSource,
  countMatches,
  detectKeyColumn,
  distinctValues,
  parseOrderFile,
  type OrderFile,
  type OrderSettings,
} from "../core/orders";
import { decodeBytes } from "../core/parse";
import type { ActivityStats } from "../core/types";
import { chip, h, selectOf } from "./dom";

export interface ActivityPanel {
  /** Configuration d'affectation correspondant aux reglages affiches. */
  buildConfig(): ActivityConfig;
  /** Met a jour le bilan et les compteurs apres chaque traitement des paiements. */
  update(stats: ActivityStats, wanted: Set<string>): void;
}

const LABELS = ACTIVITIES.map((a) => a.label);
const MAX_VALUES = 40;

// --- Reglages memorises dans le navigateur (jamais les donnees) --------------------
const STORE = {
  rules: "lettrage-auto.rules",
  manual: "lettrage-auto.manual",
  fallback: "lettrage-auto.fallback",
};

function load<T>(key: string, fallback: T): T {
  try {
    const raw = localStorage.getItem(key);
    return raw ? (JSON.parse(raw) as T) : fallback;
  } catch {
    return fallback; // stockage indisponible ou corrompu : on repart de zero
  }
}

function save(key: string, value: unknown): void {
  try {
    localStorage.setItem(key, JSON.stringify(value));
  } catch {
    // Stockage indisponible (navigation privee...) : les reglages ne seront simplement pas retenus.
  }
}

/** Regles retenues par le navigateur ; celles qui ne sont plus valides (activite disparue...) sont ecartees. */
function loadRules(): TextRule[] {
  const stored = load<unknown[]>(STORE.rules, []);
  if (!Array.isArray(stored)) return [];
  const rules: TextRule[] = [];
  for (const item of stored) {
    const x = item as Partial<TextRule> | null;
    if (!x || typeof x.value !== "string" || typeof x.activity !== "string") continue;
    if (x.field === undefined || !(x.field in RULE_FIELDS)) continue;
    if (x.op === undefined || !(x.op in RULE_OPS)) continue;
    const activity = resolveActivity(x.activity);
    if (activity) rules.push({ field: x.field, op: x.op, value: x.value, activity });
  }
  return rules;
}

// --- Export de commandes : une carte par fichier ------------------------------------
interface Card {
  file: OrderFile;
  settings: OrderSettings;
  /** Recalcule le compteur de commandes reconnues et la table valeur -> activite. */
  refresh: () => void;
}

/** Colonne du fichier dont les valeurs designent des activites (ex. "duhalle-boutique.fr"), ou -1. */
function guessActivityColumn(file: OrderFile, exclude: number): number {
  let best = -1;
  let bestScore = 0;
  for (let c = 0; c < file.headers.length; c++) {
    if (c === exclude) continue;
    const { values, truncated } = distinctValues(file, c, 12);
    if (truncated) continue;
    const score = values.filter((v) => resolveActivity(v)).length;
    if (score > bestScore) {
      best = c;
      bestScore = score;
    }
  }
  return best;
}

function defaultValueMap(file: OrderFile, col: number): Record<string, string> {
  const map: Record<string, string> = {};
  for (const v of distinctValues(file, col, MAX_VALUES).values) map[v] = resolveActivity(v) ?? "";
  return map;
}

export function createActivityPanel(root: HTMLElement, deps: { onChange: () => void }): ActivityPanel {
  let wanted = new Set<string>();
  let rules = loadRules();
  let manualText = load<string>(STORE.manual, "");
  const storedFallback = load<string>(STORE.fallback, UNASSIGNED);
  let fallback = LABELS.includes(storedFallback) ? storedFallback : UNASSIGNED;
  const cards: Card[] = [];

  const commit = (): void => deps.onChange();

  // --- En-tete : bilan -----------------------------------------------------------------
  const summary = h("div", { class: "stats" });
  const origin = h("p", { class: "hint" });
  const notice = h("div", { class: "banner" });

  const fallbackSelect = selectOf(
    [[UNASSIGNED, "À affecter (ne rien supposer)"], ...LABELS.map((l): [string, string] => [l, `Considérer comme ${l}`])],
    fallback
  );
  fallbackSelect.addEventListener("change", () => {
    fallback = fallbackSelect.value;
    save(STORE.fallback, fallback);
    commit();
  });

  // --- Exports de commandes ---------------------------------------------------------
  const cardsEl = h("div", { class: "cards" });
  const ordersError = h("div", { class: "banner error" });
  ordersError.hidden = true;
  const fileInput = h("input", { type: "file", accept: ".csv,.txt,text/csv" });
  fileInput.multiple = true;
  fileInput.hidden = true;
  const addBtn = h("button", { type: "button", class: "btn-secondary" }, "+ Ajouter un export de commandes");

  function newCard(file: OrderFile): Card {
    const key = detectKeyColumn(file, wanted).index;
    const actCol = guessActivityColumn(file, key);
    const settings: OrderSettings =
      actCol >= 0
        ? { keyColumn: key, activityColumn: actCol, valueMap: defaultValueMap(file, actCol), fixedActivity: "" }
        : { keyColumn: key, activityColumn: -1, valueMap: {}, fixedActivity: resolveActivity(file.name) ?? "" };
    return { file, settings, refresh: () => undefined };
  }

  function renderCard(card: Card): HTMLElement {
    const { file, settings } = card;
    const columns: [string, string][] = file.headers.map((name, i) => [String(i), name || `Colonne ${i + 1}`]);

    const keySelect = selectOf([["-1", "— aucune —"], ...columns], String(settings.keyColumn));
    const modeSelect = selectOf(
      [
        ["fixed:", "— choisir l'activité —"],
        ...LABELS.map((l): [string, string] => [`fixed:${l}`, `Toutes les commandes : ${l}`]),
        ["column", "Selon une colonne du fichier…"],
      ],
      settings.activityColumn >= 0 ? "column" : `fixed:${settings.fixedActivity}`
    );
    const columnSelect = selectOf(columns, String(Math.max(settings.activityColumn, 0)));
    const columnField = h("label", { class: "field" }, h("span", {}, "Colonne qui donne l'activité"), columnSelect);
    const mappingEl = h("div", { class: "value-map" });
    const matchesEl = h("span", { class: "match" });

    card.refresh = (): void => {
      const n = countMatches(file, settings.keyColumn, wanted);
      matchesEl.className = `match ${n > 0 ? "ok" : "warn"}`;
      matchesEl.textContent =
        settings.keyColumn < 0
          ? "Choisis la colonne du n° de commande."
          : n > 0
            ? `${n} commande(s) de ce fichier reconnue(s) parmi les paiements.`
            : "Aucune commande reconnue : mauvaise colonne, ou période différente de celle du CSV PayPal.";

      columnField.hidden = settings.activityColumn < 0;
      mappingEl.replaceChildren();
      if (settings.activityColumn < 0) return;
      const { values, truncated } = distinctValues(file, settings.activityColumn, MAX_VALUES);
      if (truncated) {
        mappingEl.append(
          h("p", { class: "hint" }, "Cette colonne a trop de valeurs différentes : choisis plutôt le site, la boutique ou la marque.")
        );
      }
      for (const value of values) {
        const sel = selectOf(
          [["", "— ignorer —"], ...LABELS.map((l): [string, string] => [l, l])],
          settings.valueMap[value] ?? ""
        );
        sel.addEventListener("change", () => {
          settings.valueMap[value] = sel.value;
          commit();
        });
        mappingEl.append(h("div", { class: "map-row" }, h("span", { class: "map-value" }, value), "→", sel));
      }
    };

    keySelect.addEventListener("change", () => {
      settings.keyColumn = Number(keySelect.value);
      card.refresh();
      commit();
    });
    modeSelect.addEventListener("change", () => {
      if (modeSelect.value === "column") {
        const guess = guessActivityColumn(file, settings.keyColumn);
        const other = file.headers.findIndex((_, i) => i !== settings.keyColumn);
        settings.activityColumn = guess >= 0 ? guess : other;
        if (settings.activityColumn < 0) {
          modeSelect.value = `fixed:${settings.fixedActivity}`; // pas d'autre colonne disponible
          return;
        }
        settings.valueMap = defaultValueMap(file, settings.activityColumn);
        columnSelect.value = String(settings.activityColumn);
      } else {
        settings.activityColumn = -1;
        settings.fixedActivity = modeSelect.value.slice("fixed:".length);
      }
      card.refresh();
      commit();
    });
    columnSelect.addEventListener("change", () => {
      settings.activityColumn = Number(columnSelect.value);
      settings.valueMap = defaultValueMap(file, settings.activityColumn);
      card.refresh();
      commit();
    });

    const remove = h("button", { type: "button", class: "link danger" }, "Retirer");
    remove.addEventListener("click", () => {
      cards.splice(cards.indexOf(card), 1);
      renderCards();
      commit();
    });

    const el = h(
      "div",
      { class: "card" },
      h(
        "div",
        { class: "card-head" },
        h("strong", {}, file.name),
        h("span", { class: "muted" }, `${file.rows.length} ligne(s)`),
        remove
      ),
      h(
        "div",
        { class: "card-grid" },
        h("label", { class: "field" }, h("span", {}, "Colonne du n° de commande"), keySelect),
        h("label", { class: "field" }, h("span", {}, "Activité des commandes"), modeSelect),
        columnField
      ),
      matchesEl,
      mappingEl
    );
    card.refresh();
    return el;
  }

  function renderCards(): void {
    cardsEl.replaceChildren(...cards.map(renderCard));
  }

  async function addFiles(files: File[]): Promise<void> {
    const problems: string[] = [];
    for (const f of files) {
      try {
        const file = parseOrderFile(decodeBytes(new Uint8Array(await f.arrayBuffer())), f.name);
        if (file.headers.length === 0 || file.rows.length === 0) problems.push(`${f.name} : fichier vide ou illisible`);
        else cards.push(newCard(file));
      } catch (err) {
        problems.push(`${f.name} : ${(err as Error).message}`);
      }
    }
    ordersError.hidden = problems.length === 0;
    ordersError.textContent = problems.join(" · ");
    renderCards();
    commit();
  }

  addBtn.addEventListener("click", () => fileInput.click());
  fileInput.addEventListener("change", () => {
    void addFiles([...(fileInput.files ?? [])]);
    fileInput.value = "";
  });

  const ordersBlock = h(
    "div",
    { class: "activity-block" },
    h("h3", {}, "Exports de commandes ", h("span", { class: "tag" }, "recommandé")),
    h(
      "p",
      { class: "hint" },
      "Dépose l'export des commandes d'une boutique (CSV, glisser-déposer possible). L'outil retrouve seul la colonne du n° de commande en la recoupant avec les paiements PayPal, puis rattache chaque paiement, remboursement ou litige à son activité."
    ),
    addBtn,
    fileInput,
    ordersError,
    cardsEl
  );
  ["dragenter", "dragover"].forEach((evt) =>
    ordersBlock.addEventListener(evt, (e) => {
      e.preventDefault();
      ordersBlock.classList.add("dragover");
    })
  );
  ["dragleave", "drop"].forEach((evt) =>
    ordersBlock.addEventListener(evt, (e) => {
      e.preventDefault();
      ordersBlock.classList.remove("dragover");
    })
  );
  ordersBlock.addEventListener("drop", (e) => {
    void addFiles([...((e as DragEvent).dataTransfer?.files ?? [])]);
  });

  // --- Regles de reconnaissance ---------------------------------------------------------
  const rulesEl = h("div", { class: "rules" });
  function renderRules(): void {
    rulesEl.replaceChildren(
      ...rules.map((rule, i) => {
        const field = selectOf(Object.entries(RULE_FIELDS), rule.field);
        const op = selectOf(Object.entries(RULE_OPS), rule.op);
        const value = h("input", { type: "text", placeholder: "ex. DUH-" });
        value.value = rule.value;
        const activity = selectOf(
          LABELS.map((l): [string, string] => [l, l]),
          rule.activity
        );
        const persist = (): void => {
          rule.field = field.value as RuleField;
          rule.op = op.value as RuleOp;
          rule.value = value.value;
          rule.activity = activity.value;
          save(STORE.rules, rules);
          commit();
        };
        [field, op, value, activity].forEach((el) => el.addEventListener("change", persist));
        const remove = h("button", { type: "button", class: "link danger", "aria-label": "Supprimer la règle" }, "✕");
        remove.addEventListener("click", () => {
          rules.splice(i, 1);
          save(STORE.rules, rules);
          renderRules();
          commit();
        });
        return h("div", { class: "rule-row" }, "Si", field, op, value, "→", activity, remove);
      })
    );
  }
  const addRule = h("button", { type: "button", class: "btn-secondary" }, "+ Ajouter une règle");
  addRule.addEventListener("click", () => {
    rules = [...rules, { field: "invoiceNumber", op: "startsWith", value: "", activity: LABELS[0] }];
    renderRules(); // la regle ne compte qu'une fois son texte saisi
  });
  const rulesBlock = h(
    "details",
    { class: "activity-block" },
    h("summary", {}, "Règles de reconnaissance (marqueur à la source)"),
    h(
      "p",
      { class: "hint" },
      "Pour les futurs paiements : fais ajouter un préfixe ou un libellé distinct par boutique dans la référence envoyée à PayPal (ex. « DUH- » devant le n° de facture), puis déclare-le ici. Les règles sont retenues dans ce navigateur et passent avant les exports de commandes."
    ),
    rulesEl,
    addRule
  );

  // --- Affectations manuelles -------------------------------------------------------------
  const manualArea = h("textarea", {
    rows: "5",
    spellcheck: "false",
    placeholder: "50748526 ; Duhallé\n8SE80913GR6984840 ; La Jocondienne",
  });
  manualArea.value = manualText;
  const manualInfo = h("p", { class: "hint" });
  const refreshManualInfo = (): void => {
    const m = parseManual(manualText);
    manualInfo.textContent = manualText.trim()
      ? `${m.accepted} affectation(s) reconnue(s)${m.ignored ? `, ${m.ignored} ligne(s) ignorée(s) (attendu : n° puis activité)` : ""}.`
      : "Une affectation par ligne : n° de commande (ou n° de transaction PayPal), puis l'activité. Elle l'emporte sur tout le reste.";
  };
  refreshManualInfo();
  manualArea.addEventListener("change", () => {
    manualText = manualArea.value;
    save(STORE.manual, manualText);
    refreshManualInfo();
    commit();
  });
  const manualBlock = h(
    "details",
    { class: "activity-block" },
    h("summary", {}, "Affectations manuelles (exceptions)"),
    manualArea,
    manualInfo
  );

  renderRules();

  root.className = "panel activities";
  root.append(
    h("h2", { class: "panel-title" }, `Activités : ${LABELS.join(" / ")}`),
    h(
      "p",
      { class: "hint" },
      "Le CSV PayPal ne dit pas toujours de quelle boutique vient un paiement. Chaque opération est affectée dans cet ordre : affectation manuelle, règle de reconnaissance, export de commandes, signature de la boutique, opération liée (ex. le remboursement suit son paiement), valeur par défaut."
    ),
    summary,
    origin,
    notice,
    h("label", { class: "field inline" }, h("span", {}, "Une opération non identifiée est :"), fallbackSelect),
    ordersBlock,
    rulesBlock,
    manualBlock
  );

  return {
    buildConfig(): ActivityConfig {
      return {
        rules: rules.filter((r) => r.value.trim() !== ""),
        manual: parseManual(manualText).map,
        sources: cards.map((c) => buildOrderSource(c.file, c.settings)),
        fallback,
      };
    },

    update(stats, keys): void {
      wanted = keys;
      const count = (label: string): string => String(stats.byActivity[label] ?? 0);
      summary.replaceChildren(
        ...LABELS.map((l) => chip(l, count(l))),
        chip(UNASSIGNED, count(UNASSIGNED)),
        chip(COMMON, count(COMMON))
      );

      const parts = [SOURCE.signature, SOURCE.orders, SOURCE.rule, SOURCE.manual, SOURCE.linked, SOURCE.fallback]
        .filter((k) => stats.bySource[k])
        .map((k) => `${k} : ${stats.bySource[k]}`);
      origin.textContent = parts.length ? `Origine des affectations : ${parts.join(" · ")}.` : "";

      const left = stats.byActivity[UNASSIGNED] ?? 0;
      notice.className = `banner ${left > 0 ? "warn" : "ok"}`;
      notice.textContent =
        left > 0
          ? `${left} opération(s) restent à affecter : ajoute l'export des commandes de la boutique, une règle de reconnaissance, ou choisis une valeur par défaut. Tant qu'elles ne sont pas affectées, elles n'apparaissent que dans le compte complet.`
          : "Toutes les opérations sont affectées à une activité.";
      if (stats.conflicts > 0) {
        notice.textContent += ` Attention : ${stats.conflicts} groupe(s) d'opérations liées (ex. un paiement et son remboursement) sont affectés à des activités différentes : vérifie tes règles.`;
      }

      for (const card of cards) card.refresh();
    },
  };
}
