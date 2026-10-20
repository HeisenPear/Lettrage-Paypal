# Lettrage Auto

Outil web qui transforme un **export CSV** (PayPal, relevé bancaire, etc.) en
**fichier Excel trié, catégorisé et prêt pour le lettrage comptable**.

On glisse un CSV, l'outil lit / trie / catégorise / rapproche les opérations,
puis génère un classeur `.xlsx` directement exploitable.

> 🔒 **100 % local** : tout le traitement se fait dans le navigateur. Aucune
> donnée financière n'est envoyée sur un serveur.

---

## ✨ Ce que fait l'outil

À partir d'un CSV, il produit un classeur Excel reproduisant **la structure de
lettrage de référence** — 8 feuilles, même mise en forme et mêmes codes couleur ;
**seules les données changent** :

| Feuille | Contenu |
|---|---|
| **Transactions PayPal** | Toutes les opérations, colorées **par type** (vert = paiement, crème = suspension, rose = remboursement, bleu = retrait, violet = déblocage, jaune = litige…). En fin de ligne : **Montant HT**, **TVA**, **Activité** et **Origine affectation**. |
| **Virements Banque** | Un virement PayPal → banque par retrait : période couverte, identifiant bancaire, montant viré, solde restant, nombre de transactions, à pointer (✓ Lettré). |
| **Rapprochement** | Toutes les transactions **groupées par période de virement**, avec sous-total et ligne « ▶ VIREMENT BANQUE » par période. |
| **Remboursements** | Liste des remboursements liés à leur **transaction d'origine** (date + montant du paiement initial). |
| **Anomalies & Attentions** | Sections par cas à vérifier : suspensions, déblocages, litiges, Pay Later (4X), paiements annulés ou en attente (lignes « Mémo »). |
| **Résumé** | Nombre et montant net par type ; totaux crédits / débits / solde net ; **contrôle du solde** (le solde calculé doit retomber sur celui affiché par PayPal). |
| **Chiffre d'affaires** | Par mois : CA brut TTC, remboursements, CA net TTC, **TVA collectée**, **CA net HT**, commissions, net encaissé. Puis la **répartition par activité**. |
| **Légende** | Signification de chaque couleur. |

Les montants sont de **vrais nombres** (format `#,##0.00 €`, sommables), pas du
texte.

### Fiabilité des chiffres

- **Les lignes « Mémo » ne comptent pas.** PayPal exporte aussi des lignes
  d'information (paiement annulé, en attente, doublon d'un paiement débloqué) dont
  l'impact sur le solde est « Mémo ». Elles sont exclues du chiffre d'affaires,
  du Résumé et du rapprochement, et listées à part dans « Anomalies ». Vérification :
  la somme des *Net* hors Mémo retombe exactement sur le solde PayPal.
- **TVA** : le montant brut est traité comme du **TTC**. HT = TTC / (1 + taux). Le
  taux se règle dans l'application (20 % par défaut). Comme pour une déclaration,
  HT et TVA du *total* sont calculés sur le total TTC : une ligne peut différer d'un
  centime, mais tous les tableaux affichent le même total.
- Les définitions de « vente », « remboursement », « retrait » sont **uniques**
  (`src/core/classify.ts`) : le classeur et les indicateurs de l'appli ne peuvent
  pas diverger.

### Deux activités sur un même compte PayPal

Le CSV PayPal ne dit pas toujours de quelle boutique vient un paiement (même adresse
marchand, même source de paiement, même référence). L'outil affecte donc **chaque
opération à une activité** (aujourd'hui *La Jocondienne* et *Duhallé*), dans cet ordre
de priorité :

1. **Affectation manuelle** : un n° de commande ou de transaction → une activité.
2. **Règle de reconnaissance** : « si *N° de facture* commence par `DUH-` → Duhallé ».
   C'est le **marqueur à la source** : fais ajouter un préfixe ou un libellé distinct
   par boutique dans la référence envoyée à PayPal, puis déclare-le. Les règles sont
   retenues dans le navigateur.
3. **Export de commandes** : tu déposes l'export des commandes d'une boutique (CSV).
   L'outil retrouve **seul** la colonne du n° de commande en la recoupant avec les
   paiements, puis rattache paiements, remboursements, suspensions et litiges de la
   même commande. L'activité vient soit du fichier (« toutes les commandes : Duhallé »),
   soit d'une colonne (site, boutique, marque) dont chaque valeur est associée à une
   activité.
4. **Signature de la boutique** : certaines boutiques (module PayPal de PrestaShop)
   écrivent leur nom dans « Numéro de client » (`…Shop name: La Jocondienne`).
5. **Opération liée** : un remboursement sans n° de commande suit le paiement auquel il
   se rapporte.
6. **Valeur par défaut** (au choix : *À affecter*, ou une activité).

Les **virements vers la banque** restent des mouvements du **compte commun** : ils ne
peuvent pas être partagés par activité.

Chaque ligne indique **d'où vient son affectation** (colonne « Origine affectation »),
et l'application affiche ce qui reste **à affecter**.

**Fichiers produits** (boutons de téléchargement) :

- **Compte complet** : les 8 feuilles, toutes activités, avec la répartition du CA par
  activité. À utiliser pour le lettrage bancaire.
- **Un classeur par activité** : les mêmes 8 feuilles, avec seulement les opérations de
  l'activité. « Virements Banque » et « Rapprochement » restent ceux du compte, avec une
  mention : seuls les nombres de transactions, les crédits nets et les sous-totaux
  concernent l'activité.

Ajouter une activité : une entrée dans `ACTIVITIES` (`src/core/activity.ts`).

### Multi-plateformes

L'outil n'est **pas limité à PayPal**. Il repose sur des **profils** qui
décrivent comment lire une source donnée (mapping des colonnes, format des
montants/dates, règles de catégorisation). La plateforme est **détectée
automatiquement** à partir des en-têtes, et peut être forcée manuellement.

Profils inclus :
- **PayPal** (exports « Activité » / « Transactions », FR et EN) ;
- **Banque / générique** (colonnes Date / Libellé / Montant ou Débit / Crédit).

---

## 🚀 Utilisation

### En ligne
Déployé en site statique (Vercel) : ouvrez l'URL, glissez votre CSV, réglez la TVA et
les activités si besoin, puis téléchargez le **compte complet** ou le classeur d'une
activité.

### En local
```bash
npm install
npm run dev      # serveur de dev (http://localhost:5173)
npm run build    # build de production dans dist/
npm run preview  # prévisualise le build
npm test         # tests du moteur d'activités et du classeur
```

### Démo en ligne de commande
Génère un `.xlsx` à partir du CSV d'exemple (utile pour tester le moteur) :
```bash
npm run demo
# ou avec vos propres fichiers :
npm run demo chemin/source.csv chemin/sortie.xlsx
```

---

## 🧩 Architecture

```
src/
├── core/                 # moteur (sans dépendance au navigateur, testable en Node)
│   ├── types.ts          # types partagés (Transaction, Category…)
│   ├── csv.ts            # parsing CSV (PapaParse)
│   ├── parse.ts          # montants, dates, encodage
│   ├── normalize.ts      # CSV brut -> opérations normalisées
│   ├── classify.ts       # définitions uniques : vente, remboursement, retrait, Mémo
│   ├── lettrage.ts       # rapprochement par référence (union-find) + codes
│   ├── activity.ts       # affectation des opérations aux activités
│   ├── orders.ts         # lecture des exports de commandes
│   ├── dashboard.ts      # indicateurs de l'application
│   ├── pipeline.ts       # orchestration complète
│   └── profiles/         # profils par plateforme
│       ├── profile.ts    # interface Profile + helpers
│       ├── paypal.ts
│       ├── generic.ts
│       └── index.ts      # registre + détection automatique
├── excel/                # génération du classeur (ExcelJS)
│   ├── styles.ts         # couleurs, formats
│   ├── workbook.ts       # construit les 8 feuilles (compte complet ou une activité)
│   └── download.ts       # téléchargement navigateur
├── ui/
│   ├── activityPanel.ts  # panneau des activités (exports, règles, manuel)
│   └── dom.ts            # petites aides de construction du DOM
├── main.ts               # interface (drag & drop, aperçu, options)
└── style.css
scripts/
├── demo.ts               # démo en ligne de commande
└── activity.test.ts      # tests (node:test)
```

### Ajouter une plateforme
1. Créer `src/core/profiles/ma-plateforme.ts` exportant un objet `Profile`
   (mapping des colonnes + fonctions `match` et `categorize`).
2. L'ajouter au tableau `PROFILES` dans `src/core/profiles/index.ts`.

C'est tout : le profil apparaît dans le sélecteur et entre dans la détection
automatique.

---

## 🛠️ Stack
- **Vite** + **TypeScript** (site statique, aucun back-end)
- **PapaParse** — lecture CSV robuste (séparateur auto, guillemets, encodage)
- **ExcelJS** — génération de `.xlsx` mis en forme (couleurs, filtres, totaux)

## ⚠️ Avertissement
Cet outil prépare le lettrage et le suivi du chiffre d'affaires : **vérifiez
toujours les écritures** avant intégration dans votre comptabilité. La TVA est
calculée en supposant que le montant brut encaissé est TTC, au taux saisi : à
valider selon votre régime réel.
