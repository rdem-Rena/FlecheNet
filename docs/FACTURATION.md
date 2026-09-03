# Formulaire de facturation

`UF_Facture` — attribue un numéro de facture aux interventions de **`TblInterv`** qui n'en ont pas encore.

Il s'ouvre par le bouton **Facturer** du formulaire des interventions, et se génère par **`GenererFormulaireFacture`** (module `modFact_Generateur`).

> Ce fichier est produit à partir des sources : les valeurs sont celles réellement en vigueur dans `src/`.


---

## Plan

Fenêtre de **960 × 636 points** de surface utile (≈ 1280 × 848 pixels à 96 ppp). Cinq zones, chacune portée par un vrai cadre.

| Zone | Cadre | Haut | Hauteur | Contenu |
|---|---|---|---|---|
| 1 | `fraFCarte1` | 12 | 46 | titre et année |
| 2 | `fraFCarte2` | 66 | 160 | tableau des clients à facturer |
| 3 | `fraFCarte3` | 234 | 40 | filtre des mois, « toutes les opérations », nouveau numéro |
| 4 | `fraFCarte4` | 282 | 220 | tableau des travaux, totaux, bouton Enregistrer |
| 5 | `fraFCarte5` | 510 | 76 | texte de facture et commentaires de la ligne choisie |

### Tableau des clients

9 lignes visibles, défilement virtuel : ce n'est pas la grille qui bouge mais son contenu.

| # | Colonne de `TblInterv` | Titre | Largeur |
|---|---|---|---|
| 1 | `Client_No` | N° client | 80 pt |
| 2 | `Entreprise` | Entreprise | 320 pt |
| 3 | `Titre` | Titre | 90 pt |
| 4 | `Nom` | Nom | 210 pt |
| 5 | `Prenom` | Prénom | 190 pt |
| | | | **890 pt** |

### Tableau des travaux

12 lignes visibles, défilement virtuel : ce n'est pas la grille qui bouge mais son contenu.

| # | Colonne de `TblInterv` | Titre | Largeur |
|---|---|---|---|
| 1 | `NoInterv` | N° | 46 pt |
| 2 | `Date` | Date | 64 pt |
| 3 | `Nb_Hres` | Heures | 52 pt |
| 4 | `Nb_Pers` | Pers. | 38 pt |
| 5 | `Taux/Forfait` | Taux/Forf. | 62 pt |
| 6 | `TVA` | TVA | 34 pt |
| 7 | `Forfait` | Forf. | 38 pt |
| 8 | `CA` | CA | 62 pt |
| 9 | `Texte_Facture` | Texte de facture | 200 pt |
| 10 | `Commentaires` | Commentaires | 200 pt |
| 11 | `No_Facture` | Fact. | 46 pt |
| 12 | `Select.` | Select. | 46 pt |
| | | | **888 pt** |

---

## Fonctionnement

- Le tableau du haut ne montre que les lignes **sans numéro de facture**, dédoublonnées : deux interventions d'un même client n'y font qu'une entrée.

- Choisir un client remplit le tableau du bas avec ses travaux non facturés. La case **Toutes les opérations** y ajoute ceux déjà facturés ; le menu **Mois** restreint à un mois.

- La colonne **Select.** n'existe pas dans `TblInterv` : elle se coche à l'écran, d'un clic dans la case, et décide des lignes qu'Enregistrer met à jour.

- **CA est calculé, jamais lu** : heures × personnes × taux, ou le taux seul si `Forfait` est vrai. C'est `Interv_EstimerCA`, le même calcul que le formulaire des interventions — les deux écrans montrent donc le même montant pour la même intervention.

- Les deux totaux, sous `Nb_Hres` et `CA`, portent sur les lignes **affichées**.

- **Enregistrer** écrit le numéro saisi dans `No_Facture` pour les lignes cochées, puis recharge tout : un client entièrement facturé quitte alors le tableau du haut. Sans numéro, il refuse avec *« Il manque le numéro de facture à attribuer. »*


---

## Où se règle quoi

| Réglage | Module | Valeur |
|---|---|---|
| `FA_LARGEUR` | `modFact_Theme` | 960 |
| `FA_HAUTEUR` | `modFact_Theme` | 636 |
| `FA_Z2_LIGNES` | `modFact_Theme` | 9 |
| `FA_Z4_LIGNES` | `modFact_Theme` | 12 |
| `FA_LIGNE_H` | `modFact_Theme` | ? |
| `FA_PAD_X` | `modFact_Theme` | ? |
| `FA_BARRE_L` | `modFact_Theme` | ? |

Les largeurs de colonnes sont dans `modFact_Schema`, la typographie dans la section **TYPOGRAPHIE PAR ZONE** de `modFact_Theme`.

