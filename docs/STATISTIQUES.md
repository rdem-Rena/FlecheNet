# Formulaire des statistiques

`UF_Statistiques` — un tableau de bord sur **`TblInterv`** : cinq tuiles, un graphique mensuel, une barre d'objectif et la liste des travaux, tous liés aux mêmes filtres.

Il s'ouvre par le bouton **Info** du formulaire des interventions, et se génère par **`GenererFormulaireStatistiques`** (module `modStat_Generateur`).

> Ce fichier est produit à partir des sources : les valeurs sont celles réellement en vigueur dans `src/`.


---

## Plan

Fenêtre de **960 × 680 points** de surface utile. Cinq zones, chacune portée par un vrai cadre, et une rangée de boutons.

| Zone | Cadre | Haut | Hauteur | Contenu |
|---|---|---|---|---|
| 1 | `fraSCarte1` | 12 | 46 | titre et année |
| 2 | `fraSCarte2` | 66 | 68 | les cinq tuiles, sur une seule ligne |
| 3 | `fraSCarte3` | 142 | 132 | graphique du CA mensuel, barre d'objectif à droite |
| 4 | `fraSCarte4` | 282 | 40 | les sept filtres |
| 5 | `fraSCarte5` | 330 | 300 | tableau « Liste travaux – Clients » |

### Les cinq tuiles

Toutes portent sur les lignes **affichées** : elles suivent donc les filtres.

| Tuile | Ce qu'elle totalise |
|---|---|
| Chiffre d'affaires | somme de `CA` |
| Nbr d'heures | somme de `Nb_Hres` |
| Nb d'interventions | nombre de lignes retenues |
| Facturées | somme de `CA` des lignes qui ont un `No_Facture` |
| A facturer | somme de `CA` des lignes qui n'en ont pas |

### Le tableau

20 lignes visibles, défilement virtuel.

| # | Colonne de `TblInterv` | Titre | Largeur |
|---|---|---|---|
| 1 | `Date` | Date | 64 pt |
| 2 | `Client_No` | N° client | 60 pt |
| 3 | `Entreprise` | Entreprise | 160 pt |
| 4 | `Titre` | Titre | 60 pt |
| 5 | `Nom` | Nom | 120 pt |
| 6 | `Prenom` | Prénom | 100 pt |
| 7 | `Nb_Hres` | Heures | 52 pt |
| 8 | `Nb_Pers` | Pers. | 38 pt |
| 9 | `Taux/Forfait` | Taux/Forf. | 62 pt |
| 10 | `CA` | CA | 62 pt |
| 11 | `TVA` | TVA | 34 pt |
| 12 | `Forfait` | Forf. | 38 pt |
| 13 | `No_Facture` | Fact. | 50 pt |
| | | | **900 pt** |

---

## Fonctionnement

- **Un seul jeu de lignes alimente tout.** Les filtres produisent une liste, et les tuiles, le graphique, la barre d'objectif et le tableau s'y rapportent tous : aucun ne peut contredire les autres.

- Le **graphique** se calcule à partir des lignes retenues, contrairement à celui du formulaire des interventions, qui lit une feuille de totaux déjà faite.

- La **barre d'objectif** compare le CA affiché à la cellule nommée **`?`**. Si elle manque, la barre reste vide et le dit.

- **TVA**, **Forfait** et **Facturé** sont des cases à **trois états** : grisée elle ne filtre pas, cochée elle ne garde que les lignes vraies, décochée que les fausses.

- **Entreprise**, **Nom** et **N° facture** cherchent une sous-chaîne, sans distinction de casse : taper `aeb` retrouve Aebi.

- **CA est calculé, jamais lu** : c'est `Interv_EstimerCA`, le même calcul que sur les deux autres formulaires.


---

## Où se règle quoi

| Réglage | Module | Valeur |
|---|---|---|
| `ST_LARGEUR` | `modStat_Theme` | 960 |
| `ST_HAUTEUR` | `modStat_Theme` | 680 |
| `ST_Z5_LIGNES` | `modStat_Theme` | 20 |
| `ST_NB_TUILES` | `modStat_Theme` | 5 |
| `ST_OBJ_HAUT` | `modStat_Theme` | 84 |

Les largeurs de colonnes sont dans `modStat_Schema`. La typographie et la géométrie du tableau sont **déléguées** au formulaire des interventions : `ZSEntete`, `ZSCase`, `IGR_LIGNE_H`, `IGR_PAD_X`. Changer le tableau des interventions change celui-ci.

