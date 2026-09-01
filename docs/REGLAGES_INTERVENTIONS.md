# Réglages du formulaire des interventions

Où se trouve chaque valeur, et ce qu'elle commande.

Toutes les dimensions sont en **points** (1 pt = 1/72 de pouce ; à 96 ppp, 1 pt vaut 1,333 pixel). Après toute modification, relancer **`GenererFormulaireInterventions`** : les deux formulaires sont reconstruits à partir de ces valeurs.

> Ce fichier est produit à partir des sources. Les valeurs ci-dessous sont celles réellement en vigueur dans `src/`.


---

## 1. Je veux changer…

| Ce que je veux faire | Fichier | Quoi |
|---|---|---|
| **La taille du texte du tableau** | `modInterv_Theme` | `TAILLE_GRILLE_TXT` |
| **La taille du texte des en-têtes** | `modInterv_Theme` | `TAILLE_GRILLE_ENTETE` |
| **La couleur du texte du tableau** | `modInterv_Theme` | `COUL_GRILLE_TXT` |
| **La couleur des en-têtes** | `modInterv_Theme` | `COUL_GRILLE_ENTETE` |
| **La teinte d'une ligne sur deux** | `modInterv_Theme` | `COUL_GRILLE_ZEBRE` |
| La hauteur d'une ligne du tableau | `modInterv_Theme` | `IGR_LIGNE_H` — **multiple de 0,75 obligatoire** |
| Le nombre de lignes affichées | `modInterv_Theme` | `IGR_NB_LIGNES`, plus `IT_HAUT`, `IB_TOP` et `I_HAUTEUR` |
| La largeur de la barre de défilement | `modInterv_Theme` | `IGR_BARRE_L` |
| Le retrait du texte dans une case | `modInterv_Theme` | `IGR_PAD_X` |
| La largeur d'une colonne du tableau | `modInterv_Schema` | `ILargeursListe` |
| Le titre d'une colonne du tableau | `modInterv_Schema` | `ILibellesListe` |
| L'alignement d'une colonne | `modInterv_Schema` | `IAlignementsListe` |
| Quelles colonnes le tableau montre | `modInterv_Schema` | `IColonnesListe` |
| Agrandir ou rétrécir la fenêtre | `modInterv_Theme` | `I_LARGEUR`, `I_HAUTEUR` |
| Déplacer une fiche entière | `modInterv_Theme` | `F1_TOP`, `F2_TOP`, `F3_TOP`, `IF_TOP`, `IT_TOP` |
| La disposition de la fiche de saisie | `modInterv_Theme` | `IG_BLOC`, `IG_GOUTTIERE`, `IG_ECART_REGION`, `IG_LIGNE` |
| Déplacer un champ dans la fiche | `modInterv_Schema` | colonnes *ligne*, *colonne*, *blocs*, *lignes* de sa ligne `DefInterv` |
| Renommer un libellé de champ | `modInterv_Schema` | 2ᵉ texte de sa ligne `DefInterv` |
| Verrouiller ou déverrouiller un champ | `modInterv_Schema` | argument *verrouillé* de sa ligne `DefInterv` |
| Quelles colonnes le formulaire n'écrit jamais | `modInterv_Schema` | `IColonnesNonEcrites` |
| Les champs repris du client choisi | `modInterv_Schema` | `IReportsDepuisClients` |
| Les colonnes sur lesquelles filtrer | `modInterv_Schema` | `IChampsFiltrables` |
| La disposition de la barre de filtrage | `modInterv_Theme` | constantes `IFB_…` |
| Les couleurs et tailles des boutons | `modInterv_Theme` | `IB_…`, `ICouleurBouton` |
| Le graphique du chiffre d'affaires | `modInterv_Theme` | constantes `GR_…` |
| Le sélecteur de date | `modInterv_Theme` | constantes `CAL_…` |
| La palette et la police communes aux deux formulaires | `modClients_Theme` | `POLICE`, constantes `COUL_…` |

---

## 2. Le tableau, en détail

C'est une **grille de libellés** : une case par cellule. Tout s'y règle par constantes.

| Constante | Fichier | Valeur | Ce qu'elle commande |
|---|---|---|---|
| `TAILLE_GRILLE_ENTETE` | `modInterv_Theme` | 9,5 | taille du texte des en-têtes |
| `COUL_GRILLE_ENTETE` | `modInterv_Theme` | #2C3E52 | couleur du texte des en-têtes |
| `TAILLE_GRILLE_TXT` | `modInterv_Theme` | 8,5 | taille du texte des lignes |
| `COUL_GRILLE_TXT` | `modInterv_Theme` | #4A5A6E | couleur du texte des lignes |
| `COUL_GRILLE_ZEBRE` | `modInterv_Theme` | #F1F5FA | fond d'une ligne sur deux |
| `IGR_LIGNE_H` | `modInterv_Theme` | 12,75 | hauteur d'une ligne |
| `IGR_NB_LIGNES` | `modInterv_Theme` | 17 | lignes affichées à la fois |
| `IGR_BARRE_L` | `modInterv_Theme` | 16 | largeur de la barre de défilement |
| `IGR_PAD_X` | `modInterv_Theme` | 4 | retrait du texte de chaque côté de sa colonne |
| `IT_TOP` | `modInterv_Theme` | 448 | haut du tableau |
| `IT_HAUT` | `modInterv_Theme` | 246 | hauteur totale du tableau |
| `IT_ENTETE` | `modInterv_Theme` | 22 | hauteur de la bande d'en-têtes |

Deux couleurs viennent de la palette commune, dans `modClients_Theme` : `COUL_MODIFIER` pour la ligne choisie et `COUL_ENTETE_TBL` pour la ligne survolée.

> **La hauteur de ligne doit valoir un nombre entier de pixels.** À 96 ppp, 1 pixel vaut 0,75 point : `IGR_LIGNE_H` doit donc être un multiple de 0,75 — 12 ; 12,75 ; 13,5 ; 14,25… Une valeur intermédiaire fait arrondir les positions différemment d'une ligne à l'autre et une ligne sur trois se décale d'un pixel. `simulate_interv.py` le vérifie.


### Colonnes affichées

Les quatre tableaux se lisent **position par position** et doivent rester de même longueur. Ils sont tous dans `modInterv_Schema`.

| # | `IColonnesListe` | `ILibellesListe` | `ILargeursListe` | `IAlignementsListe` |
|---|---|---|---|---|
| 1 | `Date` | Date | 64 pt | gauche |
| 2 | `Client_No` | N° client | 56 pt | gauche |
| 3 | `Entreprise` | Entreprise | 132 pt | gauche |
| 4 | `Nom` | Nom | 86 pt | gauche |
| 5 | `Prenom` | Prénom | 76 pt | gauche |
| 6 | `Nb_Hres` | Heures | 52 pt | droite |
| 7 | `Nb_Pers` | Pers. | 38 pt | centre |
| 8 | `Taux/Forfait` | Taux/Forf. | 62 pt | droite |
| 9 | `CA` | CA | 54 pt | droite |
| 10 | `Texte_Facture` | Texte de facture | 122 pt | gauche |
| 11 | `Commentaires` | Commentaires | 124 pt | gauche |
| 12 | `No_Facture` | Fact. | 40 pt | gauche |
| | | | **906 pt** | |

La somme des largeurs plus `IGR_BARRE_L` (16 pt) doit tenir dans **926 pt**, soit `I_CARTE_LARG` moins 2.


---

## 3. Modes d'emploi

### Rendre le texte du tableau plus grand

Dans `modInterv_Theme`, `TAILLE_GRILLE_TXT`. Au-delà de 12,75 pt le texte ne tiendra plus dans une ligne de 12,75 pt : augmenter aussi `IGR_LIGNE_H` — en gardant un multiple de 0,75 — et réduire `IGR_NB_LIGNES` d'autant.


### Ajouter une colonne au tableau

Dans `modInterv_Schema`, ajouter **la même position** aux quatre tableaux : le nom de la colonne dans `IColonnesListe`, son titre dans `ILibellesListe`, sa largeur dans `ILargeursListe`, son alignement dans `IAlignementsListe`. Prendre la largeur sur une autre colonne : la somme ne doit pas dépasser 910 pt.

Deux colonnes peuvent partager une case : `IC_NOM & ICL_SEPARATEUR & IC_PRENOM` affiche « Aiello Rosalba » dans une seule. Chacune garde sa colonne dans `TblInterv`.


### Afficher plus de lignes

Augmenter `IGR_NB_LIGNES`, puis ajouter autant de fois `IGR_LIGNE_H` à `IT_HAUT`, `IB_TOP` et `I_HAUTEUR` — sans quoi la grille passerait sous les boutons.


### Changer une couleur

Les constantes `COUL_…` s'écrivent en **BGR**, pas en RVB : `&HBBVVRR&`. Le bleu #1769B5 s'écrit donc `&HB56917&`. Le commentaire de fin de ligne rappelle la valeur RVB.


### Vérifier avant de régénérer

Les scripts de `scratchpad/` rejouent la construction du formulaire sans Excel et signalent chevauchements, débordements et incohérences : `check_vba.py`, `simulate_interv.py`.

