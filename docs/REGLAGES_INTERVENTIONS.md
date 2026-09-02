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
| **La graisse des en-têtes** | `modInterv_Theme` | `POLICE_ENTETE` |
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
| `POLICE_ENTETE` | `modInterv_Theme` | Segoe UI Semibold | famille des en-têtes — c'est elle qui porte la demi-graisse |
| `TAILLE_GRILLE_TXT` | `modInterv_Theme` | 9 | taille du texte des lignes |
| `COUL_GRILLE_TXT` | `modInterv_Theme` | #4A5A6E | couleur du texte des lignes |
| `COUL_GRILLE_ZEBRE` | `modInterv_Theme` | #F1F5FA | fond d'une ligne sur deux |
| `IGR_LIGNE_H` | `modInterv_Theme` | 12,75 | hauteur d'une ligne |
| `IGR_NB_LIGNES` | `modInterv_Theme` | 17 | lignes affichées à la fois |
| `IGR_BARRE_L` | `modInterv_Theme` | 16 | largeur de la barre de défilement |
| `IGR_PAD_X` | `modInterv_Theme` | 4,5 | retrait du texte de chaque côté de sa colonne |
| `IT_TOP` | `modInterv_Theme` | 448 | haut du tableau |
| `IT_HAUT` | `modInterv_Theme` | 246 | hauteur totale du tableau |
| `IT_ENTETE` | `modInterv_Theme` | 22 | hauteur de la bande d'en-têtes |

Deux couleurs viennent de la palette commune, dans `modClients_Theme` : `COUL_MODIFIER` pour la ligne choisie et `COUL_ENTETE_TBL` pour la ligne survolée.

> **Les largeurs de colonnes sont libres.** Chaque coordonnée de la grille passe par `AuPixel`, qui l'arrondit au pixel : une case posée à cheval rendrait son texte décalé et plus épais. Seule la hauteur de ligne doit être choisie juste, parce qu'elle se répète dix-sept fois et que l'erreur s'accumulerait.

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

## 3. La typographie, zone par zone

Chaque rôle de libellé a ses quatre caractéristiques dans `modInterv_Theme`, section **TYPOGRAPHIE PAR ZONE**. Pour changer l'aspect d'une zone, tout est là : il n'y a pas à ouvrir le générateur.

| Zone | Fond | Rôle | Famille | Taille | Graisse | Couleur |
|---|---|---|---|---|---|---|
| 1 | `lblCarte1` | `Z1Titre` | `POLICE` | `14` | gras | `COUL_BANDEAU` |
|  |  | `Z1Etat` | `POLICE` | `TAILLE_SOUSTITRE` | maigre | `COUL_TEXTE_DOUX` |
|  |  | `Z1Annee` | `POLICE` | `20` | gras | `COUL_MODIFIER` |
| 2 | `lblCarte2` | `Z2Section` | `POLICE` | `TAILLE_SECTION` | gras | `COUL_SECTION` |
|  |  | `Z2TuileCap` | `POLICE` | `TAILLE_TUILE_CAP` | gras | `COUL_TEXTE_DOUX` |
|  |  | `Z2TuileVal` | `POLICE` | `TAILLE_STAT` | gras | `COUL_BANDEAU` |
|  |  | `Z2GraphTitre` | `POLICE` | `TAILLE_FILTRE` | maigre | `COUL_TEXTE_DOUX` |
|  |  | `Z2GraphAxe` | `POLICE` | `TAILLE_LIBELLE` | maigre | `COUL_TEXTE_DOUX` |
| 3 | `lblCarte3` | `Z3Section` | `POLICE` | `TAILLE_SECTION` | gras | `COUL_SECTION` |
|  |  | `Z3Libelle` | `POLICE` | `TAILLE_LIBELLE` | gras | `COUL_TEXTE_DOUX` |
|  |  | `Z3Chevron` | `POLICE` | `9` | gras | `COUL_BOUTON_TXT` |
| 4 | `lblCarteFiltreI` | `Z4Libelle` | `POLICE` | `TAILLE_FILTRE` | gras | `COUL_ENTETE_TXT` |
|  |  | `Z4Lien` | `POLICE` | `TAILLE_FILTRE` | maigre | `COUL_LIEN` |
|  |  | `Z4Compteur` | `POLICE` | `TAILLE_FILTRE` | maigre | `COUL_TEXTE_DOUX` |
| 5 | `lblEnteteTableI` | `Z5Entete` | `POLICE_ENTETE` | `TAILLE_GRILLE_ENTETE` | maigre | `COUL_GRILLE_ENTETE` |
| 6 | `lblG_1_1` … `lblG_17_12` | `Z6Case` | `POLICE` | `TAILLE_GRILLE_TXT` | `GRAS_GRILLE_TXT` | `COUL_GRILLE_TXT` |
| 7 | `UF_Calendrier` | `Z7Fleche` | `POLICE` | `15` | gras | `COUL_BANDEAU_SOUS` |
|  |  | `Z7Mois` | `POLICE` | `11` | gras | `COUL_BANDEAU_TXT` |
|  |  | `Z7JourSem` | `POLICE` | `TAILLE_LIBELLE` | gras | `COUL_TEXTE_DOUX` |
|  |  | `Z7Jour` | `POLICE` | `TAILLE_CHAMP` | maigre | `COUL_TEXTE` |
|  |  | `Z7Lien` | `POLICE` | `TAILLE_FILTRE` | maigre | `COUL_LIEN` |
|  |  | `Z7Annuler` | `POLICE` | `TAILLE_FILTRE` | maigre | `COUL_TEXTE_DOUX` |

Une taille écrite en clair ne vaut que pour cette zone. Une taille **nommée** (`TAILLE_LIBELLE`, `TAILLE_FILTRE`…) vient de `modClients_Theme` et est partagée avec le formulaire des clients : la changer déplace les deux. Pour n'en bouger qu'une, remplacer le nom par un nombre.

Les couleurs, elles, sont nommées à dessein : c'est la palette du classeur, et une teinte doit rester la même partout où elle veut dire la même chose.

> `simulate_interv.py` refuse tout appel qui coderait une taille ou une couleur en dur dans le générateur : c'est ce qui garantit que ce tableau reste la seule source.


---

## 4. Modes d'emploi

### Rendre le texte du tableau plus grand

Dans `modInterv_Theme`, `TAILLE_GRILLE_TXT`. Au-delà de 12,75 pt le texte ne tiendra plus dans une ligne de 12,75 pt : augmenter aussi `IGR_LIGNE_H` — en gardant un multiple de 0,75 — et réduire `IGR_NB_LIGNES` d'autant.


### Ajouter une colonne au tableau

Dans `modInterv_Schema`, ajouter **la même position** aux quatre tableaux : le nom de la colonne dans `IColonnesListe`, son titre dans `ILibellesListe`, sa largeur dans `ILargeursListe`, son alignement dans `IAlignementsListe`. Prendre la largeur sur une autre colonne : la somme ne doit pas dépasser 910 pt.

Deux colonnes peuvent partager une case : `IC_NOM & ICL_SEPARATEUR & IC_PRENOM` affiche « Aiello Rosalba » dans une seule. Chacune garde sa colonne dans `TblInterv`.


### Afficher plus de lignes

Augmenter `IGR_NB_LIGNES`, puis ajouter autant de fois `IGR_LIGNE_H` à `IT_HAUT`, `IB_TOP` et `I_HAUTEUR` — sans quoi la grille passerait sous les boutons.


### Changer la graisse des en-têtes

MSForms ne connaît que **gras ou pas gras** : la demi-graisse se nomme, elle ne se règle pas. `POLICE_ENTETE` vaut `Segoe UI Semibold`. Pour un en-tête franchement gras, y mettre `Segoe UI` et passer les libellés en gras dans `ConstruireFiche4` ; pour un en-tête maigre, `Segoe UI Light`.


### Changer une couleur

Les constantes `COUL_…` s'écrivent en **BGR**, pas en RVB : `&HBBVVRR&`. Le bleu #1769B5 s'écrit donc `&HB56917&`. Le commentaire de fin de ligne rappelle la valeur RVB.


### Vérifier avant de régénérer

Les scripts de `scratchpad/` rejouent la construction du formulaire sans Excel et signalent chevauchements, débordements et incohérences : `check_vba.py`, `simulate_interv.py`.

