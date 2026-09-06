# Structure du formulaire

## Plan général

Formulaire sans bordure système, **840 × 532 points** (≈ 1120 × 709 px), centré
sur la fenêtre Excel.

```
┌────────────────────────────────────────────────────────────────────┐
│  Gestion des clients                                           ×   │  bandeau 48 pt
│  Tableau TblClients — feuille Clients                              │
├────────────────────────────────────────────────────────────────────┤
│ ┌────────────────────────────────────────────────────────────────┐ │
│ │ FICHE CLIENT                                                   │ │
│ │ Clef BD│Date créat.│ID Cr.│ Adresse           │No│ Natel        │ │  carte de
│ │ Entreprise               │ NPA │ Ville        │ Téléphone pro.  │ │  saisie
│ │ Titre                    │ Canton             │ Courriel        │ │  200 pt
│ │ Nom                      │ Pays  ( Suisse )   │ Texte facture   │ │
│ │ Prénom                   │ ☐TVA ☐Forf.│Taux   │ Note interne    │ │
│ └────────────────────────────────────────────────────────────────┘ │
│ ┌────────────────────────────────────────────────────────────────┐ │
│ │ Filtrer sur [Nom ▾] [__________]  Réinitialiser   n/N fiches   │ │  filtre 40 pt
│ └────────────────────────────────────────────────────────────────┘ │
│ ┌────────────────────────────────────────────────────────────────┐ │
│ │Entreprise│Titre│Nom│Prénom│Adresse│No│NPA│Ville│Canton│Taux/forf│ │  tableau
│ │ ───────────────────────────────────────────────────────────────│ │  168 pt
│ │          │Mme  │Aebi│Sylvie│Rue des Vieux Patriotes│64│2300│…   │ │
│ └────────────────────────────────────────────────────────────────┘ │
│  [ Ajouter ] [ Modifier ] [ Supprimer ] [ Effacer ]     [ Quitter ]│  boutons 30 pt
└────────────────────────────────────────────────────────────────────┘
```

## Colonnes de `TblClients` et contrôles correspondants

La fiche compte **trois blocs de cinq lignes**. Une ligne porte un à trois
champs : `bloc / ligne / rang`. Ceux qui annoncent un nombre de **caractères**
prennent la largeur qu'il faut, les autres se partagent ce qui reste.

| # | Colonne | Contrôle | Type | Bloc / ligne / rang | Largeur | Particularité |
|---|---|---|---|---|---|---|
| 1 | `Clef_BD` | `txtClef_BD` | TextBox | 1 / 1 / 1 | 8 car. | verrouillé — `CL` + n° suivant |
| 2 | `Date_Crea` | `txtDate_Crea` | TextBox | 1 / 1 / 2 | souple | verrouillé — date du jour à l'ajout |
| 3 | `ID_Cresus` | `txtID_Cresus` | TextBox | 1 / 1 / 3 | 8 car. | entier ; doublon signalé |
| 4 | `Entreprise` | `txtEntreprise` | TextBox | 1 / 2 / 1 | souple | filtrable |
| 5 | `Titre` | `cboTitre` | ComboBox | 1 / 3 / 1 | souple | liste fermée : Monsieur, Madame (+ existants) |
| 6 | `Nom` | `txtNom` | TextBox | 1 / 4 / 1 | souple | filtrable |
| 7 | `Prenom` | `txtPrenom` | TextBox | 1 / 5 / 1 | souple | |
| 8 | `Adresse` | `cboAdresse` | ComboBox | 2 / 1 / 1 | souple | liste = `Tabl_Adresses`; filtrable |
| 9 | `No` | `txtNo` | TextBox | 2 / 1 / 2 | 6 car. | texte (ex. `34B`) |
| 10 | `NoPost` | `cboNoPost` | ComboBox | 2 / 2 / 1 | 8 car. | 4 chiffres ; renseigne Ville et Cant |
| 11 | `Ville` | `txtVille` | TextBox | 2 / 2 / 2 | souple | rempli automatiquement |
| 12 | `Cant` | `txtCant` | TextBox | 2 / 3 / 1 | souple | rempli automatiquement |
| — | *(aucune)* | `txtPays` | TextBox | 2 / 4 / 1 | souple | verrouillé, figé à « Suisse » — **hors schéma** |
| 13 | `TVA` | `chkTVA` | CheckBox | 2 / 5 / 1 | 3 car. | booléen |
| 14 | `Forfait` | `chkForfait` | CheckBox | 2 / 5 / 2 | 7 car. | booléen |
| 15 | `Tx_hrs_Forf` | `txtTx_hrs_Forf` | TextBox | 2 / 5 / 3 | souple | décimal, affiché à 2 décimales |
| 16 | `Natel` | `txtNatel` | TextBox | 3 / 1 / 1 | souple | |
| 17 | `Tel_Pro` | `txtTel_Pro` | TextBox | 3 / 2 / 1 | souple | |
| 18 | `Email` | `txtEmail` | TextBox | 3 / 3 / 1 | souple | format contrôlé |
| 19 | `Texte_Facture` | `cboTexte_Facture` | ComboBox | 3 / 4 / 1 | souple | liste = `TblTxtStd` |
| 20 | `Note_Interne` | `txtNote_Interne` | TextBox | 3 / 5 / 1 | souple | |

**`Tel_Prive` n'est plus dans la fiche.** La colonne reste dans `TblClients` —
rien n'y est effacé — mais elle n'est ni affichée ni saisissable. La génération
la signale, avec les autres colonnes non couvertes.

Un champ dimensionné en caractères l'est aussi **à la saisie** : une case de
huit caractères n'en accepte pas neuf.

Le nom du contrôle est toujours `txt` / `cbo` / `chk` suivi du nom exact de la
colonne : c'est ce qui permet à `modClients_Formulaire` de faire la
correspondance sans table de conversion.

## Types écrits dans les cellules

La couche `modClients_Donnees` respecte les types déjà présents dans le tableau :

- `Clef_BD`, `NoPost`, `No`, `Cant`, `Ville` … → **texte** ;
- `Date_Crea` → **date** (format `[$-100C]d mmm yy` de la colonne) ;
- `ID_Cresus`, `Tx_hrs_Forf` → **nombre** ;
- `TVA`, `Forfait` → **booléen** (`VRAI` / `FAUX`) ;
- un champ laissé vide écrit une **cellule vide**, pas une chaîne vide.

L'écriture se fait ligne entière en un seul accès à la feuille, via
`ListRows.Add` / `ListRows(i).Range`, ce qui préserve les formats et les
formules de la colonne.

## Tableau des enregistrements

10 colonnes — le maximum d'une `ListBox` MSForms : `Entreprise`, `Titre`,
`Nom`, `Prenom`, `Adresse`, `No`, `NoPost`, `Ville`, `Cant`, `Tx_hrs_Forf`.

`Clef_BD` et `Date_Crea` n'y figurent volontairement pas : elles sont gérées par
le programme et restent lisibles dans la fiche du haut. La sélection ne s'appuie
d'ailleurs pas sur ce qui est affiché — chaque ligne du tableau garde en mémoire
le numéro de la fiche correspondante — si bien que les colonnes affichées
peuvent être changées librement.

Les autres champs (téléphones, courriel, notes, texte de facture) restent eux
aussi visibles dans la fiche dès qu'une ligne est sélectionnée.

Un clic sur un en-tête trie sur cette colonne ; un second clic inverse le sens
(▲ / ▼). Les fiches sans valeur sont renvoyées en fin de liste. Les dates sont
comparées chronologiquement, les nombres numériquement, le reste
alphabétiquement en ignorant accents et casse.

## Personnalisation

Toutes les modifications ci-dessous ne demandent qu'un lancement de
`GenererFormulaireClients` pour prendre effet.

**Changer les couleurs ou la police** → constantes de `modClients_Theme`
(les couleurs sont notées `&HBBGGRR&`, avec l'équivalent `#RRGGBB` en
commentaire).

**Redimensionner le formulaire** → `F_LARGEUR`, `F_HAUTEUR` et les constantes de
géométrie du même module. Les cartes, la grille et les boutons se placent à
partir de ces valeurs.

**Ajouter un champ** →
1. ajouter la colonne au tableau `TblClients` ;
2. dans `modClients_Schema` : incrémenter `NB_CHAMPS` et ajouter une ligne
   `DefChamp mChamps, 22, "Ma_Colonne", "Mon libellé", TYPE_TEXTE, False, 6, 1, 0, NUM_NON, "info-bulle"` ;
3. augmenter `CS_HAUT` et décaler `CF_TOP`, `CT_TOP`, `BT_TOP`, `F_HAUTEUR` de
   `GR_LIGNE` (32 pt) si une ligne de grille est ajoutée ;
4. relancer `GenererFormulaireClients`.

**Changer les champs proposés au filtrage** → `ChampsFiltrables` dans
`modClients_Schema`.

**Changer les civilités proposées** → `TitresProposes` dans le même module ; les
valeurs déjà présentes dans la colonne `Titre` s'ajoutent automatiquement.

**Changer les colonnes du tableau** → `ColonnesListe`, `LibellesListe` et
`LargeursListe` (10 entrées au maximum, largeurs en points ; leur somme doit
rester sous 790 pt pour laisser la place à la barre de défilement).

La liste complète des constantes et la position calculée de chacun des
71 contrôles sont réunies dans [`REGLAGES.md`](REGLAGES.md).

## Choix d'implémentation

**Pourquoi générer le formulaire par code ?** Un UserForm VBA stocke tous ses
contrôles dans un fichier binaire `.frx` : il n'est pas représentable en texte et
ne peut donc pas être versionné ni transmis comme du code source. Le générateur
résout cela — le formulaire est entièrement décrit par des procédures lisibles,
et se reconstruit à l'identique sur n'importe quel poste.

**Pourquoi un module de code de formulaire aussi mince ?** Chaque procédure
événementielle se contente d'appeler `modClients_Formulaire`, qui reçoit le
formulaire en `Object`. Le projet compile donc avant que le formulaire existe, et
régénérer le formulaire n'écrase jamais de logique métier.

**Rue présente dans plusieurs localités** — `Tabl_Adresses` contient des noms de
rue partagés par plusieurs communes. Si un NPA est déjà saisi et correspond à
l'une d'elles, c'est celui-là qui est conservé ; sinon la première correspondance
est utilisée. Le NPA, la ville et le canton restent modifiables à la main.
