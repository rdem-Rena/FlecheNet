# Fiche de réglages

Où se trouve chaque valeur du formulaire, et ce qu'elle commande.

Toutes les positions sont en **points** (1 pt = 1/72 de pouce ; à 96 ppp, 1 pt vaut 1,333 pixel). Après toute modification, relancer **`GenererFormulaireClients`** : le formulaire est reconstruit à partir de ces valeurs.

> Ce fichier est produit à partir des sources. Les valeurs ci-dessous sont celles réellement en vigueur dans `src/`.


---

## 1. Je veux changer…

| Ce que je veux faire | Où | Quoi |
|---|---|---|
| Agrandir ou rétrécir la fenêtre | `modClients_Theme` | `F_LARGEUR`, `F_HAUTEUR` |
| Changer une couleur | `modClients_Theme` | constantes `COUL_…` |
| Changer la police ou une taille de texte | `modClients_Theme` | `POLICE`, constantes `TAILLE_…` |
| Déplacer une zone entière (fiche, filtre, tableau, boutons) | `modClients_Theme` | `CS_TOP`, `CF_TOP`, `CT_TOP`, `BT_TOP` |
| Élargir les zones de saisie | `modClients_Theme` | `GR_BLOC`, `GR_GOUTTIERE` |
| Espacer les lignes de la grille | `modClients_Theme` | `GR_LIGNE`, `CH_LBL_HAUT`, `CH_CTL_HAUT` |
| Redimensionner les boutons | `modClients_Theme` | `BT_LARG`, `BT_HAUT`, `BT_GOUTTIERE` |
| Déplacer un champ dans la grille | `modClients_Schema` | colonnes *ligne* et *colonne* de sa ligne `DefChamp` |
| Renommer un libellé de champ | `modClients_Schema` | 2ᵉ argument de sa ligne `DefChamp` |
| Changer une info-bulle | `modClients_Schema` | dernier argument de sa ligne `DefChamp` |
| Verrouiller ou déverrouiller un champ | `modClients_Schema` | argument *verrouillé* de sa ligne `DefChamp` |
| Changer les colonnes du tableau | `modClients_Schema` | `ColonnesListe`, `LibellesListe`, `LargeursListe` |
| Changer les colonnes filtrables | `modClients_Schema` | `ChampsFiltrables` |
| Changer les civilités proposées | `modClients_Schema` | `TitresProposes` |
| Changer un texte de bouton ou son raccourci | `modClients_Generateur` | `ConstruireBoutons` |
| Changer le titre ou le sous-titre du bandeau | `modClients_Generateur` | `ConstruireBandeau` |
| Changer un message d'erreur ou de confirmation | `modClients_Formulaire` | la procédure `Clients_…` concernée |
| Changer les règles de contrôle de saisie | `modClients_Formulaire` | `Valider` |

---

## 2. Constantes de `modClients_Theme`

### Palette

| Constante | Valeur | S'applique à |
|---|---|---|
| `COUL_FOND` | `#F0F3F8` | fond de la fenêtre, autour des cartes |
| `COUL_CARTE` | `#FFFFFF` | fond des trois cartes blanches |
| `COUL_BORDURE` | `#E0E6EE` | filet qui entoure les cartes |
| `COUL_BANDEAU` | `#1B365D` | bandeau de titre |
| `COUL_BANDEAU_TXT` | `#FFFFFF` | titre du bandeau |
| `COUL_BANDEAU_SOUS` | `#96AFCD` | ligne d'état sous le titre |
| `COUL_TEXTE` | `#202D3E` | texte saisi, lignes du tableau |
| `COUL_TEXTE_DOUX` | `#7A8CA6` | libellés au-dessus des champs, compteur |
| `COUL_SECTION` | `#7A8CA6` | titre « FICHE CLIENT » |
| `COUL_CHAMP_FOND` | `#FCFDFF` | fond des zones de saisie |
| `COUL_CHAMP_BORD` | `#D6DEE9` | filet des zones de saisie au repos |
| `COUL_CHAMP_FOCUS` | `#1769B5` | filet de la zone de saisie active |
| `COUL_VERROU_FOND` | `#EEF2F7` | fond de Clef_BD et Date_Crea |
| `COUL_VERROU_TXT` | `#7C8899` | texte de Clef_BD et Date_Crea |
| `COUL_ENTETE_TBL` | `#E9EEF6` | bande d'en-têtes du tableau |
| `COUL_ENTETE_TXT` | `#44566C` | texte des en-têtes de colonnes |
| `COUL_AJOUTER` | `#008F65` | bouton Ajouter |
| `COUL_AJOUTER_H` | `#00A876` | bouton Ajouter, au survol |
| `COUL_MODIFIER` | `#1769B5` | bouton Modifier |
| `COUL_MODIFIER_H` | `#2A80D2` | bouton Modifier, au survol |
| `COUL_SUPPRIMER` | `#C73636` | bouton Supprimer |
| `COUL_SUPPRIMER_H` | `#DC4B4B` | bouton Supprimer, au survol |
| `COUL_EFFACER` | `#6C7A8C` | bouton Effacer |
| `COUL_EFFACER_H` | `#8494A8` | bouton Effacer, au survol |
| `COUL_QUITTER` | `#3A485A` | bouton Quitter |
| `COUL_QUITTER_H` | `#4E5F75` | bouton Quitter, au survol |
| `COUL_BOUTON_TXT` | `#FFFFFF` | texte des cinq boutons |
| `COUL_BOUTON_OFF` | `#C6CFD9` | Modifier et Supprimer, grisés |
| `COUL_FERMER_H` | `#E04B4B` | croix du bandeau, au survol |
| `COUL_LIEN` | `#1769B5` | lien Réinitialiser |
| `COUL_LIEN_H` | `#2A80D2` | lien Réinitialiser, au survol |

Les couleurs s'écrivent en hexadécimal VBA, c'est-à-dire `&HBBGGRR&` — bleu, vert, rouge, l'inverse de la notation web. Le commentaire en fin de ligne rappelle la valeur `#RRGGBB` correspondante ; c'est cette valeur-là qui est reprise dans le tableau ci-dessus.


### Typographie

| Constante | Valeur | S'applique à |
|---|---|---|
| `POLICE` | `Segoe UI` | tous les contrôles |
| `TAILLE_TITRE` | 12 pt | titre du bandeau |
| `TAILLE_SOUSTITRE` | 8 pt | ligne d'état du bandeau |
| `TAILLE_SECTION` | 7.5 pt | « FICHE CLIENT » |
| `TAILLE_LIBELLE` | 7.5 pt | libellés au-dessus des champs |
| `TAILLE_CHAMP` | 9.5 pt | texte saisi dans les champs |
| `TAILLE_BOUTON` | 9.5 pt | texte des boutons |
| `TAILLE_ENTETE` | 7.5 pt | en-têtes de colonnes du tableau |
| `TAILLE_LISTE` | 9 pt | lignes du tableau |
| `TAILLE_FILTRE` | 8.5 pt | barre de filtrage, cases à cocher |

### Géométrie

| Constante | Valeur | Commande |
|---|---|---|
| `F_LARGEUR` | 840 pt | largeur de la fenêtre |
| `F_HAUTEUR` | 532 pt | hauteur de la fenêtre |
| `MARGE` | 16 pt | marge gauche et droite des cartes |
| `CARTE_LARG` | 808 pt | largeur des cartes — garder `F_LARGEUR - 2 × MARGE` |
| `BAND_HAUT` | 48 pt | hauteur du bandeau de titre |
| `CS_TOP` | 58 pt | haut de la carte « fiche client » |
| `CS_HAUT` | 200 pt | hauteur de la carte « fiche client » |
| `GR_X` | 28 pt | abscisse de la 1ʳᵉ colonne de la grille |
| `GR_Y` | 86 pt | ordonnée de la 1ʳᵉ ligne de la grille |
| `GR_BLOC` | 250.5 pt | largeur d'un bloc « libellé + champ » |
| `GR_GOUTTIERE` | 16.5 pt | espace horizontal entre deux blocs |
| `GR_LIGNE` | 32 pt | pas vertical entre deux lignes de la grille |
| `CH_LBL_HAUT` | 11 pt | hauteur du libellé au-dessus d'un champ |
| `CH_CTL_HAUT` | 18 pt | hauteur d'une zone de saisie |
| `CF_TOP` | 266 pt | haut de la barre de filtrage |
| `CF_HAUT` | 40 pt | hauteur de la barre de filtrage |
| `CT_TOP` | 314 pt | haut de la carte « tableau » |
| `CT_HAUT` | 168 pt | hauteur de la carte « tableau » |
| `CT_ENTETE` | 22 pt | hauteur de la bande d'en-têtes du tableau |
| `BT_TOP` | 492 pt | haut de la rangée de boutons |
| `BT_HAUT` | 30 pt | hauteur des boutons |
| `BT_LARG` | 118 pt | largeur des boutons |
| `BT_GOUTTIERE` | 6 pt | espace entre deux boutons |

**Contrainte à respecter** : les quatre zones ne doivent pas se chevaucher. Avec les valeurs actuelles :

```
bandeau     0   ->  48  
fiche      58   ->  258    (CS_TOP -> CS_TOP + CS_HAUT)
filtre     266  ->  306    (CF_TOP -> CF_TOP + CF_HAUT)
tableau    314  ->  482    (CT_TOP -> CT_TOP + CT_HAUT)
boutons    492  ->  522    (BT_TOP -> BT_TOP + BT_HAUT)
fenêtre         532  pt de haut (F_HAUTEUR)
```

---

## 3. Position calculée de chaque contrôle

Les champs de la fiche ne sont positionnés nulle part en dur : leurs coordonnées sortent de `GrilleX(colonne)` et `GrilleY(ligne)`, à partir de la case indiquée dans `modClients_Schema`. Modifier `GR_X`, `GR_Y`, `GR_BLOC`, `GR_GOUTTIERE` ou `GR_LIGNE` les déplace tous d'un coup.

### Champs de la fiche — `ConstruireCarteSaisie`

| Contrôle | Type | Case | Gauche | Haut | Largeur | Hauteur |
|---|---|---|---|---|---|---|
| `txtClef_BD` | TextBox | B1 L1 r1 · 8 car. | 28 | 98 | 54 | 18 |
| `lblChamp_Clef_BD` | Label | B1 L1 r1 · 8 car. | 28 | 86 | 54 | 11 |
| `txtDate_Crea` | TextBox | B1 L1 r2 | 90 | 98 | 126.5 | 18 |
| `lblChamp_Date_Crea` | Label | B1 L1 r2 | 90 | 86 | 126.5 | 11 |
| `txtID_Cresus` | TextBox | B1 L1 r3 · 8 car. | 224.5 | 98 | 54 | 18 |
| `lblChamp_ID_Cresus` | Label | B1 L1 r3 · 8 car. | 224.5 | 86 | 54 | 11 |
| `txtEntreprise` | TextBox | B1 L2 r1 | 28 | 130 | 250.5 | 18 |
| `lblChamp_Entreprise` | Label | B1 L2 r1 | 28 | 118 | 250.5 | 11 |
| `cboTitre` | ComboBox | B1 L3 r1 | 28 | 162 | 250.5 | 18 |
| `lblChamp_Titre` | Label | B1 L3 r1 | 28 | 150 | 250.5 | 11 |
| `txtNom` | TextBox | B1 L4 r1 | 28 | 194 | 250.5 | 18 |
| `lblChamp_Nom` | Label | B1 L4 r1 | 28 | 182 | 250.5 | 11 |
| `txtPrenom` | TextBox | B1 L5 r1 | 28 | 226 | 250.5 | 18 |
| `lblChamp_Prenom` | Label | B1 L5 r1 | 28 | 214 | 250.5 | 11 |
| `cboAdresse` | ComboBox | B2 L1 r1 | 295 | 98 | 199.5 | 18 |
| `lblChamp_Adresse` | Label | B2 L1 r1 | 295 | 86 | 199.5 | 11 |
| `txtNo` | TextBox | B2 L1 r2 · 6 car. | 502.5 | 98 | 43 | 18 |
| `lblChamp_No` | Label | B2 L1 r2 · 6 car. | 502.5 | 86 | 43 | 11 |
| `cboNoPost` | ComboBox | B2 L2 r1 · 8 car. | 295 | 130 | 70 | 18 |
| `lblChamp_NoPost` | Label | B2 L2 r1 · 8 car. | 295 | 118 | 70 | 11 |
| `txtVille` | TextBox | B2 L2 r2 | 373 | 130 | 172.5 | 18 |
| `lblChamp_Ville` | Label | B2 L2 r2 | 373 | 118 | 172.5 | 11 |
| `txtCant` | TextBox | B2 L3 r1 | 295 | 162 | 250.5 | 18 |
| `lblChamp_Cant` | Label | B2 L3 r1 | 295 | 150 | 250.5 | 11 |
| `chkTVA` | CheckBox | B2 L5 r1 · 3 car. | 295 | 226 | 36.5 | 18 |
| `chkForfait` | CheckBox | B2 L5 r2 · 7 car. | 339.5 | 226 | 58.5 | 18 |
| `txtTx_hrs_Forf` | TextBox | B2 L5 r3 | 406 | 226 | 139.5 | 18 |
| `lblChamp_Tx_hrs_Forf` | Label | B2 L5 r3 | 406 | 214 | 139.5 | 11 |
| `txtNatel` | TextBox | B3 L1 r1 | 562 | 98 | 250.5 | 18 |
| `lblChamp_Natel` | Label | B3 L1 r1 | 562 | 86 | 250.5 | 11 |
| `txtTel_Pro` | TextBox | B3 L2 r1 | 562 | 130 | 250.5 | 18 |
| `lblChamp_Tel_Pro` | Label | B3 L2 r1 | 562 | 118 | 250.5 | 11 |
| `txtEmail` | TextBox | B3 L3 r1 | 562 | 162 | 250.5 | 18 |
| `lblChamp_Email` | Label | B3 L3 r1 | 562 | 150 | 250.5 | 11 |
| `cboTexte_Facture` | ComboBox | B3 L4 r1 | 562 | 194 | 250.5 | 18 |
| `lblChamp_Texte_Facture` | Label | B3 L4 r1 | 562 | 182 | 250.5 | 11 |
| `txtNote_Interne` | TextBox | B3 L5 r1 | 562 | 226 | 250.5 | 18 |
| `lblChamp_Note_Interne` | Label | B3 L5 r1 | 562 | 214 | 250.5 | 11 |

### Habillage et commandes — positions écrites dans le générateur

| Contrôle | Type | Procédure | Gauche | Haut | Largeur | Hauteur |
|---|---|---|---|---|---|---|
| `lblEntete` | Label | `ConstruireBandeau` | 0 | 0 | 840 | 48 |
| `lblTitre` | Label | `ConstruireBandeau` | 24 | 9 | 520 | 19 |
| `lblSousTitre` | Label | `ConstruireBandeau` | 24 | 28 | 620 | 13 |
| `lblFermer` | Label | `ConstruireBandeau` | 800 | 11 | 26 | 26 |
| `lblCarteSaisie` | Label | `ConstruireCarteSaisie` | 16 | 58 | 808 | 200 |
| `lblSectionSaisie` | Label | `ConstruireCarteSaisie` | 28 | 66 | 400 | 14 |
| `lblChamp_Pays` | Label | `ConstruireCarteSaisie` | 295 | 182 | 250.5 | 11 |
| `txtPays` | TextBox | `ConstruireCarteSaisie` | 295 | 194 | 250.5 | 18 |
| `lblChamp_Facturation` | Label | `ConstruireCarteSaisie` | — | — | — | 11 |
| `lblCarteFiltre` | Label | `ConstruireCarteFiltre` | 16 | 266 | 808 | 40 |
| `lblFiltreTitre` | Label | `ConstruireCarteFiltre` | 28 | 279 | 62 | 14 |
| `cboChampFiltre` | ComboBox | `ConstruireCarteFiltre` | 100 | 277 | 124 | 18 |
| `txtFiltre` | TextBox | `ConstruireCarteFiltre` | 232 | 277 | 284 | 18 |
| `lblResetFiltre` | Label | `ConstruireCarteFiltre` | 524 | 279 | 80 | 14 |
| `lblCompteur` | Label | `ConstruireCarteFiltre` | 612 | 279 | 200 | 14 |
| `lblCarteTable` | Label | `ConstruireCarteTableau` | 16 | 314 | 808 | 168 |
| `lblEnteteTable` | Label | `ConstruireCarteTableau` | 17 | 315 | 806 | 21 |
| `lstClients` | ListBox | `ConstruireCarteTableau` | 17 | 337 | 806 | 143 |
| `btnAjouter` | CommandButton | `ConstruireBoutons` | — | 492 | 118 | 30 |
| `btnModifier` | CommandButton | `ConstruireBoutons` | — | 492 | 118 | 30 |
| `btnSupprimer` | CommandButton | `ConstruireBoutons` | — | 492 | 118 | 30 |
| `btnEffacer` | CommandButton | `ConstruireBoutons` | — | 492 | 118 | 30 |
| `btnQuitter` | CommandButton | `ConstruireBoutons` | 706 | 492 | 118 | 30 |
| `btnAjouter` | CommandButton | `ConstruireBoutons` | 16 | 492 | 118 | 30 |
| `btnModifier` | CommandButton | `ConstruireBoutons` | 140 | 492 | 118 | 30 |
| `btnSupprimer` | CommandButton | `ConstruireBoutons` | 264 | 492 | 118 | 30 |
| `btnEffacer` | CommandButton | `ConstruireBoutons` | 388 | 492 | 118 | 30 |
| `btnQuitter` | CommandButton | `ConstruireBoutons` | 706 | 492 | 118 | 30 |
| `lblEnt_1` | Label | `ConstruireCarteTableau` | 20 | 319 | 115 | 13 |
| `lblEnt_2` | Label | `ConstruireCarteTableau` | 138 | 319 | 51 | 13 |
| `lblEnt_3` | Label | `ConstruireCarteTableau` | 192 | 319 | 99 | 13 |
| `lblEnt_4` | Label | `ConstruireCarteTableau` | 294 | 319 | 83 | 13 |
| `lblEnt_5` | Label | `ConstruireCarteTableau` | 380 | 319 | 141 | 13 |
| `lblEnt_6` | Label | `ConstruireCarteTableau` | 524 | 319 | 27 | 13 |
| `lblEnt_7` | Label | `ConstruireCarteTableau` | 554 | 319 | 39 | 13 |
| `lblEnt_8` | Label | `ConstruireCarteTableau` | 596 | 319 | 101 | 13 |
| `lblEnt_9` | Label | `ConstruireCarteTableau` | 700 | 319 | 33 | 13 |
| `lblEnt_10` | Label | `ConstruireCarteTableau` | 736 | 319 | 59 | 13 |
| `lstClients` | ListBox | `ConstruireCarteTableau` | 17 | 337 | 806 | 143 |

---

## 4. Colonnes du tableau des enregistrements

Trois fonctions de `modClients_Schema`, lues **position par position** : toute modification de l'une doit être reportée dans les deux autres.

| # | `ColonnesListe` (colonne Excel) | `LibellesListe` (en-tête affiché) | `LargeursListe` |
|---|---|---|---|
| 1 | `Entreprise` | Entreprise | 118 pt |
| 2 | `Titre` | Titre | 54 pt |
| 3 | `Nom` | Nom | 102 pt |
| 4 | `Prenom` | Prénom | 86 pt |
| 5 | `Adresse` | Adresse | 144 pt |
| 6 | `No` | No | 30 pt |
| 7 | `NoPost` | NPA | 42 pt |
| 8 | `Ville` | Ville | 104 pt |
| 9 | `Cant` | Cant. | 36 pt |
| 10 | `Tx_hrs_Forf` | Taux / forf. | 62 pt |
| | | **total** | **778 pt** |

Deux limites à respecter :

- **dix colonnes au maximum** — c'est ce qu'accepte une `ListBox` MSForms ;
- **somme des largeurs sous 790 pt** — la `ListBox` mesure 806 pt de large et la barre de défilement verticale en prend environ 16.

Les en-têtes cliquables se repositionnent tout seuls d'après `LargeursListe` : il n'y a rien d'autre à ajuster.


---

## 5. Champs de la fiche

Une ligne `DefChamp` par champ, dans `ConstruireSchema`. Les arguments, dans l'ordre :

```
DefChamp mChamps, n°, "Colonne_Excel", "Libellé", TYPE_x, verrouillé, _
         ligne, bloc, rang, caractères, NUM_x, "info-bulle"
```
| Argument | Valeurs possibles | Effet |
|---|---|---|
| n° | 1 à `NB_CHAMPS` | rang dans le schéma, et ordre de tabulation |
| Colonne_Excel | nom exact dans `TblClients` | fait le lien avec la donnée ; **sans accent ajouté** |
| Libellé | texte libre | affiché en majuscules au-dessus du champ |
| TYPE_x | `TYPE_TEXTE`, `TYPE_LISTE`, `TYPE_CASE` | zone de texte, menu déroulant, case à cocher |
| verrouillé | `True` / `False` | `True` = visible mais non saisissable, hors tabulation |
| ligne | 1 à `NB_LIGNES_FICHE` | ligne, à l'intérieur du bloc |
| bloc | 1 à `NB_BLOCS` | bloc, de gauche à droite |
| rang | 1, 2, 3 | place sur la ligne, de gauche à droite |
| caractères | 0 ou un nombre | largeur en caractères ; `0` = prend ce qui reste de la ligne |
| NUM_x | `NUM_NON`, `NUM_ENTIER`, `NUM_DECIMAL` | filtre les frappes non numériques |
| info-bulle | texte libre | s'affiche au survol |

### État actuel

| n° | Colonne | Libellé | Type | Verrouillé | Place | Largeur | Saisie |
|---|---|---|---|---|---|---|---|
| 1 | `Clef_BD` | Clef BD | TextBox | oui | B1 L1 r1 | 8 car. | libre |
| 2 | `Date_Crea` | Date de création | TextBox | oui | B1 L1 r2 | souple | libre |
| 3 | `ID_Cresus` | ID Crésus | TextBox | — | B1 L1 r3 | 8 car. | chiffres |
| 4 | `Entreprise` | Entreprise | TextBox | — | B1 L2 r1 | souple | libre |
| 5 | `Titre` | Titre | ComboBox | — | B1 L3 r1 | souple | libre |
| 6 | `Nom` | Nom | TextBox | — | B1 L4 r1 | souple | libre |
| 7 | `Prenom` | Prénom | TextBox | — | B1 L5 r1 | souple | libre |
| 8 | `Adresse` | Adresse (rue) | ComboBox | — | B2 L1 r1 | souple | libre |
| 9 | `No` | No | TextBox | — | B2 L1 r2 | 6 car. | libre |
| 10 | `NoPost` | NPA | ComboBox | — | B2 L2 r1 | 8 car. | libre |
| 11 | `Ville` | Ville | TextBox | — | B2 L2 r2 | souple | libre |
| 12 | `Cant` | Canton | TextBox | — | B2 L3 r1 | souple | libre |
| 13 | `TVA` | TVA | CheckBox | — | B2 L5 r1 | 3 car. | libre |
| 14 | `Forfait` | Forfait | CheckBox | — | B2 L5 r2 | 7 car. | libre |
| 15 | `Tx_hrs_Forf` | Taux horaire / forfait | TextBox | — | B2 L5 r3 | souple | chiffres + décimales |
| 16 | `Natel` | Natel | TextBox | — | B3 L1 r1 | souple | libre |
| 17 | `Tel_Pro` | Téléphone pro. | TextBox | — | B3 L2 r1 | souple | libre |
| 18 | `Email` | Courriel | TextBox | — | B3 L3 r1 | souple | libre |
| 19 | `Texte_Facture` | Texte de facture | ComboBox | — | B3 L4 r1 | souple | libre |
| 20 | `Note_Interne` | Note interne | TextBox | — | B3 L5 r1 | souple | libre |

---

## 6. Modes d'emploi

### Élargir la fenêtre de 60 points

Dans `modClients_Theme` : `F_LARGEUR` de 840 à 900 et `CARTE_LARG` de 808 à 868 (garder `F_LARGEUR - 2 × MARGE`). Les cartes, la barre de filtrage, le tableau et le bouton Quitter suivent automatiquement. Restent à ajuster à la main : `GR_BLOC` — passer de 250.5 à 265.5 donne 15 points de plus par colonne de saisie — et `LargeursListe`, dont la somme peut monter d'autant.

### Ajouter une 6ᵉ ligne de champs

1. `modClients_Schema` : `NB_CHAMPS` +1 et une ligne `DefChamp` en ligne 6.
2. `modClients_Theme` : ajouter `GR_LIGNE` (32 pt) à `CS_HAUT`, `CF_TOP`, `CT_TOP`, `BT_TOP` et `F_HAUTEUR`.
3. Relancer `GenererFormulaireClients`.

### Rendre le tableau plus haut de 3 lignes

Une ligne du tableau mesure environ 12,75 pt. Dans `modClients_Theme` : `CT_HAUT` de 168 à 206, puis `BT_TOP` de 492 à 530 et `F_HAUTEUR` de 532 à 570.

### Changer la couleur d'un bouton

Dans `modClients_Theme`, la constante de la couleur normale et celle du survol — par exemple `COUL_AJOUTER` et `COUL_AJOUTER_H`. Convertir la couleur web `#RRGGBB` en `&HBBGGRR&` : `#008F65` s'écrit `&H658F00&`.

### Déplacer un champ ailleurs dans la grille

Changer les arguments *ligne* et *colonne* de sa ligne `DefChamp`, et ceux du champ qui occupait la place. Aucune coordonnée n'est à calculer.

