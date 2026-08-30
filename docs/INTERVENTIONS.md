# Formulaire des interventions

`UF_Interventions` — saisie et suivi du tableau **`TblInterv`**, onglet **Interventions**, avec son sélecteur de date `UF_Calendrier`.

Les deux formulaires sont générés par **`GenererFormulaireInterventions`** (module `modInterv_Generateur`), et affichés par **`OuvrirGestionInterventions`**.

> Ce fichier est produit à partir des sources : les valeurs sont celles réellement en vigueur dans `src/`.


---

## Plan du formulaire

Fenêtre de **960 × 748 points** de surface utile (≈ 1280 × 997 pixels à 96 ppp), barre de titre Windows conservée. Quatre fiches empilées, une barre de filtrage, une rangée de boutons.

```
┌──────────────────────────────────────────────────────────────────────┐
│ Travaux effectués                                              2026  │  fiche 1
│ Intervention IN12 sélectionnée — modifiez puis cliquez sur Modifier. │  46 pt
├──────────────────────────────────────────────────────────────────────┤
│ STATISTIQUES                                                         │
│ ┌─────────────────────┐  ┌────────┐ ┌────────┐ ┌────────┐            │  fiche 2
│ │   CAGraphique       │  │CA total│ │Facturé │ │Non fact│            │  152 pt
│ │   (image exportée)  │  ├────────┤ ├────────┤ ├────────┤            │
│ └─────────────────────┘  │Mois -1 │ │Mois    │ │Mois nf │            │
├──────────────────────────────────────────────────────────────────────┤
│ FICHE INTERVENTION                                                   │
│  N° interv. │ Date      │ N° client │ Entreprise                     │  fiche 3
│  Titre      │ Nom       │ Prénom    │ Texte de facture               │  168 pt
│  Heures     │ Personnes │ ☐TVA ☐Forf│ Chiffre d'affaires             │
│  N° facture │ Commentaires ........................................  │
├──────────────────────────────────────────────────────────────────────┤
│ Filtrer sur [Entreprise ▾] [__________]  Réinit.   n/N interventions  │  38 pt
├──────────────────────────────────────────────────────────────────────┤
│ Date │N° cl.│Entreprise│Nom│Prénom│Heures│Pers.│Texte│Comment.│Fact. │  fiche 4
│ ─────────────────────────────────────────────────────────────────────│  246 pt
│ 06/05/2026 1135  …                                                   │
├──────────────────────────────────────────────────────────────────────┤
│ [Ajouter][Modifier][Supprimer][Effacer]  [Facturer][Info]  [Quitter] │  30 pt
└──────────────────────────────────────────────────────────────────────┘
```

---

## Fiche 1 — intitulé

| Élément | Source |
|---|---|
| Titre, à gauche | cellule nommée `TitreInterventions` (feuille Parametres) |
| Année, à droite | cellule nommée `AnneeEnCours`, affichée au format `aaaa` |
| Ligne d'état | message du programme : nouvelle fiche, fiche sélectionnée, résultat de la dernière opération |

La cellule `AnneeEnCours` contient une **date** (1ᵉʳ janvier de l'année) et non un nombre. Les deux cas sont traités : une valeur à quatre chiffres est prise pour une année, au-delà c'est une date dont seule l'année est affichée.


---

## Fiche 2 — statistiques

Le graphique **`CAGraphique`** de la feuille Statistiques est **exporté en GIF** dans le dossier temporaire de Windows à chaque ouverture, puis chargé dans un contrôle Image : MSForms ne sait pas afficher un graphique Excel autrement. L'export est refait après chaque ajout, modification ou suppression, pour que le graphique suive les données.

> **Pourquoi GIF et non PNG.** `LoadPicture` vient de la bibliothèque OLE, qui ne lit que bmp, ico, wmf, emf, gif et jpg. Un PNG s'exporte sans erreur mais reste ensuite illisible — le cadre du graphique restait vide, sans le moindre message. Le graphique est cherché d'abord par le nom de l'objet graphique, puis par celui de la forme, puis, s'il n'y en a qu'un sur la feuille, par défaut. Si rien n'aboutit, la raison s'affiche à la suite du titre « STATISTIQUES » plutôt que d'être passée sous silence.

Les six tuiles lisent des cellules nommées :

| Tuile | Cellule Excel |
|---|---|
| CA total | `CATotal` |
| Total facturé | `CATotalFacture` |
| Total non facturé | `CATotalNonFacture` |
| Mois précédent | `CATotalMmoisPrecedent` |
| Mois actuel | `CATotalMmoisActuel` |
| Mois act. non facturé | `CATotalMoisActuelNonFacture` |

> **Un nom corrigé.** La demande indiquait `CATotalMmoisActuel` pour « CA mois actuel non facturé », nom déjà employé par « CA mois actuel ». La cellule réellement définie dans le classeur est **`CATotalMoisActuelNonFacture`** (Statistiques!$C$24) : c'est celle-là qui est lue.

Chaque tuile est un contrôle Image portant **`CartePremium.jpg`** (sous-dossier `Images`, à côté du classeur), surmonté de deux libellés à fond blanc — l'image donne les coins arrondis, les libellés portent le texte. Si l'image est absente, la tuile retombe sur un aplat blanc et reste lisible ; `VerifierClasseurInterventions` le signale.

L'image est étirée aux dimensions de la tuile (150 × 56 pt). Une source dans ces proportions évite de déformer l'arrondi des coins.


---

## Fiche 3 — les quinze champs

| # | Colonne | Libellé | Contrôle | Case | Saisie | Particularité |
|---|---|---|---|---|---|---|
| 1 | `NoInterv` | N° intervention | `txtNoInterv` | L1 C1 | libre | verrouillé — `IN` + plus grand numéro + 1, à l'ajout |
| 2 | `Date` | Date | `txtDate` | L1 C2 | libre | date du jour par défaut ; le bouton ▾ ouvre le calendrier |
| 3 | `Client_No` | N° client (Crésus) | `txtClient_No` | L1 C3 | libre | verrouillé — repris de `ID_Cresus` du client choisi |
| 4 | `Entreprise` | Entreprise | `cboEntreprise` | L1 C4 | libre | liste filtrée au fil de la frappe sur `TblClients[Entreprise]` |
| 5 | `Titre` | Titre | `cboTitre` | L2 C1 | libre | civilités présentes dans `TblClients` |
| 6 | `Nom` | Nom | `cboNom` | L2 C2 | libre | liste filtrée au fil de la frappe sur `TblClients[Nom]` |
| 7 | `Prenom` | Prénom | `txtPrenom` | L2 C3 | libre | — |
| 8 | `Texte_Facture` | Texte de facture | `cboTexte_Facture` | L2 C4 | libre | textes standards de `TblTxtStd` |
| 9 | `Nb_Hres` | Heures | `txtNb_Hres` | L3 C1 | hhh:mm | saisie en heures : `420:00`. Excel stocke une fraction de jour. |
| 10 | `Nb_Pers` | Personnes | `txtNb_Pers` | L3 C2 | chiffres | entier de 1 à 99 |
| 11 | `TVA` | TVA | `chkTVA` | L3 C3 ½g | libre | repris du client |
| 12 | `Forfait` | Forfait | `chkForfait` | L3 C3 ½d | libre | repris du client ; change le calcul du CA |
| 13 | `CA` | Chiffre d'affaires | `txtCA` | L3 C4 | libre | **colonne calculée** — jamais écrite ; estimation affichée pendant la saisie |
| 14 | `No_Facture` | N° de facture | `txtNo_Facture` | L4 C1 | libre | verrouillé — attribué par la facturation |
| 15 | `Commentaires` | Commentaires | `txtCommentaires` | L4 C2 ×3 | libre | repris de `Note_Interne` du client |

### Report automatique depuis TblClients

Choisir une entreprise ou un nom dans sa liste remplit d'un coup :

| Colonne de `TblClients` | → | Champ du formulaire |
|---|---|---|
| `ID_Cresus` | → | `Client_No` |
| `Entreprise` | → | `Entreprise` |
| `Titre` | → | `Titre` |
| `Nom` | → | `Nom` |
| `Prenom` | → | `Prenom` |
| `TVA` | → | `TVA` |
| `Forfait` | → | `Forfait` |
| `Texte_Facture` | → | `Texte_Facture` |
| `Note_Interne` | → | `Commentaires` |

La correspondance est décrite une seule fois, dans `IReportsDepuisClients` (`modInterv_Schema`) : y ajouter une ligne suffit à reporter un champ de plus.


### Durées

Excel range `Nb_Hres` en **fraction de jour** et l'affiche au format `[h]:mm "h"` : `0,125` vaut 3 heures, `17,5` vaut 420 heures. Le formulaire, lui, parle en heures — `420:00` — et fait la conversion dans les deux sens. Une saisie sans deux-points (`420`) est complétée en `420:00` à la sortie du champ.


### Chiffre d'affaires

La colonne `CA` porte une formule dans le tableau Excel :

```
=SI([Forfait]=FAUX ; taux × [Nb_Pers] × [Nb_Hres] × 24 ; taux)
   où taux = RECHERCHEX([Client_No] ; TblClients[ID_Cresus] ; TblClients[Tx_hrs_Forf])
```
Le formulaire **n'écrit jamais cette colonne** : une écriture globale de la ligne remplacerait la formule par une valeur figée. L'écriture se fait donc cellule par cellule, en sautant toute colonne portant une formule — la détection est automatique, une autre colonne calculée ajoutée plus tard serait épargnée de la même façon.

Le champ affiche une **estimation** reprenant le même calcul, remise à jour dès qu'on change les heures, le nombre de personnes, le forfait ou le client. La valeur définitive reste celle que le tableau calcule après enregistrement.


---

## Fiche 4 — tableau des enregistrements

| # | Colonne | En-tête | Largeur |
|---|---|---|---|
| 1 | `Date` | Date | 66 pt |
| 2 | `Client_No` | N° client | 54 pt |
| 3 | `Entreprise` | Entreprise | 140 pt |
| 4 | `Nom` | Nom | 96 pt |
| 5 | `Prenom` | Prénom | 84 pt |
| 6 | `Nb_Hres` | Heures | 56 pt |
| 7 | `Nb_Pers` | Pers. | 44 pt |
| 8 | `Texte_Facture` | Texte de facture | 130 pt |
| 9 | `Commentaires` | Commentaires | 156 pt |
| 10 | `No_Facture` | Facture | 60 pt |
| | | **total** | **886 pt** |

> **Une colonne de moins que demandé.** Onze colonnes étaient souhaitées ; une `ListBox` MSForms n'en accepte que **dix**. `Titre` a été laissée de côté — c'est la moins informative dans une liste, et elle reste visible dans la fiche dès qu'une ligne est sélectionnée. Pour la réintroduire à la place d'une autre : `IColonnesListe`, `ILibellesListe` et `ILargeursListe` dans `modInterv_Schema`, lues position par position.

La hauteur retenue affiche **17 lignes** à la fois. Pour en afficher davantage, ajouter 12,75 pt par ligne à `IT_HAUT`, `IB_TOP` et `I_HAUTEUR`.

Un clic sur un en-tête trie sur cette colonne, un second inverse le sens. Dates et nombres sont comparés pour ce qu'ils sont, pas alphabétiquement.


---

## Boutons

| Bouton | Raccourci | Effet |
|---|---|---|
| **Ajouter** | `Alt+A` | Crée l'intervention ; le n° et le CA sont posés par le programme |
| **Modifier** | `Alt+M` | Enregistre les champs sur l'intervention sélectionnée |
| **Supprimer** | `Alt+S` | Supprime après confirmation, en signalant une facture éventuelle |
| **Effacer** | `Alt+E` | Vide les zones de saisie, sans toucher au tableau |
| **Facturer** | `Alt+F` | Ouvrira le formulaire de facturation — **à définir** |
| **Info** | `Alt+I` | Ouvrira le formulaire d'informations — **à définir** |
| **Quitter** | `Alt+Q / Échap` | Ferme le formulaire |

`Modifier`, `Supprimer` et `Facturer` restent grisés tant qu'aucune intervention n'est sélectionnée.

**Facturer** et **Info** sont en place mais sans destination : leurs points de branchement sont `Interv_Facturer` et `Interv_Info` dans `modInterv_Formulaire`, qui affichent pour l'instant un message. Il n'y aura qu'à remplacer ce message par l'ouverture du formulaire voulu.


---

## Sélecteur de date

`UF_Calendrier`, 224 × 216 pt, entièrement construit en contrôles standard : bandeau de navigation, sept en-têtes de jours et une grille de 42 cases.

MSForms ne fournit aucun calendrier, et le contrôle `MonthView` de Microsoft repose sur `MSCOMCT2.OCX`, absent de la plupart des postes : d'où ce calendrier maison, qui ne dépend de rien.

La date choisie et le jour même sont mis en évidence ; les jours des mois voisins apparaissent en gris clair. Fermer la fenêtre ou cliquer *Annuler* laisse la date inchangée.


---

## Modules

| Module | Rôle |
|---|---|
| [`modInterv_Schema`](../src/modInterv_Schema.bas) | Les 15 champs, les colonnes du tableau, les reports depuis TblClients. **Source de vérité.** |
| [`modInterv_Theme`](../src/modInterv_Theme.bas) | Géométrie du formulaire et couleurs des boutons Facturer et Info. |
| [`modInterv_Generateur`](../src/modInterv_Generateur.bas) | Construit UF_Interventions et UF_Calendrier, et leurs modules de code. |
| [`modInterv_Donnees`](../src/modInterv_Donnees.bas) | TblInterv, recherches dans TblClients, cellules nommées, export du graphique. |
| [`modInterv_Calendrier`](../src/modInterv_Calendrier.bas) | Logique du sélecteur de date. |
| [`modInterv_Formulaire`](../src/modInterv_Formulaire.bas) | Comportement : filtrage, tri, saisie assistée, CRUD, habillage. |
| [`modInterv_Lancement`](../src/modInterv_Lancement.bas) | `OuvrirGestionInterventions`, `VerifierClasseurInterventions`. |

**Dépendance** : la palette et la typographie viennent de `modClients_Theme`, et les utilitaires partagés — `Normaliser`, `EnTexte`, `ObtenirTable`, `TrierChaines`, `Adresses_TextesFacture` — des modules du formulaire des clients. Les sept modules `modInterv_*` s'ajoutent donc aux `modClients_*` déjà présents, ils ne les remplacent pas.


---

## Réglages

| Constante | Valeur | Commande |
|---|---|---|
| `I_LARGEUR` | 960 pt | largeur utile de la fenêtre |
| `I_HAUTEUR` | 748 pt | hauteur utile de la fenêtre |
| `I_MARGE` | 16 pt | marge des cartes |
| `I_CARTE_LARG` | 928 pt | largeur des cartes — garder `I_LARGEUR - 2 × I_MARGE` |
| `F1_TOP` | 12 pt | haut de la fiche 1 |
| `F1_HAUT` | 46 pt | hauteur de la fiche 1 |
| `F2_TOP` | 66 pt | haut de la fiche 2 |
| `F2_HAUT` | 152 pt | hauteur de la fiche 2 |
| `F2_GRAPH_LARG` | 420 pt | largeur de l'image du graphique |
| `F2_TUILE_LARG` | 150 pt | largeur d'une tuile |
| `F2_TUILE_HAUT` | 56 pt | hauteur d'une tuile |
| `F3_TOP` | 226 pt | haut de la fiche 3 |
| `F3_HAUT` | 168 pt | hauteur de la fiche 3 |
| `IG_BLOC` | 210 pt | largeur d'un bloc libellé + champ |
| `IG_GOUTTIERE` | 20 pt | espace entre deux blocs |
| `IG_LIGNE` | 32 pt | pas vertical de la grille |
| `IF_TOP` | 402 pt | haut de la barre de filtrage |
| `IF_HAUT` | 38 pt | hauteur de la barre de filtrage |
| `IT_TOP` | 448 pt | haut du tableau |
| `IT_HAUT` | 246 pt | hauteur du tableau |
| `IT_ENTETE` | 22 pt | bande d'en-têtes |
| `IB_TOP` | 704 pt | haut de la rangée de boutons |
| `IB_LARG` | 118 pt | largeur des boutons |
| `IB_ECART_GROUPE` | 24 pt | écart entre le groupe de saisie et Facturer |
| `CAL_LARGEUR` | 224 pt | largeur du calendrier |
| `CAL_HAUTEUR` | 216 pt | hauteur du calendrier |
| `CAL_JOUR_LARG` | 30 pt | largeur d'une case de jour |
| `CAL_JOUR_HAUT` | 22 pt | hauteur d'une case de jour |

Les zones ne doivent pas se chevaucher, et la dernière doit tenir sous `I_HAUTEUR` :

```
fiche 1     12 ->   58
fiche 2     66 ->  218
fiche 3    226 ->  394
filtre     402 ->  440
tableau    448 ->  694
boutons    704 ->  734
fenêtre         748 pt de haut (I_HAUTEUR)
```

Les couleurs propres à ce formulaire :

| Constante | Valeur | S'applique à |
|---|---|---|
| `COUL_FACTURER` | `#C2790B` | bouton Facturer |
| `COUL_FACTURER_H` | `#DE8F1E` | bouton Facturer, au survol |
| `COUL_INFO` | `#0E7C86` | bouton Info |
| `COUL_INFO_H` | `#1298A4` | bouton Info, au survol |

Toutes les autres couleurs viennent de `modClients_Theme` : les deux formulaires partagent une seule charte. Voir [`REGLAGES.md`](REGLAGES.md).

