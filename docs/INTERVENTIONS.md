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
│ │  ▁▄█▆█▃▂            │  │CA total│ │Facturé │ │Non fact│            │  152 pt
│ │  J F M A M J J A S  │  ├────────┤ ├────────┤ ├────────┤            │
│ └─────────────────────┘  │Mois -1 │ │Mois    │ │Mois nf │            │
├──────────────────────────────────────────────────────────────────────┤
│ FICHE INTERVENTION                                                   │
│  N° client  │ Titre     │ N° interv.│ Date      │ Texte de       │  fiche 3
│  Entreprise ............ │ N° facture│ Heures    │ facture ......  │  168 pt
│  Nom ................... │ Chiffre   │ Personnes │ Commentaires    │
│  Prénom ................ │ ☐ TVA     │ ☐ Forfait │ ..............  │
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

Le graphique du chiffre d'affaires est **dessiné en contrôles MSForms** par [`modInterv_Graphique`](../src/modInterv_Graphique.bas) : douze libellés rectangulaires, un par mois, dont la hauteur est proportionnelle au montant. Les valeurs sont lues dans la table **`Tableau7`** de la feuille Statistiques — celle-là même qui alimente le graphique Excel. Le tracé est refait après chaque ajout, modification ou suppression.

> **Pourquoi ne plus passer par une image.** Un graphique Excel exporté sort à la taille qu'il occupe sur la feuille : agrandi dans le formulaire, il devient flou. Et `LoadPicture`, qui vient de la bibliothèque OLE, ne lit que bmp, ico, wmf, emf, gif et jpg — un PNG s'exportait sans erreur puis restait illisible, le cadre demeurant vide sans le moindre message. Le tracé, lui, est net à toute taille, prend les couleurs du formulaire et ne laisse aucun fichier temporaire.

Une seule série, donc **une seule teinte et aucune légende de couleurs** : le titre nomme ce qui est représenté. La grille reste discrète et seul le sommet de l'échelle est écrit ; les autres montants se lisent **au survol d'une barre**, qui affiche « Mai 2026 — 13'260 CHF » à la place du titre. Le sommet de l'échelle est arrondi vers le haut sur les échelons 1 – 1,5 – 2 – 2,5 – 5 – 10, pour que la plus grande barre occupe rarement moins des deux tiers de la hauteur disponible.

Les six tuiles lisent des cellules nommées :

| Tuile | Cellule Excel |
|---|---|
| CA total | `CATotal` |
| Facturé | `CATotalFacture` |
| Non facturé | `CATotalNonFacture` |
| Mois précédent | `CATotalMmoisPrecedent` |
| Mois actuel | `CATotalMmoisActuel` |
| Mois non facturé | `CATotalMoisActuelNonFacture` |

> **Un nom corrigé.** La demande indiquait `CATotalMmoisActuel` pour « CA mois actuel non facturé », nom déjà employé par « CA mois actuel ». La cellule réellement définie dans le classeur est **`CATotalMoisActuelNonFacture`** (Statistiques!$C$24) : c'est celle-là qui est lue.

Chaque tuile est un contrôle Image portant **`CartePremium.jpg`** (sous-dossier `Images`, à côté du classeur), surmonté de deux libellés à fond blanc — l'image donne les coins arrondis, les libellés portent le texte. Si l'image est absente, la tuile retombe sur un aplat blanc et reste lisible ; `VerifierClasseurInterventions` le signale.

L'image est étirée aux dimensions de la tuile (118 × 56 pt). Une source dans ces proportions évite de déformer l'arrondi des coins.

Libellé et montant sont centrés dans la tuile&nbsp;: horizontalement par `TextAlign`, verticalement en posant le bloc «&nbsp;libellé + écart + montant&nbsp;» à mi-hauteur, ce qui reste juste si l'on change `F2_TUILE_HAUT`. Un libellé MSForms dessinant son texte en HAUT de son cadre, chaque cadre est taillé au plus près de sa police (`F2_TUILE_CAP_HAUT`, `F2_TUILE_VAL_HAUT`). Les filets sont retirés **et** peints en blanc, de sorte qu'aucun liseré ne puisse apparaître sur l'image de fond.


Le graphique et le bloc de tuiles se partagent la largeur de la carte&nbsp;: le graphique occupe **518 pt** à gauche, les 6 tuiles de **118 pt** sont calées sur le bord droit. Élargir les tuiles rétrécit d'autant le graphique&nbsp;; leur position se recalcule toute seule, seule `F2_GRAPH_LARG` est à ajuster en regard.


---

## Fiche 3 — les 16 champs

La grille compte quatre lignes et six colonnes de **126 pt**, groupées en **trois régions** de deux colonnes. L'écart entre régions (**42 pt**) est plus large que la gouttière intérieure (**20 pt**) : ce sont ces respirations, et non un trait, qui font lire trois groupes.

| Région | Colonnes | Ce qu'elle porte |
|---|---|---|
| 1 | 1-2 | **Le client** — n° et titre, puis entreprise, nom, prénom sur toute la largeur |
| 2 | 3-4 | **L'intervention** — n° et date, n° de facture et heures, CA et personnes, TVA et forfait, taux |
| 3 | 5-6 | **Les textes libres** — texte de facture et commentaires, hauts de deux lignes |

L'identité du client d'abord, les chiffres de l'intervention ensuite, les textes libres enfin : c'est l'ordre dans lequel une fiche se remplit. La colonne « Case » ci-dessous se lit *région, ligne, colonne*.

| # | Colonne | Libellé | Contrôle | Case | Saisie | Particularité |
|---|---|---|---|---|---|---|
| 1 | `Client_No` | N° client | `txtClient_No` | R1 L1 C1 | libre | verrouillé à la frappe, mais **enregistré** — repris de `ID_Cresus` du client choisi |
| 2 | `Titre` | Titre | `cboTitre` | R1 L1 C2 | libre | civilités présentes dans `TblClients` |
| 3 | `Entreprise` | Entreprise | `cboEntreprise` | R1 L2 C1 ×2 col. | libre | liste filtrée au fil de la frappe sur `TblClients[Entreprise]` |
| 4 | `Nom` | Nom | `cboNom` | R1 L3 C1 ×2 col. | libre | liste filtrée au fil de la frappe sur `TblClients[Nom]` |
| 5 | `Prenom` | Prénom | `txtPrenom` | R1 L4 C1 ×2 col. | libre | — |
| 6 | `NoInterv` | N° intervention | `txtNoInterv` | R2 L1 C3 | libre | verrouillé — `IN` + plus grand numéro + 1, à l'ajout |
| 7 | `Date` | Date | `txtDate` | R2 L1 C4 | libre | date du jour par défaut ; le bouton ▾ ouvre le calendrier |
| 8 | `No_Facture` | N° de facture | `txtNo_Facture` | R2 L2 C3 | libre | verrouillé — attribué par la facturation |
| 9 | `Nb_Hres` | Heures | `txtNb_Hres` | R2 L2 C4 | hhh:mm | saisie en heures : `420:00`. Excel stocke une fraction de jour. |
| 10 | `CA` | Chiffre d'affaires | `txtCA` | R2 L3 C3 | montant | calculé au fil de la saisie **et enregistré** — heures x personnes x taux |
| 11 | `Nb_Pers` | Personnes | `txtNb_Pers` | R2 L3 C4 | chiffres | entier de 1 à 99 |
| 12 | `TVA` | TVA | `chkTVA` | R2 L4 C3 ½g | libre | repris du client |
| 13 | `Forfait` | Forfait | `chkForfait` | R2 L4 C3 ½d | libre | repris du client ; coché, le CA vaut le taux seul |
| 14 | `Taux/Forfait` | Taux / forfait | `txtTauxForfait` | R2 L4 C4 | montant | repris de `Tx_hrs_Forf` du client, modifiable pour cette intervention seule |
| 15 | `Texte_Facture` | Texte de facture | `cboTexte_Facture` | R3 L1 C5 ×2 col. ×2 lig. | libre | haut de deux lignes ; textes standards de `TblTxtStd` |
| 16 | `Commentaires` | Commentaires | `txtCommentaires` | R3 L3 C5 ×2 col. ×2 lig. | libre | haut de deux lignes, multi-ligne avec ascenseur ; repris de `Note_Interne` du client |

> **Verrouillé n'est pas « non enregistré ».** Un champ verrouillé est hors de portée de la frappe, ce qui ne dit rien de son sort à l'écriture. Deux colonnes seulement échappent au formulaire — `NoInterv`, attribué à l'ajout, et `No_Facture`, attribué par la facturation ; elles sont listées dans `IColonnesNonEcrites`. Tout le reste part dans le tableau, y compris le n° de client et le chiffre d'affaires : ils sont verrouillés parce qu'on ne les tape pas, pas parce qu'ils ne comptent pas.


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
| `Tx_hrs_Forf` | → | `Taux/Forfait` |
| `Texte_Facture` | → | `Texte_Facture` |
| `Note_Interne` | → | `Commentaires` |

La correspondance est décrite une seule fois, dans `IReportsDepuisClients` (`modInterv_Schema`) : y ajouter une ligne suffit à reporter un champ de plus.


### Durées

Excel range `Nb_Hres` en **fraction de jour** et l'affiche au format `[h]:mm "h"` : `0,125` vaut 3 heures, `17,5` vaut 420 heures. Le formulaire, lui, parle en heures — `420:00` — et fait la conversion dans les deux sens. Une saisie sans deux-points (`420`) est complétée en `420:00` à la sortie du champ.


### Chiffre d'affaires

Le formulaire calcule le montant au fil de la saisie, puis l'**enregistre** dans `CA` :

```
à l'heure  CA = [Taux/Forfait] × [Nb_Pers] × [Nb_Hres] × 24
au forfait CA = [Taux/Forfait]
```
Le `× 24` convertit la fraction de jour en heures — voir Durées ci-dessus. Le calcul est refait dès qu'une des quatre valeurs change : heures, personnes, taux, case Forfait ; et aussi quand le choix d'un client rapporte son taux.

Le taux vient de la colonne `Taux/Forfait` de **l'intervention**, reprise de `TblClients[Tx_hrs_Forf]` au moment où le client est choisi. C'est ce qui change tout par rapport à une recherche faite à la volée&nbsp;: le tarif enregistré est celui qui s'appliquait le jour de la prestation, et modifier celui d'un client ne réécrit pas le passé. Il reste modifiable pour une intervention isolée.

> **Si la colonne `CA` porte encore une formule**, c'est elle qui gagne&nbsp;: l'écriture se fait cellule par cellule en sautant toute colonne calculée, et le montant du formulaire est ignoré. Les deux régimes donnent le même résultat tant que la formule est celle d'origine, mais `VerifierClasseurInterventions` dit lequel s'applique — retirer la formule laisse le formulaire enregistrer sa valeur.


---

## Fiche 4 — tableau des enregistrements

| # | Colonne | En-tête | Largeur |
|---|---|---|---|
| 1 | `Date` | Date | 66 pt |
| 2 | `Client_No` | N° client | 54 pt |
| 3 | `Entreprise` | Entreprise | 136 pt |
| 4 | `Nom` + `Prenom` | Nom et prénom | 142 pt |
| 5 | `Nb_Hres` | Heures | 56 pt |
| 6 | `Nb_Pers` | Pers. | 44 pt |
| 7 | `Taux/Forfait` | Taux/Forf. | 62 pt |
| 8 | `Texte_Facture` | Texte de facture | 130 pt |
| 9 | `Commentaires` | Commentaires | 150 pt |
| 10 | `No_Facture` | Facture | 58 pt |
| | | **total** | **898 pt** |

> **Dix cases pour seize colonnes.** Une `ListBox` MSForms n'accepte que **dix** colonnes, et la limite est dure : la onzième fait échouer l'affectation de `ColumnCount`. Deux moyens de s'en accommoder cohabitent ici. Le premier : **réunir deux colonnes dans une case**, comme le nom et le prénom, séparés par `ICL_SEPARATEUR` dans `IColonnesListe` et rassemblés à l'affichage par `Interv_ValeurListe` — chacun garde sa propre colonne dans `TblInterv` et y est enregistré séparément. Le second : **laisser de côté** ce que la fiche montre déjà — `Titre`, `NoInterv`, `CA`, `TVA` et `Forfait`. Pour changer la répartition : `IColonnesListe`, `ILibellesListe` et `ILargeursListe` dans `modInterv_Schema`, lues position par position.

La hauteur retenue affiche **17 lignes** à la fois. Pour en afficher davantage, ajouter 12,75 pt par ligne à `IT_HAUT`, `IB_TOP` et `I_HAUTEUR`.

Un clic sur un en-tête trie sur cette colonne, un second inverse le sens. Dates et nombres sont comparés pour ce qu'ils sont, pas alphabétiquement ; une case qui en réunit deux se trie sur le texte affiché, donc par nom puis par prénom.


### Filtrage

Deux filtres, qui **se cumulent**, sur la barre au-dessus du tableau :

| Filtre | Porte sur |
|---|---|
| Texte | la colonne choisie dans le menu — `Entreprise` ou `Nom` — sans tenir compte des accents ni des majuscules |
| Mois | le mois de la colonne `Date`, ou tous les mois |

« Réinitialiser » éteint les deux à la fois. Le compteur à droite dit combien d'interventions restent affichées sur le total.


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

`UF_Calendrier`, 260 × 256 pt, entièrement construit en contrôles standard : bandeau de navigation, sept en-têtes de jours et une grille de 42 cases.

MSForms ne fournit aucun calendrier, et le contrôle `MonthView` de Microsoft repose sur `MSCOMCT2.OCX`, absent de la plupart des postes : d'où ce calendrier maison, qui ne dépend de rien.

La grille est aérée — 34 × 26 pt par case — pour qu'un jour se vise sans effort. Les colonnes du samedi et du dimanche sont teintées en fond, ce qui découpe la semaine sans ajouter le moindre trait, et la case sous la souris s'éclaire : on sait toujours ce qu'un clic choisirait.

Trois états se distinguent d'un coup d'œil : le **jour choisi** (pastille bleue pleine), **aujourd'hui** (fond clair, chiffre bleu) et les **jours des mois voisins** (gris clair). Deux raccourcis en pied — *Aujourd'hui* et *Fin de mois* — donnent les deux dates les plus souvent saisies dans une facturation mensuelle. Fermer la fenêtre ou cliquer *Annuler* laisse la date inchangée.

> La barre de titre Windows reste visible au-dessus du calendrier : la supprimer demanderait un appel à l'API Windows, écarté.


---

## Modules

| Module | Rôle |
|---|---|
| [`modInterv_Schema`](../src/modInterv_Schema.bas) | Les 15 champs, les colonnes du tableau, les reports depuis TblClients. **Source de vérité.** |
| [`modInterv_Theme`](../src/modInterv_Theme.bas) | Géométrie du formulaire et couleurs des boutons Facturer et Info. |
| [`modInterv_Generateur`](../src/modInterv_Generateur.bas) | Construit UF_Interventions et UF_Calendrier, et leurs modules de code. |
| [`modInterv_Donnees`](../src/modInterv_Donnees.bas) | TblInterv, recherches dans TblClients, cellules nommées. |
| [`modInterv_Graphique`](../src/modInterv_Graphique.bas) | Trace le graphique du chiffre d'affaires en contrôles MSForms. |
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
| `F2_GRAPH_LARG` | 518 pt | largeur de la zone du graphique |
| `GR_ORIGINE_X` | 30 pt | abscisse du coin haut-gauche du graphique |
| `GR_ORIGINE_Y` | 90 pt | ordonnée du coin haut-gauche du graphique |
| `GR_MARGE_G` | 48 pt | colonne réservée à l'échelle, à gauche des barres |
| `GR_TRACE_HAUT` | 80 pt | hauteur de l'aire de tracé |
| `GR_BARRE_LARG` | 26 pt | largeur d'une barre |
| `F2_TUILE_LARG` | 118 pt | largeur d'une tuile |
| `F2_TUILE_HAUT` | 56 pt | hauteur d'une tuile |
| `F2_TUILE_INSET` | 17 pt | retrait des libellés dans une tuile |
| `F2_TUILE_CAP_HAUT` | 11 pt | hauteur du cadre du libellé de tuile |
| `F2_TUILE_VAL_HAUT` | 15 pt | hauteur du cadre du montant |
| `F2_TUILE_ECART` | 3 pt | écart entre le libellé et le montant |
| `F2_PADDING` | 14 pt | retrait intérieur de la carte des statistiques |
| `F3_TOP` | 226 pt | haut de la fiche 3 |
| `F3_HAUT` | 168 pt | hauteur de la fiche 3 |
| `IG_BLOC` | 126 pt | largeur d'une colonne de saisie |
| `IG_GOUTTIERE` | 20 pt | espace entre deux colonnes d'une même région |
| `IG_ECART_REGION` | 42 pt | espace entre deux régions |
| `IG_LIGNE` | 32 pt | pas vertical de la grille |
| `IF_TOP` | 402 pt | haut de la barre de filtrage |
| `IF_HAUT` | 38 pt | hauteur de la barre de filtrage |
| `IT_TOP` | 448 pt | haut du tableau |
| `IT_HAUT` | 246 pt | hauteur du tableau |
| `IT_ENTETE` | 22 pt | bande d'en-têtes |
| `IB_TOP` | 704 pt | haut de la rangée de boutons |
| `IB_LARG` | 118 pt | largeur des boutons |
| `IB_ECART_GROUPE` | 24 pt | écart entre le groupe de saisie et Facturer |
| `CAL_LARGEUR` | 260 pt | largeur du calendrier |
| `CAL_HAUTEUR` | 256 pt | hauteur du calendrier |
| `CAL_JOUR_LARG` | 34 pt | largeur d'une case de jour |
| `CAL_JOUR_HAUT` | 26 pt | hauteur d'une case de jour |
| `CAL_GRILLE_Y` | 66 pt | haut de la grille des jours |
| `CAL_PIED_Y` | 232 pt | ligne des raccourcis du bas |

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

