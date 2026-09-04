# Feuille d'accueil

La feuille **`Accueil`** — le menu du classeur. Quatre cartes, une par module ; un clic ouvre le formulaire correspondant.

Elle se dessine par **`GenererAccueil`** (module `modAccueil_Generateur`) et se règle par `modAccueil_Theme`.

> Ce fichier est produit à partir des sources : les valeurs sont celles réellement en vigueur dans `src/`.

---

## Pourquoi une feuille et non un cinquième formulaire

MSForms ne sait pas arrondir un coin, dégrader un fond ni porter une ombre : un menu construit en UserForm aurait l'aspect d'un panneau de contrôle de 1998. Les **formes** d'une feuille savent tout cela, et la feuille s'ouvre avec le classeur, sans macro à lancer.

Elle reste néanmoins **dessinée par code**, comme les quatre formulaires : `GenererAccueil` se relance autant de fois qu'on veut, rien ne s'accumule, et la mise en page ne se refait jamais à la main.

---

## Plan

Surface dessinée de **960 × 448 points** — la même largeur que les quatre formulaires, soit 1280 pixels à 96 ppp.

| Bande | Haut | Hauteur | Contenu |
|---|---|---|---|
| Bandeau | 0 | 104 | titre, sous-titre, année (cellule nommée `AnneeEnCours`) |
| Intitulé | 128 | 16 | « MODULES », interlettré |
| Cartes | 152 | 224 | les quatre modules, côte à côte |
| Filet | 404 | 1 | séparateur du pied |
| Pied | 414 | 16 | rappel du mode d'emploi |

Chaque carte fait **210 points** de large, séparée de la suivante par 20 points, avec 30 points de marge à gauche et à droite : `30 + 4 × 210 + 3 × 20 + 30 = 960`. `simulate_accueil.py` vérifie l'égalité.

---

## Les quatre cartes

| Pastille | Titre | Macro lancée | Accent |
|---|---|---|---|
| **C** | Clients | `OuvrirGestionClients` | `#1769B5` bleu |
| **I** | Interventions | `OuvrirGestionInterventions` | `#008F65` vert |
| **F** | Facturation | `OuvrirFacturation` | `#C77A1E` ambre |
| **S** | Statistiques | `OuvrirStatistiques` | `#6B4FA3` violet |

Le bleu et le vert sont ceux des boutons **Modifier** et **Ajouter** ; les deux autres accents ont été ajoutés pour ce menu, le rouge et l'anthracite du classeur disant déjà « supprimer » et « quitter ».

Une carte est faite de **sept formes** — fond, pastille, lettre, filet d'accent, titre, détail, lien — **groupées** : la macro est portée par le groupe, si bien que le clic est pris n'importe où sur la carte et non seulement sur son fond.

---

## Régler l'apparence

Tout est dans `modAccueil_Theme`, et rien n'est dans le générateur :

- **Textes, couleurs et macros des cartes** : la table `ConstruireCartesMenu`. C'est là qu'on renomme un module ou qu'on change la macro qu'il lance.
- **Géométrie** : les constantes `AC_*`. Changer le nombre de cartes demande de corriger `AC_NB_CARTES` **et** `AC_CARTE_LARG` — le simulateur refuse une largeur qui ne tombe pas juste.
- **Typographie** : `AC_T_*` pour les corps, `POLICE_DEMI` et `POLICE_LEGERE` pour les familles. La palette, elle, vient de `modClients_Theme` : le menu ne redéfinit aucune couleur du classeur.
- **Ombre portée** : `AC_OMBRE_FLOU`, `AC_OMBRE_DY`, `AC_OMBRE_TRANSP`.

Les formes restent modifiables à la main dans Excel — mais la prochaine génération les écrase : mieux vaut corriger le thème.

---

## Ouvrir le classeur sur l'accueil

Dans le module **`ThisWorkbook`** du VBE :

```vba
Private Sub Workbook_Open()
    AfficherAccueil
End Sub
```

`AfficherAccueil` (module `modAccueil_Generateur`) active la feuille, ou explique quoi faire si elle n'a pas encore été dessinée.

---

## Ce que fait la génération

1. Crée la feuille `Accueil` si elle manque ; sinon la déprotège et efface toutes ses formes.
2. La place en **première position** du classeur.
3. Peint le fond en `#F0F3F8`, celui des formulaires.
4. Dessine le bandeau, les quatre cartes et le pied.
5. Cache **quadrillage** et **en-têtes** de lignes et de colonnes, interdit la sélection de cellules, et protège la feuille.

Le quadrillage est une propriété de la *fenêtre* et non de la feuille : il ne se règle que sur la feuille active, d'où l'activation en fin de génération.

---

## Vérifications automatiques

`simulate_accueil.py` rejoue le dessin sans Excel et refuse :

- des cartes qui débordent la largeur ou se chevauchent ;
- un contenu qui déborde de sa carte, ou deux éléments qui se recouvrent dedans ;
- deux formes de même nom — `Shapes.Range` les désigne par leur nom, et le groupe se ferait sur la mauvaise ;
- un groupe qui n'attend pas exactement le nombre de formes nommées ;
- une carte qui lance une macro **inexistante**, **privée** ou **à arguments** : `OnAction` ne vérifie rien, l'erreur ne se verrait qu'au clic ;
- une couleur écrite en dur dans le générateur au lieu de venir du thème.
