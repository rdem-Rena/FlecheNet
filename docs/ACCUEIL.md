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

## Le pied de page et ses deux boutons

La ligne du bas porte le rappel « Cliquez sur une carte pour ouvrir le module »
à gauche, et deux petits boutons calés à droite :

| Bouton | Macro | Effet |
|---|---|---|
| **Quitter** | `Accueil_Quitter` | propose d'enregistrer, puis ferme le classeur — ou Excel s'il n'y a que lui |
| **Unlock** | `Accueil_Deverrouiller` | demande le mot de passe et rend le classeur à Excel |

Ils sont discrets par choix : ce n'est pas ce qu'on vient faire sur cette page.
Leur taille se règle par `AC_BT_LARG`, `AC_BT_HAUT` et `AC_BT_GOUT`, et le
simulateur refuse une largeur qui mangerait le texte du pied.

---

## Le mode kiosque

Verrouillé, le classeur ne montre plus que la feuille d'accueil, dans une
fenêtre de taille fixe : ni ruban, ni onglets, ni barre de formule, ni en-têtes,
ni ascenseurs, et les autres feuilles **très masquées** — un état que le menu
*Afficher* ne défait pas. L'utilisateur ne peut donc faire que ce que les quatre
formulaires lui permettent.

Tout cela vit dans `modAccueil_Verrou`, et se commande par **`AC_KIOSQUE`**
(`modAccueil_Theme`) : à `False`, le classeur s'ouvre sur l'accueil sans rien
cacher — l'état dans lequel travailler pour le modifier.

### Ce que ce verrou est, et ce qu'il n'est pas

Il met la maison en ordre, il ne la ferme pas à clef :

- **macros désactivées à l'ouverture, et rien ne se verrouille.** C'est vrai de
  tout verrou écrit en VBA ;
- la protection de structure et un mot de passe de feuille se lèvent en quelques
  minutes avec un utilitaire du commerce.

Il empêche les fausses manoeuvres, pas la malveillance.

### Le mot de passe

Dans la cellule nommée **`Mot_de_passe`**. S'il n'y en a pas, le bouton *Unlock*
déverrouille **sans rien demander** : un classeur dont on ne peut plus sortir
serait pire que pas de verrou du tout.

### Ce qui doit être rendu à Excel

Le ruban caché, la barre de formule et les onglets sont des réglages de
**l'application**, pas du classeur : un autre classeur ouvert dans la même
instance d'Excel les trouverait cachés lui aussi. `Accueil_Arreter`, appelée à
la fermeture, remet donc tout en place — et le simulateur vérifie que les trois
sorties du verrou le font.

### Installation

Le verrou part à l'ouverture du classeur, ce qui demande deux gestionnaires dans
le module `ThisWorkbook` — un module qui ne s'importe pas. **`InstallerDemarrage`**
(module `modAccueil_Verrou`) les y écrit :

```vba
Private Sub Workbook_Open()
    Accueil_Demarrer
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Accueil_Arreter
End Sub
```

Elle **n'écrase rien** : si un gestionnaire existe déjà — celui qui appelait
`AfficherAccueil`, par exemple — elle le laisse et dit la ligne à y ajouter.

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
