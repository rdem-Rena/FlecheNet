# Les deux graphiques

> Les valeurs citées ici sont celles réellement en vigueur dans
> [`../outils/construire_classeur.py`](../outils/construire_classeur.py).

Les deux sont des **feuilles graphiques** — une page entière chacune, comme dans
le fichier d'origine — et suivent les **mêmes trois réglages**, sur la feuille
`Saisie` : date de début, date de fin, jours de la moyenne lissée.

---

## Graph_Mesures — chaque relevé

Repris du fichier d'origine, à l'identique.

| Série | Source | Couleur | Marqueur |
|---|---|---|---|
| Systole | `Datas!E` | bleu `#0070C0` | × |
| Diastole | `Datas!F` | vert `#00B050` | × |
| Basse B. / Basse H. | `Datas!H` / `I` | jaune `#FFCC00` | — |
| Haute B. / Haute H. | `Datas!J` / `K` | rouge `#FF0000` | — |
| Change | `Datas!M` | gris `#BFBFBF` | barres |

Deux points en abscisse par jour : le matin et le soir. Axe des ordonnées figé
de **60 à 170 mmHg**.

**`Change`** n'est pas une mesure. C'est un repère de changement de situation,
porté à 150 ou 160 pour apparaître en barre sur toute la hauteur utile. Il était
tracé dans la couleur d'accent du thème ; il est passé au gris, pour ne plus
concurrencer les six courbes.

---

## Graph_Lissage — les deux bandes

### Ce qui est tracé

Pour chaque jour, on retient la **plus haute** et la **plus basse** des mesures
du jour — en pratique, celle du matin et celle du soir, dans un ordre ou dans
l'autre. Ces deux séries sont ensuite **lissées** sur N jours, N étant la cellule
« Jours de la moyenne lissée ».

Il reste quatre courbes :

| Courbe | Ce qu'elle vaut | Couleur |
|---|---|---|
| Systolique haute lissée | moyenne sur N jours du maximum systolique du jour | vert foncé `#548235` |
| Systolique basse lissée | moyenne sur N jours du minimum systolique du jour | vert foncé `#548235` |
| Diastolique haute lissée | moyenne sur N jours du maximum diastolique du jour | bleu `#2E75B6` |
| Diastolique basse lissée | moyenne sur N jours du minimum diastolique du jour | bleu `#2E75B6` |

L'espace entre les deux courbes **systoliques** est teinté en **vert pâle**
`#D5E8C8`, celui entre les deux **diastoliques** en **bleu pâle** `#CFE2F3`.

> **Le vert va à la systolique, alors qu'elle est bleue sur l'autre graphique.**
> C'est voulu : c'est ce qui a été demandé, et les deux graphiques ne se lisent
> pas ensemble — on regarde l'un ou l'autre. Pour inverser, échanger
> `COUL_BANDE_SYS` et `COUL_BANDE_DIA` dans le script, et les deux couleurs de
> trait avec elles.

Les quatre lignes de repère (80, 90, 120, 130) sont reprises ici aussi, dans
leurs couleurs du premier graphique.

### Comment on colorie entre deux courbes

**Excel ne sait pas le faire.** Il n'existe aucune option « remplir entre ces
deux séries ». Le classeur emploie donc le procédé habituel : une **aire
empilée** de quatre séries, dont deux sont invisibles.

Empilées de bas en haut :

| Série | Colonne de `Calcul` | Hauteur | Remplissage |
|---|---|---|---|
| `Z_Dia_Bas` | `N` | la diastolique basse lissée | aucun |
| `Z_Dia` | `O` | diastolique haute − basse | bleu pâle |
| `Z_Ecart` | `P` | systolique basse − diastolique haute | aucun |
| `Z_Sys` | `Q` | systolique haute − basse | vert pâle |

Les hauteurs s'additionnent : le sommet de la deuxième aire tombe exactement sur
la diastolique haute, le sommet de la quatrième sur la systolique haute. Les
deux aires sans remplissage ne servent qu'à porter les deux autres à la bonne
altitude, et sont retirées de la légende.

`MAX(0; …)` protège chacune : une hauteur négative — si les deux bandes venaient
à se toucher — décalerait tout ce qui est empilé au-dessus.

Les quatre aires sont dessinées **avant** les huit courbes. L'ordre des groupes
dans le fichier est l'ordre de dessin : les aplats derrière, les traits dessus.

### La moyenne lissée

```
=IFERROR(IF(ROW()<3+Nb_Jours-1,NA(),
         IF(COUNT(fenêtre)=0,NA(),SUM(fenêtre)/COUNT(fenêtre))),NA())
```

où `fenêtre` vaut `OFFSET(colonne_du_jour, 1-Nb_Jours, 0, Nb_Jours, 1)` : les
`Nb_Jours` jours qui précèdent, celui-ci compris.

Trois choses à savoir :

- **Un jour sans mesure ne casse pas la courbe.** Les extrêmes du jour rendent
  une chaîne vide, que `SOMME` et `NB` ignorent l'une comme l'autre : la
  moyenne porte sur les jours effectivement mesurés de la fenêtre.
- **La courbe commence au N-ième jour.** Tant que la fenêtre ne tient pas
  entière dans les données, la cellule rend `#N/A` et le graphique laisse un
  trou. Une moyenne sur sept jours calculée sur trois n'en serait pas une.
- **Une période non mesurée apparaît comme un trou**, et non comme un trait qui
  relie ses deux bords.

### Pourquoi aucune fonction récente

Le classeur n'emploie **aucune fonction postérieure à Excel 2007**. Ni `MAXIFS`,
ni `MINIFS`, ni `AGGREGATE`, qui auraient pourtant écrit tout cela en une ligne.

Un `.xlsx` ne les accepte que préfixées `_xlfn.` — c'est ainsi qu'Excel les
enregistre lui-même. Écrites en clair par un programme qui fabrique le fichier,
elles rendent `#NOM?` à l'ouverture, dans Excel comme ailleurs, et la faute ne
se voit qu'une fois le fichier livré.

Leur remplacement est d'ailleurs **plus rapide** : l'invariant deux-lignes-par-
jour donne les deux lignes du jour par le calcul, là où `MAXIFS` relisait les
5 368 lignes de la feuille pour chacun des 2 684 jours.

---

## Les plages nommées

Toutes sont dynamiques. Les neuf premières portent les noms du fichier
d'origine : un graphique refait ailleurs sur les mêmes noms continue de
fonctionner.

| Nom | Définition |
|---|---|
| `Plage_X`, `Systole`, `Diastole`, `Pouls`, `Basse_Bas`, `Basse_Hau`, `Haute_Bas`, `Haute_Hau`, `Change` | `OFFSET(Datas!<col>$3, Dec_M, 0, Nb_M, 1)` |
| `Jour_X`, `Jours_Nb`, `Sys_Haut_L`, `Sys_Bas_L`, `Dia_Haut_L`, `Dia_Bas_L`, `Z_Dia_Bas`, `Z_Dia`, `Z_Ecart`, `Z_Sys`, `L_Basse_Bas`, `L_Basse_Hau`, `L_Haute_Bas`, `L_Haute_Hau` | `OFFSET(Calcul!<col>$3, Dec_J, 0, Nb_J, 1)` |

`Dec_J` et `Nb_J` sont sur la feuille `Reglages`, et se calculent par
soustraction de dates. `Dec_M` et `Nb_M` en sont le double : deux lignes de
`Datas` par jour.

**Aucune n'emploie `EQUIV`.** Une recherche peut rendre `#N/A` quand la date
cherchée ne figure pas dans la colonne — c'était le cas du fichier d'origine, où
choisir une date tombant dans les 59 jours manquants cassait les neuf plages
d'un coup. Une soustraction, non.

> Le résumé de la feuille `Saisie` lit **ces mêmes plages**. Il ne peut donc pas
> annoncer autre chose que ce que les graphiques montrent.

---

## Regarder les graphiques sans Excel

LibreOffice ouvre le classeur et recalcule tout correctement, mais il **ne sait
pas résoudre une plage nommée dans une série de graphique** — sous aucune forme,
ni `[0]!Nom`, ni `Nom`, ni `fichier.xlsx!Nom`. Les deux graphiques s'y affichent
donc avec leurs axes et sans leurs courbes.

Ce n'est pas un défaut du classeur : `[0]!Nom` est ce qu'écrit Excel, et ce
qu'employait déjà le fichier d'origine. Pour contrôler le rendu hors d'Excel, il
faut recopier les graphiques avec des plages de cellules figées à la place des
noms.
