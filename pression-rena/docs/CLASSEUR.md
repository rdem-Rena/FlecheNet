# Le classeur, feuille par feuille

> Les valeurs citées ici sont celles réellement en vigueur dans
> [`../outils/construire_classeur.py`](../outils/construire_classeur.py).

---

## L'invariant

`Datas` porte **exactement deux lignes par jour calendaire**, sans trou et sans
doublon, de la première à la dernière date du gabarit. Les lignes `3 + 2k` et
`4 + 2k` sont les deux demi-journées du k-ième jour.

Trois choses en dépendent :

1. **Les plages nommées se calculent par soustraction.** Un décalage de date
   vaut deux lignes dans `Datas`, une ligne dans `Calcul`. Aucun `EQUIV`,
   aucune recherche, et donc aucun `#N/A` quand la date demandée n'est pas dans
   la colonne.
2. **`Calcul` désigne les deux lignes du jour par le calcul**, sans balayer la
   feuille.
3. **Le graphique est honnête.** Une période non mesurée apparaît comme un trou,
   alors que le fichier d'origine la faisait disparaître en rapprochant les deux
   dates qui l'encadraient.

**Ne jamais insérer ni supprimer une ligne dans `Datas`.** La feuille `Saisie`
porte un contrôle qui compare le nombre de lignes au nombre de jours et le dit
en clair. Le formulaire `UF_Pression`, lui, ne crée jamais de ligne : il remplit
celle qui existe.

---

## Saisie

La page d'accueil. Six blocs.

| Ligne | Bloc |
|---|---|
| 3 – 6 | **Réglages des deux graphiques** : date de début, date de fin, jours de la moyenne lissée. En regard : la première mesure du classeur, la dernière mesure saisie, le nombre de jours affichés |
| 9 – 11 | **Seuils de référence** : les quatre lignes de repère, en mmHg |
| 13 – 18 | **Résumé de la période affichée** |
| 20 – 36 | **Les quinze dernières mesures** |
| 38 – 39 | **Contrôle du classeur** |

Les trois réglages sont protégés par une validation : une date hors du
calendrier et un lissage hors de 1–90 sont refusés à la frappe.

Le **résumé** lit les mêmes plages nommées que les graphiques : il porte donc
exactement sur ce qu'ils montrent.

| Gauche | Droite |
|---|---|
| Mesures | Jours mesurés |
| Systolique moyenne | Systolique maximale |
| Diastolique moyenne | Diastolique maximale |
| Pouls moyen | Poids, dernier relevé |
| Au-dessus d'un seuil haut (%) | Poids moyen |

---

## Datas

Une ligne par demi-journée. Les quatorze premières colonnes sont celles du
fichier d'origine, dans le même ordre ; `Poids` est nouvelle.

| # | Colonne | Contenu |
|---|---|---|
| A | `Date` | le jour, écrit en clair sur les deux lignes |
| B | `Heure` | l'heure du relevé |
| C | `Activité` | footing, salle, séjour… |
| D | `Date-Compo` | **formule** — l'étiquette de l'axe du premier graphique |
| E | `Systole` | mmHg |
| F | `Diastole` | mmHg |
| G | `Pouls` | battements par minute |
| H – K | `Basse B.`, `Basse H.`, `Haute B.`, `Haute H.` | **formules** — les quatre lignes de repère, renvoyant aux cellules de seuil |
| L | `Demi` | **formule** — `matin` avant midi, `Soir` après |
| M | `Change` | repère de changement de situation : 150 ou 160 |
| N | `Commentaire` | texte libre |
| O | `Poids` | **nouveau** — kg, relevé le matin |

Les six colonnes en formule ne sont **jamais écrites** par le formulaire : les
écrire remplacerait leur formule par une valeur figée, et l'étiquette de l'axe
cesserait de suivre sa ligne.

### `Demi` n'est pas ce qui range les mesures

La colonne dit `matin` ou `Soir` d'après l'heure, pour l'étiquette. Mais ce qui
range une mesure, c'est le **rang de sa ligne dans sa paire** : huit jours du
fichier d'origine portent deux mesures du matin, et les ranger par l'heure en
aurait perdu une à chaque fois.

Le fichier d'origine écrivait `Erreur` dans cette colonne pour toute ligne sans
heure, ce qui remplissait de fautes toute la partie à venir du calendrier ; une
ligne vide rend désormais une chaîne vide.

---

## Calcul

Un jour par ligne. Rien à y saisir.

| # | Colonne | Contenu |
|---|---|---|
| A | `Jour` | la date |
| B | `Libellé` | l'étiquette de l'axe du graphique lissé |
| C | `Nb` | nombre de mesures du jour — calculé une fois, les quatre extrêmes s'y réfèrent |
| D – G | `Sys haut`, `Sys bas`, `Dia haut`, `Dia bas` | les extrêmes du jour ; chaîne vide si le jour n'a pas de mesure |
| H – I | `Pouls moy`, `Poids` | moyennes du jour |
| J – M | les quatre **moyennes lissées** | `#N/A` tant que la fenêtre ne tient pas entière |
| N – Q | `Z dia bas`, `Z dia`, `Z écart`, `Z sys` | les quatre séries de l'aire empilée |
| R – U | les quatre **seuils** | recopie des cellules de réglage, pour les lignes de repère |

Pourquoi les extrêmes rendent une chaîne vide et les moyennes lissées `#N/A` :
les premiers alimentent un calcul — et `SOMME` propagerait une erreur alors
qu'elle ignore le texte — les secondes sont tracées, et c'est `#N/A` qui creuse
un trou dans la courbe.

---

## Reglages

Les cellules dérivées. Aucune n'est saisie ; les plages nommées les lisent.

| Nom | Ce qu'il vaut |
|---|---|
| `Premier_Jour`, `Dernier_Jour` | les bornes du calendrier, lues sur `Datas` |
| `Derniere_Mesure` | la dernière ligne portant une systole |
| `Deb_Eff`, `Fin_Eff` | les deux dates de réglage, ramenées dans le calendrier |
| `Dec_J`, `Nb_J` | décalage et hauteur, en jours — pour `Calcul` |
| `Dec_M`, `Nb_M` | les mêmes, doublés — pour `Datas` |
| `Controle` | le contrôle de structure, en clair |

`Fin_Eff` vaut `MIN(MAX(Fin_Demandee, Deb_Eff), Dernier_Jour)` : une date de fin
antérieure au début ne rend jamais une hauteur négative.

Les quinze dernières lignes portent les **numéros de ligne** des quinze
dernières mesures, que la feuille `Saisie` affiche. Elles emploient
`RECHERCHE(2; 1/(plage<>""); …)` — le tour classique qui rend la dernière ligne
non vide — et non `AGGREGATE` en forme matricielle, que toutes les versions
n'évaluent pas de la même façon.

---

## Ce que le fichier d'origine avait, et qui a bougé

| Avant | Maintenant |
|---|---|
| Les cellules de réglage du graphique, à droite des données dans `Datas` (`N1:R5`) | Sur `Saisie` et `Reglages` |
| `Lig_Deb = MATCH(date, A:A, 0)` | Une soustraction de dates |
| 59 jours absents entre le 31.12.2020 et le 01.03.2021 | Rétablis, vides |
| Quatre colonnes de 80, 90, 120, 130 recopiés | Quatre formules renvoyant aux seuils |
| `TEXT(A3;"jjj jj.mm.aa")` | `CHOISIR` + `JOURSEM`, qui ne dépend pas de la langue |
| Pas de colonne pour le poids | `Datas!O` |
| Pas de page d'accueil | La feuille `Saisie` |
