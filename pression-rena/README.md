# Pression Rena — suivi de la pression artérielle

Un classeur Excel qui reprend un carnet de relevés existant, lui ajoute une page
de saisie, et en tire **deux graphiques** : les mesures brutes, et des moyennes
lissées dessinées en bandes.

> **Les données ne sont pas dans ce dépôt, et ne doivent pas y entrer.**
> Ce dépôt est public. Le classeur porte des milliers de mesures nominatives, le
> poids et des commentaires qui nomment des tiers. Il se **refabrique** à partir
> du fichier de relevés — voir [Fabriquer le classeur](#fabriquer-le-classeur) —
> et le `.gitignore` de ce dossier refuse tout `.xlsx` par sécurité.

---

## Ce que contient le classeur

| Feuille | Rôle |
|---|---|
| `Saisie` | Les trois réglages des graphiques, les quatre seuils, le résumé de la période, les quinze dernières mesures |
| `Datas` | Les relevés — **deux lignes par jour**, matin et soir |
| `Calcul` | Un jour par ligne : les extrêmes du jour, puis les moyennes lissées |
| `Reglages` | Les cellules dérivées, lues par les plages nommées. Rien à y modifier |
| `Graph_Mesures` | Feuille graphique : chaque relevé, comme dans le fichier d'origine |
| `Graph_Lissage` | Feuille graphique : les deux bandes lissées |

Détail complet : [`docs/CLASSEUR.md`](docs/CLASSEUR.md).

---

## Les trois réglages

Tout est piloté depuis **trois cellules** de la feuille `Saisie`, et les deux
graphiques les suivent ensemble :

| Cellule | Ce qu'elle commande |
|---|---|
| **Date de début** | Le premier jour affiché |
| **Date de fin** | Le dernier jour affiché |
| **Jours de la moyenne lissée** | La largeur de la fenêtre de lissage — 7 par défaut |

Une date hors du calendrier est ramenée à ses bornes plutôt que de vider le
graphique. Quatre cellules de plus règlent les **seuils** : les lignes de repère
à 80, 90, 120 et 130 mmHg se déplacent avec elles, sur les deux graphiques.

---

## Les deux graphiques

**`Graph_Mesures`** reprend celui du fichier d'origine : la systolique en bleu et
la diastolique en vert, chacune avec ses marqueurs, les quatre lignes de repère,
et les barres grises de la colonne `Change`.

**`Graph_Lissage`** est le nouveau. Pour chaque jour, on retient la **plus haute**
et la **plus basse** des mesures du jour, puis on lisse chacune sur N jours. Il
reste quatre courbes, et l'espace entre les deux courbes systoliques est teinté
en **vert pâle**, celui entre les deux diastoliques en **bleu pâle**.

Comment on colorie entre deux courbes, et pourquoi le vert va à la systolique
alors qu'elle est bleue sur l'autre graphique : [`docs/GRAPHIQUES.md`](docs/GRAPHIQUES.md).

---

## Ce qui a changé par rapport au fichier d'origine

- **Une colonne `Poids`** a été ajoutée — elle n'existait pas, alors que le poids
  était relevé chaque matin.
- **59 jours manquants** ont été rétablis, vides : le fichier d'origine passait
  du 31.12.2020 au 01.03.2021 sans rien dire, et le graphique rapprochait
  simplement les deux dates. Le trou se voit maintenant.
- **Les étiquettes de date** ne dépendent plus de la langue d'Excel. Elles
  affichent la même chose qu'avant — `lun 31.08.20` — mais partout.
- **Les quatre lignes de repère** ne sont plus quatre colonnes de nombres
  recopiés : elles renvoient aux cellules de réglage.
- Les cellules qui pilotaient le graphique, posées à droite des données dans
  `Datas`, sont passées sur `Saisie` et `Reglages`.

Les 3 613 mesures sont reprises **à l'identique**, date et rang compris.

---

## La page de saisie

Deux façons de saisir, au choix.

**Sans macro** — la feuille `Datas` porte déjà le calendrier jusqu'en 2028 :
aller à la ligne du jour et remplir les cellules. La feuille `Saisie` montre en
face les quinze dernières mesures et le résumé de la période.

**Avec le formulaire** — `UF_Pression` est un formulaire de saisie construit par
code, comme ceux de FlècheNet : date, demi-journée, heure, activité, systole,
diastole, pouls, poids, commentaire, et un repère de couleur qui dit si la
mesure passe au-dessus des seuils. Il faut pour cela importer les modules de
[`src/`](src) et enregistrer le classeur en `.xlsm` :
[`docs/INSTALLATION.md`](docs/INSTALLATION.md).

Le formulaire **ne crée jamais de ligne** : il remplit celle qui correspond au
jour et à la demi-journée. C'est ce qui garde l'invariant du classeur.

---

## L'invariant

`Datas` porte **exactement deux lignes par jour**, sans trou ni doublon. Les
lignes `3 + 2k` et `4 + 2k` sont les deux demi-journées du k-ième jour.

C'est ce qui permet aux plages nommées des graphiques de se calculer par une
soustraction de dates, sans `EQUIV` ni recherche : un jour vaut deux lignes dans
`Datas`, une ligne dans `Calcul`.

**Ne jamais insérer ni supprimer une ligne dans `Datas`** : tout se décalerait,
et les deux graphiques montreraient les bonnes courbes aux mauvaises dates. La
feuille `Saisie` porte un contrôle qui le vérifie et le dit en clair.

---

## Fabriquer le classeur

```sh
pip install openpyxl

# à partir du carnet de relevés
python3 outils/construire_classeur.py Pression_Rena_Suivi.xlsx Pression_Rena.xlsx
```

### Vérifier ce qui a été fabriqué

Le classeur ne contient que des formules, sans valeur en cache : il faut le
faire recalculer avant de pouvoir le contrôler.

```sh
soffice --headless --convert-to xlsx --outdir recalc Pression_Rena.xlsx
python3 outils/verifier_classeur.py Pression_Rena_Suivi.xlsx recalc/Pression_Rena.xlsx
```

Le script recalcule tout de son côté, en Python, et compare : les mesures
reprises, l'invariant, les extrêmes du jour, les moyennes lissées, les quatre
séries empilées, le résumé, les étiquettes, et l'absence de toute valeur
d'erreur autre que `#N/A`. Environ 32 000 contrôles.

### Vérifier les modules VBA

```sh
python3 outils/verifier_vba.py src
```

Noms publics déclarés deux fois, contrôles appelés mais jamais créés, variable
locale qui masque une procédure, géométrie qui déborde, caractère qui ne
s'écrit pas en Windows-1252 — autant de fautes que VBA ne signale pas à la
compilation.

---

## Encodage des fichiers

Les `.bas` sont enregistrés en **Windows-1252 / CRLF**, l'encodage attendu par
l'éditeur VBA. Ne pas les convertir en UTF-8 : les accents des libellés et des
messages seraient déformés à l'import.
