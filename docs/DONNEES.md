# Deux fichiers : les données d'un côté, l'application de l'autre

Jusqu'ici tout vivait dans un seul `FlecheNettoyageSA2026.xlsm` : les macros, la
feuille d'accueil, et les six tableaux. Un seul fichier, donc une seule personne
à la fois, et une mise à jour du programme qui obligeait à recopier les données.

Désormais :

| | Fichier | Où | Contient |
|---|---|---|---|
| **Données** | `FlecheNettoyageSA-2026.xlsx` | OneDrive, dossier `FlecheNettoyageSA` | 5 onglets, **aucune macro** |
| **Application** | `FlecheNettoyageSA.xlsm` | le disque de chaque poste | l'accueil et les modules, **aucune donnée** |

**Un fichier de données par année.** `-2026`, puis `-2027`, puis `-2028` : les
interventions ne sont plus remises à zéro, l'année précédente reste entière et
se consulte quand on veut.

**L'application se remplace en écrasant un fichier.** Elle ne contient rien
qu'on puisse perdre.

---

## Le fichier de données

Cinq onglets, et rien d'autre :

| Onglet | Tableau |
|---|---|
| `Interventions` | `TblInterv` |
| `Clients` | `TblClients` |
| `Adresses` | `Tabl_Adresses` |
| `Liste_NPA_Suisse` | `Tabl_Villes_CH` |
| `Parametres` | `TblTxtStd` — les textes de facture |

Et deux **cellules nommées**, dans l'onglet `Parametres` :

| Nom | Contenu |
|---|---|
| `Datas_Version` | `1` — le numéro de version du schéma |
| `Objectif_annuel_CA` | l'objectif de chiffre d'affaires de **cette année-là** |

`Datas_Version` est comparé à ce que l'application attend. Le jour où une
colonne s'ajoutera à `TblInterv`, un poste resté sur une vieille application
recevra un message clair au lieu de voir les choses se casser en silence.

`Objectif_annuel_CA` est du côté des données parce qu'il change d'une année à
l'autre : chaque fichier porte le sien.

> **Le nom du fichier fait foi pour l'année.** Il n'y a plus de cellule
> `AnneeEnCours` : elle disait ce qu'on avait bien voulu y écrire, et recopier le
> classeur de l'an passé en oubliant de la changer suffisait à la faire mentir.

## Le fichier d'application

La feuille `Accueil`, les trente modules, et un onglet `Parametres` qui ne porte
que trois **cellules nommées** :

| Nom | Contenu |
|---|---|
| `Mot_de_passe` | le mot de passe du bouton *Unlock*, et de la prise de main sur les données |
| `TitreInterventions` | l'intitulé du bandeau du formulaire des interventions |
| `Dossier_Donnees` | *facultatif* — le chemin du dossier des données, écrit à la main |

> **Ne mettez pas ces cellules dans la feuille `Accueil`.** `GenererAccueil`
> efface son contenu à chaque exécution : elles disparaîtraient au premier
> redessin.

`Dossier_Donnees` est la porte de sortie. Laissée vide, l'application cherche
toute seule un dossier nommé `FlecheNettoyageSA` sous les racines de
synchronisation OneDrive du poste — variables d'environnement `OneDrive`,
`OneDriveCommercial`, `OneDriveConsumer`, plus deux niveaux de sous-dossiers,
ce qui couvre aussi les bibliothèques SharePoint. Si OneDrive n'est pas là où on
l'attend, écrivez le chemin dans cette cellule et l'application s'y tient.

---

## Faire la séparation, une fois

0. **Déverrouiller d'abord** : bouton **Unlock** de la feuille d'accueil. C'est
   l'étape qu'on oublie, et elle n'est pas facultative — voir l'encadré
   ci-dessous.
1. **Copier** `FlecheNettoyageSA2026.xlsm`. La copie deviendra le fichier de
   données, l'original l'application.
2. Dans la copie : supprimer la feuille `Accueil`, puis
   **Fichier ▸ Enregistrer sous ▸ Classeur Excel (`.xlsx`)** sous le nom
   `FlecheNettoyageSA-2026.xlsx`. Le format `.xlsx` retire les macros, c'est
   voulu : ce fichier ne doit plus en contenir.
3. Le déposer dans OneDrive, dans un dossier nommé **`FlecheNettoyageSA`**, et le
   partager avec les postes concernés.
4. Dans l'onglet `Parametres` de ce fichier : nommer une cellule
   **`Datas_Version`** et y écrire **`1`** ; vérifier que `Objectif_annuel_CA`
   y est bien.
5. Dans l'original : supprimer les cinq onglets de données, ne garder que
   `Accueil` et un onglet `Parametres` vide portant `Mot_de_passe` et
   `TitreInterventions`. Enregistrer sous `FlecheNettoyageSA.xlsm`.
6. Réimporter les modules de [`src/`](../src) — voir
   [`INSTALLATION.md`](INSTALLATION.md), **§ Remplacer un module** : VBA ne
   remplace jamais un module, il en ajoute un second.
7. Relancer **`GenererAccueil`** : le bandeau gagne sa ligne d'état et son
   sélecteur d'année.
8. Copier `FlecheNettoyageSA.xlsm` sur chaque poste.

Si quelque chose ne se trouve pas, lancez **`DiagnostiquerDatas`** (`Alt + F8`) :
il dit où l'application a cherché, ce qu'elle a trouvé, et qui tient le verrou.

### Pourquoi déverrouiller avant de copier

**Trois des réglages du kiosque sont enregistrés dans le classeur**, et non dans
Excel. Copier l'application pendant qu'elle est verrouillée les emporte donc
dans le fichier de données :

| Réglage | Portée | Dans le fichier ? |
|---|---|---|
| ruban caché (`SHOW.TOOLBAR`) | Excel | non |
| barre de formule, barre d'état | Excel | non |
| dimensions de la fenêtre | Excel | non |
| onglets, quadrillage, en-têtes, ascenseurs | **fenêtre du classeur** | **oui** |
| feuilles `xlSheetVeryHidden` | **feuille** | **oui** |
| structure protégée par mot de passe | **classeur** | **oui** |

Ouvert à la main dans Excel, un tel fichier ne montre **aucun onglet**, aucune
donnée — la surface est nue — et la moitié du ruban est grisée. Et
« très masquée » **ne se défait pas par le menu Afficher** : seul le code, ou
l'éditeur VBA, y revient. Dans un fichier sans macro, c'est déroutant.

Le bouton *Unlock* (`Accueil_Deverrouiller`) défait les trois d'un coup :
il réaffiche les feuilles, lève la protection de structure et rend à la fenêtre
ses onglets. **Copier ensuite.**

> **L'application le répare aussi toute seule.** `NormaliserDonnees` remet le
> fichier d'aplomb à chaque ouverture en écriture, et `Datas_Fermer`
> l'enregistre : un fichier passé une fois par l'application s'ouvre ensuite
> normalement dans Excel. Ce fichier doit rester consultable **sans**
> l'application — c'est tout l'intérêt de n'y avoir mis aucune macro.

## Le fichier de données s'ouvre et on ne voit RIEN

Fond gris ou noir, aucun onglet, aucune donnée, et presque tout le ruban grisé.
**La cause la plus fréquente, et de loin : le fichier a été enregistré fenêtre
masquée.**

L'application ouvre le classeur de données **fenêtre masquée**, pour qu'il ne
tombe pas par-dessus le kiosque avec ses onglets et son quadrillage. Or l'état
de la fenêtre est écrit **dans le fichier**, en clair :

```xml
<workbookView visibility="hidden" ... showSheetTabs="0" ...>
```

Un classeur enregistré ainsi s'ouvre ensuite **sans aucune fenêtre**. Il n'y a
donc rien à montrer, et faute de classeur actif Excel grise la quasi-totalité du
ruban. Le fichier paraît mort, et rien dans Excel ne dit pourquoi.

**La réparation tient en un geste :**

> **Affichage ▸ Fenêtre ▸ Afficher** *(Unhide)* → choisir le classeur → **OK**,
> puis **enregistrer**.

Ce n'est pas *Affichage ▸ Afficher les feuilles* : c'est la **fenêtre** du
classeur qui est masquée, pas ses feuilles.

L'application le répare aussi toute seule : `MontrerFenetre` rend sa fenêtre au
classeur **avant chaque enregistrement**, et `NormaliserDonnees` la rend à un
fichier qui arrive déjà masqué. Laissez l'application l'ouvrir une fois en
écriture et la fermer par *Quitter*.

### Ce qui pouvait s'enregistrer masqué

Deux chemins, tous deux bouchés :

- **à la fermeture**, `Datas_Fermer` enregistrait sans rendre la fenêtre ;
- **par Excel lui-même**, quand le classeur de données restait ouvert sans que
  l'application sache encore qu'il était là. `mClasseur` est une variable de
  module : réimporter un module, taper *Fin* dans l'éditeur, ou une erreur non
  interceptée, et **VBA remet à zéro tout l'état du projet**. Le classeur restait
  alors ouvert, fenêtre masquée, orphelin. Excel finissait par demander s'il
  fallait l'enregistrer — et un « oui » gravait la fenêtre masquée dans le
  fichier.

> **Deux questions « enregistrer ? » en quittant le kiosque, c'est le signe de
> cet orphelin.** Normalement il n'y en a aucune : l'application enregistre et
> ferme les données elle-même. `ClasseurOrphelin` retrouve maintenant le
> classeur par son nom quand la variable l'a perdu.

### Si les feuilles aussi sont masquées

Cela arrive quand le fichier de données a été **copié depuis l'application
verrouillée** — voir l'encadré plus haut. Les feuilles sont alors
*très masquées*, et **le menu Afficher ne les propose pas** :

1. **Révision ▸ Protéger le classeur** → décocher. Le mot de passe est celui de
   la cellule `Mot_de_passe`.
2. **Alt + F11**, puis **Ctrl + G** pour la fenêtre *Exécution*. Taper cette
   ligne et **Entrée** — elle fonctionne dans un `.xlsx`, qui ne peut pas
   *stocker* de macro mais sait toujours en exécuter une tapée là :

   ```vba
   For Each s In ActiveWorkbook.Sheets: s.Visible = True: Next
   ```

3. **Fichier ▸ Options ▸ Options avancées ▸ Afficher les options pour ce
   classeur** → cocher **Afficher les onglets de classeur**.
4. **Affichage** → cocher **Quadrillage** et **Titres**.
5. Enregistrer.

### Si c'est Excel lui-même qui reste nu

Ruban réduit aux seuls noms d'onglets, pas de barre de formule, **et pour tous
les classeurs** : ces trois réglages-là appartiennent à **Excel**, pas au
fichier, et Excel les garde d'une séance à l'autre.

- **Ctrl + F1** ramène le ruban ;
- **Affichage ▸ Barre de formule** la barre de formule.

`Accueil_Arreter` les rend désormais **en premier**, avant même de fermer les
données : une erreur en fermant les données ne doit pas pouvoir coûter le ruban
de tous vos classeurs.

---

## Un seul poste à la fois

Excel sait refuser la seconde ouverture d'un classeur — mais pas sur OneDrive.
Là, il ouvre à chacun sa copie, laisse les deux postes écrire, puis fabrique un
« fichier en conflit » que personne ne relit jamais. Le verrou d'Excel ne protège
donc de rien ici, et celui de OneDrive encore moins.

D'où **notre propre verrou** : un fichier posé à côté des données.

```
FlecheNettoyageSA-2026.xlsx   ->   FlecheNettoyageSA-2026.verrou
```

Il tient sur une ligne — qui, sur quel poste, depuis quand :

```
rdem|PC-BUREAU|01.10.2026 14:32:05| 46296.6056
```

- **Sa seule présence interdit l'écriture aux autres postes.** Le second
  utilisateur voit qui travaille et depuis quand, et se voit proposer
  **d'ouvrir en consultation seule** plutôt que d'être renvoyé sans rien.
- **Il se périme en quatre heures.** Excel qui se ferme mal, un poste qu'on
  éteint, et il resterait là pour toujours. Quatre heures suffisent parce qu'**il
  se renouvelle à chaque écriture** : un poste qui travaille vraiment le
  rajeunit sans cesse, seul un poste parti le laisse vieillir.
- **Un verrou laissé par VOTRE poste se reprend sans rien demander.** Même
  utilisateur, même machine : la séance qui l'a posé n'existe plus, et aucun
  collègue ne peut être derrière. Poser la question « voulez-vous prendre la
  main ? » à quelqu'un qui redémarre sa propre application ne lui apprend rien,
  et l'habitue à répondre oui sans lire — précisément ce qu'il ne faut pas, le
  jour où le verrou sera vraiment celui d'un autre.
- **On peut le forcer**, derrière le mot de passe, quand on sait que l'autre
  poste a été éteint sans fermer l'application.
- Il est **relu après écriture** : deux postes peuvent l'écrire presque en même
  temps, c'est OneDrive qui tranche, après coup et sans le dire. On relit donc
  ce qu'on vient d'écrire, et si la signature n'est plus la nôtre, c'est l'autre
  qui a la main.

Ce verrou n'empêche pas quelqu'un qui ouvrirait le fichier de données à la main,
et ne voit rien d'un poste resté hors ligne. C'était le besoin : empêcher deux
personnes de se marcher dessus, pas se défendre.

### Si un verrou reste quand même en travers

**Supprimez le fichier `.verrou`** dans le dossier partagé, à côté des données.
C'est un fichier texte ordinaire : ouvrez-le d'abord dans le Bloc-notes pour
voir qui le tient, et depuis quand.

`DiagnostiquerDatas` le dit aussi, en une ligne — y compris quand il a été
laissé par une séance précédente de votre poste et sera donc repris tout seul.

> **Un verrou qui ne se rend pas bloque d'abord VOTRE poste**, et c'est à quoi on
> le reconnaît : au démarrage, l'application vous demande si vous voulez prendre
> la main sur vous-même. Deux défauts l'avaient provoqué, tous deux corrigés —
> `Datas_Fermer` sortait avant de rendre le verrou quand elle ne trouvait pas de
> classeur à fermer, ce qui est justement le cas où il reste ; et `Verrou_Rendre`
> exigeait la marque exacte, perdue dès que l'état du projet VBA est remis à
> zéro.
>
> **Tant qu'un verrou traîne, le fichier de données ne se répare pas** :
> `NormaliserDonnees` ne travaille qu'en écriture, et le verrou force la
> consultation seule.

---

## Le sélecteur d'année

Dans le bandeau de l'accueil, **l'année et la ligne qui la commente sont
cliquables**.

```
                                                      2026
            données ouvertes en écriture — cliquer pour changer d'année
```

- **L'année ouverte par défaut est le fichier le plus récent présent** dans le
  dossier — et non `Year(Date)`. Le 1er janvier, l'application pointerait sinon
  vers un fichier qui n'existe pas encore ; elle ouvre ainsi toujours quelque
  chose de réel, et le sélecteur ne s'ouvre que si on le demande.
- **L'année la plus récente seule s'ouvre en écriture**, par une personne à la
  fois.
- **Une année passée s'ouvre en consultation seule, sans verrou** : personne n'y
  écrit, plusieurs personnes peuvent donc la relire en même temps.

En consultation seule, la ligne d'état passe à l'ambre et **les boutons qui
écrivent restent gris** — *Ajouter*, *Modifier*, *Supprimer*, *Facturer*. Le
refus doit se voir avant la saisie, pas après.

---

## Changer d'année

1. Copier le fichier de l'année écoulée sous le nom de la nouvelle :
   `FlecheNettoyageSA-2027.xlsx`.
2. Y vider `TblInterv` de ses lignes, et ajuster `Objectif_annuel_CA`.
3. C'est tout. Au prochain démarrage, l'application ouvre `-2027`, parce que
   c'est le fichier le plus récent ; `-2026` reste consultable d'un clic sur
   l'année.

Ne supprimez pas les anciens fichiers : ils ne gênent rien, et ce sont eux que
le sélecteur propose.

---

## Quand cela ne va pas

| Symptôme | Où regarder |
|---|---|
| « Le dossier des données est introuvable » | `DiagnostiquerDatas` ; sinon écrire le chemin dans `Dossier_Donnees` |
| « Ces données ne sont pas de la même version » | l'application du poste est plus ancienne que le fichier : la remplacer |
| Les données s'ouvrent toujours en consultation | quelqu'un d'autre les tient — `DiagnostiquerDatas` dit qui |
| « Prendre la main ? » alors que c'est votre propre poste | un verrou laissé par une séance précédente : il se reprend seul ; sinon supprimez le fichier `.verrou` |
| Les formulaires s'ouvrent vides | le fichier de données est ouvert mais ses tableaux manquent : `VerifierClasseur` |
| Le fichier de données s'ouvre sans onglets, ruban grisé | sa **fenêtre** est masquée : *Affichage ▸ Fenêtre ▸ Afficher* |
| Deux questions « enregistrer ? » en quittant le kiosque | un classeur de données orphelin — voir la section ci-dessus |
| Excel reste nu pour tous les classeurs | **Ctrl + F1**, et *Affichage ▸ Barre de formule* |
| Les images des tuiles manquent | `DiagnostiquerChemins` — c'est le dossier de l'**application**, pas celui des données |
