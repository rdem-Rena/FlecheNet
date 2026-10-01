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
- **On peut le forcer**, derrière le mot de passe, quand on sait que l'autre
  poste a été éteint sans fermer l'application.
- Il est **relu après écriture** : deux postes peuvent l'écrire presque en même
  temps, c'est OneDrive qui tranche, après coup et sans le dire. On relit donc
  ce qu'on vient d'écrire, et si la signature n'est plus la nôtre, c'est l'autre
  qui a la main.

Ce verrou n'empêche pas quelqu'un qui ouvrirait le fichier de données à la main,
et ne voit rien d'un poste resté hors ligne. C'était le besoin : empêcher deux
personnes de se marcher dessus, pas se défendre.

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
| Les formulaires s'ouvrent vides | le fichier de données est ouvert mais ses tableaux manquent : `VerifierClasseur` |
| Les images des tuiles manquent | `DiagnostiquerChemins` — c'est le dossier de l'**application**, pas celui des données |
