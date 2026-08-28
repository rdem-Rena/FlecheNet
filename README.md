# FlècheNet — Formulaire de gestion des clients

Procédures VBA qui **génèrent** et font fonctionner le UserForm `UF_Clients`, un
formulaire de saisie complet pour le tableau **`TblClients`** de l'onglet
**Clients** du classeur `FlecheNettoyageSACL01.xlsm`.

Le formulaire n'est pas dessiné à la main : il est construit par code à partir du
schéma décrit dans `modClients_Schema`. Un champ ajouté au schéma devient une
zone de saisie dans le formulaire à la régénération suivante.

---

## Démarrage rapide

1. Dans Excel : **Fichier ▸ Options ▸ Centre de gestion de la confidentialité ▸
   Paramètres du Centre de gestion de la confidentialité ▸ Paramètres des macros**
   → cocher **« Accès approuvé au modèle d'objet du projet VBA »**, puis fermer
   et rouvrir le classeur.
2. **Alt + F11** pour ouvrir l'éditeur VBA, puis **Fichier ▸ Importer un
   fichier…** pour importer les 6 modules du dossier [`src/`](src).
3. Lancer la macro **`GenererFormulaireClients`** → le formulaire `UF_Clients`
   est créé.
4. Lancer la macro **`OuvrirGestionClients`** → le formulaire s'affiche.

Détails et dépannage : [`docs/INSTALLATION.md`](docs/INSTALLATION.md).

---

## Ce que fait le formulaire

**En haut** — la fiche client : les 21 colonnes de `TblClients`, réparties sur une
grille de 4 colonnes × 5 lignes.

- `Clef_BD` et `Date_Crea` sont **visibles mais non saisissables** : le programme
  attribue la clef (`CL1`, `CL2`, … `CL` + plus grand numéro existant + 1) et la
  date du jour au moment de l'ajout.
- `Titre` est un **menu déroulant** proposant *Monsieur* et *Madame* (les
  civilités déjà présentes dans le tableau — *Famille*, *Docteur*,
  *Monsieur et Madame* — restent sélectionnables pour ne pas perdre l'existant).
- `Adresse` et `NoPost` sont alimentés depuis l'onglet **Adresses**
  (`Tabl_Adresses`) : choisir une rue renseigne automatiquement **NoPost**,
  **Ville** et **Cant**. Un NPA absent de `Tabl_Adresses` est cherché dans
  `Tabl_Villes_CH` (onglet *Liste_NPA_Suisse*).
- `Texte_Facture` propose les textes standards de l'onglet **Parametres**
  (`TblTxtStd`).
- `TVA` et `Forfait` sont des cases à cocher (valeurs booléennes, comme dans le
  tableau).

**Au milieu** — la barre de filtrage : un menu déroulant choisit la colonne
(**Entreprise**, **Nom**, **Adresse**) et une zone de texte filtre le tableau au
fur et à mesure de la frappe, sans tenir compte des accents ni de la casse.

**En dessous** — le tableau des enregistrements : `Entreprise`, `Titre`, `Nom`,
`Prenom`, `Adresse`, `No`, `NoPost`, `Ville`, `Cant`, `Tx_hrs_Forf` — tri par clic
sur un en-tête, sélection d'une ligne pour charger la fiche dans les champs du
haut.

**En bas** — les cinq boutons demandés :

| Bouton | Raccourci | Effet |
|---|---|---|
| **Ajouter** | `Alt+A` | Crée une fiche avec les champs saisis |
| **Modifier** | `Alt+M` | Enregistre les modifications de la fiche sélectionnée |
| **Supprimer** | `Alt+S` | Supprime la fiche après confirmation |
| **Effacer** | `Alt+E` | Vide les zones de saisie, sans toucher au tableau |
| **Quitter** | `Alt+Q` / `Échap` | Ferme le formulaire |

Contrôles avant écriture : nom **ou** entreprise obligatoire, NPA à 4 chiffres,
ID Cresus et taux horaire numériques, adresse de courriel plausible,
avertissement en cas d'ID Cresus déjà utilisé. À la suppression, le programme
signale les interventions de `TblInterv` qui référencent le client.

---

## Organisation du code

| Module | Rôle |
|---|---|
| [`modClients_Schema`](src/modClients_Schema.bas) | Description des 21 champs : colonne, libellé, type de contrôle, position, contraintes. **Source de vérité du formulaire.** |
| [`modClients_Theme`](src/modClients_Theme.bas) | Charte graphique : palette, typographie, géométrie. |
| [`modClients_Generateur`](src/modClients_Generateur.bas) | Crée le UserForm, ses 71 contrôles et son module de code. |
| [`modClients_Donnees`](src/modClients_Donnees.bas) | Lecture / écriture de `TblClients`, clefs, contrôles d'intégrité. |
| [`modClients_Adresses`](src/modClients_Adresses.bas) | Recherches dans `Tabl_Adresses`, `Tabl_Villes_CH`, `TblTxtStd`. |
| [`modClients_Formulaire`](src/modClients_Formulaire.bas) | Comportement du formulaire : filtrage, tri, CRUD, habillage. |
| [`modClients_Lancement`](src/modClients_Lancement.bas) | `OuvrirGestionClients`, `VerifierClasseur`. |

Chaque procédure porte un en-tête qui dit ce qu'elle fait, ce qu'elle attend et
ce qu'elle renvoie ; les passages où VBA se comporte de façon inattendue sont
commentés sur place.

Le module de code du UserForm ne contient **que** des procédures
événementielles d'une ligne, écrites par le générateur ; toute la logique reste
dans `modClients_Formulaire`. Conséquence pratique : le projet compile même
avant que le formulaire existe, et le formulaire peut être régénéré à volonté
sans rien réécrire.

Mapping colonnes ↔ contrôles et plan du formulaire :
[`docs/STRUCTURE.md`](docs/STRUCTURE.md).

Pour ajuster l'aspect du formulaire — couleurs, tailles, position de chaque
contrôle — la fiche de réglages [`docs/REGLAGES.md`](docs/REGLAGES.md) recense
chaque valeur, dit ce qu'elle commande et où elle se trouve.

---

## Encodage des fichiers

Les fichiers `.bas` sont enregistrés en **Windows-1252 / CRLF**, l'encodage
attendu par l'éditeur VBA. Ne pas les convertir en UTF-8 : les accents des
libellés et des messages seraient déformés à l'import.
