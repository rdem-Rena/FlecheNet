# Installation

## 1. Ouvrir le classeur

`Pression_Rena.xlsx` s'ouvre tel quel, et **tout fonctionne sans macro** : les
deux graphiques, les trois réglages, le résumé, les quinze dernières mesures.

Pour saisir sans le formulaire : aller sur `Datas`, à la ligne du jour — le
calendrier est déjà écrit jusqu'au 05.01.2028 — et remplir `Heure`, `Systole`,
`Diastole`, `Pouls`, et `Poids` le matin.

À la première ouverture, Excel recalcule l'ensemble : quelques secondes, une
seule fois.

---

## 2. Ajouter le formulaire de saisie

Le formulaire `UF_Pression` demande des macros, donc un classeur `.xlsm`.

### a. Autoriser l'accès au projet VBA

**Fichier > Options > Centre de gestion de la confidentialité > Paramètres du
Centre de gestion de la confidentialité > Paramètres des macros** → cocher
**« Accès approuvé au modèle d'objet du projet VBA »**. Fermer et rouvrir Excel.

C'est ce qui permet au générateur de créer le formulaire. Une fois le formulaire
créé, l'option n'est plus nécessaire pour s'en servir.

### b. Enregistrer en `.xlsm`

**Fichier > Enregistrer sous**, type **Classeur Excel prenant en charge les
macros (\*.xlsm)**. Un `.xlsx` ne peut pas porter de macro, et Excel les retire
en silence à l'enregistrement.

### c. Importer les modules

**Alt + F11** pour ouvrir l'éditeur VBA, puis **Fichier > Importer un
fichier…** pour importer les six modules du dossier [`../src`](../src) :

| Module | Rôle |
|---|---|
| `modPression_Schema` | Feuilles, colonnes, et la conversion date → ligne |
| `modPression_Theme` | Charte graphique et géométrie du formulaire |
| `modPression_Donnees` | Lecture et écriture de `Datas`, conversions, contrôles |
| `modPression_Generateur` | Construit le formulaire et son module de code |
| `modPression_Formulaire` | Comportement du formulaire |
| `modPression_Lancement` | `OuvrirSaisiePression`, `VerifierClasseur` |

L'ordre n'a pas d'importance : les modules ne s'appellent qu'à l'exécution.

### d. Générer le formulaire

**Alt + F8**, puis lancer **`GenererFormulairePression`**. Un message annonce le
nombre de contrôles posés.

### e. S'en servir

**Alt + F8** → **`OuvrirSaisiePression`**. Ou, plus commode, poser un bouton sur
la feuille `Saisie` : **Développeur > Insérer > Bouton**, et lui affecter
`OuvrirSaisiePression`.

---

## Ce que fait le formulaire

**La date et la demi-journée désignent une ligne**, et une seule. En changer
recharge la fiche depuis cette ligne : ce qui est affiché est toujours ce que
porte le classeur, et **Enregistrer** écrase cette ligne-là. Aucune ligne n'est
jamais insérée ni supprimée.

- **Aujourd'hui** pose la date, la demi-journée et l'heure de maintenant. Avant
  midi c'est le matin, après c'est le soir.
- L'**activité** propose celles déjà employées, sans doublon et rangées par
  ordre alphabétique — `Ext - 9.0 km` se retrouve sans le retaper, et les
  orthographes ne se multiplient pas. Une nouvelle peut être tapée.
- Sous les valeurs, un **repère de couleur** dit si la mesure passe au-dessus
  des seuils du classeur. Ce n'est pas un diagnostic : le formulaire rapproche
  deux nombres de deux autres, et les seuils sont ceux des lignes de repère des
  graphiques, réglés sur la feuille `Saisie`.
- Un clic sur une ligne du tableau du bas ramène cette demi-journée dans la
  fiche.
- **Effacer** vide la fiche, **pas le classeur**. Pour effacer une mesure,
  vider les champs puis **Enregistrer**.

Les bornes de saisie — systole 60–260, diastole 30–160, pouls 30–220, poids
30–250 — n'ont rien de médical : elles écartent les fautes de frappe. Le fichier
d'origine porte un pouls à 557, qui est un 57 tapé trop longtemps.

---

## Régénérer le formulaire

Changer une couleur ou une dimension dans `modPression_Theme`, puis relancer
`GenererFormulairePression` : le formulaire est vidé et reconstruit sur place.

**Ne pas le supprimer avant.** VBA ne rend le nom d'un composant supprimé qu'au
prochain chargement du classeur : une suppression à la main oblige à fermer puis
rouvrir le classeur avant de pouvoir régénérer. Le générateur le détecte et le
dit, plutôt que de laisser passer une erreur 75 incompréhensible.

---

## Quand quelque chose ne va pas

Lancer **`VerifierClasseur`** (Alt + F8). Il dit :

- si les quatre feuilles sont là ;
- les bornes du calendrier et le nombre de lignes ;
- **si l'invariant tient** — deux lignes par jour, sans trou ;
- si les plages nommées des graphiques répondent ;
- si le formulaire est généré.

Le message du **contrôle de structure** est aussi affiché en haut du formulaire
et sur la feuille `Saisie` : si une ligne a été insérée ou supprimée dans
`Datas`, les deux graphiques sont décalés, et c'est le seul endroit où cela se
voit.

---

## Encodage des fichiers

Les `.bas` sont enregistrés en **Windows-1252 / CRLF**, l'encodage attendu par
l'éditeur VBA. Ne pas les convertir en UTF-8 : les accents des libellés et des
messages seraient déformés à l'import.
