# Installation pas à pas

## 1. Autoriser l'accès au projet VBA

Le générateur crée le UserForm en pilotant l'éditeur VBA. Excel bloque cet accès
par défaut ; il faut l'autoriser **une seule fois**, pour l'installation de
Microsoft 365 (le réglage n'est pas enregistré dans le classeur).

**Fichier ▸ Options ▸ Centre de gestion de la confidentialité ▸ Paramètres du
Centre de gestion de la confidentialité… ▸ Paramètres des macros** →
cocher **« Accès approuvé au modèle d'objet du projet VBA »** → **OK**.

Fermer puis rouvrir le classeur : le réglage n'est lu qu'au démarrage.

> Si vous préférez ne pas laisser cette option active, elle peut être décochée
> après la génération. Le formulaire, une fois créé, fonctionne sans elle.

## 2. Importer les modules

`Alt + F11` pour ouvrir l'éditeur VBA, puis pour chacun des six fichiers du
dossier `src/` : **Fichier ▸ Importer un fichier…** (`Ctrl + M`).

Ordre indifférent :

```
modClients_Theme.bas
modClients_Schema.bas
modClients_Donnees.bas
modClients_Adresses.bas
modClients_Formulaire.bas
modClients_Generateur.bas
modClients_Lancement.bas
```

Vérification : **Débogage ▸ Compiler VBAProject** ne doit signaler aucune erreur.
Le projet compile même si le UserForm n'existe pas encore.

## 3. Vérifier le classeur

Lancer la macro **`VerifierClasseur`** (`Alt + F8`). Elle contrôle la présence
des tableaux et des 21 colonnes attendues et indique si le formulaire est déjà
généré. Résultat attendu sur le classeur d'origine :

```
[OK] Tableau TblClients : feuille Clients, 354 fiches, 21 colonnes
[OK] Les 21 colonnes du schéma sont présentes
[OK] Tableau Tabl_Adresses : 500 lignes - recherche d'adresses
[OK] Tableau Tabl_Villes_CH : 5743 lignes - recherche de NPA (repli)
[OK] Tableau TblTxtStd : 4 lignes - textes de facture
[OK] Tableau TblInterv : 26 lignes - contrôle avant suppression
```

## 4. Générer le formulaire

Lancer **`GenererFormulaireClients`**. La procédure :

1. contrôle que `TblClients` existe et signale toute colonne absente du schéma ;
2. supprime une éventuelle version précédente de `UF_Clients` (après
   confirmation) ;
3. crée le UserForm, ses 71 contrôles et leur mise en forme ;
4. écrit le module de code du formulaire (78 procédures événementielles).

Un message confirme la création.

## 5. Ouvrir le formulaire

Lancer **`OuvrirGestionClients`**.

Pour un accès permanent, placer un bouton sur une feuille :
**Développeur ▸ Insérer ▸ Bouton (contrôle de formulaire)**, puis affecter la
macro `OuvrirGestionClients`.

## 6. Enregistrer

Enregistrer le classeur au format **`.xlsm`** (classeur Excel prenant en charge
les macros). Le formulaire et les modules y sont conservés : la génération n'est
à refaire que si le schéma ou la charte graphique changent.

---

## Dépannage

**« Excel refuse l'accès au projet VBA »**
L'étape 1 n'a pas été faite, ou le classeur n'a pas été rouvert depuis.

**Le formulaire s'affiche mais reste vide**
`TblClients` est introuvable ou vide. Lancer `VerifierClasseur`.

**Une colonne du tableau n'apparaît pas dans le formulaire**
Elle n'est pas décrite dans `modClients_Schema` ; le générateur le signale avant
de construire le formulaire. Voir « Ajouter un champ » dans
[`STRUCTURE.md`](STRUCTURE.md).

**Accents déformés dans les libellés**
Les fichiers `.bas` ont été convertis en UTF-8. Les réimporter depuis le dépôt
sans conversion (ils sont en Windows-1252).

**La fenêtre ne se déplace pas**
Le formulaire est sans barre de titre : il se déplace en faisant glisser le
bandeau bleu foncé du haut. Le `×` en haut à droite et la touche `Échap` le
ferment.

**Repartir de zéro**
`SupprimerFormulaireClients` retire `UF_Clients` du projet ; relancer ensuite
`GenererFormulaireClients`.
