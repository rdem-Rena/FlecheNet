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
modClients_Theme.bas        modInterv_Theme.bas
modClients_Schema.bas       modInterv_Schema.bas
modClients_Donnees.bas      modInterv_Donnees.bas
modClients_Adresses.bas     modInterv_Calendrier.bas
modClients_Formulaire.bas   modInterv_Formulaire.bas
modClients_Generateur.bas   modInterv_Generateur.bas
modClients_Lancement.bas    modInterv_Lancement.bas
```

Les modules `modInterv_*` s'appuient sur les `modClients_*` — palette,
typographie et utilitaires partagés : importer les sept de gauche même si seul
le formulaire des interventions vous intéresse.

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

## Le classeur dans un dossier OneDrive ou SharePoint

Le classeur fonctionne dans un dossier partagé et synchronisé, à une condition :
**l'ouvrir depuis l'Explorateur**, dans le dossier synchronisé, et non depuis le
site web ou depuis Excel en ligne.

Pourquoi cette condition : dans un dossier synchronisé, `ThisWorkbook.Path` ne
rend pas un chemin mais une adresse web —
`https://contoso.sharepoint.com/sites/Equipe/Documents partagés/Flèche` — que ni
`Dir$`, ni `Open`, ni `LoadPicture` ne savent lire. Le module **`modChemins`**
retrouve le chemin local correspondant en confrontant la fin de cette adresse
aux racines de synchronisation du poste, puis tout le reste du code passe par
lui. C'est notamment ce qui permet à l'image de fond des tuiles, dans
`Images\CartePremium.jpg`, de s'afficher.

Si une image ne s'affiche toujours pas, lancer **`DiagnostiquerChemins`**
(module `modChemins`) : le message dit ce que le classeur croit être son
dossier, ce que le programme en a tiré, et si l'image s'y trouve.

Trois causes possibles quand le dossier local reste introuvable :

- le classeur a été ouvert depuis le navigateur — le rouvrir depuis
  l'Explorateur ;
- le dossier n'est pas synchronisé sur ce poste — cliquer **Synchroniser** dans
  la bibliothèque SharePoint ;
- le client OneDrive n'est pas installé — les variables d'environnement
  `OneDrive` et `OneDriveCommercial` sont alors absentes, et il n'y a aucune
  racine où chercher.

Le dossier `Images` doit rester **à côté du classeur**, comme en local.

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

**« La génération a échoué : 75 » ou « objet spécifié introuvable »**
Ce message venait d'une version antérieure du générateur, qui supprimait le
formulaire avant de le recréer sous le même nom. VBA ne libère le nom d'un
composant supprimé qu'au retour à Excel : la première génération réussissait,
toutes les suivantes échouaient. Le générateur reconstruit désormais le
formulaire sur place, sans jamais le supprimer.

Si vous rencontrez encore ce message avec la version à jour :
1. fermez puis rouvrez le classeur — cela solde les suppressions en attente ;
2. lancez `NettoyerFormulairesOrphelins`, qui retire les `UserForm1`,
   `UserForm2`… qu'une tentative interrompue aurait laissés (seuls les
   formulaires au nom automatique, sans contrôle ni code, sont concernés) ;
3. relancez `GenererFormulaireClients`.

**Repartir de zéro**
`SupprimerFormulaireClients` retire `UF_Clients` du projet ; relancer ensuite
`GenererFormulaireClients`.
