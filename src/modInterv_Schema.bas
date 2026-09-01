Attribute VB_Name = "modInterv_Schema"
Option Explicit
'==============================================================================
' modInterv_Schema
'------------------------------------------------------------------------------
' Description du tableau TblInterv : source de vérité à partir de laquelle le
' formulaire UF_Interventions est généré, un champ par entrée ci-dessous.
'
' Pour ajouter un champ : ajouter la colonne au tableau Excel, ajouter une ligne
' DefInterv, incrémenter NB_CHAMPS_INTERV, puis relancer
' GenererFormulaireInterventions.
'==============================================================================

'--- Noms des objets du classeur ----------------------------------------------
Public Const NOM_TABLE_INTERVENTIONS As String = "TblInterv"
Public Const NOM_FEUILLE_STATS As String = "Statistiques"
Public Const NOM_TABLE_GRAPH As String = "Tableau7"   ' valeurs mensuelles, feuille Statistiques
Public Const NOM_FORM_INTERV As String = "UF_Interventions"
Public Const NOM_FORM_CALENDRIER As String = "UF_Calendrier"

'--- Cellules nommées ---------------------------------------------------------
Public Const CEL_TITRE As String = "TitreInterventions"
Public Const CEL_ANNEE As String = "AnneeEnCours"

'--- Colonnes de TblInterv ----------------------------------------------------
Public Const IC_NO As String = "NoInterv"
Public Const IC_DATE As String = "Date"
Public Const IC_CLIENT As String = "Client_No"
Public Const IC_ENTREPRISE As String = "Entreprise"
Public Const IC_TITRE As String = "Titre"
Public Const IC_NOM As String = "Nom"
Public Const IC_PRENOM As String = "Prenom"
Public Const IC_HEURES As String = "Nb_Hres"
Public Const IC_CA As String = "CA"
Public Const IC_TAUX As String = "Taux/Forfait"
Public Const IC_PERS As String = "Nb_Pers"
Public Const IC_TVA As String = "TVA"
Public Const IC_FORFAIT As String = "Forfait"
Public Const IC_TEXTE As String = "Texte_Facture"
Public Const IC_COMMENT As String = "Commentaires"
Public Const IC_FACTURE As String = "No_Facture"

Public Const PREFIXE_NO_INTERV As String = "IN"

'--- Types de contrôle --------------------------------------------------------
' Aux trois types du formulaire des clients s'ajoute la zone de date, qui est
' une zone de texte doublée d'un bouton ouvrant le calendrier.
Public Const ITYPE_TEXTE As String = "T"
Public Const ITYPE_LISTE As String = "C"      ' menu déroulant simple
Public Const ITYPE_CASE As String = "K"
Public Const ITYPE_DATE As String = "D"       ' zone de date + sélecteur
Public Const ITYPE_AUTO As String = "A"       ' liste filtrée au fil de la frappe

'--- Contraintes de saisie ----------------------------------------------------
Public Const ISAISIE_LIBRE As Long = 0
Public Const ISAISIE_ENTIER As Long = 1       ' chiffres seuls
Public Const ISAISIE_HEURES As Long = 2       ' hhh:mm
Public Const ISAISIE_MONTANT As Long = 3      ' chiffres et un séparateur décimal

Public Const NB_CHAMPS_INTERV As Long = 16

'--- Tableau du formulaire ----------------------------------------------------
' Réunit dans une même case deux colonnes de TblInterv : « Nom+Prenom ». Chacune
' garde sa propre colonne dans le tableau Excel, seul l'affichage les rassemble.
Public Const ICL_SEPARATEUR As String = "+"

' Premier choix du filtre de mois, celui qui ne filtre rien.
Public Const TOUS_LES_MOIS As String = "Tous les mois"

'==============================================================================
' Définition d'un champ du formulaire
'==============================================================================
Public Type ChampInterv
    Colonne As String       ' nom exact de la colonne dans TblInterv
    Libelle As String
    TypeCtrl As String
    Verrouille As Boolean   ' True = géré par le programme, non saisissable
    Ligne As Long           ' ligne de grille, 1 à 4
    Col As Long             ' colonne de grille, 1 à 6
    Blocs As Long           ' nombre de colonnes occupées en largeur (1 par défaut)
    NbLignes As Long        ' nombre de lignes occupées en hauteur (1 par défaut)
    Moitie As Long          ' 0 = colonne entière, 1 = moitié gauche, 2 = moitié droite
                            ' (deux cases à cocher tiennent dans une colonne)
    Saisie As Long          ' ISAISIE_*
    Aide As String
End Type

Private mChamps() As ChampInterv
Private mCharges As Boolean

'==============================================================================
' Schéma complet du formulaire
'==============================================================================
Public Function ObtenirChampsInterv() As ChampInterv()
    If Not mCharges Then ConstruireSchemaInterv
    ObtenirChampsInterv = mChamps
End Function

'------------------------------------------------------------------------------
' Remplit le schéma. C'est LA table à modifier pour déplacer, renommer ou
' verrouiller un champ ; le générateur et le formulaire s'y adaptent seuls.
'
'   n°  colonne, libellé, type, verrouillé, ligne, colonne, blocs, lignes,
'       moitié, contrainte de saisie, info-bulle
'
' La grille compte quatre lignes et six colonnes, groupées en trois régions de
' deux colonnes. Le schéma est en caractères ASCII : les traits fins d'Unicode
' n'existent pas en Windows-1252, format dans lequel ces fichiers sont stockés.
'
'   +-- région 1 : le client --+-- région 2 : l'intervention -+- région 3 ---+
'   | N° client   | Titre      | N° intervention | Date       | Texte de     |
'   | Entreprise .............>| N° de facture   | Heures     | facture ....>|
'   | Nom ....................>| Chiffre d'aff.  | Personnes  | Commentaires |
'   | Prénom .................>| [ ] TVA         | [ ] Forfait| ............>|
'   +--------------------------+------------------------------+--------------+
'
' Les pointillés marquent un champ qui s'étale sur les deux colonnes de sa
' région ; dans la troisième, les deux champs occupent aussi deux lignes.
'
' L'identité du client d'abord, puis les chiffres de l'intervention, puis les
' textes libres : c'est l'ordre dans lequel une fiche se remplit.
'------------------------------------------------------------------------------
Private Sub ConstruireSchemaInterv()
    ReDim mChamps(1 To NB_CHAMPS_INTERV)

    ' --- région 1 : le client, colonnes 1 et 2 -------------------------------
    DefInterv mChamps, 1, IC_CLIENT, "N° client", ITYPE_TEXTE, True, 1, 1, 1, 1, 0, ISAISIE_LIBRE, _
        "Renseigné automatiquement en choisissant une entreprise ou un nom"
    DefInterv mChamps, 2, IC_TITRE, "Titre", ITYPE_LISTE, False, 1, 2, 1, 1, 0, ISAISIE_LIBRE, _
        "Civilité du client"
    DefInterv mChamps, 3, IC_ENTREPRISE, "Entreprise", ITYPE_AUTO, False, 2, 1, 2, 1, 0, ISAISIE_LIBRE, _
        "Tapez les premières lettres : la liste des entreprises de TblClients se filtre au fil de la frappe"
    DefInterv mChamps, 4, IC_NOM, "Nom", ITYPE_AUTO, False, 3, 1, 2, 1, 0, ISAISIE_LIBRE, _
        "Tapez les premières lettres : la liste des noms de TblClients se filtre au fil de la frappe"
    DefInterv mChamps, 5, IC_PRENOM, "Prénom", ITYPE_TEXTE, False, 4, 1, 2, 1, 0, ISAISIE_LIBRE, _
        "Prénom du client"

    ' --- région 2 : l'intervention, colonnes 3 et 4 ---------------------------
    DefInterv mChamps, 6, IC_NO, "N° intervention", ITYPE_TEXTE, True, 1, 3, 1, 1, 0, ISAISIE_LIBRE, _
        "Index du tableau, attribué automatiquement par le programme"
    DefInterv mChamps, 7, IC_DATE, "Date", ITYPE_DATE, False, 1, 4, 1, 1, 0, ISAISIE_LIBRE, _
        "Date de l'intervention. Cliquez sur l'icône pour ouvrir le calendrier."
    DefInterv mChamps, 8, IC_FACTURE, "N° de facture", ITYPE_TEXTE, True, 2, 3, 1, 1, 0, ISAISIE_LIBRE, _
        "Attribué par la facturation, pas depuis ce formulaire"
    DefInterv mChamps, 9, IC_HEURES, "Heures", ITYPE_TEXTE, False, 2, 4, 1, 1, 0, ISAISIE_HEURES, _
        "Durée au format heures:minutes, par exemple 420:00"
    DefInterv mChamps, 10, IC_CA, "Chiffre d'affaires", ITYPE_TEXTE, True, 3, 3, 1, 1, 0, ISAISIE_MONTANT, _
        "Calculé au fil de la saisie : heures x personnes x taux, ou le taux seul au forfait"
    DefInterv mChamps, 11, IC_PERS, "Personnes", ITYPE_TEXTE, False, 3, 4, 1, 1, 0, ISAISIE_ENTIER, _
        "Nombre de personnes intervenues, de 1 à 99"

    ' TVA et Forfait se partagent la colonne 3 : deux cases courtes n'ont pas
    ' besoin d'une colonne chacune, et les garder côte à côte libère la
    ' colonne 4 pour le taux, qu'elles commandent.
    DefInterv mChamps, 12, IC_TVA, "TVA", ITYPE_CASE, False, 4, 3, 1, 1, 1, ISAISIE_LIBRE, _
        "Le client est assujetti à la TVA"
    DefInterv mChamps, 13, IC_FORFAIT, "Forfait", ITYPE_CASE, False, 4, 3, 1, 1, 2, ISAISIE_LIBRE, _
        "Au forfait, le chiffre d'affaires vaut le taux seul, sans multiplier par les heures"
    DefInterv mChamps, 14, IC_TAUX, "Taux / forfait", ITYPE_TEXTE, False, 4, 4, 1, 1, 0, ISAISIE_MONTANT, _
        "Repris de Tx_hrs_Forf du client, modifiable pour cette intervention seule"

    ' --- région 3 : les textes libres, colonnes 5 et 6 ------------------------
    ' Deux champs seulement, hauts de deux lignes chacun : ils remplissent la
    ' région et laissent la place d'écrire.
    DefInterv mChamps, 15, IC_TEXTE, "Texte de facture", ITYPE_LISTE, False, 1, 5, 2, 2, 0, ISAISIE_LIBRE, _
        "Texte repris sur la facture, repris du client ou choisi dans la liste"
    DefInterv mChamps, 16, IC_COMMENT, "Commentaires", ITYPE_TEXTE, False, 3, 5, 2, 2, 0, ISAISIE_LIBRE, _
        "Remarque libre sur l'intervention"

    mCharges = True
End Sub

'------------------------------------------------------------------------------
' Écrit une ligne du schéma. Sert uniquement à rendre ConstruireSchemaInterv
' lisible : sans elle il faudrait dix affectations par champ.
'------------------------------------------------------------------------------
Private Sub DefInterv(ByRef tb() As ChampInterv, ByVal idx As Long, ByVal colonne As String, _
                      ByVal libelle As String, ByVal typeCtrl As String, ByVal verrouille As Boolean, _
                      ByVal ligne As Long, ByVal col As Long, ByVal blocs As Long, _
                      ByVal nbLignes As Long, ByVal moitie As Long, ByVal saisie As Long, _
                      ByVal aide As String)
    tb(idx).Colonne = colonne
    tb(idx).Libelle = libelle
    tb(idx).TypeCtrl = typeCtrl
    tb(idx).Verrouille = verrouille
    tb(idx).Ligne = ligne
    tb(idx).Col = col
    tb(idx).Blocs = blocs
    tb(idx).NbLignes = nbLignes
    tb(idx).Moitie = moitie
    tb(idx).Saisie = saisie
    tb(idx).Aide = aide
End Sub

'==============================================================================
' Nommage des contrôles générés
'------------------------------------------------------------------------------
' Préfixe de type suivi du nom exact de la colonne : c'est cette règle, et elle
' seule, qui relie une colonne de TblInterv à son contrôle.
'==============================================================================
Public Function INomControle(ByRef ch As ChampInterv) As String
    Select Case ch.TypeCtrl
        Case ITYPE_LISTE, ITYPE_AUTO: INomControle = "cbo" & INomSur(ch.Colonne)
        Case ITYPE_CASE:              INomControle = "chk" & INomSur(ch.Colonne)
        Case Else:                    INomControle = "txt" & INomSur(ch.Colonne)
    End Select
End Function

'------------------------------------------------------------------------------
' Nom de colonne réduit à ce qu'un identifiant VBA accepte.
'
' Une colonne Excel peut s'appeler « Taux/Forfait » ; un contrôle, non. Tout ce
' qui n'est ni lettre, ni chiffre, ni soulignement est retiré, ce qui donne
' txtTauxForfait. Deux colonnes ne doivent donc pas se réduire au même nom —
' simulate_interv.py le vérifie.
'------------------------------------------------------------------------------
Public Function INomSur(ByVal nomColonne As String) As String
    Dim i As Long, c As String, r As String

    For i = 1 To Len(nomColonne)
        c = Mid$(nomColonne, i, 1)
        If (c >= "A" And c <= "Z") Or (c >= "a" And c <= "z") _
           Or (c >= "0" And c <= "9") Or c = "_" Then
            r = r & c
        End If
    Next i
    INomSur = r
End Function

'------------------------------------------------------------------------------
' Nom du libellé placé au-dessus d'un champ.
' Le préfixe lblI_ évite toute collision avec les libellés de l'habillage et avec
' ceux du formulaire des clients.
'------------------------------------------------------------------------------
Public Function INomLibelle(ByRef ch As ChampInterv) As String
    INomLibelle = "lblI_" & INomSur(ch.Colonne)
End Function

'------------------------------------------------------------------------------
' Nom du contrôle portant une colonne donnée ; chaîne vide si elle est inconnue.
'------------------------------------------------------------------------------
Public Function INomControleColonne(ByVal nomColonne As String) As String
    Dim ch() As ChampInterv, i As Long
    ch = ObtenirChampsInterv()
    For i = LBound(ch) To UBound(ch)
        If StrComp(ch(i).Colonne, nomColonne, vbTextCompare) = 0 Then
            INomControleColonne = INomControle(ch(i))
            Exit Function
        End If
    Next i
End Function

'==============================================================================
' Colonnes que le formulaire n'écrit JAMAIS dans TblInterv.
'------------------------------------------------------------------------------
' « Verrouillé » et « non enregistré » sont deux choses différentes, et les
' confondre a coûté un défaut : le n° de client est verrouillé — l'utilisateur
' ne le tape pas, il vient du client choisi — mais il doit bel et bien partir
' dans le tableau, sinon la colonne reste vide et le chiffre d'affaires, qui
' cherche ce numéro dans TblClients, ne trouve rien.
'
' Deux colonnes seulement échappent au formulaire :
'   NoInterv   attribué par IntervBD_Ajouter, jamais retouché ensuite
'   No_Facture attribué par la facturation, pas depuis cette fiche
'
' Le CA en faisait partie tant qu'il était calculé par une formule Excel. Il
' est maintenant calculé au fil de la saisie, à partir du taux enregistré sur
' l'intervention, et écrit comme les autres colonnes. Si la colonne porte
' encore une formule, EcrireCellules la laisse tranquille : la formule gagne,
' et VerifierClasseurInterventions dit lequel des deux régimes s'applique.
'==============================================================================
Public Function IColonnesNonEcrites() As Variant
    IColonnesNonEcrites = Array(IC_NO, IC_FACTURE)
End Function

'------------------------------------------------------------------------------
' True si le formulaire ne doit pas écrire cette colonne.
'------------------------------------------------------------------------------
Public Function IColonneNonEcrite(ByVal colonne As String) As Boolean
    Dim v As Variant, i As Long
    v = IColonnesNonEcrites()
    For i = LBound(v) To UBound(v)
        If StrComp(colonne, CStr(v(i)), vbTextCompare) = 0 Then
            IColonneNonEcrite = True
            Exit Function
        End If
    Next i
End Function

'==============================================================================
' Colonnes du tableau des enregistrements
'------------------------------------------------------------------------------
' Le tableau est une GRILLE DE LIBELLÉS, plus une ListBox : le plafond de dix
' colonnes est levé, et il n'y a plus rien à sacrifier ni à fusionner. Douze
' colonnes tiennent dans la largeur, chiffre d'affaires compris.
'
' Une colonne affichée peut malgré tout en réunir plusieurs, séparées par
' ICL_SEPARATEUR : « Nom+Prenom » donnerait « Aiello Rosalba » dans une seule
' case. La possibilité reste ouverte si l'on veut resserrer un jour ; chacune
' garde de toute façon sa propre colonne dans TblInterv.
'
' Ne restent dehors que Titre, NoInterv, TVA et Forfait, qui n'apprennent rien
' dans une liste et que la fiche montre dès qu'on sélectionne une ligne.
'
' Les quatre tableaux ci-dessous se lisent position par position et doivent
' rester de même longueur.
'==============================================================================
Public Function IColonnesListe() As Variant
    IColonnesListe = Array(IC_DATE, IC_CLIENT, IC_ENTREPRISE, IC_NOM, IC_PRENOM, _
                           IC_HEURES, IC_PERS, IC_TAUX, IC_CA, _
                           IC_TEXTE, IC_COMMENT, IC_FACTURE)
End Function

'------------------------------------------------------------------------------
' True si la colonne affichée réunit plusieurs colonnes du tableau.
'------------------------------------------------------------------------------
Public Function IColonneComposee(ByVal cle As String) As Boolean
    IColonneComposee = (InStr(cle, ICL_SEPARATEUR) > 0)
End Function

'------------------------------------------------------------------------------
' Colonnes de TblInterv derrière une colonne affichée.
'   renvoie : un tableau d'un seul nom, ou de plusieurs si la case est partagée
'------------------------------------------------------------------------------
Public Function IColonnesSources(ByVal cle As String) As Variant
    If IColonneComposee(cle) Then
        IColonnesSources = Split(cle, ICL_SEPARATEUR)
    Else
        IColonnesSources = Array(cle)
    End If
End Function

'------------------------------------------------------------------------------
' Alignement de chaque colonne — ce qu'une ListBox ne savait pas faire. Durées
' et montants se lisent alignés à droite, les ordres de grandeur se comparant
' alors d'un coup d'oeil ; le nombre de personnes, à un ou deux chiffres, est
' centré.
'------------------------------------------------------------------------------
Public Function IAlignementsListe() As Variant
    IAlignementsListe = Array(MSF_TextAlignLeft, MSF_TextAlignLeft, MSF_TextAlignLeft, _
                              MSF_TextAlignLeft, MSF_TextAlignLeft, MSF_TextAlignRight, _
                              MSF_TextAlignCenter, MSF_TextAlignRight, MSF_TextAlignRight, _
                              MSF_TextAlignLeft, MSF_TextAlignLeft, MSF_TextAlignLeft)
End Function

'------------------------------------------------------------------------------
Public Function ILibellesListe() As Variant
    ILibellesListe = Array("Date", "N° client", "Entreprise", "Nom", "Prénom", _
                           "Heures", "Pers.", "Taux/Forf.", "Chiffre d'aff.", _
                           "Texte de facture", "Commentaires", "Fact.")
End Function

'------------------------------------------------------------------------------
' Largeurs en points, dans le même ordre. Leur somme doit rester sous 910 pt :
' la ListBox mesure 926 pt de large et la barre de défilement en prend 16.
'------------------------------------------------------------------------------
Public Function ILargeursListe() As Variant
    ' Facture ne porte jamais plus de six chiffres : 40 points lui suffisent, et
    ' ce qu'elle rend va aux colonnes qui tronquaient.
    ILargeursListe = Array(64, 50, 132, 86, 76, 52, 34, 58, 68, 122, 124, 40)
End Function

'==============================================================================
' Champs proposés par le menu déroulant de filtrage
'==============================================================================
Public Function IChampsFiltrables() As Variant
    IChampsFiltrables = Array(IC_ENTREPRISE, IC_NOM)
End Function

'==============================================================================
' Colonnes recopiées depuis TblClients quand une entreprise ou un nom est choisi.
'   élément 0 = colonne source dans TblClients
'   élément 1 = colonne cible dans TblInterv
'==============================================================================
Public Function IReportsDepuisClients() As Variant
    IReportsDepuisClients = Array( _
        Array("ID_Cresus", IC_CLIENT), _
        Array("Entreprise", IC_ENTREPRISE), _
        Array("Titre", IC_TITRE), _
        Array("Nom", IC_NOM), _
        Array("Prenom", IC_PRENOM), _
        Array("TVA", IC_TVA), _
        Array("Forfait", IC_FORFAIT), _
        Array("Tx_hrs_Forf", IC_TAUX), _
        Array("Texte_Facture", IC_TEXTE), _
        Array("Note_Interne", IC_COMMENT))
End Function

'==============================================================================
' Accès au tableau des interventions
'==============================================================================
Public Function TableInterventions() As ListObject
    Set TableInterventions = ObtenirTable(NOM_TABLE_INTERVENTIONS)
End Function
