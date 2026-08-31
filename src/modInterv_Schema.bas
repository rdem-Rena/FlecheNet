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

Public Const NB_CHAMPS_INTERV As Long = 15

'==============================================================================
' Définition d'un champ du formulaire
'==============================================================================
Public Type ChampInterv
    Colonne As String       ' nom exact de la colonne dans TblInterv
    Libelle As String
    TypeCtrl As String
    Verrouille As Boolean   ' True = géré par le programme, non saisissable
    Ligne As Long           ' ligne de grille, 1 à 4
    Col As Long             ' colonne de grille, 1 à 4
    Moitie As Long          ' 0 = bloc entier, 1 = moitié gauche, 2 = moitié droite
    Blocs As Long           ' nombre de blocs occupés en largeur (1 par défaut)
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
'   n°  colonne, libellé, type, verrouillé, ligne, colonne, moitié, blocs,
'       contrainte de saisie, info-bulle
'------------------------------------------------------------------------------
Private Sub ConstruireSchemaInterv()
    ReDim mChamps(1 To NB_CHAMPS_INTERV)

    DefInterv mChamps, 1, IC_NO, "N° intervention", ITYPE_TEXTE, True, 1, 1, 0, 1, ISAISIE_LIBRE, _
        "Index du tableau, attribué automatiquement par le programme"
    DefInterv mChamps, 2, IC_DATE, "Date", ITYPE_DATE, False, 1, 2, 0, 1, ISAISIE_LIBRE, _
        "Date de l'intervention. Cliquez sur l'icône pour ouvrir le calendrier."
    DefInterv mChamps, 3, IC_CLIENT, "N° client (Crésus)", ITYPE_TEXTE, True, 1, 3, 0, 1, ISAISIE_LIBRE, _
        "Renseigné automatiquement en choisissant une entreprise ou un nom"
    DefInterv mChamps, 4, IC_ENTREPRISE, "Entreprise", ITYPE_AUTO, False, 1, 4, 0, 1, ISAISIE_LIBRE, _
        "Tapez les premières lettres : la liste des entreprises de TblClients se filtre au fil de la frappe"

    DefInterv mChamps, 5, IC_TITRE, "Titre", ITYPE_LISTE, False, 2, 1, 0, 1, ISAISIE_LIBRE, _
        "Civilité du client"
    DefInterv mChamps, 6, IC_NOM, "Nom", ITYPE_AUTO, False, 2, 2, 0, 1, ISAISIE_LIBRE, _
        "Tapez les premières lettres : la liste des noms de TblClients se filtre au fil de la frappe"
    DefInterv mChamps, 7, IC_PRENOM, "Prénom", ITYPE_TEXTE, False, 2, 3, 0, 1, ISAISIE_LIBRE, _
        "Prénom du client"
    DefInterv mChamps, 8, IC_TEXTE, "Texte de facture", ITYPE_LISTE, False, 2, 4, 0, 1, ISAISIE_LIBRE, _
        "Texte repris sur la facture, repris du client ou choisi dans la liste"

    DefInterv mChamps, 9, IC_HEURES, "Heures", ITYPE_TEXTE, False, 3, 1, 0, 1, ISAISIE_HEURES, _
        "Durée au format heures:minutes, par exemple 420:00"
    DefInterv mChamps, 10, IC_PERS, "Personnes", ITYPE_TEXTE, False, 3, 2, 0, 1, ISAISIE_ENTIER, _
        "Nombre de personnes intervenues, de 1 à 99"
    DefInterv mChamps, 11, IC_TVA, "TVA", ITYPE_CASE, False, 3, 3, 1, 1, ISAISIE_LIBRE, _
        "Le client est assujetti à la TVA"
    DefInterv mChamps, 12, IC_FORFAIT, "Forfait", ITYPE_CASE, False, 3, 3, 2, 1, ISAISIE_LIBRE, _
        "Facturation au forfait plutôt qu'à l'heure"
    DefInterv mChamps, 13, IC_CA, "Chiffre d'affaires", ITYPE_TEXTE, True, 3, 4, 0, 1, ISAISIE_LIBRE, _
        "Calculé par le tableau Excel. La valeur affichée ici est une estimation, mise à jour au fil de la saisie."

    DefInterv mChamps, 14, IC_FACTURE, "N° de facture", ITYPE_TEXTE, True, 4, 1, 0, 1, ISAISIE_LIBRE, _
        "Attribué par la facturation, pas depuis ce formulaire"
    DefInterv mChamps, 15, IC_COMMENT, "Commentaires", ITYPE_TEXTE, False, 4, 2, 0, 3, ISAISIE_LIBRE, _
        "Remarque libre sur l'intervention"

    mCharges = True
End Sub

'------------------------------------------------------------------------------
' Écrit une ligne du schéma. Sert uniquement à rendre ConstruireSchemaInterv
' lisible : sans elle il faudrait dix affectations par champ.
'------------------------------------------------------------------------------
Private Sub DefInterv(ByRef tb() As ChampInterv, ByVal idx As Long, ByVal colonne As String, _
                      ByVal libelle As String, ByVal typeCtrl As String, ByVal verrouille As Boolean, _
                      ByVal ligne As Long, ByVal col As Long, ByVal moitie As Long, _
                      ByVal blocs As Long, ByVal saisie As Long, ByVal aide As String)
    tb(idx).Colonne = colonne
    tb(idx).Libelle = libelle
    tb(idx).TypeCtrl = typeCtrl
    tb(idx).Verrouille = verrouille
    tb(idx).Ligne = ligne
    tb(idx).Col = col
    tb(idx).Moitie = moitie
    tb(idx).Blocs = blocs
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
        Case ITYPE_LISTE, ITYPE_AUTO: INomControle = "cbo" & ch.Colonne
        Case ITYPE_CASE:              INomControle = "chk" & ch.Colonne
        Case Else:                    INomControle = "txt" & ch.Colonne
    End Select
End Function

'------------------------------------------------------------------------------
' Nom du libellé placé au-dessus d'un champ.
' Le préfixe lblI_ évite toute collision avec les libellés de l'habillage et avec
' ceux du formulaire des clients.
'------------------------------------------------------------------------------
Public Function INomLibelle(ByRef ch As ChampInterv) As String
    INomLibelle = "lblI_" & ch.Colonne
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
' Trois colonnes seulement échappent au formulaire :
'   NoInterv   attribué par IntervBD_Ajouter, jamais retouché ensuite
'   CA         porte une formule ; l'écrire l'effacerait
'   No_Facture attribué par la facturation, pas depuis cette fiche
'==============================================================================
Public Function IColonnesNonEcrites() As Variant
    IColonnesNonEcrites = Array(IC_NO, IC_CA, IC_FACTURE)
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
' Dix au maximum : c'est ce qu'accepte une ListBox MSForms. Titre a été laissé
' de côté — c'est la colonne la moins informative dans une liste, et elle reste
' visible dans la fiche dès qu'une ligne est sélectionnée. Pour la réintégrer,
' remplacer ici une autre colonne.
'==============================================================================
Public Function IColonnesListe() As Variant
    IColonnesListe = Array(IC_DATE, IC_CLIENT, IC_ENTREPRISE, IC_NOM, IC_PRENOM, _
                           IC_HEURES, IC_PERS, IC_TEXTE, IC_COMMENT, IC_FACTURE)
End Function

'------------------------------------------------------------------------------
' En-têtes affichés au-dessus du tableau, un par colonne de IColonnesListe et
' dans le même ordre. C'est ici que se mettent accents et abréviations, le nom
' réel de la colonne Excel restant dans IColonnesListe.
'------------------------------------------------------------------------------
Public Function ILibellesListe() As Variant
    ILibellesListe = Array("Date", "N° client", "Entreprise", "Nom", "Prénom", _
                           "Heures", "Pers.", "Texte de facture", "Commentaires", "Facture")
End Function

'------------------------------------------------------------------------------
' Largeurs en points, dans le même ordre. Leur somme doit rester sous 910 pt :
' la ListBox mesure 926 pt de large et la barre de défilement en prend 16.
'------------------------------------------------------------------------------
Public Function ILargeursListe() As Variant
    ILargeursListe = Array(66, 54, 140, 96, 84, 56, 44, 130, 156, 60)
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
        Array("Texte_Facture", IC_TEXTE), _
        Array("Note_Interne", IC_COMMENT))
End Function

'==============================================================================
' Accès au tableau des interventions
'==============================================================================
Public Function TableInterventions() As ListObject
    Set TableInterventions = ObtenirTable(NOM_TABLE_INTERVENTIONS)
End Function
