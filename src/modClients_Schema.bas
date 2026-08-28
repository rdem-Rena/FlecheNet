Attribute VB_Name = "modClients_Schema"
Option Explicit
'==============================================================================
' modClients_Schéma
'------------------------------------------------------------------------------
' Description du tableau TblClients : c'est LA source de vérité à partir de
' laquelle le UserForm est généré (un champ = une entrée ci-dessous).
'
' Pour ajouter un champ au formulaire : ajouter une colonne au tableau Excel,
' ajouter une ligne Def(...) dans ObtenirChamps, augmenter NB_CHAMPS, puis
' relancer GenererFormulaireClients.
'==============================================================================

'--- Noms des objets du classeur ----------------------------------------------
Public Const NOM_TABLE_CLIENTS As String = "TblClients"
Public Const NOM_TABLE_ADRESSES As String = "Tabl_Adresses"
Public Const NOM_TABLE_VILLES As String = "Tabl_Villes_CH"
Public Const NOM_TABLE_TEXTES As String = "TblTxtStd"
Public Const NOM_TABLE_INTERV As String = "TblInterv"
Public Const NOM_FORMULAIRE As String = "UF_Clients"

'--- Colonnes gérées par le programme -----------------------------------------
Public Const COL_CLEF As String = "Clef_BD"
Public Const COL_DATE As String = "Date_Crea"
Public Const PREFIXE_CLEF As String = "CL"

'--- Colonnes utilisées par les recherches d'adresse --------------------------
Public Const COL_ADRESSE As String = "Adresse"
Public Const COL_NPA As String = "NoPost"
Public Const COL_VILLE As String = "Ville"
Public Const COL_CANTON As String = "Cant"
Public Const COL_TEXTE_FACTURE As String = "Texte_Facture"
Public Const COL_ID_CRESUS As String = "ID_Cresus"

'--- Types de controle --------------------------------------------------------
Public Const TYPE_TEXTE As String = "T"     ' TextBox
Public Const TYPE_LISTE As String = "C"     ' ComboBox
Public Const TYPE_CASE As String = "K"      ' CheckBox

'--- Contraintes numériques ---------------------------------------------------
Public Const NUM_NON As Long = 0
Public Const NUM_ENTIER As Long = 1
Public Const NUM_DECIMAL As Long = 2

Public Const NB_CHAMPS As Long = 21

'==============================================================================
' Définition d'un champ du formulaire
'==============================================================================
Public Type ChampClient
    Colonne As String       ' nom exact de la colonne dans TblClients
    Libelle As String       ' libellé affiche au-dessus de la zone de saisie
    TypeCtrl As String      ' TYPE_TEXTE / TYPE_LISTE / TYPE_CASE
    Verrouille As Boolean   ' True = géré par le programme, non saisissable
    Ligne As Long           ' ligne dans la grille (1 à 5)
    Col As Long             ' colonne dans la grille (1 à 4)
    Moitie As Long          ' 0 = bloc entier, 1 = moitie gauche, 2 = moitie droite
    Numerique As Long       ' NUM_NON / NUM_ENTIER / NUM_DECIMAL
    Aide As String          ' info-bulle
End Type

Private mChamps() As ChampClient
Private mCharges As Boolean

'==============================================================================
' Schéma complet du formulaire
'==============================================================================
Public Function ObtenirChamps() As ChampClient()
    If Not mCharges Then ConstruireSchema
    ObtenirChamps = mChamps
End Function

Private Sub ConstruireSchema()
    ReDim mChamps(1 To NB_CHAMPS)

    '           idx  colonne          libellé                type         verr.  lig col moit  num          info-bulle
    DefChamp mChamps, 1, COL_CLEF, "Clef BD", TYPE_TEXTE, True, 1, 1, 0, NUM_NON, _
        "Index du tableau, attribué automatiquement par le programme"
    DefChamp mChamps, 2, COL_DATE, "Date de création", TYPE_TEXTE, True, 1, 2, 0, NUM_NON, _
        "Date de création de la fiche, attribuée automatiquement"
    DefChamp mChamps, 3, COL_ID_CRESUS, "ID Crésus", TYPE_TEXTE, False, 1, 3, 0, NUM_ENTIER, _
        "Identifiant du client dans Crésus (référence des interventions)"
    DefChamp mChamps, 4, "Entreprise", "Entreprise", TYPE_TEXTE, False, 1, 4, 0, NUM_NON, _
        "Raison sociale, si le client est une entreprise"

    DefChamp mChamps, 5, "Titre", "Titre", TYPE_LISTE, False, 2, 1, 0, NUM_NON, _
        "Civilité du client"
    DefChamp mChamps, 6, "Nom", "Nom", TYPE_TEXTE, False, 2, 2, 0, NUM_NON, _
        "Nom de famille du client"
    DefChamp mChamps, 7, "Prenom", "Prénom", TYPE_TEXTE, False, 2, 3, 0, NUM_NON, _
        "Prénom du client"
    DefChamp mChamps, 8, "Email", "Courriel", TYPE_TEXTE, False, 2, 4, 0, NUM_NON, _
        "Adresse de courriel"

    DefChamp mChamps, 9, COL_ADRESSE, "Adresse (rue)", TYPE_LISTE, False, 3, 1, 0, NUM_NON, _
        "Rue : la liste provient de l'onglet Adresses. Le NPA, la ville et le canton se remplissent automatiquement."
    DefChamp mChamps, 10, "No", "No", TYPE_TEXTE, False, 3, 2, 0, NUM_NON, _
        "Numéro dans la rue (peut contenir une lettre, ex. 34B)"
    DefChamp mChamps, 11, COL_NPA, "NPA", TYPE_LISTE, False, 3, 3, 0, NUM_NON, _
        "Numéro postal : renseigne la ville et le canton"
    DefChamp mChamps, 12, COL_VILLE, "Ville", TYPE_TEXTE, False, 3, 4, 0, NUM_NON, _
        "Localité, renseignée automatiquement depuis le NPA"

    DefChamp mChamps, 13, COL_CANTON, "Canton", TYPE_TEXTE, False, 4, 1, 0, NUM_NON, _
        "Canton, renseigné automatiquement depuis le NPA"
    DefChamp mChamps, 14, "Tel_Prive", "Téléphone privé", TYPE_TEXTE, False, 4, 2, 0, NUM_NON, _
        "Téléphone privé"
    DefChamp mChamps, 15, "Tel_Pro", "Téléphone pro.", TYPE_TEXTE, False, 4, 3, 0, NUM_NON, _
        "Téléphone professionnel"
    DefChamp mChamps, 16, "Natel", "Natel", TYPE_TEXTE, False, 4, 4, 0, NUM_NON, _
        "Téléphone mobile"

    DefChamp mChamps, 17, "Tx_hrs_Forf", "Taux horaire / forfait", TYPE_TEXTE, False, 5, 1, 0, NUM_DECIMAL, _
        "Taux horaire ou montant forfaitaire, en CHF"
    DefChamp mChamps, 18, "TVA", "TVA", TYPE_CASE, False, 5, 2, 1, NUM_NON, _
        "Le client est assujetti à la TVA"
    DefChamp mChamps, 19, "Forfait", "Forfait", TYPE_CASE, False, 5, 2, 2, NUM_NON, _
        "La facturation se fait au forfait"
    DefChamp mChamps, 20, COL_TEXTE_FACTURE, "Texte de facture", TYPE_LISTE, False, 5, 3, 0, NUM_NON, _
        "Texte standard repris sur les factures (onglet Parametres)"
    DefChamp mChamps, 21, "Note_Interne", "Note interne", TYPE_TEXTE, False, 5, 4, 0, NUM_NON, _
        "Remarque interne, non imprimée"

    mCharges = True
End Sub

Private Sub DefChamp(ByRef tb() As ChampClient, ByVal idx As Long, ByVal colonne As String, _
                ByVal libelle As String, ByVal typeCtrl As String, ByVal verrouille As Boolean, _
                ByVal ligne As Long, ByVal col As Long, ByVal moitie As Long, _
                ByVal numerique As Long, ByVal aide As String)
    tb(idx).Colonne = colonne
    tb(idx).Libelle = libelle
    tb(idx).TypeCtrl = typeCtrl
    tb(idx).Verrouille = verrouille
    tb(idx).Ligne = ligne
    tb(idx).Col = col
    tb(idx).Moitie = moitie
    tb(idx).Numerique = numerique
    tb(idx).Aide = aide
End Sub

'==============================================================================
' Nommage des contrôles générés
'==============================================================================
Public Function NomControle(ByRef ch As ChampClient) As String
    Select Case ch.TypeCtrl
        Case TYPE_LISTE: NomControle = "cbo" & ch.Colonne
        Case TYPE_CASE:  NomControle = "chk" & ch.Colonne
        Case Else:       NomControle = "txt" & ch.Colonne
    End Select
End Function

Public Function NomLibelle(ByRef ch As ChampClient) As String
    NomLibelle = "lblChamp_" & ch.Colonne
End Function

Public Function NomControleColonne(ByVal nomColonne As String) As String
    Dim ch() As ChampClient, i As Long
    ch = ObtenirChamps()
    For i = LBound(ch) To UBound(ch)
        If StrComp(ch(i).Colonne, nomColonne, vbTextCompare) = 0 Then
            NomControleColonne = NomControle(ch(i))
            Exit Function
        End If
    Next i
End Function

'==============================================================================
' Colonnes affichées dans le tableau des enregistrements
' (10 colonnes au maximum : limite d'une ListBox MSForms)
'==============================================================================
Public Function ColonnesListe() As Variant
    ColonnesListe = Array(COL_CLEF, COL_DATE, "Entreprise", "Titre", "Nom", _
                          "Prenom", COL_ADRESSE, "No", COL_NPA, COL_VILLE)
End Function

Public Function LibellesListe() As Variant
    LibellesListe = Array("Clef BD", "Création", "Entreprise", "Titre", "Nom", _
                          "Prénom", "Adresse", "No", "NPA", "Ville")
End Function

Public Function LargeursListe() As Variant
    LargeursListe = Array(46, 62, 112, 56, 100, 82, 142, 30, 42, 106)
End Function

'==============================================================================
' Champs proposés par le menu déroulant de filtrage
'==============================================================================
Public Function ChampsFiltrables() As Variant
    ChampsFiltrables = Array("Entreprise", "Nom", COL_ADRESSE)
End Function

'==============================================================================
' Valeurs proposées par le menu déroulant Titre
'==============================================================================
Public Function TitresProposes() As Variant
    TitresProposes = Array("Monsieur", "Madame")
End Function

'==============================================================================
' Accès aux tableaux structures du classeur
'==============================================================================
Public Function ObtenirTable(ByVal nomTable As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, nomTable, vbTextCompare) = 0 Then
                Set ObtenirTable = lo
                Exit Function
            End If
        Next lo
    Next ws
End Function

Public Function TableClients() As ListObject
    Set TableClients = ObtenirTable(NOM_TABLE_CLIENTS)
End Function

'------------------------------------------------------------------------------
' Index (base 1) d'une colonne dans un tableau structure ; 0 si absente.
'------------------------------------------------------------------------------
Public Function IndexColonne(ByVal lo As ListObject, ByVal nomColonne As String) As Long
    Dim lc As ListColumn
    If lo Is Nothing Then Exit Function
    For Each lc In lo.ListColumns
        If StrComp(lc.Name, nomColonne, vbTextCompare) = 0 Then
            IndexColonne = lc.Index
            Exit Function
        End If
    Next lc
End Function
