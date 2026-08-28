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
Public Const COL_TAUX As String = "Tx_hrs_Forf"

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
'------------------------------------------------------------------------------
' Schéma complet du formulaire : un élément par champ, dans l'ordre de saisie.
'   renvoie : un tableau de ChampClient indexé de 1 à NB_CHAMPS
' Le schéma n'est construit qu'une fois par session puis conservé en mémoire.
'------------------------------------------------------------------------------
Public Function ObtenirChamps() As ChampClient()
    If Not mCharges Then ConstruireSchema
    ObtenirChamps = mChamps
End Function

'------------------------------------------------------------------------------
' Remplit le schéma. C'est LA table à modifier pour ajouter, déplacer,
' renommer ou verrouiller un champ ; le générateur et le formulaire s'y adaptent
' sans autre changement.
'
' Les colonnes ci-dessous se lisent :
'   n°  nom de la colonne Excel, libellé affiché, type de contrôle, verrouillé,
'       ligne de grille, colonne de grille, moitié de bloc, contrainte numérique,
'       info-bulle.
'------------------------------------------------------------------------------
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

    DefChamp mChamps, 17, COL_TAUX, "Taux horaire / forfait", TYPE_TEXTE, False, 5, 1, 0, NUM_DECIMAL, _
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

'------------------------------------------------------------------------------
' Écrit une ligne du schéma. Sert uniquement à rendre ConstruireSchema lisible :
' sans elle il faudrait neuf affectations par champ.
'------------------------------------------------------------------------------
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
'------------------------------------------------------------------------------
' Nom du contrôle qui porte un champ, dans le formulaire.
'   renvoie : txt / cbo / chk suivi du nom exact de la colonne Excel
' C'est cette règle, et elle seule, qui relie une colonne à son contrôle : il n'y
' a aucune table de correspondance à tenir à jour.
'------------------------------------------------------------------------------
Public Function NomControle(ByRef ch As ChampClient) As String
    Select Case ch.TypeCtrl
        Case TYPE_LISTE: NomControle = "cbo" & ch.Colonne
        Case TYPE_CASE:  NomControle = "chk" & ch.Colonne
        Case Else:       NomControle = "txt" & ch.Colonne
    End Select
End Function

'------------------------------------------------------------------------------
' Nom du libellé placé au-dessus d'un champ.
' Le préfixe lblChamp_ évite toute collision avec les libellés de l'habillage
' (lblTitre serait sinon à la fois le titre du formulaire et le libellé du champ
' Titre).
'------------------------------------------------------------------------------
Public Function NomLibelle(ByRef ch As ChampClient) As String
    NomLibelle = "lblChamp_" & ch.Colonne
End Function

'------------------------------------------------------------------------------
' Nom du contrôle correspondant à une colonne, cherché dans le schéma.
'   renvoie : le nom du contrôle, ou une chaîne vide si la colonne est inconnue
' Version pratique de NomControle quand on n'a que le nom de la colonne sous la
' main, par exemple NomControleColonne(COL_NPA).
'------------------------------------------------------------------------------
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
' Colonnes du tableau des enregistrements
'==============================================================================
'------------------------------------------------------------------------------
' Colonnes du tableau des enregistrements, de gauche à droite.
'   renvoie : les noms EXACTS des colonnes de TblClients
'
' Dix au maximum : c'est la limite d'une ListBox MSForms. Les champs qui n'y
' figurent pas restent visibles dans la fiche du haut dès qu'une ligne est
' sélectionnée.
'
' Toute modification doit être reportée à l'identique dans LibellesListe et
' LargeursListe, qui sont lues position par position.
'------------------------------------------------------------------------------
Public Function ColonnesListe() As Variant
    ColonnesListe = Array("Entreprise", "Titre", "Nom", "Prenom", COL_ADRESSE, _
                          "No", COL_NPA, COL_VILLE, COL_CANTON, COL_TAUX)
End Function

'------------------------------------------------------------------------------
' En-têtes affichés au-dessus du tableau. Un libellé par colonne de
' ColonnesListe, dans le même ordre : c'est ici qu'on met les accents et les
' abréviations, le nom réel de la colonne Excel restant dans ColonnesListe.
'------------------------------------------------------------------------------
Public Function LibellesListe() As Variant
    LibellesListe = Array("Entreprise", "Titre", "Nom", "Prénom", "Adresse", _
                          "No", "NPA", "Ville", "Cant.", "Taux / forf.")
End Function

'------------------------------------------------------------------------------
' Largeur de chaque colonne, en points, dans le même ordre.
' Leur somme doit rester sous 790 pt : la ListBox mesure 806 pt de large et il
' faut laisser la place à la barre de défilement verticale (environ 16 pt).
' Les en-têtes cliquables se repositionnent automatiquement d'après ces valeurs.
'------------------------------------------------------------------------------
Public Function LargeursListe() As Variant
    LargeursListe = Array(118, 54, 102, 86, 144, 30, 42, 104, 36, 62)
End Function

'==============================================================================
' Champs proposés par le menu déroulant de filtrage
'==============================================================================
'------------------------------------------------------------------------------
' Colonnes proposées par le menu déroulant de filtrage.
' Ces valeurs servent à la fois d'étiquette affichée et de clef de recherche :
' il faut donc y écrire le nom exact de la colonne Excel, sans accent ajouté.
'------------------------------------------------------------------------------
Public Function ChampsFiltrables() As Variant
    ChampsFiltrables = Array("Entreprise", "Nom", COL_ADRESSE)
End Function

'==============================================================================
' Valeurs proposées par le menu déroulant Titre
'==============================================================================
'------------------------------------------------------------------------------
' Civilités proposées d'office par le menu déroulant Titre.
' Les civilités déjà présentes dans la colonne Titre du tableau viennent s'y
' ajouter automatiquement à l'ouverture : aucune fiche existante ne devient
' inaffichable si sa civilité n'est pas dans cette liste.
'------------------------------------------------------------------------------
Public Function TitresProposes() As Variant
    TitresProposes = Array("Monsieur", "Madame")
End Function

'==============================================================================
' Accès aux tableaux structurés du classeur
'==============================================================================
'------------------------------------------------------------------------------
' Cherche un tableau structuré dans tout le classeur, feuille par feuille.
'   nomTable : nom du tableau, par exemple TblClients
'   renvoie  : le ListObject, ou Nothing s'il n'existe pas
'
' Passer par le nom du tableau plutôt que par un nom de feuille ou une plage de
' cellules : le tableau peut être déplacé, renommé de feuille ou décalé, le code
' continue de le trouver.
'------------------------------------------------------------------------------
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

'------------------------------------------------------------------------------
' Raccourci vers le tableau TblClients.
'------------------------------------------------------------------------------
Public Function TableClients() As ListObject
    Set TableClients = ObtenirTable(NOM_TABLE_CLIENTS)
End Function

'------------------------------------------------------------------------------
' Position d'une colonne dans un tableau structuré.
'   renvoie : le numéro de colonne (1 = première), ou 0 si elle n'existe pas
' La comparaison ignore la casse.
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
