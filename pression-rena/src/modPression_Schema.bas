Attribute VB_Name = "modPression_Schema"
Option Explicit
'==============================================================================
' modPression_Schema
'------------------------------------------------------------------------------
' Description du classeur « Pression Rena » : ses feuilles, ses colonnes, et
' l'invariant sur lequel repose tout le reste.
'
' L'INVARIANT. La feuille Datas porte EXACTEMENT deux lignes par jour
' calendaire, sans trou et sans doublon : ligne 3 + 2k et ligne 4 + 2k sont le
' matin et le soir du k-ième jour. Une date se convertit donc en numéro de
' ligne par une soustraction — voir LigneDe — et le formulaire n'a jamais à
' chercher, ni à insérer, ni à supprimer une ligne. Il remplit celle qui
' existe déjà.
'
' C'est aussi ce qui permet aux plages nommées des deux graphiques de se
' calculer par un décalage. Insérer une ligne dans Datas les décalerait toutes :
' la cellule Controle de la feuille Reglages le signale, et VerifierClasseur le
' répète.
'==============================================================================

'--- Feuilles -----------------------------------------------------------------
Public Const FEUILLE_DATAS As String = "Datas"
Public Const FEUILLE_SAISIE As String = "Saisie"
Public Const FEUILLE_CALCUL As String = "Calcul"
Public Const FEUILLE_REGLAGES As String = "Reglages"

Public Const NOM_FORM_PRESSION As String = "UF_Pression"

'--- Repères de la feuille Datas ----------------------------------------------
Public Const LIGNE_ENTETE As Long = 2
Public Const LIGNE_DEB As Long = 3

'--- Colonnes de Datas --------------------------------------------------------
Public Const PC_DATE As Long = 1
Public Const PC_HEURE As Long = 2
Public Const PC_ACTIVITE As Long = 3
Public Const PC_COMPO As Long = 4
Public Const PC_SYS As Long = 5
Public Const PC_DIA As Long = 6
Public Const PC_POULS As Long = 7
Public Const PC_BB As Long = 8
Public Const PC_BH As Long = 9
Public Const PC_HB As Long = 10
Public Const PC_HH As Long = 11
Public Const PC_DEMI As Long = 12
Public Const PC_CHANGE As Long = 13
Public Const PC_COMMENT As Long = 14
Public Const PC_POIDS As Long = 15

' Les colonnes que le formulaire n'écrit JAMAIS : elles portent une formule.
' Date-Compo, Demi et les quatre lignes de repère se recalculent seules ; les
' écrire remplacerait leur formule par une valeur figée, et l'étiquette de
' l'axe cesserait de suivre la ligne.
Public Const PC_PREMIERE_FORMULE As Long = PC_COMPO

'--- Les deux demi-journées ---------------------------------------------------
' Ce n'est pas l'heure qui décide, mais le RANG de la ligne dans sa paire.
' Huit jours du fichier d'origine portent deux mesures du matin : les ranger
' par l'heure en aurait perdu une à chaque fois.
Public Const DEMI_MATIN As Long = 0
Public Const DEMI_SOIR As Long = 1

'--- Cellules nommées lues par le formulaire ----------------------------------
Public Const CEL_CONTROLE As String = "Controle"
Public Const CEL_SEUIL_SYS As String = "Seuil_Sys_Haut"
Public Const CEL_SEUIL_DIA As String = "Seuil_Dia_Haut"

'--- Bornes de saisie ---------------------------------------------------------
' Elles n'ont rien de médical : elles écartent les fautes de frappe évidentes,
' le 557 pour 57 qu'on trouve dans le fichier d'origine par exemple.
Public Const SYS_MIN As Long = 60
Public Const SYS_MAX As Long = 260
Public Const DIA_MIN As Long = 30
Public Const DIA_MAX As Long = 160
Public Const POULS_MIN As Long = 30
Public Const POULS_MAX As Long = 220
Public Const POIDS_MIN As Double = 30
Public Const POIDS_MAX As Double = 250

'==============================================================================
' ACCÈS AUX FEUILLES
'==============================================================================
'------------------------------------------------------------------------------
' Une feuille du classeur, par son nom.
'   renvoie : la feuille, ou Nothing si elle n'existe pas
'------------------------------------------------------------------------------
Public Function ObtenirFeuille(ByVal nom As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, nom, vbTextCompare) = 0 Then
            Set ObtenirFeuille = ws
            Exit Function
        End If
    Next ws
End Function

Public Function FeuilleDatas() As Worksheet
    Set FeuilleDatas = ObtenirFeuille(FEUILLE_DATAS)
End Function

'------------------------------------------------------------------------------
' Valeur d'une cellule nommée du classeur.
'   renvoie : la valeur, ou Empty si le nom n'existe pas
'------------------------------------------------------------------------------
Public Function ValeurNommee(ByVal nom As String) As Variant
    On Error Resume Next
    ValeurNommee = ThisWorkbook.Names(nom).RefersToRange.Value
    On Error GoTo 0
End Function

'==============================================================================
' L'INVARIANT : UNE DATE EST UN NUMÉRO DE LIGNE
'==============================================================================
'------------------------------------------------------------------------------
' Premier jour du calendrier, lu sur la feuille Datas.
'   renvoie : la date, ou 0 si la feuille manque ou est vide
'------------------------------------------------------------------------------
Public Function PremierJour() As Date
    Dim ws As Worksheet
    Set ws = FeuilleDatas()
    If ws Is Nothing Then Exit Function
    If Not IsDate(ws.Cells(LIGNE_DEB, PC_DATE).Value) Then Exit Function
    PremierJour = CDate(ws.Cells(LIGNE_DEB, PC_DATE).Value)
End Function

'------------------------------------------------------------------------------
' Dernière ligne du calendrier.
'------------------------------------------------------------------------------
Public Function DerniereLigne() As Long
    Dim ws As Worksheet
    Set ws = FeuilleDatas()
    If ws Is Nothing Then Exit Function
    DerniereLigne = ws.Cells(ws.Rows.Count, PC_DATE).End(xlUp).Row
End Function

'------------------------------------------------------------------------------
' Dernier jour du calendrier.
'------------------------------------------------------------------------------
Public Function DernierJour() As Date
    Dim ws As Worksheet, lg As Long
    Set ws = FeuilleDatas()
    If ws Is Nothing Then Exit Function
    lg = DerniereLigne()
    If lg < LIGNE_DEB Then Exit Function
    If Not IsDate(ws.Cells(lg, PC_DATE).Value) Then Exit Function
    DernierJour = CDate(ws.Cells(lg, PC_DATE).Value)
End Function

'------------------------------------------------------------------------------
' Numéro de ligne d'une demi-journée.
'   jour    : la date de la mesure
'   demi    : DEMI_MATIN ou DEMI_SOIR
'   renvoie : le numéro de ligne dans Datas, ou 0 si la date est hors calendrier
'
' Toute la conversion tient dans cette soustraction. C'est ce que garantit
' l'invariant, et ce qui évite au formulaire de chercher une date dans 5368
' lignes à chaque frappe.
'------------------------------------------------------------------------------
Public Function LigneDe(ByVal jour As Date, ByVal demi As Long) As Long
    Dim premier As Date, ecart As Long

    premier = PremierJour()
    If premier = 0 Then Exit Function

    ecart = CLng(Int(CDbl(jour)) - Int(CDbl(premier)))
    If ecart < 0 Then Exit Function

    LigneDe = LIGNE_DEB + ecart * 2 + demi
    If LigneDe > DerniereLigne() Then LigneDe = 0
End Function

'------------------------------------------------------------------------------
' La date que porte une ligne, et sa demi-journée.
'------------------------------------------------------------------------------
Public Function JourDeLaLigne(ByVal ligne As Long) As Date
    Dim premier As Date
    premier = PremierJour()
    If premier = 0 Or ligne < LIGNE_DEB Then Exit Function
    JourDeLaLigne = premier + (ligne - LIGNE_DEB) \ 2
End Function

Public Function DemiDeLaLigne(ByVal ligne As Long) As Long
    DemiDeLaLigne = (ligne - LIGNE_DEB) Mod 2
End Function

'------------------------------------------------------------------------------
' Le nom d'une demi-journée, tel que l'écrit la colonne Demi.
'------------------------------------------------------------------------------
Public Function NomDemi(ByVal demi As Long) As String
    NomDemi = IIf(demi = DEMI_MATIN, "matin", "Soir")
End Function

'==============================================================================
' LISTES PROPOSÉES PAR LE FORMULAIRE
'==============================================================================
'------------------------------------------------------------------------------
' Les valeurs de la colonne Change. Elles ne mesurent rien : ce sont des
' repères, portés à la hauteur où ils doivent apparaître sur le graphique.
'------------------------------------------------------------------------------
Public Function ChangesProposes() As Variant
    ChangesProposes = Array("", "150", "160")
End Function
