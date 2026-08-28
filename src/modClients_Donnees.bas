Attribute VB_Name = "modClients_Donnees"
Option Explicit
'==============================================================================
' modClients_Donnees
'------------------------------------------------------------------------------
' Couche d'accès aux données du tableau TblClients (feuille "Clients").
' Le formulaire ne touche jamais directement aux cellules : il passe par les
' procédures de ce module, qui travaillent sur un cache mémoire rechargé après
' chaque écriture.
'==============================================================================

Private mDonnees As Variant     ' cache : (1 To n, 1 To nbColonnes)
Private mNbLignes As Long
Private mNbColonnes As Long
Private mIdxColonne As Object   ' Dictionary : nom de colonne -> index

'==============================================================================
' Chargement du cache depuis le tableau structure
'==============================================================================
Public Sub Donnees_Charger()
    Dim lo As ListObject, lc As ListColumn

    Set lo = TableClients()
    If lo Is Nothing Then
        Err.Raise vbObjectError + 513, "modClients_Donnees", _
                  "Le tableau " & NOM_TABLE_CLIENTS & " est introuvable dans ce classeur."
    End If

    Set mIdxColonne = CreateObject("Scripting.Dictionary")
    mIdxColonne.CompareMode = 1                     ' comparaison insensible à la casse
    For Each lc In lo.ListColumns
        If Not mIdxColonne.Exists(lc.Name) Then mIdxColonne.Add lc.Name, lc.Index
    Next lc
    mNbColonnes = lo.ListColumns.Count

    If lo.ListRows.Count = 0 Then
        mNbLignes = 0
        mDonnees = Empty
    Else
        mNbLignes = lo.ListRows.Count
        If mNbLignes = 1 And mNbColonnes = 1 Then
            ' cas dégénéré : Range.Value ne renvoie pas un tableau
            Dim tmp() As Variant
            ReDim tmp(1 To 1, 1 To 1)
            tmp(1, 1) = lo.DataBodyRange.Value
            mDonnees = tmp
        Else
            mDonnees = lo.DataBodyRange.Value
        End If
    End If
End Sub

Private Sub AssurerCache()
    If mIdxColonne Is Nothing Then Donnees_Charger
End Sub

'==============================================================================
' Interrogation du cache
'==============================================================================
Public Function Donnees_NbLignes() As Long
    AssurerCache
    Donnees_NbLignes = mNbLignes
End Function

Public Function Donnees_IndexColonne(ByVal nomColonne As String) As Long
    AssurerCache
    If mIdxColonne.Exists(nomColonne) Then Donnees_IndexColonne = mIdxColonne(nomColonne)
End Function

'------------------------------------------------------------------------------
' Valeur brute d'une cellule du cache.
'------------------------------------------------------------------------------
Public Function Donnees_Valeur(ByVal ligne As Long, ByVal nomColonne As String) As Variant
    Dim ic As Long
    AssurerCache
    ic = Donnees_IndexColonne(nomColonne)
    If ic = 0 Or ligne < 1 Or ligne > mNbLignes Then
        Donnees_Valeur = Empty
    Else
        Donnees_Valeur = mDonnees(ligne, ic)
    End If
End Function

'------------------------------------------------------------------------------
' Valeur mise en forme pour l'affichage (dates, montants, cases à cocher).
'------------------------------------------------------------------------------
Public Function Donnees_ValeurAffichee(ByVal ligne As Long, ByVal nomColonne As String) As String
    Dim v As Variant
    v = Donnees_Valeur(ligne, nomColonne)

    If IsEmpty(v) Or IsNull(v) Then
        Donnees_ValeurAffichee = vbNullString
    ElseIf VarType(v) = vbBoolean Then
        Donnees_ValeurAffichee = IIf(v, "Oui", "Non")
    ElseIf StrComp(nomColonne, COL_DATE, vbTextCompare) = 0 Then
        If IsNumeric(v) Then
            If CDbl(v) > 0 Then Donnees_ValeurAffichee = Format$(CDate(CDbl(v)), "dd/mm/yyyy")
        ElseIf IsDate(v) Then
            Donnees_ValeurAffichee = Format$(CDate(v), "dd/mm/yyyy")
        End If
    Else
        Donnees_ValeurAffichee = CStr(v)
    End If
End Function

'------------------------------------------------------------------------------
' Numéro de ligne (dans le cache) correspondant à une Clef_BD ; 0 si absente.
'------------------------------------------------------------------------------
Public Function Donnees_TrouverLigne(ByVal clef As String) As Long
    Dim i As Long, ic As Long
    AssurerCache
    If Len(clef) = 0 Then Exit Function
    ic = Donnees_IndexColonne(COL_CLEF)
    If ic = 0 Then Exit Function
    For i = 1 To mNbLignes
        If StrComp(CStr(mDonnees(i, ic) & ""), clef, vbTextCompare) = 0 Then
            Donnees_TrouverLigne = i
            Exit Function
        End If
    Next i
End Function

'==============================================================================
' Génération de la clef : PREFIXE_CLEF + plus grand numéro existant + 1
'==============================================================================
Public Function Donnees_NouvelleClef() As String
    Dim i As Long, ic As Long, maxi As Long, n As Long, s As String
    AssurerCache
    ic = Donnees_IndexColonne(COL_CLEF)
    If ic > 0 Then
        For i = 1 To mNbLignes
            s = CStr(mDonnees(i, ic) & "")
            n = PartieNumerique(s)
            If n > maxi Then maxi = n
        Next i
    End If
    Donnees_NouvelleClef = PREFIXE_CLEF & CStr(maxi + 1)
End Function

Private Function PartieNumerique(ByVal s As String) As Long
    Dim i As Long, c As String, res As String
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If c >= "0" And c <= "9" Then res = res & c
    Next i
    If Len(res) > 0 And Len(res) <= 9 Then PartieNumerique = CLng(res)
End Function

'==============================================================================
' Écritures
'------------------------------------------------------------------------------
' "valeurs" est un Dictionary : nom de colonne -> valeur déjà typée.
' Les colonnes Clef_BD et Date_Crea sont ignorées : elles sont posées ici.
'==============================================================================
Public Function Donnees_Ajouter(ByVal valeurs As Object) As String
    Dim lo As ListObject, lr As ListRow, ligne() As Variant
    Dim clef As String

    AssurerCache
    Set lo = TableClients()
    clef = Donnees_NouvelleClef()

    ReDim ligne(1 To 1, 1 To mNbColonnes)
    RemplirLigne ligne, valeurs
    PoserValeur ligne, COL_CLEF, clef
    PoserValeur ligne, COL_DATE, Date

    On Error GoTo Fin
    Application.EnableEvents = False
    Set lr = lo.ListRows.Add
    lr.Range.Value = ligne

Fin:
    Application.EnableEvents = True
    If Err.Number <> 0 Then Err.Raise Err.Number, "Donnees_Ajouter", Err.Description

    Donnees_Charger
    Donnees_Ajouter = clef
End Function

Public Function Donnees_Modifier(ByVal clef As String, ByVal valeurs As Object) As Boolean
    Dim lo As ListObject, i As Long, ligne() As Variant, ic As Long, c As Variant

    AssurerCache
    i = Donnees_TrouverLigne(clef)
    If i = 0 Then Exit Function

    Set lo = TableClients()

    ' on repart de la ligne existante pour ne toucher qu'aux colonnes fournies
    ReDim ligne(1 To 1, 1 To mNbColonnes)
    For ic = 1 To mNbColonnes
        ligne(1, ic) = mDonnees(i, ic)
    Next ic
    RemplirLigne ligne, valeurs
    PoserValeur ligne, COL_CLEF, clef      ' la clef reste inchangée

    On Error GoTo Fin
    Application.EnableEvents = False
    lo.ListRows(i).Range.Value = ligne

Fin:
    Application.EnableEvents = True
    If Err.Number <> 0 Then Err.Raise Err.Number, "Donnees_Modifier", Err.Description

    Donnees_Charger
    Donnees_Modifier = True
End Function

Public Function Donnees_Supprimer(ByVal clef As String) As Boolean
    Dim lo As ListObject, i As Long

    AssurerCache
    i = Donnees_TrouverLigne(clef)
    If i = 0 Then Exit Function

    Set lo = TableClients()

    On Error GoTo Fin
    Application.EnableEvents = False
    lo.ListRows(i).Delete

Fin:
    Application.EnableEvents = True
    If Err.Number <> 0 Then Err.Raise Err.Number, "Donnees_Supprimer", Err.Description

    Donnees_Charger
    Donnees_Supprimer = True
End Function

Private Sub RemplirLigne(ByRef ligne() As Variant, ByVal valeurs As Object)
    Dim k As Variant, ic As Long
    For Each k In valeurs.Keys
        If StrComp(CStr(k), COL_CLEF, vbTextCompare) <> 0 _
           And StrComp(CStr(k), COL_DATE, vbTextCompare) <> 0 Then
            ic = Donnees_IndexColonne(CStr(k))
            If ic > 0 Then ligne(1, ic) = valeurs(k)
        End If
    Next k
End Sub

Private Sub PoserValeur(ByRef ligne() As Variant, ByVal nomColonne As String, ByVal v As Variant)
    Dim ic As Long
    ic = Donnees_IndexColonne(nomColonne)
    If ic > 0 Then ligne(1, ic) = v
End Sub

'==============================================================================
' Intégrité : interventions rattachées à un client (TblInterv.Client_No
' contient l'ID Cresus du client).
'==============================================================================
Public Function Donnees_NbInterventions(ByVal idCresus As String) As Long
    Dim lo As ListObject, v As Variant, i As Long, ic As Long, n As Long

    If Len(Trim$(idCresus)) = 0 Then Exit Function
    Set lo = ObtenirTable(NOM_TABLE_INTERV)
    If lo Is Nothing Then Exit Function
    If lo.ListRows.Count = 0 Then Exit Function

    ic = IndexColonne(lo, "Client_No")
    If ic = 0 Then Exit Function

    v = lo.DataBodyRange.Value
    If Not IsArray(v) Then Exit Function
    For i = LBound(v, 1) To UBound(v, 1)
        If StrComp(Trim$(CStr(v(i, ic) & "")), Trim$(idCresus), vbTextCompare) = 0 Then n = n + 1
    Next i
    Donnees_NbInterventions = n
End Function

'==============================================================================
' Doublon d'ID Cresus (hors fiche en cours d'édition)
'==============================================================================
Public Function Donnees_ClefAvecMemeIdCresus(ByVal idCresus As String, ByVal clefExclue As String) As String
    Dim i As Long, icId As Long, icClef As Long, s As String

    AssurerCache
    If Len(Trim$(idCresus)) = 0 Then Exit Function
    icId = Donnees_IndexColonne(COL_ID_CRESUS)
    icClef = Donnees_IndexColonne(COL_CLEF)
    If icId = 0 Or icClef = 0 Then Exit Function

    For i = 1 To mNbLignes
        s = CStr(mDonnees(i, icClef) & "")
        If StrComp(s, clefExclue, vbTextCompare) <> 0 Then
            If StrComp(Trim$(CStr(mDonnees(i, icId) & "")), Trim$(idCresus), vbTextCompare) = 0 Then
                Donnees_ClefAvecMemeIdCresus = s
                Exit Function
            End If
        End If
    Next i
End Function

'==============================================================================
' Valeurs distinctes d'une colonne (utilise pour compléter la liste des titres)
'==============================================================================
Public Function Donnees_ValeursDistinctes(ByVal nomColonne As String) As Variant
    Dim d As Object, i As Long, ic As Long, s As String

    AssurerCache
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    ic = Donnees_IndexColonne(nomColonne)
    If ic > 0 Then
        For i = 1 To mNbLignes
            s = Trim$(CStr(mDonnees(i, ic) & ""))
            If Len(s) > 0 Then
                If Not d.Exists(s) Then d.Add s, 1
            End If
        Next i
    End If
    Donnees_ValeursDistinctes = d.Keys
End Function

'==============================================================================
' Comparaison insensible à la casse et aux accents
'==============================================================================
Public Function Normaliser(ByVal s As String) As String
    Const AVEC As String = "àáâãäåçèéêëìíîïñòóôõöùúûüýÿÀÁÂÃÄÅÇÈÉÊËÌÍÎÏÑÒÓÔÕÖÙÚÛÜÝ"
    Const SANS As String = "aaaaaaceeeeiiiinooooouuuuyyAAAAAACEEEEIIIINOOOOOUUUUY"
    Dim i As Long, p As Long, c As String, res As String

    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        p = InStr(1, AVEC, c, vbBinaryCompare)
        If p > 0 Then c = Mid$(SANS, p, 1)
        res = res & c
    Next i
    Normaliser = LCase$(res)
End Function

'==============================================================================
' Conversions utilitaires
'==============================================================================
Public Function EnNombre(ByVal s As String, ByRef ok As Boolean) As Double
    Dim t As String
    t = Trim$(Replace$(Replace$(s, " ", ""), ",", "."))
    t = Replace$(t, "'", vbNullString)          ' séparateur de milliers suisse
    If Len(t) = 0 Then
        ok = False
        Exit Function
    End If
    If IsNumeric(t) Then
        ok = True
        EnNombre = Val(t)
    Else
        ok = False
    End If
End Function

Public Function EnTexte(ByVal v As Variant) As String
    If IsEmpty(v) Or IsNull(v) Then
        EnTexte = vbNullString
    Else
        EnTexte = CStr(v)
    End If
End Function
