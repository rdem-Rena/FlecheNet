Attribute VB_Name = "modPression_Donnees"
Option Explicit
'==============================================================================
' modPression_Donnees
'------------------------------------------------------------------------------
' Lecture et écriture de la feuille Datas.
'
' RIEN N'EST JAMAIS INSÉRÉ NI SUPPRIMÉ. Le calendrier est déjà écrit, deux
' lignes par jour jusqu'en 2028 : enregistrer une mesure, c'est remplir des
' cellules d'une ligne qui existe. C'est ce qui garde l'invariant, donc les
' plages nommées des deux graphiques, et ce qui distingue ce module de celui
' des interventions de FlècheNet, où une ligne s'ajoute au tableau.
'
' Les six colonnes calculées — Date-Compo, Demi, et les quatre lignes de
' repère — ne sont jamais écrites : leur formule doit survivre à la saisie.
'==============================================================================

'==============================================================================
' LECTURE
'==============================================================================
'------------------------------------------------------------------------------
' Valeur brute d'une cellule de Datas.
'   renvoie : la valeur, ou Empty si la ligne est hors du calendrier
'------------------------------------------------------------------------------
Public Function Pression_Valeur(ByVal ligne As Long, ByVal colonne As Long) As Variant
    Dim ws As Worksheet
    Set ws = FeuilleDatas()
    If ws Is Nothing Then Exit Function
    If ligne < LIGNE_DEB Or ligne > DerniereLigne() Then Exit Function
    Pression_Valeur = ws.Cells(ligne, colonne).Value
End Function

'------------------------------------------------------------------------------
' Valeur d'une cellule, telle qu'elle doit s'afficher dans un tableau.
'------------------------------------------------------------------------------
Public Function Pression_Affichee(ByVal ligne As Long, ByVal colonne As Long) As String
    Dim v As Variant
    v = Pression_Valeur(ligne, colonne)
    If IsEmpty(v) Then Exit Function
    If IsError(v) Then Exit Function

    Select Case colonne
        Case PC_DATE
            If IsDate(v) Then Pression_Affichee = Format$(v, "dd.mm.yy")
        Case PC_HEURE
            If IsNumeric(v) Or IsDate(v) Then Pression_Affichee = Format$(v, "hh:nn")
        Case PC_POIDS
            If IsNumeric(v) Then Pression_Affichee = Format$(v, "0.0")
        Case Else
            Pression_Affichee = CStr(v)
    End Select
End Function

'------------------------------------------------------------------------------
' Une demi-journée porte-t-elle déjà une mesure ?
'------------------------------------------------------------------------------
Public Function Pression_EstRemplie(ByVal ligne As Long) As Boolean
    Dim v As Variant
    v = Pression_Valeur(ligne, PC_SYS)
    Pression_EstRemplie = IsNumeric(v) And Not IsEmpty(v)
End Function

'------------------------------------------------------------------------------
' Les lignes des dernières mesures saisies, de la plus récente à la plus
' ancienne.
'   combien : nombre de lignes voulues
'   renvoie : un tableau de numéros de ligne, vide si le classeur n'a rien
'
' La recherche remonte depuis la fin du calendrier. Le calendrier court
' jusqu'en 2028 et se termine donc par des centaines de lignes vides : on part
' de la dernière ligne et on remonte, ce qui coûte quelques milliers de
' lectures au pire, contre 5368 pour un parcours complet.
'------------------------------------------------------------------------------
Public Function Pression_DernieresLignes(ByVal combien As Long) As Variant
    Dim res() As Long, n As Long, ligne As Long

    ReDim res(1 To combien)
    For ligne = DerniereLigne() To LIGNE_DEB Step -1
        If Pression_EstRemplie(ligne) Then
            n = n + 1
            res(n) = ligne
            If n = combien Then Exit For
        End If
    Next ligne

    If n = 0 Then
        Pression_DernieresLignes = Array()
    Else
        ReDim Preserve res(1 To n)
        Pression_DernieresLignes = res
    End If
End Function

'------------------------------------------------------------------------------
' Les activités déjà employées, sans doublon et rangées par ordre alphabétique.
' Elles alimentent le menu déroulant : l'auteur retrouve « Ext - 9.0 km » sans
' le retaper, et les orthographes ne se multiplient pas.
'------------------------------------------------------------------------------
Public Function Pression_Activites() As Variant
    Dim ws As Worksheet, vues As Object, ligne As Long, v As Variant
    Dim liste() As String, n As Long, i As Long, j As Long, tampon As String

    Set ws = FeuilleDatas()
    If ws Is Nothing Then
        Pression_Activites = Array()
        Exit Function
    End If

    Set vues = CreateObject("Scripting.Dictionary")
    vues.CompareMode = 1
    For ligne = LIGNE_DEB To DerniereLigne()
        v = ws.Cells(ligne, PC_ACTIVITE).Value
        If Not IsEmpty(v) Then
            If Len(Trim$(CStr(v))) > 0 Then
                If Not vues.Exists(Trim$(CStr(v))) Then vues.Add Trim$(CStr(v)), True
            End If
        End If
    Next ligne

    If vues.Count = 0 Then
        Pression_Activites = Array()
        Exit Function
    End If

    n = vues.Count
    ReDim liste(1 To n)
    i = 1
    For Each v In vues.Keys
        liste(i) = CStr(v)
        i = i + 1
    Next v

    ' Tri à bulles : la liste compte une centaine d'entrées, et un tri écrit
    ' sur place se relit mieux qu'un appel à une routine générale.
    For i = 1 To n - 1
        For j = 1 To n - i
            If StrComp(liste(j), liste(j + 1), vbTextCompare) > 0 Then
                tampon = liste(j)
                liste(j) = liste(j + 1)
                liste(j + 1) = tampon
            End If
        Next j
    Next i

    Pression_Activites = liste
End Function

'==============================================================================
' ÉCRITURE
'==============================================================================
'------------------------------------------------------------------------------
' Écrit une demi-journée.
'   ligne   : la ligne de Datas, telle que LigneDe l'a calculée
'   valeurs : Dictionary numéro de colonne -> valeur déjà typée. Une colonne
'             absente du dictionnaire garde ce qu'elle porte ; une valeur vide
'             efface la cellule.
'
' Les colonnes calculées sont écartées ici, et non chez l'appelant : c'est le
' seul endroit qui écrit, donc le seul où l'oubli serait possible.
'------------------------------------------------------------------------------
Public Sub Pression_Ecrire(ByVal ligne As Long, ByVal valeurs As Object)
    Dim ws As Worksheet, k As Variant, colonne As Long, v As Variant

    Set ws = FeuilleDatas()
    If ws Is Nothing Then Err.Raise vbObjectError + 700, "Pression_Ecrire", _
        "La feuille " & FEUILLE_DATAS & " est introuvable."
    If ligne < LIGNE_DEB Or ligne > DerniereLigne() Then Err.Raise vbObjectError + 701, _
        "Pression_Ecrire", "La ligne " & ligne & " est hors du calendrier."

    On Error GoTo Fin
    Application.EnableEvents = False
    For Each k In valeurs.Keys
        colonne = CLng(k)
        If Not Pression_EstCalculee(colonne) Then
            v = valeurs(k)
            If IsEmpty(v) Then
                ws.Cells(ligne, colonne).ClearContents
            ElseIf VarType(v) = vbString Then
                If Len(v) = 0 Then
                    ws.Cells(ligne, colonne).ClearContents
                Else
                    ws.Cells(ligne, colonne).Value = v
                End If
            Else
                ws.Cells(ligne, colonne).Value = v
            End If
        End If
    Next k

Fin:
    Application.EnableEvents = True
    If Err.Number <> 0 Then Err.Raise Err.Number, "Pression_Ecrire", Err.Description
End Sub

'------------------------------------------------------------------------------
' True si la colonne porte une formule et ne doit donc jamais être écrite.
'------------------------------------------------------------------------------
Public Function Pression_EstCalculee(ByVal colonne As Long) As Boolean
    Select Case colonne
        Case PC_COMPO, PC_DEMI, PC_BB, PC_BH, PC_HB, PC_HH
            Pression_EstCalculee = True
    End Select
End Function

'==============================================================================
' CONVERSIONS ET CONTRÔLES
'==============================================================================
'------------------------------------------------------------------------------
' Lit un entier saisi, entre deux bornes.
'   ok : mis à False si le texte n'est pas un entier, ou sort des bornes
' Un texte VIDE est accepté et rend 0 avec ok à True : c'est ainsi que le
' formulaire distingue « pas saisi » de « mal saisi ».
'------------------------------------------------------------------------------
Public Function Pression_Entier(ByVal texte As String, ByVal mini As Long, _
                                ByVal maxi As Long, ByRef ok As Boolean) As Long
    Dim v As Double

    ok = True
    texte = Trim$(texte)
    If Len(texte) = 0 Then Exit Function

    If Not IsNumeric(texte) Then
        ok = False
        Exit Function
    End If
    v = CDbl(texte)
    If v <> Int(v) Or v < mini Or v > maxi Then
        ok = False
        Exit Function
    End If
    Pression_Entier = CLng(v)
End Function

'------------------------------------------------------------------------------
' Lit un nombre décimal saisi, entre deux bornes. La virgule et le point sont
' acceptés l'un comme l'autre : sur un clavier suisse, le pavé numérique rend
' un point là où Excel attend une virgule.
'------------------------------------------------------------------------------
Public Function Pression_Nombre(ByVal texte As String, ByVal mini As Double, _
                                ByVal maxi As Double, ByRef ok As Boolean) As Double
    Dim v As Double

    ok = True
    texte = Trim$(Replace$(texte, ".", Application.DecimalSeparator))
    texte = Replace$(texte, ",", Application.DecimalSeparator)
    If Len(texte) = 0 Then Exit Function

    If Not IsNumeric(texte) Then
        ok = False
        Exit Function
    End If
    v = CDbl(texte)
    If v < mini Or v > maxi Then
        ok = False
        Exit Function
    End If
    Pression_Nombre = v
End Function

'------------------------------------------------------------------------------
' Lit une heure saisie, sous la forme « 7:42 », « 07:42 » ou « 742 ».
'   ok : mis à False si le texte ne s'interprète pas
'------------------------------------------------------------------------------
Public Function Pression_Heure(ByVal texte As String, ByRef ok As Boolean) As Double
    Dim heures As Long, minutes As Long, p As Long

    ok = True
    texte = Trim$(texte)
    If Len(texte) = 0 Then Exit Function

    p = InStr(texte, ":")
    If p = 0 Then p = InStr(texte, ".")
    If p = 0 Then p = InStr(texte, "h")

    If p > 0 Then
        If Not IsNumeric(Left$(texte, p - 1)) Then GoTo Faux
        heures = CLng(Left$(texte, p - 1))
        If Len(Mid$(texte, p + 1)) = 0 Then
            minutes = 0
        ElseIf IsNumeric(Mid$(texte, p + 1)) Then
            minutes = CLng(Mid$(texte, p + 1))
        Else
            GoTo Faux
        End If
    ElseIf IsNumeric(texte) And Len(texte) >= 3 Then
        ' « 742 » vaut 7:42, « 2020 » vaut 20:20
        heures = CLng(Left$(texte, Len(texte) - 2))
        minutes = CLng(Right$(texte, 2))
    ElseIf IsNumeric(texte) Then
        heures = CLng(texte)
    Else
        GoTo Faux
    End If

    If heures < 0 Or heures > 23 Or minutes < 0 Or minutes > 59 Then GoTo Faux
    Pression_Heure = (heures * 60 + minutes) / 1440
    Exit Function

Faux:
    ok = False
End Function

'------------------------------------------------------------------------------
' Lit une date saisie. Accepte « 14.09.26 », « 14.09.2026 » et « 14/09/2026 ».
'------------------------------------------------------------------------------
Public Function Pression_Date(ByVal texte As String, ByRef ok As Boolean) As Date
    ok = True
    texte = Trim$(texte)
    If Len(texte) = 0 Then
        ok = False
        Exit Function
    End If

    texte = Replace$(texte, "/", ".")
    texte = Replace$(texte, "-", ".")
    If Not IsDate(texte) Then
        ok = False
        Exit Function
    End If
    Pression_Date = Int(CDate(texte))
End Function
