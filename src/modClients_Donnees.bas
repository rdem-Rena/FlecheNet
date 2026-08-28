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
' Chargement du cache
'==============================================================================
'------------------------------------------------------------------------------
' Relit tout le tableau TblClients et le range en mémoire.
' Un seul accès à la feuille suffit à charger les 354 fiches : le reste du code
' travaille ensuite sur ce tableau en mémoire, ce qui rend le filtrage et le tri
' instantanés. Appelée à l'ouverture du formulaire et après chaque écriture.
'------------------------------------------------------------------------------
Public Sub Donnees_Charger()
    Dim lo As ListObject, lc As ListColumn

    Set lo = TableClients()
    If lo Is Nothing Then
        Err.Raise vbObjectError + 513, "modClients_Donnees", _
                  "Le tableau " & NOM_TABLE_CLIENTS & " est introuvable dans ce classeur."
    End If

    ' Un dictionnaire nom de colonne -> numéro évite de reparcourir les
    ' en-têtes à chaque lecture. CompareMode doit être posé AVANT le premier
    ' ajout : il ne peut plus être changé ensuite.
    Set mIdxColonne = CreateObject("Scripting.Dictionary")
    mIdxColonne.CompareMode = 1                     ' 1 = insensible à la casse
    For Each lc In lo.ListColumns
        If Not mIdxColonne.Exists(lc.Name) Then mIdxColonne.Add lc.Name, lc.Index
    Next lc
    mNbColonnes = lo.ListColumns.Count

    If lo.ListRows.Count = 0 Then
        mNbLignes = 0
        mDonnees = Empty
    Else
        mNbLignes = lo.ListRows.Count
        ' Range.Value rend un tableau à deux dimensions, SAUF pour une plage
        ' d'une seule cellule où il rend la valeur brute : ce cas est reconstruit
        ' à la main pour que le reste du code n'ait qu'une forme à traiter.
        If mNbLignes = 1 And mNbColonnes = 1 Then
            Dim tmp() As Variant
            ReDim tmp(1 To 1, 1 To 1)
            tmp(1, 1) = lo.DataBodyRange.Value
            mDonnees = tmp
        Else
            mDonnees = lo.DataBodyRange.Value
        End If
    End If
End Sub

'------------------------------------------------------------------------------
' Charge le cache s'il ne l'est pas déjà. Placée en tête des procédures de
' lecture pour qu'aucune ne dépende de l'ordre des appels.
'------------------------------------------------------------------------------
Private Sub AssurerCache()
    If mIdxColonne Is Nothing Then Donnees_Charger
End Sub

'==============================================================================
' Interrogation du cache
'==============================================================================
'------------------------------------------------------------------------------
' Nombre de fiches dans le tableau.
'------------------------------------------------------------------------------
Public Function Donnees_NbLignes() As Long
    AssurerCache
    Donnees_NbLignes = mNbLignes
End Function

'------------------------------------------------------------------------------
' Position d'une colonne dans le cache.
'   renvoie : le numéro de colonne, ou 0 si elle n'existe pas
'------------------------------------------------------------------------------
Public Function Donnees_IndexColonne(ByVal nomColonne As String) As Long
    AssurerCache
    If mIdxColonne.Exists(nomColonne) Then Donnees_IndexColonne = mIdxColonne(nomColonne)
End Function

'------------------------------------------------------------------------------
' Valeur brute d'une cellule, telle qu'Excel la rend : nombre, date, booléen
' ou texte.
'   ligne      : numéro de fiche dans le cache, à partir de 1
'   nomColonne : nom de la colonne dans TblClients
' Renvoie Empty si la ligne ou la colonne n'existe pas, plutôt que de déclencher
' une erreur : les appelants n'ont pas à se protéger.
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
' Valeur convertie en texte lisible, pour le tableau et les zones de saisie.
' Les dates deviennent jj/mm/aaaa, le taux prend deux décimales, les booléens
' deviennent Oui / Non, et une cellule vide devient une chaîne vide.
'------------------------------------------------------------------------------
Public Function Donnees_ValeurAffichee(ByVal ligne As Long, ByVal nomColonne As String) As String
    Dim v As Variant
    v = Donnees_Valeur(ligne, nomColonne)

    If IsEmpty(v) Or IsNull(v) Then
        Donnees_ValeurAffichee = vbNullString
    ElseIf VarType(v) = vbBoolean Then
        Donnees_ValeurAffichee = IIf(v, "Oui", "Non")
    ElseIf StrComp(nomColonne, COL_DATE, vbTextCompare) = 0 Then
        ' Excel rend une date soit en Date, soit en numéro de série : les deux
        ' cas mènent au même affichage jj/mm/aaaa.
        If IsNumeric(v) Then
            If CDbl(v) > 0 Then Donnees_ValeurAffichee = Format$(CDate(CDbl(v)), "dd/mm/yyyy")
        ElseIf IsDate(v) Then
            Donnees_ValeurAffichee = Format$(CDate(v), "dd/mm/yyyy")
        End If
    ElseIf StrComp(nomColonne, COL_TAUX, vbTextCompare) = 0 And IsNumeric(v) Then
        ' Montant : deux décimales, avec le séparateur décimal de Windows.
        Donnees_ValeurAffichee = Format$(CDbl(v), "0.00")
    Else
        Donnees_ValeurAffichee = CStr(v)
    End If
End Function

'------------------------------------------------------------------------------
' Retrouve une fiche par sa Clef_BD.
'   renvoie : le numéro de ligne dans le cache, ou 0 si la clef est introuvable
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
' Génération de la clef
'==============================================================================
'------------------------------------------------------------------------------
' Fabrique la clef de la prochaine fiche : CL suivi du plus grand numéro
' présent, augmenté de 1.
' Le plus grand numéro est recalculé à chaque fois plutôt que déduit du nombre de
' lignes : supprimer une fiche au milieu du tableau ne provoque donc jamais de
' clef en double.
'------------------------------------------------------------------------------
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

'------------------------------------------------------------------------------
' Extrait les chiffres d'une chaîne : CL349 donne 349.
' Renvoie 0 si la chaîne n'en contient aucun, ou plus de neuf (au-delà, la
' conversion en Long déborderait).
'------------------------------------------------------------------------------
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
'------------------------------------------------------------------------------
' Ajoute une fiche en fin de tableau.
'   valeurs : Dictionary nom de colonne -> valeur déjà typée
'   renvoie : la Clef_BD attribuée
'
' Clef_BD et Date_Crea sont posées ici, jamais par l'appelant : c'est ce qui
' garantit qu'elles restent cohérentes quoi que contienne le formulaire.
'------------------------------------------------------------------------------
Public Function Donnees_Ajouter(ByVal valeurs As Object) As String
    Dim lo As ListObject, lr As ListRow, ligne() As Variant
    Dim clef As String

    AssurerCache
    Set lo = TableClients()
    clef = Donnees_NouvelleClef()

    ' La ligne est préparée entièrement en mémoire, puis écrite en un seul
    ' accès à la feuille : c'est rapide, et les formats et formules de colonne
    ' du tableau structuré sont préservés.
    ReDim ligne(1 To 1, 1 To mNbColonnes)
    RemplirLigne ligne, valeurs
    PoserValeur ligne, COL_CLEF, clef
    PoserValeur ligne, COL_DATE, Date

    ' EnableEvents est remis à True par l'étiquette Fin même en cas d'erreur :
    ' le laisser à False figerait le recalcul dans tout Excel.
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

'------------------------------------------------------------------------------
' Met à jour une fiche existante.
'   clef    : Clef_BD de la fiche à modifier
'   valeurs : Dictionary nom de colonne -> valeur ; les colonnes absentes du
'             dictionnaire gardent leur valeur actuelle
'   renvoie : True si la fiche a été trouvée et écrite
'------------------------------------------------------------------------------
Public Function Donnees_Modifier(ByVal clef As String, ByVal valeurs As Object) As Boolean
    Dim lo As ListObject, i As Long, ligne() As Variant, ic As Long, c As Variant

    AssurerCache
    i = Donnees_TrouverLigne(clef)
    If i = 0 Then Exit Function

    Set lo = TableClients()

    ' on repart de la ligne existante pour ne toucher qu'aux colonnes fournies
    ' On repart de la ligne telle qu'elle est dans le tableau, puis on n'écrase
    ' que les colonnes fournies : une colonne absente du formulaire, ou ajoutée
    ' plus tard dans TblClients, conserve ainsi sa valeur.
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

'------------------------------------------------------------------------------
' Supprime la ligne du tableau correspondant à une Clef_BD.
'   renvoie : True si la fiche existait
' La confirmation de l'utilisateur est demandée en amont, par le formulaire.
'------------------------------------------------------------------------------
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

'------------------------------------------------------------------------------
' Reporte les valeurs d'un Dictionary dans un tableau d'une ligne, prêt à être
' écrit dans la feuille.
' Clef_BD et Date_Crea sont volontairement ignorées : elles sont posées juste
' après par l'appelant, ce qui empêche le formulaire de les écraser.
'------------------------------------------------------------------------------
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

'------------------------------------------------------------------------------
' Écrit une valeur dans la ligne en préparation, en silence si la colonne
' n'existe pas dans ce classeur.
'------------------------------------------------------------------------------
Private Sub PoserValeur(ByRef ligne() As Variant, ByVal nomColonne As String, ByVal v As Variant)
    Dim ic As Long
    ic = Donnees_IndexColonne(nomColonne)
    If ic > 0 Then ligne(1, ic) = v
End Sub

'==============================================================================
' Intégrité référentielle
'==============================================================================
'------------------------------------------------------------------------------
' Compte les interventions rattachées à un client.
'   idCresus : identifiant Crésus du client
'   renvoie  : le nombre de lignes de TblInterv qui le référencent
'
' La colonne Client_No de TblInterv contient l'ID Crésus, et non la Clef_BD.
' Sert à prévenir avant une suppression, sans jamais l'empêcher.
'------------------------------------------------------------------------------
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
' Doublons
'==============================================================================
'------------------------------------------------------------------------------
' Cherche une autre fiche portant le même ID Crésus.
'   clefExclue : fiche en cours d'édition, à ne pas comparer avec elle-même
'   renvoie    : la Clef_BD de la fiche en doublon, ou une chaîne vide
'------------------------------------------------------------------------------
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
' Valeurs distinctes
'==============================================================================
'------------------------------------------------------------------------------
' Valeurs différentes présentes dans une colonne, sans les vides.
' Sert à compléter le menu déroulant Titre avec les civilités déjà utilisées.
'------------------------------------------------------------------------------
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
' Normalisation
'==============================================================================
'------------------------------------------------------------------------------
' Met une chaîne en minuscules et remplace les caractères accentués par leur
' équivalent sans accent.
'   renvoie : la chaîne comparable
'
' C'est ce qui permet de trouver « Apothéloz » en tapant « apotheloz », et de
' trier « Éric » à sa vraie place alphabétique.
'------------------------------------------------------------------------------
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
'------------------------------------------------------------------------------
' Convertit une saisie en nombre, en acceptant les conventions locales.
'   ok      : passé à True si la conversion a réussi
'   renvoie : le nombre, ou 0 en cas d'échec
' La virgule décimale et l'apostrophe des milliers (1'250.50) sont admises.
'------------------------------------------------------------------------------
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

'------------------------------------------------------------------------------
' Convertit en texte sans déclencher d'erreur sur Empty ou Null.
' À utiliser partout où la valeur vient d'un contrôle ou d'une cellule.
'------------------------------------------------------------------------------
Public Function EnTexte(ByVal v As Variant) As String
    If IsEmpty(v) Or IsNull(v) Then
        EnTexte = vbNullString
    Else
        EnTexte = CStr(v)
    End If
End Function
