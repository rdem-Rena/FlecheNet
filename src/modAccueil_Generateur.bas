Attribute VB_Name = "modAccueil_Generateur"
Option Explicit
'==============================================================================
' modAccueil_Generateur - DESSINE LA FEUILLE D'ACCUEIL
'------------------------------------------------------------------------------
' UN SEUL POINT D'ENTRÉE : GenererAccueil. Il crée la feuille si elle manque,
' efface tout ce qu'elle porte, et la redessine. On peut donc la relancer autant
' de fois qu'on veut : rien ne s'accumule.
'
' Le dessin est fait de FORMES, pas de contrôles : une forme sait arrondir ses
' coins, porter une ombre et recevoir une macro au clic. Les formes d'une même
' carte sont GROUPÉES, si bien que le clic est pris n'importe où sur la carte et
' non seulement sur son fond.
'
' Les mesures, les couleurs et le texte des quatre cartes sont dans
' modAccueil_Theme ; ce module ne décide de rien.
'==============================================================================

' L'étape en cours, pour que le message d'erreur dise où le dessin a échoué.
Private mEtape As String

'==============================================================================
' POINT D'ENTRÉE
'==============================================================================
Public Sub GenererAccueil()
    Dim ws As Worksheet, n As Long

    On Error GoTo Erreur

    mEtape = "préparation de la feuille"
    Set ws = PreparerFeuille()

    mEtape = "bandeau"
    DessinerBandeau ws

    mEtape = "intitulé de section"
    Etiquette ws, "lblAcSection", AC_MARGE, AC_SECTION_TOP, 400, 16, _
              "MODULES", POLICE_DEMI, AC_T_SECTION, COUL_TEXTE_DOUX, _
              MSO_ALIGN_GAUCHE, AC_SECTION_ESPACE

    mEtape = "cartes"
    n = DessinerCartes(ws)

    mEtape = "pied de page"
    DessinerPied ws

    mEtape = "mise en page"
    HabillerFenetre ws

    mEtape = vbNullString
    MsgBox "La feuille " & NOM_FEUILLE_ACCUEIL & " est prête." & vbCrLf & vbCrLf & _
           n & " cartes, " & ws.Shapes.Count & " formes." & vbCrLf & vbCrLf & _
           "Cliquez sur une carte pour ouvrir le module correspondant.", _
           vbInformation, "Feuille d'accueil"
    Exit Sub

Erreur:
    MsgBox "Le dessin de l'accueil a échoué :" & vbCrLf & vbCrLf & _
           Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
           IIf(Len(mEtape) > 0, "Étape : " & mEtape & vbCrLf & vbCrLf, "") & _
           "La feuille est peut-être à moitié dessinée : relancez " & _
           "GenererAccueil, qui repart toujours d'une feuille vide.", _
           vbCritical, "Feuille d'accueil"
End Sub

'------------------------------------------------------------------------------
' Ramène la feuille d'accueil à l'avant du classeur et l'affiche.
' À appeler depuis Workbook_Open pour que le classeur s'ouvre dessus.
'------------------------------------------------------------------------------
Public Sub AfficherAccueil()
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(NOM_FEUILLE_ACCUEIL)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "La feuille " & NOM_FEUILLE_ACCUEIL & " n'existe pas encore." & _
               vbCrLf & vbCrLf & "Lancez GenererAccueil (module " & _
               "modAccueil_Generateur).", vbExclamation, "Feuille d'accueil"
        Exit Sub
    End If

    ws.Activate
    ws.Range("A1").Select
End Sub

'==============================================================================
' LA FEUILLE
'------------------------------------------------------------------------------
' Créée si elle manque, vidée si elle existe. Le vidage ne se contente pas des
' formes : la protection est levée d'abord, faute de quoi rien ne s'efface.
'==============================================================================
Private Function PreparerFeuille() As Worksheet
    Dim ws As Worksheet, i As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(NOM_FEUILLE_ACCUEIL)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = NOM_FEUILLE_ACCUEIL
    Else
        ws.Unprotect
        For i = ws.Shapes.Count To 1 Step -1
            ws.Shapes(i).Delete
        Next i
    End If

    ' l'accueil est la première feuille : c'est là qu'on revient
    If ws.Index > 1 Then ws.Move Before:=ThisWorkbook.Worksheets(1)

    ws.Cells.Clear
    ws.Cells.Interior.Color = COUL_FOND

    Set PreparerFeuille = ws
End Function

'------------------------------------------------------------------------------
' Cache ce qui trahit le tableur — quadrillage, en-têtes de lignes et de
' colonnes — et empêche de sélectionner une cellule : l'accueil doit se
' présenter comme un écran, pas comme une feuille de calcul.
'
' Le quadrillage est une propriété de la FENÊTRE et non de la feuille : elle ne
' se pose que sur la feuille active, d'où l'activation.
'------------------------------------------------------------------------------
Private Sub HabillerFenetre(ws As Worksheet)
    ws.Activate
    On Error Resume Next
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    On Error GoTo 0

    ws.Range("A1").Select
    ws.EnableSelection = XL_AUCUNE_SELECTION
    ws.Protect UserInterfaceOnly:=True
End Sub

'==============================================================================
' LE BANDEAU
'------------------------------------------------------------------------------
' Pleine largeur et collé au coin : sans quadrillage ni en-têtes, il ne reste
' aucune bordure de tableur pour le trahir.
'==============================================================================
Private Sub DessinerBandeau(ws As Worksheet)
    Dim sh As Object, annee As String

    Set sh = Forme(ws, MSO_RECT, "shAcBandeau", 0, 0, AC_LARGEUR, AC_BAND_HAUT)
    sh.Fill.ForeColor.RGB = COUL_BANDEAU
    sh.Line.Visible = MSO_FAUX

    Etiquette ws, "lblAcTitre", AC_MARGE, AC_TITRE_TOP, 620, 34, _
              "Flèche Nettoyage SA", POLICE_DEMI, AC_T_TITRE, COUL_BANDEAU_TXT, _
              MSO_ALIGN_GAUCHE, 0

    Etiquette ws, "lblAcSous", AC_MARGE + 2, AC_SOUS_TOP, 620, 20, _
              "Clients, interventions, facturation et statistiques", _
              POLICE, AC_T_SOUS, COUL_BANDEAU_SOUS, MSO_ALIGN_GAUCHE, 0

    annee = AnneeAccueil()
    If Len(annee) > 0 Then
        Etiquette ws, "lblAcAnnee", AC_ANNEE_X, AC_ANNEE_TOP, AC_ANNEE_LARG, 46, _
                  annee, POLICE_LEGERE, AC_T_ANNEE, COUL_BANDEAU_SOUS, _
                  MSO_ALIGN_DROITE, 0
    End If
End Sub

'------------------------------------------------------------------------------
' L'année affichée, prise dans la cellule nommée du classeur — la même que celle
' des quatre formulaires. Vide si la cellule manque : le bandeau s'en passe.
'------------------------------------------------------------------------------
Private Function AnneeAccueil() As String
    Dim v As Variant

    v = Interv_CelluleNommee(CEL_ANNEE)
    If IsEmpty(v) Then Exit Function

    On Error Resume Next
    If IsDate(v) Then
        AnneeAccueil = Format$(v, "yyyy")
    ElseIf IsNumeric(v) Then
        AnneeAccueil = Format$(CLng(v), "0000")
    Else
        AnneeAccueil = EnTexte(v)
    End If
    On Error GoTo 0
End Function

'==============================================================================
' LES CARTES
'------------------------------------------------------------------------------
' Cinq formes chacune, puis un groupe : le fond, la pastille, sa lettre, le
' filet d'accent, le titre, le détail et le lien. Le GROUPE porte la macro, si
' bien qu'un clic n'importe où sur la carte l'ouvre.
'
'   renvoie : le nombre de cartes dessinées
'==============================================================================
Private Function DessinerCartes(ws As Worksheet) As Long
    Dim ct() As CarteMenu, i As Long, x As Single, y As Single
    Dim sh As Object, grp As Object, noms(0 To 6) As Variant, suf As String

    ct = ObtenirCartesMenu()
    y = AC_CARTE_TOP

    For i = LBound(ct) To UBound(ct)
        mEtape = "carte " & ct(i).Titre
        x = AcCarteX(i)
        suf = CStr(i)

        ' --- le fond ----------------------------------------------------------
        Set sh = Forme(ws, MSO_RECT_ARRONDI, "shAcFond_" & suf, _
                       x, y, AC_CARTE_LARG, AC_CARTE_HAUT)
        sh.Adjustments(1) = 0.045
        sh.Fill.ForeColor.RGB = COUL_CARTE
        sh.Line.ForeColor.RGB = COUL_BORDURE
        sh.Line.Weight = 0.75
        Ombrer sh
        noms(0) = sh.Name

        ' --- la pastille et sa lettre -----------------------------------------
        Set sh = Forme(ws, MSO_OVALE, "shAcPast_" & suf, _
                       x + AC_CARTE_PAD, y + AC_PASTILLE_TOP, _
                       AC_PASTILLE_D, AC_PASTILLE_D)
        sh.Fill.ForeColor.RGB = ct(i).Couleur
        sh.Line.Visible = MSO_FAUX
        noms(1) = sh.Name

        Set sh = Etiquette(ws, "lblAcLettre_" & suf, _
                           x + AC_CARTE_PAD, y + AC_PASTILLE_TOP, _
                           AC_PASTILLE_D, AC_PASTILLE_D, _
                           ct(i).Lettre, POLICE_DEMI, AC_T_PASTILLE, _
                           COUL_BOUTON_TXT, MSO_ALIGN_CENTRE, 0)
        sh.TextFrame2.VerticalAnchor = MSO_ANCRE_MILIEU
        noms(2) = sh.Name

        ' --- le filet d'accent, entre la pastille et le titre ------------------
        Set sh = Forme(ws, MSO_RECT, "shAcTrait_" & suf, _
                       x + AC_CARTE_PAD, y + AC_TRAIT_TOP, _
                       AC_TRAIT_LARG, AC_TRAIT_HAUT)
        sh.Fill.ForeColor.RGB = ct(i).Couleur
        sh.Line.Visible = MSO_FAUX
        noms(3) = sh.Name

        ' --- le texte ---------------------------------------------------------
        Set sh = Etiquette(ws, "lblAcTitre_" & suf, _
                           x + AC_CARTE_PAD, y + AC_TITRE_C_TOP, _
                           AC_CARTE_LARG - 2 * AC_CARTE_PAD, 26, _
                           ct(i).Titre, POLICE_DEMI, AC_T_TITRE_C, COUL_TEXTE, _
                           MSO_ALIGN_GAUCHE, 0)
        noms(4) = sh.Name

        Set sh = Etiquette(ws, "lblAcDetail_" & suf, _
                           x + AC_CARTE_PAD, y + AC_DETAIL_TOP, _
                           AC_CARTE_LARG - 2 * AC_CARTE_PAD, AC_DETAIL_HAUT, _
                           ct(i).Detail, POLICE, AC_T_DETAIL, COUL_TEXTE_DOUX, _
                           MSO_ALIGN_GAUCHE, 0)
        sh.TextFrame2.WordWrap = MSO_VRAI
        sh.TextFrame2.TextRange.ParagraphFormat.SpaceWithin = 1.15
        noms(5) = sh.Name

        Set sh = Etiquette(ws, "lblAcLien_" & suf, _
                           x + AC_CARTE_PAD, y + AC_LIEN_TOP, _
                           AC_CARTE_LARG - 2 * AC_CARTE_PAD, 18, _
                           "Ouvrir " & ChrW(8594), POLICE_DEMI, AC_T_LIEN, _
                           ct(i).Couleur, MSO_ALIGN_GAUCHE, 0)
        noms(6) = sh.Name

        ' --- le groupe, qui reçoit la macro ------------------------------------
        Set grp = ws.Shapes.Range(noms).Group
        grp.Name = "mnuCarte_" & suf
        grp.Placement = XL_FLOTTANT
        grp.OnAction = ct(i).Macro
    Next i

    DessinerCartes = UBound(ct) - LBound(ct) + 1
End Function

'==============================================================================
' LE PIED DE PAGE
'==============================================================================
Private Sub DessinerPied(ws As Worksheet)
    Dim sh As Object

    Set sh = Forme(ws, MSO_RECT, "shAcFilet", AC_MARGE, AC_FILET_TOP, _
                   AC_LARGEUR - 2 * AC_MARGE, 1)
    sh.Fill.ForeColor.RGB = COUL_BORDURE
    sh.Line.Visible = MSO_FAUX

    Etiquette ws, "lblAcPied", AC_MARGE, AC_PIED_TOP, _
              AC_LARGEUR - 2 * AC_MARGE, 16, _
              "Cliquez sur une carte pour ouvrir le module." & _
              "     " & ChrW(183) & "     " & _
              "Les formulaires se régénèrent par les modules Generateur.", _
              POLICE, AC_T_PIED, COUL_TEXTE_DOUX, MSO_ALIGN_GAUCHE, 0
End Sub

'==============================================================================
' OUTILS DE DESSIN
'==============================================================================

'------------------------------------------------------------------------------
' Une forme pleine, nommée et détachée des cellules : sans Placement, la
' redimensionner reviendrait à redimensionner une colonne.
'------------------------------------------------------------------------------
Private Function Forme(ws As Worksheet, ByVal genre As Long, ByVal nom As String, _
                       ByVal gauche As Single, ByVal haut As Single, _
                       ByVal largeur As Single, ByVal hauteur As Single) As Object
    Dim sh As Object

    Set sh = ws.Shapes.AddShape(genre, gauche, haut, largeur, hauteur)
    sh.Name = nom
    sh.Placement = XL_FLOTTANT
    sh.Fill.Solid
    Set Forme = sh
End Function

'------------------------------------------------------------------------------
' Une zone de texte sans fond ni filet.
'
' Les marges intérieures d'une zone de texte valent par défaut deux à trois
' points : remises à zéro, la position écrite est celle du texte, et non celle
' d'une boîte qui l'entoure de loin.
'
'   espacement : interlettrage en points, 0 pour le laisser tel quel
'------------------------------------------------------------------------------
Private Function Etiquette(ws As Worksheet, ByVal nom As String, _
                           ByVal gauche As Single, ByVal haut As Single, _
                           ByVal largeur As Single, ByVal hauteur As Single, _
                           ByVal texte As String, ByVal police As String, _
                           ByVal taille As Single, ByVal couleur As Long, _
                           ByVal alignement As Long, ByVal espacement As Single) As Object
    Dim sh As Object

    Set sh = ws.Shapes.AddTextbox(MSO_TEXTE_HORIZONTAL, gauche, haut, largeur, hauteur)
    sh.Name = nom
    sh.Placement = XL_FLOTTANT
    sh.Fill.Visible = MSO_FAUX
    sh.Line.Visible = MSO_FAUX

    With sh.TextFrame2
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        .MarginBottom = 0
        .WordWrap = MSO_FAUX
        .AutoSize = MSO_AUTOSIZE_AUCUN
        .VerticalAnchor = MSO_ANCRE_HAUT
        .TextRange.Text = texte
        .TextRange.ParagraphFormat.Alignment = alignement
        .TextRange.Font.Name = police
        .TextRange.Font.Size = taille
        .TextRange.Font.Fill.ForeColor.RGB = couleur
        If espacement <> 0 Then .TextRange.Font.Spacing = espacement
    End With

    Set Etiquette = sh
End Function

'------------------------------------------------------------------------------
' L'ombre portée d'une carte : très floue, très transparente, décalée vers le
' bas seulement. Elle doit se sentir sans se voir.
'
' Sous On Error Resume Next : les versions anciennes n'exposent pas Blur, et
' une ombre un peu moins douce vaut mieux qu'un dessin qui s'arrête.
'------------------------------------------------------------------------------
Private Sub Ombrer(sh As Object)
    On Error Resume Next
    With sh.Shadow
        .Visible = MSO_VRAI
        .Style = MSO_OMBRE_EXTERIEURE
        .ForeColor.RGB = COUL_BANDEAU
        .Transparency = AC_OMBRE_TRANSP
        .Blur = AC_OMBRE_FLOU
        .OffsetX = 0
        .OffsetY = AC_OMBRE_DY
    End With
    On Error GoTo 0
End Sub
