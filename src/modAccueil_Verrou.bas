Attribute VB_Name = "modAccueil_Verrou"
Option Explicit
'==============================================================================
' modAccueil_Verrou - LE CLASSEUR EN MODE KIOSQUE
'------------------------------------------------------------------------------
' Verrouillé, le classeur ne montre plus que la feuille d'accueil, dans une
' fenêtre de taille fixe : ni ruban, ni onglets, ni barre de formule, ni
' en-têtes, ni ascenseurs, et les autres feuilles très masquées. L'utilisateur
' ne peut donc faire que ce que les quatre formulaires lui permettent.
'
' CE QUE CE VERROU EST, ET CE QU'IL N'EST PAS. Il met la maison en ordre, il ne
' la ferme pas à clef :
'
'   - macros désactivées à l'ouverture, et rien ne se verrouille : le classeur
'     s'ouvre nu. C'est vrai de tout verrou écrit en VBA ;
'   - la protection de structure et le mot de passe d'une feuille se lèvent en
'     quelques minutes avec un utilitaire du commerce.
'
' Il empêche les fausses manoeuvres, pas la malveillance. C'était le besoin.
'
' CES RÉGLAGES SONT CEUX D'EXCEL, PAS DU CLASSEUR. Le ruban caché l'est pour
' toute l'application : un autre classeur ouvert dans la même instance le
' trouverait caché lui aussi. D'où Accueil_Arreter, appelée à la fermeture, qui
' remet tout en place quoi qu'il arrive.
'
' LE MOT DE PASSE est dans la cellule nommée « Mot_de_passe ». S'il n'y en a
' pas, le bouton déverrouille SANS RIEN DEMANDER : un classeur dont on ne peut
' plus sortir serait pire que pas de verrou du tout.
'==============================================================================

' L'état courant, pour ne pas verrouiller deux fois ni restaurer dans le vide.
Private mVerrouille As Boolean

'==============================================================================
' LES DEUX POINTS D'ENTRÉE DU CLASSEUR
'------------------------------------------------------------------------------
' À appeler depuis ThisWorkbook. InstallerDemarrage écrit ces deux appels.
'==============================================================================

'------------------------------------------------------------------------------
' À l'ouverture : la feuille d'accueil, puis le verrou si AC_KIOSQUE le veut.
'------------------------------------------------------------------------------
Public Sub Accueil_Demarrer()
    If FeuilleAccueilManquante() Then Exit Sub

    AfficherAccueil
    If AC_KIOSQUE Then Accueil_Verrouiller
End Sub

'------------------------------------------------------------------------------
' À la fermeture : rendre à Excel son ruban et ses barres.
'
' Les feuilles restent masquées, elles : leur état est enregistré dans le
' fichier, et c'est ainsi qu'on veut le retrouver à la prochaine ouverture.
'------------------------------------------------------------------------------
Public Sub Accueil_Arreter()
    PoserInterface True
    mVerrouille = False
End Sub

'==============================================================================
' VERROUILLER ET DÉVERROUILLER
'==============================================================================
Public Sub Accueil_Verrouiller()
    Dim ws As Worksheet

    ' SANS FEUILLE D'ACCUEIL, ON NE VERROUILLE RIEN. Masquer toutes les feuilles
    ' sans en laisser une visible laisse un classeur qu'Excel refuse d'afficher.
    If FeuilleAccueilManquante() Then Exit Sub

    AfficherAccueil

    On Error Resume Next
    ThisWorkbook.Unprotect MotDePasse()
    On Error GoTo 0

    ' « Très masquée » ne se défait pas par le menu Afficher : seul le code, ou
    ' l'éditeur VBA, y revient.
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, NOM_FEUILLE_ACCUEIL, vbTextCompare) <> 0 Then
            On Error Resume Next
            ws.Visible = xlSheetVeryHidden
            On Error GoTo 0
        End If
    Next ws

    CadrerFenetre
    PoserInterface False

    On Error Resume Next
    ThisWorkbook.Protect MotDePasse(), Structure:=True, Windows:=False
    On Error GoTo 0

    mVerrouille = True
End Sub

'------------------------------------------------------------------------------
' Le bouton « Unlock » de la feuille d'accueil.
'------------------------------------------------------------------------------
Public Sub Accueil_Deverrouiller()
    Dim mdp As String, saisi As String, ws As Worksheet

    mdp = MotDePasse()
    If Len(mdp) > 0 Then
        saisi = InputBox("Mot de passe :", "Déverrouiller le classeur")
        If Len(saisi) = 0 Then Exit Sub                  ' annulé, ou vide
        If StrComp(saisi, mdp, vbBinaryCompare) <> 0 Then
            MsgBox "Mot de passe incorrect.", vbExclamation, "Déverrouiller le classeur"
            Exit Sub
        End If
    End If

    On Error Resume Next
    ThisWorkbook.Unprotect mdp
    On Error GoTo 0

    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        ws.Visible = xlSheetVisible
        On Error GoTo 0
    Next ws

    PoserInterface True
    On Error Resume Next
    Application.WindowState = xlMaximized
    On Error GoTo 0

    mVerrouille = False

    MsgBox "Le classeur est déverrouillé." & vbCrLf & vbCrLf & _
           IIf(Len(mdp) = 0, "Aucun mot de passe n'est défini dans la cellule " & _
                             "nommée " & CEL_MOT_DE_PASSE & "." & vbCrLf & vbCrLf, "") & _
           "Il se reverrouillera à la prochaine ouverture. Pour le refaire tout " & _
           "de suite : Accueil_Verrouiller.", _
           vbInformation, "Déverrouiller le classeur"
End Sub

'------------------------------------------------------------------------------
' Le bouton « Quitter » de la feuille d'accueil.
'
' Le ruban est rendu AVANT de fermer : si l'utilisateur annule, Excel reste
' utilisable, et s'il ferme, l'instance suivante ne trouve pas un ruban disparu.
'------------------------------------------------------------------------------
Public Sub Accueil_Quitter()
    Dim rep As Long

    rep = MsgBox("Enregistrer les modifications avant de quitter ?", _
                 vbQuestion + vbYesNoCancel, "Quitter")
    If rep = vbCancel Then Exit Sub

    On Error Resume Next
    If rep = vbYes Then
        ThisWorkbook.Save
    Else
        ThisWorkbook.Saved = True        ' pour qu'Excel ne redemande pas
    End If
    On Error GoTo 0

    PoserInterface True

    On Error Resume Next
    If Workbooks.Count <= 1 Then
        Application.Quit
    Else
        ThisWorkbook.Close SaveChanges:=False
    End If
    On Error GoTo 0
End Sub

'==============================================================================
' L'INTERFACE D'EXCEL
'==============================================================================

'------------------------------------------------------------------------------
' Montre ou cache tout ce qui entoure la feuille.
'
' Le ruban n'a pas de propriété : il se cache par la macro Excel 4 SHOW.TOOLBAR,
' la seule voie qui marche de 2007 à aujourd'hui sans personnaliser le ruban en
' XML — ce qui demanderait de rouvrir le classeur pour chaque changement.
'
' Tout est sous On Error Resume Next : selon la version et l'état de la fenêtre,
' l'une ou l'autre de ces propriétés refuse d'être posée, et ce n'est jamais une
' raison d'interrompre l'ouverture du classeur.
'------------------------------------------------------------------------------
Private Sub PoserInterface(ByVal visible As Boolean)
    On Error Resume Next

    Application.DisplayFormulaBar = visible
    Application.DisplayStatusBar = visible
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon""," & _
                                   IIf(visible, "True", "False") & ")"

    ActiveWindow.DisplayWorkbookTabs = visible
    ActiveWindow.DisplayHorizontalScrollBar = visible
    ActiveWindow.DisplayVerticalScrollBar = visible
    ActiveWindow.DisplayHeadings = visible
    ActiveWindow.DisplayGridlines = visible

    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' La fenêtre d'Excel à dimensions fixes, centrée.
'
' L'ÉCRAN SE MESURE EN AGRANDISSANT la fenêtre puis en lisant sa taille : c'est
' la seule façon de le connaître sans appeler l'API Windows. Et on ne dépasse
' jamais ce que l'écran offre — sur un portable, une fenêtre trop grande mettrait
' le bouton Unlock hors de portée.
'------------------------------------------------------------------------------
Private Sub CadrerFenetre()
    Dim ecranL As Single, ecranH As Single, l As Single, h As Single

    On Error Resume Next

    Application.WindowState = xlMaximized
    ecranL = Application.Width
    ecranH = Application.Height
    Application.WindowState = xlNormal

    l = AC_FEN_LARGEUR
    h = AC_FEN_HAUTEUR
    If ecranL > 0 And l > ecranL Then l = ecranL
    If ecranH > 0 And h > ecranH Then h = ecranH

    Application.Width = l
    Application.Height = h
    Application.Left = (ecranL - l) / 2
    Application.Top = (ecranH - h) / 2

    On Error GoTo 0
End Sub

'==============================================================================
' OUTILS
'==============================================================================
Private Function MotDePasse() As String
    Dim v As Variant

    v = Interv_CelluleNommee(CEL_MOT_DE_PASSE)
    If IsEmpty(v) Then Exit Function
    MotDePasse = Trim$(EnTexte(v))
End Function

'------------------------------------------------------------------------------
' True si la feuille d'accueil n'existe pas, avec le message qui va avec.
'------------------------------------------------------------------------------
Private Function FeuilleAccueilManquante() As Boolean
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(NOM_FEUILLE_ACCUEIL)
    On Error GoTo 0

    If Not ws Is Nothing Then Exit Function

    FeuilleAccueilManquante = True
    MsgBox "La feuille " & NOM_FEUILLE_ACCUEIL & " n'existe pas : le classeur " & _
           "reste ouvert tel quel." & vbCrLf & vbCrLf & _
           "Lancez GenererAccueil (module modAccueil_Generateur), puis rouvrez " & _
           "le classeur.", vbExclamation, "Verrouillage"
End Function

'==============================================================================
' INSTALLER LES DEUX APPELS DANS ThisWorkbook
'------------------------------------------------------------------------------
' Le verrou doit partir à l'ouverture du classeur, et se défaire à sa fermeture.
' Ces deux gestionnaires ne peuvent vivre que dans le module ThisWorkbook, qui
' ne s'importe pas : cette procédure les y écrit.
'
' ELLE N'ÉCRASE RIEN. Si un gestionnaire existe déjà — celui qui appelait
' AfficherAccueil, par exemple — elle le laisse et dit la ligne à y ajouter.
'==============================================================================
Public Sub InstallerDemarrage()
    Dim vbComp As Object, code As Object, txt As String, msg As String

    On Error GoTo Erreur

    Set vbComp = ThisWorkbook.VBProject.VBComponents(ThisWorkbook.CodeName)
    Set code = vbComp.CodeModule
    If code.CountOfLines > 0 Then txt = code.Lines(1, code.CountOfLines)

    If InStr(1, txt, "Workbook_Open", vbTextCompare) = 0 Then
        code.AddFromString _
            "Private Sub Workbook_Open()" & vbCrLf & _
            "    Accueil_Demarrer" & vbCrLf & _
            "End Sub"
        msg = msg & "[pose] Workbook_Open appelle Accueil_Demarrer" & vbCrLf
    Else
        msg = msg & "[a faire] Workbook_Open existe deja : ajoutez-y la ligne" & _
              vbCrLf & "          Accueil_Demarrer" & vbCrLf
    End If

    If InStr(1, txt, "Workbook_BeforeClose", vbTextCompare) = 0 Then
        code.AddFromString _
            "Private Sub Workbook_BeforeClose(Cancel As Boolean)" & vbCrLf & _
            "    Accueil_Arreter" & vbCrLf & _
            "End Sub"
        msg = msg & "[pose] Workbook_BeforeClose appelle Accueil_Arreter" & vbCrLf
    Else
        msg = msg & "[a faire] Workbook_BeforeClose existe deja : ajoutez-y la " & _
              "ligne" & vbCrLf & "          Accueil_Arreter" & vbCrLf
    End If

    MsgBox Replace$(msg, "[pose]", ChrW(10003)) & vbCrLf & _
           "Enregistrez, fermez puis rouvrez le classeur pour voir le verrou " & _
           "agir.", vbInformation, "Demarrage du classeur"
    Exit Sub

Erreur:
    MsgBox "Installation impossible :" & vbCrLf & vbCrLf & _
           Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
           "Si le message parle d'accès au projet VBA, cochez « Accès approuvé " & _
           "au modèle d'objet du projet VBA » dans les paramètres des macros, " & _
           "puis rouvrez le classeur.", vbCritical, "Demarrage du classeur"
End Sub
