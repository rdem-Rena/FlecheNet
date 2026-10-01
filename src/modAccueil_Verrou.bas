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

' vbext_pk_Proc : une procédure ordinaire, par opposition à un accesseur de
' propriété. Déclarée ici pour ne dépendre d'aucune référence, comme les
' constantes MSF_* et MSO_* du reste du classeur.
Private Const PROC_NORMALE As Long = 0

'==============================================================================
' LES DEUX POINTS D'ENTRÉE DU CLASSEUR
'------------------------------------------------------------------------------
' À appeler depuis ThisWorkbook. InstallerDemarrage écrit ces deux appels.
'==============================================================================

'------------------------------------------------------------------------------
' À l'ouverture : la feuille d'accueil, puis le verrou si AC_KIOSQUE le veut.
'------------------------------------------------------------------------------
Public Sub Accueil_Demarrer()
    On Error GoTo Erreur

    If FeuilleAccueilManquante() Then Exit Sub

    AfficherAccueil
    If AC_KIOSQUE Then Accueil_Verrouiller

    ' LES DONNÉES EN DERNIER, LE KIOSQUE POSÉ. Leur ouverture peut poser une
    ' question — verrou tenu par un autre poste, version du schéma — et une
    ' boîte de dialogue devant un classeur à moitié verrouillé laisserait
    ' l'utilisateur devant ses onglets, à se demander ce qui se passe.
    Datas_Ouvrir 0, False
    Accueil_MajBandeau
    Exit Sub

Erreur:
    ' Une erreur ici laisserait le classeur ouvert SANS verrou et sans rien
    ' dire : l'utilisateur verrait ses onglets et croirait le verrou inactif.
    ' On rend d'abord l'interface, puis on explique.
    PoserInterface True
    MsgBox "Le verrouillage n'a pas pu se faire :" & vbCrLf & vbCrLf & _
           Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
           "Le classeur reste ouvert normalement. Lancez DiagnostiquerVerrou " & _
           "(module modAccueil_Verrou) pour savoir ce qui manque.", _
           vbExclamation, "Verrouillage"
End Sub

'------------------------------------------------------------------------------
' À la fermeture : rendre à Excel son ruban et ses barres.
'
' Les feuilles restent masquées, elles : leur état est enregistré dans le
' fichier, et c'est ainsi qu'on veut le retrouver à la prochaine ouverture.
'------------------------------------------------------------------------------
Public Sub Accueil_Arreter()
    ' EXCEL RETROUVE SON RUBAN D'ABORD, ET QUOI QU'IL ARRIVE ENSUITE. Le ruban
    ' caché, la barre de formule et la barre d'état sont des réglages
    ' D'EXCEL : s'ils ne sont pas rendus, l'utilisateur se retrouve devant un
    ' Excel où il ne reste que les noms des onglets, pour tous ses classeurs et
    ' aux séances suivantes. Une erreur en fermant les données ne doit pas
    ' pouvoir coûter cela — c'est déjà ce qui est arrivé une fois.
    PoserInterface True
    mVerrouille = False

    ' PUIS LES DONNÉES : Datas_Fermer les enregistre et REND LE VERROU. Les
    ' laisser ouvertes derrière un classeur fermé tiendrait tous les autres
    ' postes à l'écart jusqu'à la péremption des quatre heures — et Excel
    ' finirait par demander s'il faut les enregistrer, question à laquelle un
    ' « oui » grave la fenêtre masquée dans le fichier.
    On Error Resume Next
    Datas_Fermer
    On Error GoTo 0
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

    ' CE FICHIER NE PORTE PLUS AUCUNE DONNÉE : la question de l'enregistrement
    ' ne se pose que s'il a lui-même changé — une feuille regénérée, un module
    ' modifié — ce qui ne peut arriver que verrou levé. Les données, elles,
    ' s'enregistrent toutes seules à leur fermeture.
    If ThisWorkbook.Saved Then
        If MsgBox("Quitter l'application ?", vbQuestion + vbOKCancel, _
                  "Quitter") <> vbOK Then Exit Sub
    Else
        rep = MsgBox("Ce fichier d'application a été modifié." & vbCrLf & vbCrLf & _
                     "L'enregistrer avant de quitter ?", _
                     vbQuestion + vbYesNoCancel, "Quitter")
        If rep = vbCancel Then Exit Sub
        On Error Resume Next
        If rep = vbYes Then
            ThisWorkbook.Save
        Else
            ThisWorkbook.Saved = True    ' pour qu'Excel ne redemande pas
        End If
        On Error GoTo 0
    End If

    ' AVANT DE COMPTER LES CLASSEURS. Tant que les données sont ouvertes,
    ' Workbooks.Count vaut deux : Excel ne se fermerait pas, et il resterait une
    ' fenêtre vide, sans ruban, dont on ne sait plus sortir.
    Datas_Fermer

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
' LE SÉLECTEUR D'ANNÉE
'------------------------------------------------------------------------------
' Lancé en cliquant sur l'année du bandeau, ou sur la ligne qui la commente.
'
' UNE BOÎTE DE SAISIE ET NON UN FORMULAIRE. Un cinquième UserForm généré par
' code, avec son module de thème et son générateur, pour choisir entre deux ou
' trois nombres : la liste des années tient sur une ligne, et l'année en cours
' est déjà proposée. On tape rarement autre chose qu'Entrée.
'==============================================================================
Public Sub Accueil_ChoisirAnnee()
    Dim annees As Variant, liste As String, i As Long
    Dim saisi As String, an As Long, recente As Long

    annees = Datas_Annees()
    If Not IsArray(annees) Then Exit Sub
    If UBound(annees) < LBound(annees) Then
        MsgBox "Aucun fichier de données n'a été trouvé." & vbCrLf & vbCrLf & _
               "Lancez DiagnostiquerDatas (module modDatas_Classeur) : il dit " & _
               "où l'application a cherché.", vbExclamation, "Changer d'année"
        Exit Sub
    End If

    recente = CLng(annees(LBound(annees)))
    For i = LBound(annees) To UBound(annees)
        liste = liste & IIf(Len(liste) > 0, "    ", "") & CStr(annees(i))
    Next i

    saisi = Trim$(InputBox( _
            "Années disponibles :" & vbCrLf & vbCrLf & "      " & liste & vbCrLf & _
            vbCrLf & CStr(recente) & " est l'année en cours : elle seule " & _
            "s'ouvre en écriture, et par une personne à la fois." & vbCrLf & _
            "Les précédentes s'ouvrent en consultation, et plusieurs personnes " & _
            "peuvent les lire en même temps." & vbCrLf & vbCrLf & _
            "Quelle année ouvrir ?", "Changer d'année", CStr(Interv_AnneeAffichee())))

    If Len(saisi) = 0 Then Exit Sub                  ' annulé, ou vide

    If Len(saisi) <> 4 Or Not QueDesChiffres(saisi) Then
        MsgBox "Donnez une année sur quatre chiffres.", vbExclamation, _
               "Changer d'année"
        Exit Sub
    End If

    an = CLng(saisi)
    If Not DansLaListe(an, annees) Then
        MsgBox "Il n'y a pas de fichier pour " & CStr(an) & "." & vbCrLf & vbCrLf & _
               "Années disponibles :    " & liste, vbExclamation, "Changer d'année"
        Exit Sub
    End If

    If an = Datas_Annee() Then Exit Sub              ' déjà celle-là

    Datas_Ouvrir an, False
    Accueil_MajBandeau
End Sub

Private Function DansLaListe(ByVal an As Long, ByRef annees As Variant) As Boolean
    Dim i As Long

    For i = LBound(annees) To UBound(annees)
        If CLng(annees(i)) = an Then
            DansLaListe = True
            Exit Function
        End If
    Next i
End Function

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
    Dim fen As Object

    On Error Resume Next

    Application.DisplayFormulaBar = visible
    Application.DisplayStatusBar = visible
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon""," & _
                                   IIf(visible, "True", "False") & ")"

    ' LA FENÊTRE DU CLASSEUR, et non ActiveWindow : à l'ouverture, quand Excel
    ' tournait déjà, la fenêtre active peut encore être celle d'un autre
    ' classeur — on lui cacherait ses onglets, et pas les nôtres.
    Set fen = ThisWorkbook.Windows(1)
    fen.DisplayWorkbookTabs = visible
    fen.DisplayHorizontalScrollBar = visible
    fen.DisplayVerticalScrollBar = visible
    fen.DisplayHeadings = visible
    fen.DisplayGridlines = visible

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
'------------------------------------------------------------------------------
' Le mot de passe, dans la cellule nommée de l'APPLICATION.
'
' CÔTÉ APPLICATION ET NON CÔTÉ DONNÉES : chaque poste a son fichier
' d'application, et le mot de passe n'a donc pas à voyager dans le dossier
' partagé, où n'importe qui le lirait en ouvrant le classeur.
'
' Publique parce que modDatas_Verrou la demande aussi : prendre la main sur les
' données d'un autre poste est un geste du même ordre que déverrouiller le
' classeur, et les deux doivent répondre au même mot de passe.
'------------------------------------------------------------------------------
Public Function MotDePasse() As String
    Dim v As Variant

    v = App_CelluleNommee(CEL_MOT_DE_PASSE)
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
    Dim vbComp As Object, code As Object, msg As String

    On Error GoTo Erreur

    Set vbComp = ThisWorkbook.VBProject.VBComponents(ThisWorkbook.CodeName)
    Set code = vbComp.CodeModule

    msg = AjouterAppel(code, "Workbook_Open", "Accueil_Demarrer", _
                       "Private Sub Workbook_Open()") & vbCrLf
    msg = msg & AjouterAppel(code, "Workbook_BeforeClose", "Accueil_Arreter", _
                             "Private Sub Workbook_BeforeClose(Cancel As Boolean)")

    MsgBox msg & vbCrLf & vbCrLf & _
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

'------------------------------------------------------------------------------
' Fait en sorte qu'une procédure événementielle de ThisWorkbook appelle une de
' nos procédures — en la créant si elle manque, en GLISSANT L'APPEL DANS LA
' SIENNE si elle existe déjà.
'
' C'est ce dernier cas qui compte : Workbook_Open existait presque sûrement,
' pour appeler AfficherAccueil. L'ancienne version se contentait alors de dire
' la ligne à ajouter — un message qu'on lit une fois sur deux, et le verrou ne
' partait jamais.
'
'   renvoie : la ligne de compte rendu à afficher
'------------------------------------------------------------------------------
Private Function AjouterAppel(code As Object, ByVal proc As String, _
                              ByVal appel As String, ByVal signature As String) As String
    Dim txt As String, ligne As Long, vu As String

    vu = ChrW(10003) & " "
    If code.CountOfLines > 0 Then txt = code.Lines(1, code.CountOfLines)

    If InStr(1, txt, appel, vbTextCompare) > 0 Then
        AjouterAppel = vu & proc & " appelait déjà " & appel
        Exit Function
    End If

    If InStr(1, txt, proc, vbTextCompare) = 0 Then
        code.AddFromString signature & vbCrLf & "    " & appel & vbCrLf & "End Sub"
        AjouterAppel = vu & proc & " créée, elle appelle " & appel
        Exit Function
    End If

    ' PROC_NORMALE : la procédure existe, on insère l'appel juste sous sa
    ' signature. Ni Workbook_Open ni Workbook_BeforeClose n'ont de signature
    ' coupée en deux, la ligne suivante est donc bien le début du corps.
    ligne = 0
    On Error Resume Next
    ligne = code.ProcBodyLine(proc, PROC_NORMALE)
    On Error GoTo 0

    If ligne = 0 Then
        AjouterAppel = "! " & proc & " existe mais reste introuvable : " & _
                       "ajoutez-y à la main la ligne   " & appel
        Exit Function
    End If

    code.InsertLines ligne + 1, "    " & appel
    AjouterAppel = vu & appel & " ajouté à " & proc & ", qui existait déjà"
End Function

'==============================================================================
' DIAGNOSTIC
'------------------------------------------------------------------------------
' À lancer si le classeur s'ouvre encore avec ses onglets : le message dit
' laquelle des quatre conditions n'est pas remplie.
'==============================================================================
Public Sub DiagnostiquerVerrou()
    Dim code As Object, txt As String, msg As String, ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(NOM_FEUILLE_ACCUEIL)
    Set code = ThisWorkbook.VBProject.VBComponents(ThisWorkbook.CodeName).CodeModule
    If Not code Is Nothing Then
        If code.CountOfLines > 0 Then txt = code.Lines(1, code.CountOfLines)
    End If
    On Error GoTo 0

    msg = Etat(Not ws Is Nothing, "la feuille " & NOM_FEUILLE_ACCUEIL & " existe") & vbCrLf
    msg = msg & Etat(AC_KIOSQUE, "AC_KIOSQUE vaut True (modAccueil_Theme)") & vbCrLf
    msg = msg & Etat(Not code Is Nothing, "le module ThisWorkbook est lisible") & vbCrLf
    msg = msg & Etat(InStr(1, txt, "Accueil_Demarrer", vbTextCompare) > 0, _
                     "Workbook_Open appelle Accueil_Demarrer") & vbCrLf
    msg = msg & Etat(InStr(1, txt, "Accueil_Arreter", vbTextCompare) > 0, _
                     "Workbook_BeforeClose appelle Accueil_Arreter") & vbCrLf
    msg = msg & Etat(Len(MotDePasse()) > 0, _
                     "la cellule " & CEL_MOT_DE_PASSE & " porte un mot de passe")

    MsgBox msg & vbCrLf & vbCrLf & _
           "Une croix sur les deux appels : lancez InstallerDemarrage." & vbCrLf & _
           "Une croix sur le mot de passe : le bouton Unlock ne demandera rien.", _
           vbInformation, "Verrouillage : etat"
End Sub

Private Function Etat(ByVal ok As Boolean, ByVal quoi As String) As String
    Etat = IIf(ok, ChrW(10003), ChrW(215)) & "  " & quoi
End Function
