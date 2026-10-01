Attribute VB_Name = "modDatas_Classeur"
Option Explicit
'==============================================================================
' modDatas_Classeur - LE CLASSEUR DE DONNÉES, OUVERT À CÔTÉ DE L'APPLICATION
'------------------------------------------------------------------------------
' L'application et les données vivent désormais dans deux fichiers :
'
'   l'APPLICATION  un .xlsm par poste, qui ne contient AUCUNE donnée — la
'                  feuille d'accueil, les modules, et deux cellules nommées.
'                  Le remplacer, c'est écraser un fichier, rien de plus.
'
'   les DONNÉES    FlecheNettoyageSA-AAAA.xlsx, un par année, dans un dossier
'                  partagé sur OneDrive. Cinq onglets, aucune macro.
'
' TOUT L'ACCÈS AUX DONNÉES PASSE PAR ICI. ObtenirTable parcourt les feuilles du
' classeur que ce module tient ouvert, et non plus celles de ThisWorkbook :
' c'est cette seule ligne qui fait basculer trente modules d'un fichier à
' l'autre.
'
' L'ANNÉE PAR DÉFAUT EST LE FICHIER LE PLUS RÉCENT PRÉSENT dans le dossier, et
' non Year(Date) : le 1er janvier, l'application pointerait sinon vers un
' -AAAA qui n'existe pas encore, et s'ouvrirait sur un message d'erreur. Elle
' ouvre ainsi toujours quelque chose de réel, et le sélecteur ne s'ouvre que si
' on le demande.
'
' UNE ANNÉE PASSÉE S'OUVRE EN CONSULTATION SEULE, sans verrou : personne n'y
' écrit, plusieurs personnes peuvent donc la relire en même temps. Seule
' l'année la plus récente s'ouvre en exclusivité — voir modDatas_Verrou.
'==============================================================================

'--- Le fichier ---------------------------------------------------------------
Public Const PREFIXE_DATAS As String = "FlecheNettoyageSA-"
Public Const EXT_DATAS As String = ".xlsx"

' Le dossier partagé, cherché sous les racines OneDrive du poste.
Public Const DOSSIER_DATAS As String = "FlecheNettoyageSA"

' Cellule nommée de l'APPLICATION : un chemin écrit à la main, qui court-circuite
' la recherche. La porte de sortie si OneDrive n'est pas là où on l'attend.
Public Const CEL_DOSSIER_DATAS As String = "Dossier_Donnees"

'--- La version du schéma -----------------------------------------------------
' Cellule nommée des DONNÉES, comparée à ce que cette application attend. Le
' jour où une colonne s'ajoutera à TblInterv, un poste resté sur une vieille
' application recevra un message clair, au lieu de voir les choses se casser en
' silence.
Public Const CEL_DATAS_VERSION As String = "Datas_Version"
Public Const DATAS_VERSION As String = "1"

'--- L'état courant -----------------------------------------------------------
Private mClasseur As Workbook
Private mAnnee As Long
Private mLecture As Boolean

' LE CHEMIN PAR LEQUEL ON A OUVERT, et non mClasseur.FullName : sur OneDrive,
' FullName rend une adresse https, dont ni Kill ni Open ne savent que faire. Le
' verrou se poserait alors sur un chemin et se lèverait sur un autre.
Private mChemin As String

' Garde-fou : Datas_Classeur appelle Datas_Ouvrir, qui appelle des fonctions
' susceptibles de rappeler Datas_Classeur. Sans lui, l'ouverture qui échoue
' repart indéfiniment.
Private mEnOuverture As Boolean

'==============================================================================
' CE QUE LE RESTE DU CLASSEUR DEMANDE
'==============================================================================

'------------------------------------------------------------------------------
' Le classeur de données, ouvert au besoin.
'   renvoie : le classeur, ou Nothing s'il n'a pas pu s'ouvrir
'------------------------------------------------------------------------------
Public Function Datas_Classeur() As Workbook
    If mClasseur Is Nothing And Not mEnOuverture Then
        Datas_Ouvrir 0, False
    End If
    Set Datas_Classeur = mClasseur
End Function

Public Function Datas_Ouverte() As Boolean
    Datas_Ouverte = Not (mClasseur Is Nothing)
End Function

Public Function Datas_Annee() As Long
    Datas_Annee = mAnnee
End Function

'------------------------------------------------------------------------------
' True si les données sont ouvertes en consultation : aucune écriture n'est
' permise, et les boutons qui écrivent se grisent.
'------------------------------------------------------------------------------
Public Function Datas_LectureSeule() As Boolean
    Datas_LectureSeule = mLecture
End Function

'==============================================================================
' LE GARDE-FOU DES ÉCRITURES
'------------------------------------------------------------------------------
' LES SEPT FONCTIONS QUI ÉCRIVENT COMMENCENT PAR L'APPELER. Elle répond à deux
' questions d'un coup :
'
'   - les données sont-elles ouvertes en écriture ?
'   - le verrou est-il TOUJOURS le nôtre ?
'
' La seconde n'est pas de la méfiance mal placée : une application laissée
' ouverte toute la nuit voit son verrou périmer, et quelqu'un d'autre le
' reprendre au matin. Mieux vaut l'apprendre avant d'écrire qu'après — et le
' passage en consultation seule, ici, évite de reposer la question à chaque
' ligne.
'
'   renvoie : True si l'on peut écrire ; sinon elle a déjà tout expliqué
'==============================================================================
Public Function Datas_PeutEcrire() As Boolean
    Dim msg As String

    If mClasseur Is Nothing Then
        MsgBox "Les données ne sont pas ouvertes.", vbExclamation, "Données"
        Exit Function
    End If

    If mLecture Then
        MsgBox Datas_RefusEcriture(), vbExclamation, "Consultation seule"
        Exit Function
    End If

    If Verrou_Confirmer(mChemin, msg) Then
        Datas_PeutEcrire = True
        Exit Function
    End If

    ' LE VERROU EST PERDU. On bascule en consultation plutôt que de laisser
    ' croire que la prochaine tentative passera — et le bandeau de l'accueil le
    ' dit aussitôt, sans quoi il annoncerait l'écriture pendant toute la suite
    ' de la séance.
    mLecture = True
    Accueil_MajBandeau
    MsgBox msg & vbCrLf & vbCrLf & _
           "Les données passent en CONSULTATION SEULE. Fermez puis rouvrez " & _
           "l'application pour reprendre la main quand ce poste aura fini.", _
           vbCritical, "Verrou perdu"
End Function

'------------------------------------------------------------------------------
' Le message affiché quand une écriture est refusée.
'------------------------------------------------------------------------------
Public Function Datas_RefusEcriture() As String
    Datas_RefusEcriture = "Les données " & CStr(mAnnee) & " sont ouvertes en " & _
                          "CONSULTATION SEULE." & vbCrLf & vbCrLf & _
                          "Une année passée ne se modifie pas, et l'année en " & _
                          "cours peut être prise par quelqu'un d'autre. Le " & _
                          "bandeau de l'accueil le rappelle."
End Function

'==============================================================================
' OUVRIR ET FERMER
'==============================================================================

'------------------------------------------------------------------------------
' Ouvre les données d'une année.
'   annee   : l'année voulue, 0 pour la plus récente présente
'   forcer  : True pour ouvrir en consultation même si l'écriture était possible
'   renvoie : True si les données sont ouvertes
'------------------------------------------------------------------------------
Public Function Datas_Ouvrir(ByVal annee As Long, ByVal forcer As Boolean) As Boolean
    Dim chemin As String, dossier As String, recente As Long
    Dim lecture As Boolean, msg As String

    If mEnOuverture Then Exit Function
    mEnOuverture = True

    dossier = Datas_Dossier()
    If Len(dossier) = 0 Then
        MsgBox "Le dossier des données est introuvable." & vbCrLf & vbCrLf & _
               "Lancez DiagnostiquerDatas (module modDatas_Classeur) : il dit " & _
               "où il a cherché, et ce qu'il a trouvé.", vbCritical, "Données"
        GoTo Fin
    End If

    recente = Datas_AnneeLaPlusRecente()
    If annee = 0 Then annee = recente
    If annee = 0 Then
        MsgBox "Aucun fichier " & PREFIXE_DATAS & "AAAA" & EXT_DATAS & " dans" & _
               vbCrLf & dossier, vbCritical, "Données"
        GoTo Fin
    End If

    chemin = dossier & Application.PathSeparator & PREFIXE_DATAS & _
             Format$(annee, "0000") & EXT_DATAS
    If Not FichierExiste(chemin) Then
        MsgBox "Fichier introuvable :" & vbCrLf & chemin, vbCritical, "Données"
        GoTo Fin
    End If

    Datas_Fermer

    ' UNE ANNÉE PASSÉE EST TOUJOURS EN CONSULTATION. Seule la plus récente
    ' demande le verrou ; s'il est pris, on PROPOSE la consultation plutôt que
    ' de refuser sèchement — on peut au moins regarder en attendant son tour.
    lecture = forcer Or (annee < recente)
    If Not lecture Then
        If Not Verrou_Prendre(chemin, msg) Then
            If MsgBox(msg & vbCrLf & vbCrLf & "Ouvrir en consultation seule ?", _
                      vbQuestion + vbYesNo, "Données déjà ouvertes") <> vbYes Then
                GoTo Fin
            End If
            lecture = True
        End If
    End If

    If Not OuvrirFichier(chemin, lecture) Then
        If Not lecture Then Verrou_Rendre chemin
        GoTo Fin
    End If

    mChemin = chemin
    mAnnee = annee
    mLecture = lecture Or mClasseur.ReadOnly      ' Excel a le dernier mot

    If Not VersionCompatible() Then
        Datas_Fermer
        GoTo Fin
    End If

    Datas_Ouvrir = True

Fin:
    mEnOuverture = False
End Function

'==============================================================================
' FERMER LES DONNÉES ET RENDRE LE VERROU
'------------------------------------------------------------------------------
' ON ENREGISTRE AVANT DE FERMER. Les formulaires écrivent directement dans les
' cellules : tout ce qui est dans le classeur à cet instant y a été mis
' volontairement, et poser la question « enregistrer ? » reviendrait à offrir
' de perdre la journée par mégarde.
'
' LA FENÊTRE REDEVIENT VISIBLE AVANT L'ENREGISTREMENT, et ce n'est pas un détail
' d'affichage : un classeur enregistré FENÊTRE MASQUÉE s'ouvre ensuite SANS
' AUCUNE FENÊTRE. Excel affiche son fond gris, aucun onglet, et grise la
' quasi-totalité du ruban faute de classeur actif. Le fichier paraît mort.
'
' SI mClasseur A ÉTÉ PERDU, ON RETROUVE LE CLASSEUR PAR SON NOM. C'est le cas qui
' a abouti au fichier illisible : voir ClasseurOrphelin.
'==============================================================================
Public Sub Datas_Fermer()
    Dim wb As Workbook, chemin As String, maj As Boolean

    Set wb = mClasseur
    If wb Is Nothing Then Set wb = ClasseurOrphelin()
    If wb Is Nothing Then
        Oublier
        Exit Sub
    End If

    chemin = mChemin
    If Len(chemin) = 0 Then chemin = wb.FullName

    maj = Application.ScreenUpdating
    On Error Resume Next
    Application.ScreenUpdating = False
    If Not mLecture Then
        MontrerFenetre wb
        wb.Save
    End If
    wb.Close SaveChanges:=False
    Application.ScreenUpdating = maj
    On Error GoTo 0

    Set mClasseur = Nothing
    If Not mLecture Then Verrou_Rendre chemin
    Oublier
End Sub

'------------------------------------------------------------------------------
' LE CLASSEUR DE DONNÉES RETROUVÉ SANS LA VARIABLE QUI LE TENAIT.
'
' mClasseur est une variable de module. Réimporter un module, taper Fin dans
' l'éditeur, ou une erreur non interceptée, et VBA REMET À ZÉRO TOUT L'ÉTAT DU
' PROJET. Le classeur reste alors ouvert, fenêtre masquée, et plus personne ne
' sait qu'il est là : Datas_Fermer croyait n'avoir rien à fermer.
'
' Excel finit par demander s'il faut l'enregistrer, et un « oui » écrit alors la
' fenêtre masquée dans le fichier. C'est ainsi que le premier fichier de données
' est devenu illisible, et c'est pourquoi on ne laisse plus d'orphelin derrière
' soi.
'
' AnneeDuNom plutôt que le préfixe seul : c'est le même test que pour lister les
' années, et il ne peut pas confondre l'application avec ses données.
'------------------------------------------------------------------------------
Private Function ClasseurOrphelin() As Workbook
    Dim wb As Workbook

    For Each wb In Workbooks
        If Not wb Is ThisWorkbook Then
            If AnneeDuNom(wb.Name) > 0 Then
                Set ClasseurOrphelin = wb
                Exit Function
            End If
        End If
    Next wb
End Function

'------------------------------------------------------------------------------
' La fenêtre du classeur de données, cachée pendant qu'on travaille, rendue
' avant tout enregistrement.
'
' MASQUÉE PENDANT LE TRAVAIL : visible, elle apparaîtrait comme une seconde
' fenêtre Excel, avec ses onglets et son quadrillage, et ferait tomber le
' kiosque, qui ne règle que la sienne.
'
' RENDUE AVANT D'ENREGISTRER : l'état de la fenêtre est écrit DANS le fichier —
' <workbookView visibility="hidden"> — et ce qui est masqué au moment de
' l'enregistrement le reste pour tous ceux qui ouvriront le fichier ensuite.
'------------------------------------------------------------------------------
Private Sub MasquerFenetre(wb As Workbook)
    On Error Resume Next
    wb.Windows(1).Visible = False
    On Error GoTo 0
End Sub

Private Sub MontrerFenetre(wb As Workbook)
    On Error Resume Next
    wb.Windows(1).Visible = True
    On Error GoTo 0
End Sub

Private Sub Oublier()
    Set mClasseur = Nothing
    mAnnee = 0
    mLecture = False
    mChemin = vbNullString
    Datas_PerimerCaches
End Sub

'==============================================================================
' LES CACHES ONT CHANGÉ DE CLASSEUR
'------------------------------------------------------------------------------
' Trois modules gardent un tableau en mémoire pour ne pas relire la feuille à
' chaque frappe. Changer d'année sans les prévenir, ce serait afficher les
' clients de 2026 au-dessus des interventions de 2025 — sans la moindre erreur,
' et sans que rien ne se voie.
'
' APPELÉE DEPUIS Oublier, c'est-à-dire à chaque fermeture, et donc aussi au début
' de chaque ouverture, qui ferme d'abord. Un cache qui survivrait à la fermeture
' rendrait des données qu'on n'a plus le droit de lire.
'
' LES TROIS INVALIDATEURS SONT PARESSEUX : ils posent un drapeau, ils ne
' relisent rien. C'est ce qui permet de les appeler alors que le classeur de
' données vient justement de se fermer.
'==============================================================================
Public Sub Datas_PerimerCaches()
    Donnees_Recharger
    Interv_ToutRecharger
    Adresses_Recharger
End Sub

'------------------------------------------------------------------------------
' Ouvre le fichier, le remet d'aplomb, puis masque sa fenêtre.
'
' DANS CET ORDRE. NormaliserDonnees rend sa fenêtre à un fichier qu'on avait
' enregistré masqué ; la masquer ensuite est notre affaire, le temps de la
' séance, et MontrerFenetre la rendra avant d'enregistrer.
'
' TOUT SOUS ScreenUpdating À FAUX : entre l'ouverture et le masquage, la fenêtre
' du classeur de données existe, et elle apparaîtrait par-dessus le kiosque.
'
' AddToMru:=False : le classeur de données n'a rien à faire dans la liste des
' documents récents, où un double clic l'ouvrirait hors de l'application, sans
' verrou.
'------------------------------------------------------------------------------
Private Function OuvrirFichier(ByVal chemin As String, ByVal lecture As Boolean) As Boolean
    Dim wb As Workbook, maj As Boolean

    maj = Application.ScreenUpdating
    On Error GoTo Erreur
    Application.ScreenUpdating = False

    Set wb = Workbooks.Open(Filename:=chemin, ReadOnly:=lecture, UpdateLinks:=0, _
                            Notify:=False, AddToMru:=False, _
                            IgnoreReadOnlyRecommended:=True)

    If Not (lecture Or wb.ReadOnly) Then NormaliserDonnees wb
    MasquerFenetre wb

    Set mClasseur = wb
    OuvrirFichier = True

Erreur:
    If Err.Number <> 0 Then
        MsgBox "Ouverture impossible :" & vbCrLf & vbCrLf & _
               Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & chemin, _
               vbCritical, "Données"
    End If
    Application.ScreenUpdating = maj
End Function

'==============================================================================
' UN FICHIER DE DONNÉES DOIT S'OUVRIR SEUL, DANS EXCEL, SANS L'APPLICATION
'------------------------------------------------------------------------------
' C'est tout l'intérêt de n'y avoir mis aucune macro, et ce n'est donc pas de la
' cosmétique. Trois choses peuvent l'en empêcher, et TOUTES LES TROIS SONT
' ENREGISTRÉES DANS LE CLASSEUR, pas dans Excel.
'
' 1. LA FENÊTRE MASQUÉE. C'est celle qui a réellement cassé le premier fichier,
'    et c'est NOUS qui la masquons, pour qu'elle ne tombe pas par-dessus le
'    kiosque. L'état de la fenêtre s'écrit dans le fichier, en clair :
'
'        <workbookView visibility="hidden" ... showSheetTabs="0" ...>
'
'    Un classeur enregistré ainsi s'ouvre ensuite SANS AUCUNE FENÊTRE : fond
'    gris, aucun onglet, et presque tout le ruban grisé faute de classeur actif.
'    Rien dans Excel ne dit pourquoi. On la rend donc avant tout enregistrement
'    — voir MontrerFenetre — et on la rend ici à un fichier qui arrive déjà
'    masqué.
'
' 2. LES FEUILLES « TRÈS MASQUÉES » et 3. LA PROTECTION DE STRUCTURE, qui
'    viennent du kiosque : copier l'application pendant qu'elle était
'    verrouillée les emporte dans la copie. Et « très masquée » ne se défait
'    pas par le menu Afficher : seul le code, ou l'éditeur VBA, y revient — dans
'    un fichier sans macro.
'
' EN ÉCRITURE SEULEMENT. Un classeur ouvert en lecture seule ne garderait rien de
' ces corrections, et les y appliquer ne ferait que le marquer modifié.
'
' On ne touche pas à la protection des FEUILLES : le kiosque n'en protège
' qu'une, celle de l'accueil, qui n'est pas dans ce fichier. Une feuille de
' données protégée l'a donc été exprès, et ce n'est pas à nous de la rouvrir.
'==============================================================================
Private Sub NormaliserDonnees(wb As Workbook)
    Dim ws As Worksheet, fen As Object

    On Error Resume Next

    ' avec le mot de passe de l'application, puis sans : elle n'en avait
    ' peut-être pas, et Unprotect refuse un mot de passe de trop
    wb.Unprotect MotDePasse()
    wb.Unprotect

    For Each ws In wb.Worksheets
        If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
    Next ws

    ' LA FENÊTRE D'ABORD. OuvrirFichier la remasque juste après, pour la durée de
    ' la séance seulement.
    MontrerFenetre wb

    Set fen = wb.Windows(1)
    fen.DisplayWorkbookTabs = True
    fen.DisplayHorizontalScrollBar = True
    fen.DisplayVerticalScrollBar = True
    fen.DisplayHeadings = True
    fen.DisplayGridlines = True

    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' La version du schéma des données doit être celle que cette application attend.
'------------------------------------------------------------------------------
Private Function VersionCompatible() As Boolean
    Dim v As String

    v = Trim$(EnTexte(Datas_CelluleNommee(CEL_DATAS_VERSION)))
    If v = DATAS_VERSION Then
        VersionCompatible = True
        Exit Function
    End If

    MsgBox "Ces données ne sont pas de la même version que l'application." & _
           vbCrLf & vbCrLf & _
           "données     : " & IIf(Len(v) > 0, v, "(cellule " & _
                                  CEL_DATAS_VERSION & " absente)") & vbCrLf & _
           "application : " & DATAS_VERSION & vbCrLf & vbCrLf & _
           "Demandez la version à jour de l'application avant d'ouvrir ces " & _
           "données : les écrire avec une application d'une autre version " & _
           "abîmerait le fichier sans rien dire.", vbCritical, "Données"
End Function

'==============================================================================
' LE DOSSIER ET LES ANNÉES DISPONIBLES
'==============================================================================

'------------------------------------------------------------------------------
' Le dossier des données : la cellule nommée de l'application si elle en porte
' un, sinon le dossier cherché sous les racines OneDrive du poste.
'------------------------------------------------------------------------------
Public Function Datas_Dossier() As String
    Dim ecrit As String

    ecrit = Trim$(EnTexte(App_CelluleNommee(CEL_DOSSIER_DATAS)))
    If Len(ecrit) > 0 Then
        ' ÉCRIT À LA MAIN, DONC SANS APPEL : si ce chemin est faux, on ne va pas
        ' chercher ailleurs en douce — le diagnostic doit pouvoir le dire.
        If DossierExiste(ecrit) Then Datas_Dossier = ecrit
        Exit Function
    End If

    Datas_Dossier = DossierOneDrive(DOSSIER_DATAS)
End Function

'------------------------------------------------------------------------------
' Les années pour lesquelles un fichier existe, de la plus récente à la plus
' ancienne.
'   renvoie : un tableau de Long de base 1, vide s'il n'y en a aucun
'------------------------------------------------------------------------------
Public Function Datas_Annees() As Variant
    Dim dossier As String, n As String, res() As Long, nb As Long
    Dim i As Long, j As Long, t As Long, an As Long, garde As Long

    Datas_Annees = Array()

    ' AVANT LA BOUCLE Dir$, et non dedans : Dir$ a un état interne unique, et
    ' Datas_Dossier s'en sert aussi pour parcourir les racines OneDrive.
    dossier = Datas_Dossier()
    If Len(dossier) = 0 Then Exit Function

    ReDim res(1 To 200)
    On Error Resume Next
    n = Dir$(dossier & Application.PathSeparator & PREFIXE_DATAS & "*" & EXT_DATAS)
    Do While Len(n) > 0 And garde < 200
        an = AnneeDuNom(n)
        If an > 0 Then
            nb = nb + 1
            res(nb) = an
        End If
        n = Dir$()
        garde = garde + 1
    Loop
    On Error GoTo 0

    If nb = 0 Then Exit Function

    ' Tri décroissant. Quelques éléments : un tri par sélection suffit.
    For i = 1 To nb - 1
        For j = i + 1 To nb
            If res(j) > res(i) Then
                t = res(i): res(i) = res(j): res(j) = t
            End If
        Next j
    Next i

    ReDim Preserve res(1 To nb)
    Datas_Annees = res
End Function

Public Function Datas_AnneeLaPlusRecente() As Long
    Dim a As Variant

    a = Datas_Annees()
    If Not IsArray(a) Then Exit Function
    If UBound(a) >= LBound(a) Then Datas_AnneeLaPlusRecente = CLng(a(LBound(a)))
End Function

'------------------------------------------------------------------------------
' L'année lue dans un nom de fichier, 0 si ce n'en est pas un des nôtres.
'
' Le filtre de Dir$ laisse passer FlecheNettoyageSA-sauvegarde.xlsx et
' FlecheNettoyageSA-2026 (copie).xlsx : c'est ici qu'on les écarte.
'------------------------------------------------------------------------------
Private Function AnneeDuNom(ByVal nomFichier As String) As Long
    Dim coeur As String

    If StrComp(Left$(nomFichier, Len(PREFIXE_DATAS)), PREFIXE_DATAS, _
               vbTextCompare) <> 0 Then Exit Function
    coeur = Mid$(nomFichier, Len(PREFIXE_DATAS) + 1)
    If StrComp(Right$(coeur, Len(EXT_DATAS)), EXT_DATAS, vbTextCompare) <> 0 Then Exit Function
    coeur = Left$(coeur, Len(coeur) - Len(EXT_DATAS))

    If Len(coeur) <> 4 Then Exit Function
    If Not QueDesChiffres(coeur) Then Exit Function
    AnneeDuNom = CLng(coeur)
End Function

'==============================================================================
' LES CELLULES NOMMÉES, DE PART ET D'AUTRE
'------------------------------------------------------------------------------
' Deux fichiers, donc deux fonctions. Les confondre, ce serait aller chercher le
' mot de passe dans les données, ou l'objectif annuel dans l'application — et
' l'une comme l'autre rendrait une cellule vide, sans la moindre erreur.
'==============================================================================
Public Function Datas_CelluleNommee(ByVal nom As String) As Variant
    If mClasseur Is Nothing Then Exit Function
    On Error Resume Next
    Datas_CelluleNommee = mClasseur.Names(nom).RefersToRange.Value
    On Error GoTo 0
End Function

Public Function App_CelluleNommee(ByVal nom As String) As Variant
    On Error Resume Next
    App_CelluleNommee = ThisWorkbook.Names(nom).RefersToRange.Value
    On Error GoTo 0
End Function

'==============================================================================
' DIAGNOSTIC
'------------------------------------------------------------------------------
' À lancer quand l'application ne trouve pas les données, ou refuse d'écrire :
' le message dit où elle a cherché, ce qu'elle a trouvé, et qui tient le verrou.
'==============================================================================
Public Sub DiagnostiquerDatas()
    Dim dossier As String, a As Variant, msg As String, i As Long, liste As String

    dossier = Datas_Dossier()
    msg = "Dossier des données :" & vbCrLf & _
          IIf(Len(dossier) > 0, dossier, "(introuvable)") & vbCrLf & vbCrLf

    If Len(dossier) > 0 Then
        a = Datas_Annees()
        If IsArray(a) Then
            If UBound(a) >= LBound(a) Then
                For i = LBound(a) To UBound(a)
                    liste = liste & IIf(Len(liste) > 0, ", ", "") & CStr(a(i))
                Next i
            End If
        End If
        msg = msg & "Années trouvées : " & IIf(Len(liste) > 0, liste, "aucune") & _
              vbCrLf & vbCrLf
    Else
        msg = msg & "Cherché dans la cellule nommée " & CEL_DOSSIER_DATAS & _
              " de cette application, puis sous les racines OneDrive du poste " & _
              "— un dossier nommé " & DOSSIER_DATAS & "." & vbCrLf & vbCrLf
    End If

    If Datas_Ouverte() Then
        msg = msg & "Ouvert : " & CStr(mAnnee) & _
              IIf(mLecture, "   (consultation seule)", "   (écriture)") & vbCrLf & _
              "Version des données : " & _
              EnTexte(Datas_CelluleNommee(CEL_DATAS_VERSION)) & _
              "   -   attendue : " & DATAS_VERSION & vbCrLf & _
              "Verrou : " & Verrou_Etat(mChemin)
    Else
        msg = msg & "Aucune année ouverte pour l'instant."
    End If

    MsgBox msg, vbInformation, "Données : état"
End Sub
