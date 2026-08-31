Attribute VB_Name = "modInterv_Calendrier"
Option Explicit
'==============================================================================
' modInterv_Calendrier
'------------------------------------------------------------------------------
' Sélecteur de date du formulaire des interventions.
'
' MSForms ne fournit aucun calendrier, et le contrôle MonthView de Microsoft
' repose sur MSCOMCT2.OCX, absent de la plupart des postes. Le calendrier est
' donc entièrement construit en contrôles standard : 7 en-têtes de jours et une
' grille de 42 cases, générés comme le reste du formulaire.
'
' Ce que fait la mise en forme, et pourquoi :
'   - la grille est aérée (34 x 26 points par case) : on vise une case du bout
'     de la souris sans effort ;
'   - les colonnes du samedi et du dimanche sont teintées en fond, ce qui
'     découpe la semaine sans ajouter le moindre trait ;
'   - la case sous la souris s'éclaire, donc on sait toujours ce qu'un clic
'     choisirait ;
'   - trois états se distinguent d'un coup d'oeil : le jour choisi (pastille
'     pleine), aujourd'hui (fond clair), les jours des mois voisins (gris) ;
'   - deux raccourcis en pied : « Aujourd'hui » et « Fin de mois », les deux
'     dates que l'on saisit le plus souvent dans une facturation mensuelle.
'
' Emploi :
'     d = Calendrier_Choisir(dateDeDepart, gauche, haut, ok)
'     If ok Then ... d contient la date choisie
'==============================================================================

Private mMois As Date               ' 1er jour du mois affiché
Private mChoisie As Date
Private mValide As Boolean          ' True si l'utilisateur a cliqué une date
Private mChargement As Boolean

Private Const NB_CASES As Long = 42 ' 6 semaines de 7 jours

' État de chaque case, calculé une fois par Cal_Dessiner puis relu par le
' survol : repeindre une case demande de savoir à quoi la ramener.
Private Const ETAT_NORMAL As Long = 0
Private Const ETAT_HORS As Long = 1     ' jour d'un mois voisin
Private Const ETAT_AUJOURDHUI As Long = 2
Private Const ETAT_CHOISIE As Long = 3

Private mJours(1 To NB_CASES) As Date
Private mEtat(1 To NB_CASES) As Long
Private mSurvol As Long             ' case sous la souris, 0 si aucune

'==============================================================================
' Ouvre le calendrier et renvoie la date choisie.
'   depart  : date affichée à l'ouverture
'   gauche, haut : position à l'écran, en points
'   ok      : passé à True si une date a été choisie, False si l'utilisateur
'             a renoncé
'==============================================================================
Public Function Calendrier_Choisir(ByVal depart As Date, ByVal gauche As Single, _
                                   ByVal haut As Single, ByRef ok As Boolean) As Date
    Dim f As Object

    ok = False
    mValide = False
    ' l'heure est retirée : les cases de la grille sont comparées à la date
    ' choisie, et 31/08/2026 08:30 ne serait égal à aucune d'elles
    mChoisie = DateSerial(Year(depart), Month(depart), Day(depart))
    mMois = DateSerial(Year(depart), Month(depart), 1)

    On Error GoTo Erreur
    Set f = UserForms.Add(NOM_FORM_CALENDRIER)
    f.Left = gauche
    f.Top = haut
    f.Show

    ok = mValide
    Calendrier_Choisir = mChoisie
    Exit Function

Erreur:
    MsgBox "Le calendrier " & NOM_FORM_CALENDRIER & " n'a pas pu s'ouvrir." & vbCrLf & _
           "Lancez GenererFormulaireInterventions pour le créer.", _
           vbExclamation, "Sélecteur de date"
End Function

'==============================================================================
' Préparation à l'ouverture, appelée par UserForm_Initialize du calendrier.
'
' Les initiales du samedi et du dimanche prennent la teinte douce des libellés :
' l'en-tête dit la même chose que le fond teinté des deux colonnes, en plus
' discret encore.
'==============================================================================
Public Sub Cal_Initialiser(f As Object)
    Dim i As Long, jours As Variant, lb As Object

    mChargement = True
    mSurvol = 0

    ' Width et Height d'un UserForm désignent ses dimensions EXTÉRIEURES, barre
    ' de titre comprise : sans cette correction, la rangée de raccourcis du bas
    ' se retrouverait rognée d'une vingtaine de points, comme sur le formulaire
    ' principal.
    On Error Resume Next
    f.Width = f.Width + (CAL_LARGEUR - f.InsideWidth)
    f.Height = f.Height + (CAL_HAUTEUR - f.InsideHeight)
    On Error GoTo 0

    jours = Array("L", "M", "M", "J", "V", "S", "D")
    For i = 0 To 6
        Set lb = CalCtl(f, "lblJS_" & CStr(i + 1))
        If Not lb Is Nothing Then
            lb.Caption = jours(i)
            lb.ForeColor = IIf(i >= 5, COUL_CHAMP_BORD, COUL_TEXTE_DOUX)
        End If
    Next i
    mChargement = False

    Cal_Dessiner f
End Sub

'==============================================================================
' Remplit la grille pour le mois courant.
'
' La grille commence au lundi de la semaine contenant le 1er du mois : les
' derniers jours du mois précédent et les premiers du suivant sont affichés en
' gris clair, comme dans tout calendrier.
'
' Les dates et les états sont mémorisés au passage : le survol s'en sert pour
' repeindre une seule case, sans avoir à recalculer le mois entier.
'==============================================================================
Public Sub Cal_Dessiner(f As Object)
    Dim i As Long, decalage As Long, jour As Date, lb As Object

    Set lb = CalCtl(f, "lblCalMois")
    If Not lb Is Nothing Then lb.Caption = NomDuMois(Month(mMois)) & " " & Format$(mMois, "yyyy")

    ' Weekday(..., vbMonday) vaut 1 le lundi : la première case du calendrier
    ' recule donc d'autant de jours.
    decalage = Weekday(mMois, vbMonday) - 1
    mSurvol = 0

    For i = 1 To NB_CASES
        jour = mMois - decalage + (i - 1)
        mJours(i) = jour

        If jour = mChoisie Then
            mEtat(i) = ETAT_CHOISIE
        ElseIf jour = Date Then
            mEtat(i) = ETAT_AUJOURDHUI
        ElseIf Month(jour) <> Month(mMois) Or Year(jour) <> Year(mMois) Then
            mEtat(i) = ETAT_HORS
        Else
            mEtat(i) = ETAT_NORMAL
        End If

        Set lb = CalCtl(f, "lblJ_" & CStr(i))
        If Not lb Is Nothing Then lb.Caption = CStr(Day(jour))
        PeindreCase f, i, False
    Next i
End Sub

'------------------------------------------------------------------------------
' Donne à une case son aspect, d'après son état et la présence de la souris.
'   survole : True si la souris se trouve dessus
'
' Une case au repos est TRANSPARENTE, jamais blanche : c'est ainsi que la teinte
' des colonnes du week-end, posée en fond derrière la grille, reste visible.
'------------------------------------------------------------------------------
Private Sub PeindreCase(f As Object, ByVal i As Long, ByVal survole As Boolean)
    Dim lb As Object

    If i < 1 Or i > NB_CASES Then Exit Sub
    Set lb = CalCtl(f, "lblJ_" & CStr(i))
    If lb Is Nothing Then Exit Sub

    Select Case mEtat(i)

        Case ETAT_CHOISIE               ' pastille pleine : la date retenue
            lb.BackStyle = MSF_BackStyleOpaque
            lb.BackColor = IIf(survole, COUL_MODIFIER_H, COUL_MODIFIER)
            lb.ForeColor = COUL_BOUTON_TXT
            lb.Font.Bold = True

        Case ETAT_AUJOURDHUI            ' fond clair : le jour même
            lb.BackStyle = MSF_BackStyleOpaque
            lb.BackColor = IIf(survole, COUL_CAL_SURVOL, COUL_ENTETE_TBL)
            lb.ForeColor = COUL_MODIFIER
            lb.Font.Bold = True

        Case ETAT_HORS                  ' mois voisin : présent mais en retrait
            lb.BackStyle = IIf(survole, MSF_BackStyleOpaque, MSF_BackStyleTransparent)
            lb.BackColor = COUL_CAL_SURVOL
            lb.ForeColor = COUL_CHAMP_BORD
            lb.Font.Bold = False

        Case Else
            lb.BackStyle = IIf(survole, MSF_BackStyleOpaque, MSF_BackStyleTransparent)
            lb.BackColor = COUL_CAL_SURVOL
            lb.ForeColor = COUL_TEXTE
            lb.Font.Bold = survole
    End Select
End Sub

'==============================================================================
' Survol d'une case.
'   indice : numéro de case, ou 0 pour n'en survoler aucune
'
' Deux cases repeintes au maximum — celle que l'on quitte, celle que l'on prend.
' Repasser les 42 à chaque déplacement de souris ferait scintiller la grille.
'==============================================================================
Public Sub Cal_SurvolJour(f As Object, ByVal indice As Long)
    If mChargement Then Exit Sub
    If indice = mSurvol Then Exit Sub

    If mSurvol >= 1 Then PeindreCase f, mSurvol, False
    mSurvol = indice
    If indice >= 1 Then PeindreCase f, indice, True
End Sub

'==============================================================================
' Navigation
'==============================================================================
Public Sub Cal_MoisPrecedent(f As Object)
    mMois = DateAdd("m", -1, mMois)
    Cal_Dessiner f
End Sub

'------------------------------------------------------------------------------
' Affiche le mois suivant.
'------------------------------------------------------------------------------
Public Sub Cal_MoisSuivant(f As Object)
    mMois = DateAdd("m", 1, mMois)
    Cal_Dessiner f
End Sub

'------------------------------------------------------------------------------
' Ramène l'affichage sur le mois en cours sans rien choisir : l'utilisateur
' garde la main pour cliquer, ou non, la date du jour.
'------------------------------------------------------------------------------
Public Sub Cal_Aujourdhui(f As Object)
    mMois = DateSerial(Year(Date), Month(Date), 1)
    mChoisie = Date
    Cal_Dessiner f
End Sub

'------------------------------------------------------------------------------
' Choisit d'emblée le dernier jour du mois affiché et referme : c'est la date
' que porte la plupart des interventions facturées au mois.
'
' Le 0e jour du mois suivant est le dernier du mois affiché — la façon usuelle
' de l'obtenir sans se soucier des mois de 28, 30 ou 31 jours.
'------------------------------------------------------------------------------
Public Sub Cal_FinMois(f As Object)
    mChoisie = DateSerial(Year(mMois), Month(mMois) + 1, 0)
    mValide = True
    Unload f
End Sub

'==============================================================================
' Clic sur une case : la date est lue dans le tableau rempli par Cal_Dessiner.
'==============================================================================
Public Sub Cal_ChoisirJour(f As Object, ByVal indice As Long)
    If mChargement Then Exit Sub
    If indice < 1 Or indice > NB_CASES Then Exit Sub

    mChoisie = mJours(indice)
    mValide = True
    Unload f
End Sub

'------------------------------------------------------------------------------
' Referme le calendrier sans rien choisir : la date du formulaire reste
' inchangée. Fermer la fenêtre par sa croix produit le même effet, le drapeau de
' validation restant à False.
'------------------------------------------------------------------------------
Public Sub Cal_Annuler(f As Object)
    mValide = False
    Unload f
End Sub

'==============================================================================
' Survol des flèches de navigation du bandeau : elles s'éclaircissent jusqu'au
' blanc du bandeau.
'==============================================================================
Public Sub Cal_SurvolLien(f As Object, ByVal nom As String, ByVal actif As Boolean)
    Dim lb As Object
    Set lb = CalCtl(f, nom)
    If lb Is Nothing Then Exit Sub
    lb.ForeColor = IIf(actif, COUL_BANDEAU_TXT, COUL_BANDEAU_SOUS)
End Sub

'------------------------------------------------------------------------------
' Survol des raccourcis du pied : soulignés le temps du passage, comme un lien.
'   annuler : True pour « Annuler », qui vit sur la teinte discrète et vire au
'             rouge plutôt qu'au bleu
'------------------------------------------------------------------------------
Public Sub Cal_SurvolPied(f As Object, ByVal nom As String, ByVal actif As Boolean, _
                          ByVal annuler As Boolean)
    Dim lb As Object

    Set lb = CalCtl(f, nom)
    If lb Is Nothing Then Exit Sub

    If annuler Then
        lb.ForeColor = IIf(actif, COUL_SUPPRIMER, COUL_TEXTE_DOUX)
    Else
        lb.ForeColor = IIf(actif, COUL_LIEN_H, COUL_LIEN)
    End If
    lb.Font.Underline = actif
End Sub

'------------------------------------------------------------------------------
' Souris hors de tout élément actif : on éteint le survol de la grille et des
' liens en une fois.
'------------------------------------------------------------------------------
Public Sub Cal_SurvolRepos(f As Object)
    Cal_SurvolJour f, 0
    Cal_SurvolLien f, "lblCalPrec", False
    Cal_SurvolLien f, "lblCalSuiv", False
    Cal_SurvolPied f, "lblCalAujourdhui", False, False
    Cal_SurvolPied f, "lblCalFinMois", False, False
    Cal_SurvolPied f, "lblCalAnnuler", False, True
End Sub

'==============================================================================
' Nom du mois en français, indépendamment des paramètres régionaux de Windows :
' le formulaire doit s'afficher pareil sur tous les postes.
'==============================================================================
Public Function NomDuMois(ByVal m As Long) As String
    Dim noms As Variant
    noms = Array("Janvier", "Février", "Mars", "Avril", "Mai", "Juin", _
                 "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre")
    If m >= 1 And m <= 12 Then NomDuMois = noms(m - 1)
End Function

'------------------------------------------------------------------------------
' Mois abrégé, pour l'axe du graphique.
'
' Pas un Left$(nom, 3) : « Juin » et « Juillet » donneraient tous deux « Jui »,
' et deux colonnes voisines porteraient la même étiquette.
'------------------------------------------------------------------------------
Public Function MoisAbrege(ByVal m As Long) As String
    Dim noms As Variant
    noms = Array("Jan", "F" & ChrW(233) & "v", "Mars", "Avr", "Mai", "Juin", _
                 "Juil", "Ao" & ChrW(251) & "t", "Sep", "Oct", "Nov", "D" & ChrW(233) & "c")
    If m >= 1 And m <= 12 Then MoisAbrege = noms(m - 1)
End Function

'------------------------------------------------------------------------------
' Accès sûr à un contrôle du calendrier.
'------------------------------------------------------------------------------
Private Function CalCtl(f As Object, ByVal nom As String) As Object
    On Error Resume Next
    Set CalCtl = f.Controls(nom)
    On Error GoTo 0
End Function
