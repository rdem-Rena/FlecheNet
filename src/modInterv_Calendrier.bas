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
' Emploi :
'     d = Calendrier_Choisir(dateDeDepart, gauche, haut, ok)
'     If ok Then ... d contient la date choisie
'==============================================================================

Private mMois As Date               ' 1er jour du mois affiché
Private mChoisie As Date
Private mValide As Boolean          ' True si l'utilisateur a cliqué une date
Private mChargement As Boolean

Private Const NB_CASES As Long = 42 ' 6 semaines de 7 jours

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
    mChoisie = depart
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
'==============================================================================
Public Sub Cal_Initialiser(f As Object)
    Dim i As Long, jours As Variant, lb As Object

    mChargement = True
    jours = Array("L", "M", "M", "J", "V", "S", "D")
    For i = 0 To 6
        Set lb = CalCtl(f, "lblJS_" & CStr(i + 1))
        If Not lb Is Nothing Then lb.Caption = jours(i)
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
'==============================================================================
Public Sub Cal_Dessiner(f As Object)
    Dim i As Long, premier As Date, decalage As Long, jour As Date
    Dim lb As Object, dansLeMois As Boolean, estAujourdhui As Boolean, estChoisie As Boolean

    Set lb = CalCtl(f, "lblCalMois")
    If Not lb Is Nothing Then lb.Caption = NomDuMois(Month(mMois)) & " " & Format$(mMois, "yyyy")

    premier = mMois
    ' Weekday(..., vbMonday) vaut 1 le lundi : la première case du calendrier
    ' recule donc d'autant de jours.
    decalage = Weekday(premier, vbMonday) - 1

    For i = 1 To NB_CASES
        Set lb = CalCtl(f, "lblJ_" & CStr(i))
        If Not lb Is Nothing Then
            jour = premier - decalage + (i - 1)
            dansLeMois = (Month(jour) = Month(mMois) And Year(jour) = Year(mMois))
            estAujourdhui = (jour = Date)
            estChoisie = (jour = mChoisie)

            lb.Caption = CStr(Day(jour))
            lb.Tag = Format$(jour, "yyyy-mm-dd")

            If estChoisie Then
                lb.BackStyle = MSF_BackStyleOpaque
                lb.BackColor = COUL_MODIFIER
                lb.ForeColor = COUL_BOUTON_TXT
                lb.Font.Bold = True
            ElseIf estAujourdhui Then
                lb.BackStyle = MSF_BackStyleOpaque
                lb.BackColor = COUL_ENTETE_TBL
                lb.ForeColor = COUL_TEXTE
                lb.Font.Bold = True
            Else
                lb.BackStyle = MSF_BackStyleTransparent
                lb.ForeColor = IIf(dansLeMois, COUL_TEXTE, COUL_CHAMP_BORD)
                lb.Font.Bold = False
            End If
        End If
    Next i
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

'==============================================================================
' Clic sur une case : la date est lue dans la propriété Tag du libellé, où
' Cal_Dessiner l'a rangée au format aaaa-mm-jj.
'==============================================================================
Public Sub Cal_ChoisirJour(f As Object, ByVal indice As Long)
    Dim lb As Object, t As String

    If mChargement Then Exit Sub
    Set lb = CalCtl(f, "lblJ_" & CStr(indice))
    If lb Is Nothing Then Exit Sub

    t = lb.Tag
    If Len(t) <> 10 Then Exit Sub

    mChoisie = DateSerial(CInt(Left$(t, 4)), CInt(Mid$(t, 6, 2)), CInt(Right$(t, 2)))
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
' Survol des flèches de navigation du bandeau.
'
' Les cases de jours n'ont volontairement pas d'effet de survol : le repeindre
' à chaque déplacement de souris ferait scintiller les 42 libellés, pour un
' gain d'agrément mince — la date choisie et le jour même restent, eux,
' nettement mis en évidence.
'==============================================================================
Public Sub Cal_SurvolLien(f As Object, ByVal nom As String, ByVal actif As Boolean)
    Dim lb As Object
    Set lb = CalCtl(f, nom)
    If lb Is Nothing Then Exit Sub
    lb.ForeColor = IIf(actif, COUL_BANDEAU_TXT, COUL_BANDEAU_SOUS)
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
' Accès sûr à un contrôle du calendrier.
'------------------------------------------------------------------------------
Private Function CalCtl(f As Object, ByVal nom As String) As Object
    On Error Resume Next
    Set CalCtl = f.Controls(nom)
    On Error GoTo 0
End Function
