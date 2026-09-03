Attribute VB_Name = "modStat_Formulaire"
Option Explicit
'==============================================================================
' modStat_Formulaire
'------------------------------------------------------------------------------
' Le comportement de UF_Statistiques.
'
' UN SEUL JEU DE LIGNES ALIMENTE TOUT. Les filtres produisent une liste
' d'index, et les tuiles, le graphique, la barre d'objectif et le tableau s'y
' rapportent tous. Aucun des quatre ne peut donc contredire les autres, et
' changer un filtre les met à jour d'un même mouvement.
'
' Le tableau est dessiné en libellés avec défilement virtuel : ce ne sont pas
' les vingt lignes qui bougent mais leur contenu.
'==============================================================================

Private mLignes As Variant       ' index de lignes de TblInterv, apres filtrage
Private mPremiere As Long        ' premiere ligne affichee, 1 = la premiere
Private mSel As Long             ' ligne choisie DANS mLignes, 0 = aucune
Private mAjuste As Boolean
Private mChargement As Boolean   ' vrai pendant le remplissage des filtres

'==============================================================================
' OUVERTURE
'------------------------------------------------------------------------------
' Appelée par le bouton « Info » du formulaire des interventions.
'==============================================================================
Public Sub OuvrirStatistiques()
    Dim f As Object

    If TableInterventions() Is Nothing Then
        MsgBox "Le tableau " & NOM_TABLE_INTERVENTIONS & " est introuvable dans ce classeur.", _
               vbCritical, "Statistiques"
        Exit Sub
    End If

    On Error GoTo Erreur
    Set f = UserForms.Add(NOM_FORM_STAT)
    f.Show
    Exit Sub

Erreur:
    If Err.Number = 424 Or Err.Number = 5 Then
        MsgBox "Le formulaire " & NOM_FORM_STAT & " n'existe pas encore dans ce classeur." & _
               vbCrLf & vbCrLf & "Lancez d'abord GenererFormulaireStatistiques " & _
               "(module modStat_Generateur).", vbExclamation, "Statistiques"
    Else
        MsgBox "Ouverture impossible :" & vbCrLf & vbCrLf & _
               Err.Number & " - " & Err.Description, vbCritical, "Statistiques"
    End If
End Sub

'==============================================================================
' CYCLE DE VIE
'==============================================================================
Public Sub Stat_Initialiser(f As Object)
    Dim c As Object, i As Long

    ' Le module survit à la fermeture du formulaire : sans ces remises à zéro,
    ' une deuxième ouverture reprendrait l'état de la première.
    mAjuste = False
    mPremiere = 1
    mSel = 0
    mChargement = True

    Set c = SCtl(f, "lblSAnnee")
    If Not c Is Nothing Then c.Caption = SAnneeAffichee()

    ' le mois n'est pas un nombre à l'écran mais un nom : la position dans la
    ' liste sert de numéro, le premier choix valant « pas de filtre »
    Set c = SCtl(f, "cboSMois")
    If Not c Is Nothing Then
        c.Clear
        c.AddItem TOUS_LES_MOIS_S
        For i = 1 To 12
            c.AddItem Format$(DateSerial(2000, i, 1), "mmmm")
        Next i
        c.ListIndex = 0
    End If

    SPoserImageTuiles f
    mChargement = False
    Stat_Rafraichir f
End Sub

' La hauteur d'un UserForm est sa hauteur EXTÉRIEURE, barre de titre comprise :
' la surface utile se corrige donc à l'ouverture, une seule fois.
Public Sub Stat_Activer(f As Object)
    If mAjuste Then Exit Sub
    mAjuste = True
    On Error Resume Next
    f.Width = f.Width + (ST_LARGEUR - f.InsideWidth)
    f.Height = f.Height + (ST_HAUTEUR - f.InsideHeight)
    On Error GoTo 0
End Sub

Public Sub Stat_Quitter(f As Object)
    Unload f
End Sub

'==============================================================================
' REMETTRE LES FILTRES DANS LEUR ÉTAT D'ORIGINE
'------------------------------------------------------------------------------
' Les sept d'un coup, et un seul recalcul à la fin : mChargement l'empêche de
' se déclencher sept fois pendant qu'on vide les contrôles.
'==============================================================================
Public Sub Stat_Reinitialiser(f As Object)
    Dim c As Object, i As Long
    Dim vides As Variant, cases As Variant

    mChargement = True

    Set c = SCtl(f, "cboSMois")
    If Not c Is Nothing Then c.ListIndex = 0

    vides = Array("txtSEntreprise", "txtSNom", "txtSNoFacture")
    For i = LBound(vides) To UBound(vides)
        Set c = SCtl(f, CStr(vides(i)))
        If Not c Is Nothing Then c.Text = vbNullString
    Next i

    cases = Array("chkSTVA", "chkSForfait", "chkSFacture")
    For i = LBound(cases) To UBound(cases)
        Set c = SCtl(f, CStr(cases(i)))
        If Not c Is Nothing Then c.Value = False
    Next i

    mPremiere = 1
    mSel = 0
    mChargement = False
    Stat_Rafraichir f
End Sub

'==============================================================================
' LE RECALCUL COMPLET
'------------------------------------------------------------------------------
' Un seul chemin : les filtres donnent les lignes, les lignes donnent tout le
' reste. Appelé par les sept filtres.
'==============================================================================
Public Sub Stat_Rafraichir(f As Object)
    Dim crit As FiltreStat

    If mChargement Then Exit Sub

    crit = SLireFiltres(f)
    mLignes = Stat_Lignes(crit)
    If mSel > SNb(mLignes) Then mSel = 0
    If mPremiere > SNb(mLignes) Then mPremiere = 1

    PeindreTuiles f
    PeindreGraphique f
    PeindreObjectif f
    PeindreGrille f
End Sub

'------------------------------------------------------------------------------
' Ce que les sept contrôles de filtrage disent, réuni en un objet.
'------------------------------------------------------------------------------
Private Function SLireFiltres(f As Object) As FiltreStat
    Dim c As Object, r As FiltreStat

    Set c = SCtl(f, "cboSMois")
    If Not c Is Nothing Then
        If c.ListIndex > 0 Then r.Mois = c.ListIndex
    End If
    r.Entreprise = STexteDe(f, "txtSEntreprise")
    r.Nom = STexteDe(f, "txtSNom")
    r.NoFacture = STexteDe(f, "txtSNoFacture")
    r.TVA = STriEtat(f, "chkSTVA")
    r.Forfait = STriEtat(f, "chkSForfait")
    r.Facture = STriEtat(f, "chkSFacture")

    SLireFiltres = r
End Function

'==============================================================================
' LES CINQ TUILES
'==============================================================================
Private Sub PeindreTuiles(f As Object)
    Dim tu As Variant, i As Long, c As Object, cle As String

    tu = STuiles()
    For i = LBound(tu) To UBound(tu)
        Set c = SCtl(f, "lblSTuileVal_" & CStr(i + 1))
        If Not c Is Nothing Then
            cle = CStr(tu(i)(1))
            c.Caption = SFormatTuile(cle, Stat_Total(mLignes, cle))
        End If
    Next i
End Sub

'------------------------------------------------------------------------------
' Chaque tuile a son unité : un compte est entier, des heures s'écrivent en
' heures et minutes, un montant porte deux décimales.
'------------------------------------------------------------------------------
Private Function SFormatTuile(ByVal cle As String, ByVal v As Double) As String
    Select Case cle
        Case TU_NB:     SFormatTuile = Format$(v, "#,##0")
        Case TU_HEURES: SFormatTuile = Interv_HeuresVersTexte(v)
        Case Else:      SFormatTuile = Format$(v, "#,##0.00")
    End Select
End Function

'------------------------------------------------------------------------------
' L'image de fond des tuiles, si le classeur en a une à côté de lui.
'------------------------------------------------------------------------------
Private Sub SPoserImageTuiles(f As Object)
    Dim chemin As String, i As Long, c As Object

    chemin = Interv_CheminImageTuile()
    If Len(chemin) = 0 Then Exit Sub

    On Error Resume Next
    For i = 1 To ST_NB_TUILES
        Set c = SCtl(f, "imgSTuile_" & CStr(i))
        If Not c Is Nothing Then Set c.Picture = LoadPicture(chemin)
    Next i
    On Error GoTo 0
End Sub

'==============================================================================
' LE GRAPHIQUE MENSUEL
'------------------------------------------------------------------------------
' Les barres partent de ST_GR_X et ST_GR_Y, la MÊME origine que le décor posé
' par le générateur : leurs coordonnées comptent depuis le cadre, comme lui.
'==============================================================================
Private Sub PeindreGraphique(f As Object)
    Dim ca As Variant, i As Long, maxi As Double, echelle As Double
    Dim gauche As Single, largeur As Single, base As Single, pas As Single
    Dim c As Object, h As Single

    ca = Stat_CAParMois(mLignes)
    For i = 1 To 12
        If ca(i) > maxi Then maxi = ca(i)
    Next i
    echelle = SEchelleHaute(maxi)

    Set c = SCtl(f, "lblSGrMax")
    If Not c Is Nothing Then c.Caption = Format$(echelle, "#,##0")

    gauche = ST_GR_X + GR_MARGE_G
    largeur = ST_GR_LARG - GR_MARGE_G
    base = ST_GR_Y + GR_TRACE_TOP + GR_TRACE_HAUT
    pas = largeur / GR_NB_MOIS

    For i = 1 To GR_NB_MOIS
        h = 0
        If echelle > 0 Then h = CSng(ca(i) / echelle) * GR_TRACE_HAUT
        If h < 1 Then h = 1                      ' un trait reste visible à zéro

        Set c = SCtl(f, "lblSGrBarre_" & CStr(i))
        If Not c Is Nothing Then
            c.Left = AuPixel(gauche + (i - 1) * pas + (pas - GR_BARRE_LARG) / 2)
            c.Top = AuPixel(base - h)
            c.Height = AuPixel(h)
            c.ControlTipText = Format$(DateSerial(2000, i, 1), "mmmm") & " : " & _
                               Format$(ca(i), "#,##0.00")
        End If

        Set c = SCtl(f, "lblSGrMois_" & CStr(i))
        If Not c Is Nothing Then
            c.Left = AuPixel(gauche + (i - 1) * pas)
            c.Width = AuPixel(pas)
            c.Caption = UCase$(Left$(Format$(DateSerial(2000, i, 1), "mmm"), 3))
        End If
    Next i
End Sub

'------------------------------------------------------------------------------
' Une échelle ronde au-dessus du plus grand mois : 1, 2 ou 5 fois une puissance
' de dix. Une échelle collée au maximum donnerait une barre pleine hauteur, sans
' repère pour la lire.
'------------------------------------------------------------------------------
Private Function SEchelleHaute(ByVal maxi As Double) As Double
    Dim p As Double, m As Double

    If maxi <= 0 Then
        SEchelleHaute = 0
        Exit Function
    End If
    p = 10 ^ Int(Log(maxi) / Log(10#))
    m = maxi / p
    If m <= 1 Then
        SEchelleHaute = p
    ElseIf m <= 2 Then
        SEchelleHaute = 2 * p
    ElseIf m <= 5 Then
        SEchelleHaute = 5 * p
    Else
        SEchelleHaute = 10 * p
    End If
End Function

'==============================================================================
' LA BARRE D'OBJECTIF
'------------------------------------------------------------------------------
' Verticale, elle se remplit du BAS vers le haut : c'est le Top du remplissage
' qui descend, sa base restant collée au fond de la cuve.
'==============================================================================
Private Sub PeindreObjectif(f As Object)
    Dim c As Object, ca As Double, objectif As Double, part As Double, h As Single

    ca = Stat_Total(mLignes, TU_CA)
    objectif = Stat_ObjectifAnnuel()

    If objectif > 0 Then
        part = ca / objectif
        If part < 0 Then part = 0
        If part > 1 Then part = 1               ' la cuve ne déborde pas
        h = CSng(part) * (ST_OBJ_HAUT - 2)
    End If
    If h < 1 Then h = 1

    Set c = SCtl(f, "lblSObjPlein")
    If Not c Is Nothing Then
        c.Height = AuPixel(h)
        c.Top = AuPixel(ST_OBJ_TOP + ST_OBJ_HAUT - 1 - h)
    End If

    Set c = SCtl(f, "lblSObjPct")
    If Not c Is Nothing Then
        If objectif > 0 Then
            c.Caption = Format$(ca / objectif, "0 %")
        Else
            c.Caption = ChrW(8212)               ' tiret cadratin : rien à calculer
        End If
    End If

    Set c = SCtl(f, "lblSObjDetail")
    If Not c Is Nothing Then
        If objectif > 0 Then
            c.Caption = Format$(ca, "#,##0") & " sur " & Format$(objectif, "#,##0")
        Else
            c.Caption = "La cellule " & CEL_OBJECTIF & " est absente ou vide."
        End If
    End If
End Sub

'==============================================================================
' LE TABLEAU
'==============================================================================
Public Sub Stat_GrilleClic(f As Object, ByVal ligneEcran As Long)
    Dim idx As Long

    idx = mPremiere + ligneEcran - 1
    If idx < 1 Or idx > SNb(mLignes) Then Exit Sub
    mSel = idx
    PeindreGrille f
End Sub

Public Sub Stat_Defiler(f As Object)
    Dim c As Object

    Set c = SCtl(f, "sbStat")
    If c Is Nothing Then Exit Sub
    mPremiere = c.Value + 1
    PeindreGrille f
End Sub

Private Sub PeindreGrille(f As Object)
    Dim cols As Variant, r As Long, i As Long, idx As Long, lb As Object
    Dim choisie As Boolean, vide As Boolean

    cols = SColonnes()
    For r = 1 To ST_Z5_LIGNES
        idx = mPremiere + r - 1
        vide = (idx < 1 Or idx > SNb(mLignes))
        choisie = (Not vide) And (idx = mSel)

        Set lb = SCtl(f, "lblSL_" & CStr(r))
        If Not lb Is Nothing Then lb.BackColor = SFondLigne(r, choisie)

        For i = LBound(cols) To UBound(cols)
            Set lb = SCtl(f, "lblS_" & CStr(r) & "_" & CStr(i - LBound(cols) + 1))
            If Not lb Is Nothing Then
                If vide Then
                    lb.Caption = vbNullString
                Else
                    lb.Caption = SValeur(CLng(mLignes(idx)), CStr(cols(i)))
                End If
                lb.ForeColor = IIf(choisie, COUL_BOUTON_TXT, COUL_GRILLE_TXT)
            End If
        Next i
    Next r

    SMajBarre f
    SMajCompteur f
End Sub

'------------------------------------------------------------------------------
' Ce qu'affiche une case. Le CA est CALCULÉ et jamais lu, les deux colonnes
' booléennes montrent un « vu » ou rien.
'------------------------------------------------------------------------------
Private Function SValeur(ByVal ligne As Long, ByVal colonne As String) As String
    If colonne = IC_CA Then
        SValeur = Format$(Fact_CADeLaLigne(ligne), "#,##0.00")
    ElseIf SEstCase(colonne) Then
        SValeur = IIf(Fact_EnBooleen(Interv_Valeur(ligne, colonne)), FCoche(), vbNullString)
    Else
        SValeur = Interv_ValeurAffichee(ligne, colonne)
    End If
End Function

' La barre suit le nombre de lignes EN TROP, jamais le total : sa plage vaut
' zéro quand tout tient à l'écran, et le curseur disparaît alors.
Private Sub SMajBarre(f As Object)
    Dim c As Object, maxi As Long, premiere As Long

    Set c = SCtl(f, "sbStat")
    If c Is Nothing Then Exit Sub

    maxi = SNb(mLignes) - ST_Z5_LIGNES
    If maxi < 0 Then maxi = 0
    premiere = mPremiere
    If premiere - 1 > maxi Then premiere = maxi + 1

    On Error Resume Next
    c.Max = maxi
    c.Enabled = (maxi > 0)
    c.Value = premiere - 1
    On Error GoTo 0
End Sub

Private Sub SMajCompteur(f As Object)
    Dim c As Object

    Set c = SCtl(f, "lblSCompteur")
    If c Is Nothing Then Exit Sub
    c.Caption = SNb(mLignes) & " intervention(s) sur " & Interv_NbLignes() & _
                " retenue(s) par les filtres"
End Sub

'==============================================================================
' PETITS SERVICES
'==============================================================================

Private Function SNb(ByVal t As Variant) As Long
    On Error Resume Next
    If IsArray(t) Then SNb = UBound(t) - LBound(t) + 1
    On Error GoTo 0
    If SNb < 0 Then SNb = 0
End Function

' La ligne choisie d'abord, sinon une ligne sur deux teintée : sur treize
' colonnes, l'oeil suit une ligne bien mieux ainsi.
Private Function SFondLigne(ByVal r As Long, ByVal choisie As Boolean) As Long
    If choisie Then
        SFondLigne = COUL_MODIFIER
    ElseIf r Mod 2 = 0 Then
        SFondLigne = COUL_GRILLE_ZEBRE
    Else
        SFondLigne = COUL_CARTE
    End If
End Function

Private Function STexteDe(f As Object, ByVal nom As String) As String
    Dim c As Object

    Set c = SCtl(f, nom)
    If Not c Is Nothing Then STexteDe = Trim$(EnTexte(c.Text))
End Function

' L'état d'une case à cocher : cochée garde les lignes vraies, décochée les
' fausses.
'
' Null ne devrait plus se présenter — les cases n'ont que deux états depuis que
' l'état grisé a été retiré — mais un formulaire resté ouvert d'une version
' précédente pourrait en porter un. Il se lit alors comme décoché, ce qui est
' précisément le comportement demandé.
Private Function STriEtat(f As Object, ByVal nom As String) As Long
    Dim c As Object

    STriEtat = TRI_FAUX
    Set c = SCtl(f, nom)
    If c Is Nothing Then Exit Function

    On Error Resume Next
    If IsNull(c.Value) Then
        STriEtat = TRI_FAUX
    ElseIf c.Value = True Then
        STriEtat = TRI_VRAI
    Else
        STriEtat = TRI_FAUX
    End If
    On Error GoTo 0
End Function

Private Function SAnneeAffichee() As String
    Dim v As Variant

    v = Interv_CelluleNommee(CEL_ANNEE)
    If IsEmpty(v) Then Exit Function
    On Error Resume Next
    If IsDate(v) Then
        SAnneeAffichee = Format$(v, "yyyy")
    ElseIf IsNumeric(v) Then
        SAnneeAffichee = Format$(CLng(v), "0000")
    Else
        SAnneeAffichee = EnTexte(v)
    End If
    On Error GoTo 0
End Function

' La collection Controls d'un UserForm est plate : elle trouve aussi ce qui est
' dans un cadre.
Private Function SCtl(f As Object, ByVal nom As String) As Object
    On Error Resume Next
    Set SCtl = f.Controls(nom)
    On Error GoTo 0
End Function
