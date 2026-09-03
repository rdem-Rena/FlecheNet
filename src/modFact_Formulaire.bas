Attribute VB_Name = "modFact_Formulaire"
Option Explicit
'==============================================================================
' modFact_Formulaire
'------------------------------------------------------------------------------
' Le comportement de UF_Facture. Le module de code du formulaire ne fait
' qu'appeler ici : régénérer le formulaire n'écrase donc jamais rien.
'
' DEUX GRILLES DESSINÉES EN LIBELLÉS, chacune avec son défilement virtuel. Ce
' ne sont pas les lignes qui bougent mais leur contenu : neuf lignes de cases
' pour le tableau du haut, douze pour celui du bas, repeintes à chaque
' défilement. Un client qui aurait mille interventions ne coûte donc pas un
' contrôle de plus.
'
' L'état vit dans les variables ci-dessous, jamais dans les contrôles : la
' sélection, la position du défilement et les cases cochées survivent ainsi à
' un repeint.
'==============================================================================

Private mLignesC As Variant          ' index de lignes de TblInterv, tableau du haut
Private mLignesT As Variant          ' index de lignes de TblInterv, tableau du bas
Private mPremierC As Long            ' première ligne affichée, 1 = la première
Private mPremierT As Long
Private mSelC As Long                ' ligne choisie DANS mLignesC, 0 = aucune
Private mSelT As Long
Private mCoches As Object            ' index de ligne -> True, les cases Select.
Private mAjuste As Boolean

'==============================================================================
' OUVERTURE
'------------------------------------------------------------------------------
' Appelée par le bouton « Facturer » du formulaire des interventions.
'==============================================================================
Public Sub OuvrirFacturation()
    Dim f As Object

    If TableInterventions() Is Nothing Then
        MsgBox "Le tableau " & NOM_TABLE_INTERVENTIONS & " est introuvable dans ce classeur.", _
               vbCritical, "Facturation"
        Exit Sub
    End If

    On Error GoTo Erreur
    Set f = UserForms.Add(NOM_FORM_FACTURE)
    f.Show
    Exit Sub

Erreur:
    If Err.Number = 424 Or Err.Number = 5 Then
        MsgBox "Le formulaire " & NOM_FORM_FACTURE & " n'existe pas encore dans ce classeur." & _
               vbCrLf & vbCrLf & "Lancez d'abord GenererFormulaireFacture " & _
               "(module modFact_Generateur).", vbExclamation, "Facturation"
    Else
        MsgBox "Ouverture impossible :" & vbCrLf & vbCrLf & _
               Err.Number & " - " & Err.Description, vbCritical, "Facturation"
    End If
End Sub

'==============================================================================
' CYCLE DE VIE
'==============================================================================
Public Sub Fact_Initialiser(f As Object)
    Dim c As Object, i As Long

    Set mCoches = CreateObject("Scripting.Dictionary")
    ' Le module survit à la fermeture du formulaire : sans cette remise à zéro,
    ' une deuxième ouverture ne corrigerait plus sa hauteur.
    mAjuste = False
    mPremierC = 1
    mPremierT = 1
    mSelC = 0
    mSelT = 0

    Set c = FCtl(f, "lblFAnnee")
    If Not c Is Nothing Then c.Caption = FAnneeAffichee()

    ' le mois n'est pas un nombre à l'écran mais un nom : la position dans la
    ' liste sert de numéro de mois, le premier choix valant « pas de filtre »
    Set c = FCtl(f, "cboFMois")
    If Not c Is Nothing Then
        c.Clear
        c.AddItem TOUS_LES_MOIS_F
        For i = 1 To 12
            c.AddItem Format$(DateSerial(2000, i, 1), "mmmm")
        Next i
        c.ListIndex = 0
    End If

    Fact_ToutRecharger f
End Sub

' La hauteur d'un UserForm est sa hauteur EXTÉRIEURE, barre de titre comprise :
' la surface utile se corrige donc à l'ouverture, une seule fois.
Public Sub Fact_Activer(f As Object)
    If mAjuste Then Exit Sub
    mAjuste = True
    On Error Resume Next
    f.Width = f.Width + (FA_LARGEUR - f.InsideWidth)
    f.Height = f.Height + (FA_HAUTEUR - f.InsideHeight)
    On Error GoTo 0
End Sub

Public Sub Fact_Quitter(f As Object)
    Unload f
End Sub

'==============================================================================
' CHARGEMENT
'------------------------------------------------------------------------------
' Relit TblInterv, refait la liste des clients à facturer et repeint tout. À
' appeler après chaque écriture dans le classeur.
'==============================================================================
Public Sub Fact_ToutRecharger(f As Object)
    Interv_ToutRecharger
    mLignesC = Fact_ClientsNonFactures()
    If mSelC > FNb(mLignesC) Then mSelC = 0
    If mPremierC > FNb(mLignesC) Then mPremierC = 1
    PeindreClients f
    Fact_RafraichirTravaux f
End Sub

'==============================================================================
' TABLEAU DU HAUT — les clients
'==============================================================================
Public Sub Fact_ClientClic(f As Object, ByVal ligneEcran As Long)
    Dim idx As Long

    idx = mPremierC + ligneEcran - 1
    If idx < 1 Or idx > FNb(mLignesC) Then Exit Sub

    mSelC = idx
    mSelT = 0
    mPremierT = 1
    PeindreClients f
    Fact_RafraichirTravaux f
End Sub

Public Sub Fact_DefilerClients(f As Object)
    Dim c As Object

    Set c = FCtl(f, "sbFClients")
    If c Is Nothing Then Exit Sub
    mPremierC = c.Value + 1
    PeindreClients f
End Sub

Private Sub PeindreClients(f As Object)
    Dim cols As Variant, r As Long, i As Long, idx As Long, lb As Object
    Dim choisie As Boolean, vide As Boolean

    cols = FClientsColonnes()
    For r = 1 To FA_Z2_LIGNES
        idx = mPremierC + r - 1
        vide = (idx < 1 Or idx > FNb(mLignesC))
        choisie = (Not vide) And (idx = mSelC)

        Set lb = FCtl(f, "lblFCL_" & CStr(r))
        If Not lb Is Nothing Then lb.BackColor = FFondLigne(r, choisie)

        For i = LBound(cols) To UBound(cols)
            Set lb = FCtl(f, "lblFC_" & CStr(r) & "_" & CStr(i - LBound(cols) + 1))
            If Not lb Is Nothing Then
                If vide Then
                    lb.Caption = vbNullString
                Else
                    lb.Caption = Interv_ValeurAffichee(CLng(mLignesC(idx)), CStr(cols(i)))
                End If
                lb.ForeColor = IIf(choisie, COUL_BOUTON_TXT, COUL_GRILLE_TXT)
            End If
        Next i
    Next r

    MajBarre f, "sbFClients", FNb(mLignesC), FA_Z2_LIGNES, mPremierC
End Sub

'==============================================================================
' TABLEAU DU BAS — les travaux du client choisi
'==============================================================================
Public Sub Fact_RafraichirTravaux(f As Object)
    Dim clientNo As String

    If mSelC >= 1 And mSelC <= FNb(mLignesC) Then
        clientNo = EnTexte(Interv_Valeur(CLng(mLignesC(mSelC)), IC_CLIENT))
    End If

    mLignesT = Fact_TravauxDuClient(clientNo, FCochee(f, "chkFToutes"), FMoisChoisi(f))
    If mSelT > FNb(mLignesT) Then mSelT = 0
    If mPremierT > FNb(mLignesT) Then mPremierT = 1
    PeindreTravaux f
End Sub

'------------------------------------------------------------------------------
' Un clic dans le tableau du bas.
'   colonne = 0 : la bande de fond, entre deux cases
'   colonne = celle de Select. : la case bascule, la ligne ne change pas
'   sinon : la ligne devient la ligne choisie
'------------------------------------------------------------------------------
Public Sub Fact_TravailClic(f As Object, ByVal ligneEcran As Long, ByVal colonne As Long)
    Dim idx As Long, cle As String

    idx = mPremierT + ligneEcran - 1
    If idx < 1 Or idx > FNb(mLignesT) Then Exit Sub

    If colonne = FIndexTravaux(FC_SELECT) Then
        cle = CStr(CLng(mLignesT(idx)))
        If mCoches.Exists(cle) Then
            mCoches.Remove cle
        Else
            mCoches.Add cle, True
        End If
    Else
        mSelT = idx
    End If
    PeindreTravaux f
End Sub

Public Sub Fact_DefilerTravaux(f As Object)
    Dim c As Object

    Set c = FCtl(f, "sbFTravaux")
    If c Is Nothing Then Exit Sub
    mPremierT = c.Value + 1
    PeindreTravaux f
End Sub

Private Sub PeindreTravaux(f As Object)
    Dim cols As Variant, r As Long, i As Long, idx As Long, lb As Object
    Dim choisie As Boolean, vide As Boolean

    cols = FTravauxColonnes()
    For r = 1 To FA_Z4_LIGNES
        idx = mPremierT + r - 1
        vide = (idx < 1 Or idx > FNb(mLignesT))
        choisie = (Not vide) And (idx = mSelT)

        Set lb = FCtl(f, "lblFTL_" & CStr(r))
        If Not lb Is Nothing Then lb.BackColor = FFondLigne(r, choisie)

        For i = LBound(cols) To UBound(cols)
            Set lb = FCtl(f, "lblFT_" & CStr(r) & "_" & CStr(i - LBound(cols) + 1))
            If Not lb Is Nothing Then
                If vide Then
                    lb.Caption = vbNullString
                Else
                    lb.Caption = ValeurTravail(CLng(mLignesT(idx)), CStr(cols(i)))
                End If
                lb.ForeColor = IIf(choisie, COUL_BOUTON_TXT, COUL_GRILLE_TXT)
            End If
        Next i
    Next r

    MajBarre f, "sbFTravaux", FNb(mLignesT), FA_Z4_LIGNES, mPremierT
    MajTotaux f

    ' Les deux zones de texte du bas suivent la ligne choisie, et se vident
    ' quand il n'y en a plus. C'est fait ICI, à chaque repeint, et non au clic :
    ' changer de client ou de filtre remet mSelT à zéro sans qu'on ait cliqué,
    ' et les deux champs gardaient alors le texte du record précédent — celui
    ' d'un client qui n'est même plus à l'écran.
    If mSelT >= 1 And mSelT <= FNb(mLignesT) Then
        MontrerTextes f, CLng(mLignesT(mSelT))
    Else
        MontrerTextes f, 0
    End If
End Sub

'------------------------------------------------------------------------------
' Ce qu'affiche une case du tableau du bas.
'
' Trois colonnes ne se lisent pas telles quelles : le CA est CALCULÉ et jamais
' lu, les cases à cocher montrent une croix ou rien, et Select. n'existe pas
' dans TblInterv — son état vit dans mCoches.
'------------------------------------------------------------------------------
Private Function ValeurTravail(ByVal ligne As Long, ByVal colonne As String) As String
    If colonne = FC_SELECT Then
        ValeurTravail = IIf(mCoches.Exists(CStr(ligne)), FCoche(), vbNullString)
    ElseIf colonne = IC_CA Then
        ValeurTravail = Format$(Fact_CADeLaLigne(ligne), "#,##0.00")
    ElseIf FEstCase(colonne) Then
        ValeurTravail = IIf(Fact_EnBooleen(Interv_Valeur(ligne, colonne)), FCoche(), vbNullString)
    Else
        ValeurTravail = Interv_ValeurAffichee(ligne, colonne)
    End If
End Function

'------------------------------------------------------------------------------
' Les deux totaux, sous les colonnes qu'ils additionnent.
'------------------------------------------------------------------------------
Private Sub MajTotaux(f As Object)
    Dim c As Object

    Set c = FCtl(f, "lblFTotHeures")
    If Not c Is Nothing Then _
        c.Caption = Interv_HeuresVersTexte(Fact_Somme(mLignesT, IC_HEURES))

    Set c = FCtl(f, "lblFTotCA")
    If Not c Is Nothing Then _
        c.Caption = Format$(Fact_Somme(mLignesT, IC_CA), "#,##0.00")
End Sub

'------------------------------------------------------------------------------
' Le texte de facture et les commentaires de la ligne choisie.
'------------------------------------------------------------------------------
'   ligne = 0 : aucune ligne choisie, les deux champs se vident.
Private Sub MontrerTextes(f As Object, ByVal ligne As Long)
    Dim c As Object

    Set c = FCtl(f, "txtFTexte")
    If Not c Is Nothing Then _
        c.Text = IIf(ligne = 0, vbNullString, EnTexte(Interv_Valeur(ligne, IC_TEXTE)))

    Set c = FCtl(f, "txtFComm")
    If Not c Is Nothing Then _
        c.Text = IIf(ligne = 0, vbNullString, EnTexte(Interv_Valeur(ligne, IC_COMMENT)))
End Sub

'==============================================================================
' ENREGISTRER
'------------------------------------------------------------------------------
' Attribue le numéro saisi aux lignes cochées, puis recharge tout : les clients
' entièrement facturés disparaissent alors du tableau du haut.
'==============================================================================
Public Sub Fact_Enregistrer_Clic(f As Object)
    Dim c As Object, numero As String, cibles() As Long, n As Long, i As Long, faites As Long

    Set c = FCtl(f, "txtFNum")
    If Not c Is Nothing Then numero = Trim$(EnTexte(c.Text))

    If Len(numero) = 0 Then
        MsgBox "Il manque le numéro de facture à attribuer.", vbExclamation, "Facturation"
        If Not c Is Nothing Then c.SetFocus
        Exit Sub
    End If

    ' les lignes cochées, et elles seules
    If FNb(mLignesT) > 0 Then
        ReDim cibles(1 To FNb(mLignesT))
        For i = 1 To FNb(mLignesT)
            If mCoches.Exists(CStr(CLng(mLignesT(i)))) Then
                n = n + 1
                cibles(n) = CLng(mLignesT(i))
            End If
        Next i
    End If

    If n = 0 Then
        MsgBox "Aucune ligne n'est cochée dans la colonne " & FC_SELECT & "." & vbCrLf & _
               vbCrLf & "Cochez les travaux à porter sur la facture " & numero & ".", _
               vbExclamation, "Facturation"
        Exit Sub
    End If

    ReDim Preserve cibles(1 To n)
    faites = Fact_Enregistrer(numero, cibles)

    mCoches.RemoveAll
    mSelT = 0
    Fact_ToutRecharger f

    MsgBox faites & " intervention(s) portée(s) sur la facture " & numero & ".", _
           vbInformation, "Facturation"
End Sub

'==============================================================================
' PETITS SERVICES
'==============================================================================

' Nombre d'éléments d'un tableau qui peut être vide.
Private Function FNb(ByVal t As Variant) As Long
    On Error Resume Next
    If IsArray(t) Then FNb = UBound(t) - LBound(t) + 1
    On Error GoTo 0
    If FNb < 0 Then FNb = 0
End Function

' Le fond d'une ligne : la ligne choisie d'abord, sinon une ligne sur deux
' teintée — sur douze colonnes, l'oeil suit une ligne bien mieux ainsi.
Private Function FFondLigne(ByVal r As Long, ByVal choisie As Boolean) As Long
    If choisie Then
        FFondLigne = COUL_MODIFIER
    ElseIf r Mod 2 = 0 Then
        FFondLigne = COUL_GRILLE_ZEBRE
    Else
        FFondLigne = COUL_CARTE
    End If
End Function

' La barre suit le nombre de lignes en trop, jamais le nombre total : sa plage
' vaut zéro quand tout tient à l'écran, et le curseur disparaît alors.
Private Sub MajBarre(f As Object, ByVal nom As String, ByVal total As Long, _
                     ByVal visibles As Long, ByVal premiere As Long)
    Dim c As Object, maxi As Long

    Set c = FCtl(f, nom)
    If c Is Nothing Then Exit Sub

    maxi = total - visibles
    If maxi < 0 Then maxi = 0
    On Error Resume Next
    c.Max = maxi
    c.Enabled = (maxi > 0)
    If premiere - 1 > maxi Then premiere = maxi + 1
    c.Value = premiere - 1
    On Error GoTo 0
End Sub

Private Function FCochee(f As Object, ByVal nom As String) As Boolean
    Dim c As Object

    Set c = FCtl(f, nom)
    If Not c Is Nothing Then FCochee = (c.Value = True)
End Function

' Le mois choisi, 1 à 12 ; 0 pour « tous les mois », qui est le premier choix.
Private Function FMoisChoisi(f As Object) As Long
    Dim c As Object

    Set c = FCtl(f, "cboFMois")
    If c Is Nothing Then Exit Function
    If c.ListIndex > 0 Then FMoisChoisi = c.ListIndex
End Function

Private Function FAnneeAffichee() As String
    Dim v As Variant

    v = Interv_CelluleNommee(CEL_ANNEE)
    If IsEmpty(v) Then Exit Function
    On Error Resume Next
    If IsDate(v) Then
        FAnneeAffichee = Format$(v, "yyyy")
    ElseIf IsNumeric(v) Then
        FAnneeAffichee = Format$(CLng(v), "0000")
    Else
        FAnneeAffichee = EnTexte(v)
    End If
    On Error GoTo 0
End Function

' Un contrôle du formulaire, ou Nothing s'il n'existe pas. La collection
' Controls d'un UserForm est plate : elle trouve aussi ce qui est dans un cadre.
Private Function FCtl(f As Object, ByVal nom As String) As Object
    On Error Resume Next
    Set FCtl = f.Controls(nom)
    On Error GoTo 0
End Function
