Attribute VB_Name = "modPression_Formulaire"
Option Explicit
'==============================================================================
' modPression_Formulaire
'------------------------------------------------------------------------------
' Comportement du formulaire UF_Pression. Le module de code du formulaire ne
' contient que des appels d'une ligne vers ici.
'
' LE PRINCIPE. La date et la demi-journée désignent une ligne de Datas, et une
' seule — LigneDe le calcule. Changer l'une ou l'autre RECHARGE la fiche depuis
' cette ligne : ce qui est affiché est donc toujours ce que porte le classeur,
' et enregistrer écrase cette ligne-là. Rien n'est jamais inséré.
'==============================================================================

' Vrai pendant que le programme remplit les contrôles. Les événements Change
' rappelleraient sinon Pression_Cibler au milieu d'un chargement, qui
' rechargerait à son tour : la fiche se viderait toute seule.
Private mEnCours As Boolean

'==============================================================================
' OUVERTURE
'==============================================================================
Public Sub Pression_Initialiser(f As Object)
    Dim activites As Variant, changes As Variant, i As Long

    mEnCours = True

    f.cboPDemi.Clear
    f.cboPDemi.AddItem "matin"
    f.cboPDemi.AddItem "Soir"

    changes = ChangesProposes()
    f.cboPChange.Clear
    For i = LBound(changes) To UBound(changes)
        f.cboPChange.AddItem CStr(changes(i))
    Next i

    activites = Pression_Activites()
    f.cboPActivite.Clear
    If IsArray(activites) Then
        For i = LBound(activites) To UBound(activites)
            f.cboPActivite.AddItem CStr(activites(i))
        Next i
    End If

    f.lblPEtatClasseur.Caption = EtatDuClasseur()

    mEnCours = False

    Pression_Aujourdhui f
    RemplirDernieres f
End Sub

'------------------------------------------------------------------------------
' Ce que dit le bandeau : la dernière mesure, et l'état de la structure.
'------------------------------------------------------------------------------
Private Function EtatDuClasseur() As String
    Dim lignes As Variant, controle As Variant, msg As String

    lignes = Pression_DernieresLignes(1)
    If IsArray(lignes) Then
        If UBound(lignes) >= LBound(lignes) Then
            msg = "Derni" & ChrW(232) & "re mesure : " & _
                  Pression_Affichee(CLng(lignes(LBound(lignes))), PC_DATE) & " " & _
                  NomDemi(DemiDeLaLigne(CLng(lignes(LBound(lignes))))) & "  " & _
                  Pression_Affichee(CLng(lignes(LBound(lignes))), PC_SYS) & "/" & _
                  Pression_Affichee(CLng(lignes(LBound(lignes))), PC_DIA)
        End If
    End If
    If Len(msg) = 0 Then msg = "Aucune mesure dans ce classeur."

    ' Le contrôle de structure est porté par la feuille Reglages : s'il est
    ' tombé, les deux graphiques sont décalés et le dire ici est le seul moyen
    ' que quelqu'un le voie.
    controle = ValeurNommee(CEL_CONTROLE)
    If Not IsEmpty(controle) Then
        If Left$(CStr(controle), 9) <> "Structure" Then
            msg = ChrW(9888) & " " & CStr(controle)
        End If
    End If

    EtatDuClasseur = msg
End Function

'==============================================================================
' POSER UNE VALEUR DANS UN MENU DÉROULANT
'------------------------------------------------------------------------------
' UN MENU FERMÉ REFUSE .Text. Un ComboBox en fmStyleDropDownList — cboPDemi et
' cboPChange — répond « erreur 380, valeur de propriété non valide » dès qu'on
' lui affecte une chaîne, et la chaîne VIDE la déclenche à tous les coups :
' pour MSForms, « rien » ne figure jamais dans la liste, même quand on y a
' ajouté une entrée vide.
'
' Rien ne le signale à la compilation. La faute n'apparaît qu'à l'ouverture du
' formulaire, sur une ligne d'apparence irréprochable — ici la reprise du
' repère Change d'une demi-journée qui n'en porte pas, c'est-à-dire presque
' toutes. verifier_vba.py refuse désormais tout retour en arrière.
'
' ListIndex, lui, accepte -1 pour « rien de choisi ».
'------------------------------------------------------------------------------
Private Sub PoserListe(c As Object, ByVal valeur As String)
    Dim i As Long

    ' Un menu OUVERT accepte n'importe quel texte, y compris vide : c'est tout
    ' l'intérêt de cboPActivite, où une activité nouvelle se tape.
    If c.Style = MSF_StyleDropDownCombo Then
        c.Text = valeur
        Exit Sub
    End If

    If Len(valeur) = 0 Then
        c.ListIndex = -1
        Exit Sub
    End If

    For i = 0 To c.ListCount - 1
        If StrComp(CStr(c.List(i)), valeur, vbTextCompare) = 0 Then
            c.ListIndex = i
            Exit Sub
        End If
    Next i

    ' La valeur vient du classeur et ne figure pas dans la liste proposée : on
    ' l'ajoute plutôt que de la perdre. Un Change à 140 saisi à la main dans la
    ' feuille reste ainsi visible dans le formulaire, et se réenregistre tel
    ' quel.
    c.AddItem valeur
    c.ListIndex = c.ListCount - 1
End Sub

'==============================================================================
' DÉSIGNER UNE DEMI-JOURNÉE
'==============================================================================
'------------------------------------------------------------------------------
' Recalcule la ligne visée d'après la date et la demi-journée, et recharge la
' fiche depuis cette ligne.
'------------------------------------------------------------------------------
Public Sub Pression_Cibler(f As Object)
    Dim jour As Date, ok As Boolean, ligne As Long

    If mEnCours Then Exit Sub

    jour = Pression_Date(f.txtPDate.Text, ok)
    If Not ok Then
        f.txtPLigne.Text = vbNullString
        f.lblPMessage.Caption = "Date non reconnue."
        Exit Sub
    End If

    ligne = LigneDe(jour, DemiChoisie(f))
    If ligne = 0 Then
        f.txtPLigne.Text = vbNullString
        f.lblPMessage.Caption = "Le " & Format$(jour, "dd.mm.yyyy") & _
                                " est hors du calendrier (" & _
                                Format$(PremierJour(), "dd.mm.yyyy") & " " & ChrW(8211) & " " & _
                                Format$(DernierJour(), "dd.mm.yyyy") & ")."
        Exit Sub
    End If

    f.lblPMessage.Caption = vbNullString
    ChargerLigne f, ligne
End Sub

' La LECTURE de .Text reste permise sur un menu fermé — c'est l'écriture qui ne
' l'est pas. Mais ListIndex dit la même chose sans dépendre de l'orthographe
' des deux entrées, et vaut -1 tant que rien n'est choisi : on prend alors le
' matin, qui est le premier de la journée.
Private Function DemiChoisie(f As Object) As Long
    DemiChoisie = IIf(f.cboPDemi.ListIndex = 1, DEMI_SOIR, DEMI_MATIN)
End Function

'------------------------------------------------------------------------------
' Remplit la fiche depuis une ligne de Datas.
'------------------------------------------------------------------------------
Private Sub ChargerLigne(f As Object, ByVal ligne As Long)
    mEnCours = True

    f.txtPLigne.Text = CStr(ligne) & "  " & ChrW(8212) & "  " & _
                       Format$(JourDeLaLigne(ligne), "ddd dd.mm.yy") & " " & _
                       NomDemi(DemiDeLaLigne(ligne))
    f.txtPHeure.Text = Pression_Affichee(ligne, PC_HEURE)
    PoserListe f.cboPActivite, Pression_Affichee(ligne, PC_ACTIVITE)
    f.txtPSys.Text = Pression_Affichee(ligne, PC_SYS)
    f.txtPDia.Text = Pression_Affichee(ligne, PC_DIA)
    f.txtPPouls.Text = Pression_Affichee(ligne, PC_POULS)
    f.txtPPoids.Text = Pression_Affichee(ligne, PC_POIDS)
    PoserListe f.cboPChange, Pression_Affichee(ligne, PC_CHANGE)
    f.txtPComment.Text = Pression_Affichee(ligne, PC_COMMENT)

    mEnCours = False
    Pression_Juger f
End Sub

'------------------------------------------------------------------------------
' Pose la date, la demi-journée et l'heure de maintenant.
' Avant midi, c'est le matin ; après, le soir.
'------------------------------------------------------------------------------
Public Sub Pression_Aujourdhui(f As Object)
    mEnCours = True
    f.txtPDate.Text = Format$(Date, "dd.mm.yy")
    PoserListe f.cboPDemi, IIf(Hour(Now) < 12, "matin", "Soir")
    mEnCours = False

    Pression_Cibler f

    mEnCours = True
    If Len(Trim$(f.txtPHeure.Text)) = 0 Then f.txtPHeure.Text = Format$(Now, "hh:nn")
    mEnCours = False
End Sub

'==============================================================================
' L'ÉTAT DE LA MESURE SAISIE
'==============================================================================
'------------------------------------------------------------------------------
' Compare ce qui est tapé aux deux seuils hauts du classeur, et le dit.
'
' CE N'EST PAS UN DIAGNOSTIC. Les seuils sont ceux des lignes de repère des
' graphiques, réglés sur la feuille Saisie : le formulaire ne fait que
' rapprocher deux nombres de deux autres.
'------------------------------------------------------------------------------
Public Sub Pression_Juger(f As Object)
    Dim sys As Long, dia As Long, ok As Boolean, okDia As Boolean
    Dim seuilSys As Variant, seuilDia As Variant

    If mEnCours Then Exit Sub

    sys = Pression_Entier(f.txtPSys.Text, SYS_MIN, SYS_MAX, ok)
    dia = Pression_Entier(f.txtPDia.Text, DIA_MIN, DIA_MAX, okDia)

    If Not ok Or Not okDia Then
        f.lblPEtat.ForeColor = PCOUL_ALERTE
        f.lblPEtat.Caption = "Valeur hors des bornes de saisie."
        Exit Sub
    End If
    If sys = 0 Or dia = 0 Then
        f.lblPEtat.Caption = vbNullString
        Exit Sub
    End If

    seuilSys = ValeurNommee(CEL_SEUIL_SYS)
    seuilDia = ValeurNommee(CEL_SEUIL_DIA)
    If Not IsNumeric(seuilSys) Then seuilSys = 130
    If Not IsNumeric(seuilDia) Then seuilDia = 90

    If sys > CDbl(seuilSys) Or dia > CDbl(seuilDia) Then
        f.lblPEtat.ForeColor = PCOUL_ALERTE
        f.lblPEtat.Caption = sys & "/" & dia & "  " & ChrW(8212) & "  au-dessus du rep" & _
                             ChrW(232) & "re " & CLng(seuilSys) & "/" & CLng(seuilDia)
    Else
        f.lblPEtat.ForeColor = PCOUL_BON
        f.lblPEtat.Caption = sys & "/" & dia & "  " & ChrW(8212) & "  sous le rep" & _
                             ChrW(232) & "re " & CLng(seuilSys) & "/" & CLng(seuilDia)
    End If
End Sub

'==============================================================================
' ENREGISTRER
'==============================================================================
Public Sub Pression_Enregistrer(f As Object)
    Dim jour As Date, ligne As Long, ok As Boolean
    Dim sys As Long, dia As Long, pouls As Long
    Dim poids As Double, heure As Double
    Dim valeurs As Object, faute As String

    jour = Pression_Date(f.txtPDate.Text, ok)
    If Not ok Then
        Avertir f, "La date n'est pas reconnue."
        Exit Sub
    End If
    ligne = LigneDe(jour, DemiChoisie(f))
    If ligne = 0 Then
        Avertir f, "Cette date est hors du calendrier de la feuille " & FEUILLE_DATAS & "."
        Exit Sub
    End If

    ' Chaque contrôle rend sa faute plutôt que de s'arrêter au premier : celui
    ' qui a tapé trois valeurs de travers les voit toutes d'un coup.
    sys = Pression_Entier(f.txtPSys.Text, SYS_MIN, SYS_MAX, ok)
    If Not ok Then faute = Ajouter(faute, "Systole : un entier entre " & SYS_MIN & " et " & SYS_MAX)
    dia = Pression_Entier(f.txtPDia.Text, DIA_MIN, DIA_MAX, ok)
    If Not ok Then faute = Ajouter(faute, "Diastole : un entier entre " & DIA_MIN & " et " & DIA_MAX)
    pouls = Pression_Entier(f.txtPPouls.Text, POULS_MIN, POULS_MAX, ok)
    If Not ok Then faute = Ajouter(faute, "Pouls : un entier entre " & POULS_MIN & " et " & POULS_MAX)
    poids = Pression_Nombre(f.txtPPoids.Text, POIDS_MIN, POIDS_MAX, ok)
    If Not ok Then faute = Ajouter(faute, "Poids : un nombre entre " & POIDS_MIN & " et " & POIDS_MAX)
    heure = Pression_Heure(f.txtPHeure.Text, ok)
    If Not ok Then faute = Ajouter(faute, "Heure : 7:42, ou 742")

    If Len(faute) > 0 Then
        MsgBox "La mesure n'a pas été enregistrée :" & vbCrLf & vbCrLf & faute, _
               vbExclamation, "Saisie d'une mesure"
        Exit Sub
    End If

    ' Une systole sans diastole, ou l'inverse, est presque toujours une frappe
    ' interrompue : mieux vaut le demander que d'écrire une demi-mesure.
    If (sys = 0) <> (dia = 0) Then
        If MsgBox("Une seule des deux valeurs est renseignée." & vbCrLf & vbCrLf & _
                  "Enregistrer quand même ?", vbQuestion + vbYesNo, _
                  "Saisie d'une mesure") <> vbYes Then Exit Sub
    End If

    Set valeurs = CreateObject("Scripting.Dictionary")
    valeurs.Add PC_HEURE, IIf(heure = 0, Empty, heure)
    valeurs.Add PC_ACTIVITE, Trim$(f.cboPActivite.Text)
    valeurs.Add PC_SYS, IIf(sys = 0, Empty, sys)
    valeurs.Add PC_DIA, IIf(dia = 0, Empty, dia)
    valeurs.Add PC_POULS, IIf(pouls = 0, Empty, pouls)
    valeurs.Add PC_POIDS, IIf(poids = 0, Empty, poids)
    valeurs.Add PC_CHANGE, ValeurChange(f)
    valeurs.Add PC_COMMENT, Trim$(f.txtPComment.Text)

    On Error GoTo Erreur
    Pression_Ecrire ligne, valeurs

    f.lblPEtatClasseur.Caption = EtatDuClasseur()
    f.lblPMessage.Caption = "Enregistr" & ChrW(233) & " " & ChrW(8212) & " ligne " & ligne
    RemplirDernieres f

    ' Le formulaire n'est pas modal : l'auteur peut vouloir enchaîner. La fiche
    ' reste donc telle quelle, et c'est le message qui dit que c'est fait.
    Exit Sub

Erreur:
    MsgBox "L'enregistrement a échoué :" & vbCrLf & vbCrLf & _
           Err.Number & " - " & Err.Description, vbCritical, "Saisie d'une mesure"
End Sub

Private Function ValeurChange(f As Object) As Variant
    Dim texte As String
    texte = Trim$(f.cboPChange.Text)
    If Len(texte) = 0 Then
        ValeurChange = Empty
    ElseIf IsNumeric(texte) Then
        ValeurChange = CLng(texte)
    Else
        ValeurChange = texte
    End If
End Function

Private Function Ajouter(ByVal liste As String, ByVal ligne As String) As String
    Ajouter = liste & IIf(Len(liste) > 0, vbCrLf, "") & ChrW(8226) & " " & ligne
End Function

Private Sub Avertir(f As Object, ByVal message As String)
    f.lblPMessage.Caption = message
    MsgBox message, vbExclamation, "Saisie d'une mesure"
End Sub

'==============================================================================
' EFFACER, QUITTER
'==============================================================================
'------------------------------------------------------------------------------
' Vide les zones de saisie. NE TOUCHE PAS AU CLASSEUR : c'est la fiche qui se
' vide, pas la ligne. Pour effacer une mesure du classeur, il faut vider les
' champs puis Enregistrer.
'------------------------------------------------------------------------------
Public Sub Pression_Effacer(f As Object)
    mEnCours = True
    f.txtPHeure.Text = vbNullString
    PoserListe f.cboPActivite, vbNullString
    f.txtPSys.Text = vbNullString
    f.txtPDia.Text = vbNullString
    f.txtPPouls.Text = vbNullString
    f.txtPPoids.Text = vbNullString
    PoserListe f.cboPChange, vbNullString
    f.txtPComment.Text = vbNullString
    f.lblPEtat.Caption = vbNullString
    mEnCours = False
    f.lblPMessage.Caption = "Fiche vid" & ChrW(233) & "e " & ChrW(8212) & _
                            " le classeur n'a pas chang" & ChrW(233) & "."
End Sub

Public Sub Pression_Quitter(f As Object)
    Unload f
End Sub

'==============================================================================
' LE TABLEAU DES DERNIÈRES MESURES
'==============================================================================
Private Sub RemplirDernieres(f As Object)
    Dim lignes As Variant, cols As Variant, r As Long, i As Long, ligne As Long

    lignes = Pression_DernieresLignes(P_NB_DERNIERES)
    cols = PDernieresColonnes()

    For r = 1 To P_NB_DERNIERES
        ligne = 0
        If IsArray(lignes) Then
            If r + LBound(lignes) - 1 <= UBound(lignes) Then
                ligne = CLng(lignes(r + LBound(lignes) - 1))
            End If
        End If
        f.Controls("lblPL_" & CStr(r)).Tag = CStr(ligne)
        For i = LBound(cols) To UBound(cols)
            f.Controls("lblP_" & CStr(r) & "_" & CStr(i + 1)).Caption = _
                IIf(ligne = 0, vbNullString, Pression_Affichee(ligne, CLng(cols(i))))
        Next i
    Next r
End Sub

'------------------------------------------------------------------------------
' Un clic sur une ligne du tableau ramène cette demi-journée dans la fiche.
' Le numéro de ligne est rangé dans le Tag de la bande de fond : la case
' cliquée ne sait dire que son rang à l'écran.
'------------------------------------------------------------------------------
Public Sub Pression_Reprendre(f As Object, ByVal rang As Long)
    Dim ligne As Long, tag As String

    tag = f.Controls("lblPL_" & CStr(rang)).Tag
    If Len(tag) = 0 Then Exit Sub
    ligne = CLng(tag)
    If ligne = 0 Then Exit Sub

    mEnCours = True
    f.txtPDate.Text = Format$(JourDeLaLigne(ligne), "dd.mm.yy")
    PoserListe f.cboPDemi, NomDemi(DemiDeLaLigne(ligne))
    mEnCours = False

    ChargerLigne f, ligne
    f.lblPMessage.Caption = "Reprise de la ligne " & ligne & "."
End Sub
