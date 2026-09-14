Attribute VB_Name = "modPression_Generateur"
Option Explicit
'==============================================================================
' modPression_Generateur
'------------------------------------------------------------------------------
' Génère de toutes pièces le formulaire de saisie UF_Pression.
'
'   >>> Procédure à lancer : GenererFormulairePression
'
' PRÉREQUIS : Fichier > Options > Centre de gestion de la confidentialité >
' Paramètres du Centre de gestion de la confidentialité > Paramètres des
' macros > « Accès approuvé au modèle d'objet du projet VBA ». Puis fermer et
' rouvrir le classeur.
'
' IL NE FAUT PAS SUPPRIMER LE FORMULAIRE avant de régénérer : la génération le
' vide et le reconstruit sur place. VBA ne rend le nom d'un composant supprimé
' qu'au prochain chargement du classeur, et une suppression à la main oblige
' donc à fermer le classeur avant de pouvoir regénérer.
'
' Le module de code du formulaire ne contient QUE des procédures
' événementielles d'une ligne, écrites ici : toute la logique reste dans
' modPression_Formulaire. Conséquence pratique — le projet compile avant même
' que le formulaire existe, et régénérer n'écrase jamais de comportement.
'==============================================================================

Private Const CT_MSFORM As Long = 3              ' vbext_ct_MSForm
Private Const ERR_NOM_OCCUPE As Long = vbObjectError + 713
Private Const ERR_FORM_NON_VIDE As Long = vbObjectError + 714

Private mCode As String

'==============================================================================
' POINT D'ENTRÉE
'==============================================================================
Public Sub GenererFormulairePression()
    Dim vbProj As Object, nbCtrl As Long

    If FeuilleDatas() Is Nothing Then
        MsgBox "La feuille " & FEUILLE_DATAS & " est introuvable dans ce classeur." & _
               vbCrLf & "Le formulaire de saisie ne peut pas être généré.", _
               vbCritical, "Génération du formulaire"
        Exit Sub
    End If

    On Error Resume Next
    Set vbProj = ThisWorkbook.VBProject
    On Error GoTo 0
    If vbProj Is Nothing Then
        MsgBox "Excel refuse l'accès au projet VBA." & vbCrLf & vbCrLf & _
               "Activez l'option " & Chr$(34) & "Accès approuvé au modèle d'objet du " & _
               "projet VBA" & Chr$(34) & ", fermez puis rouvrez le classeur, et " & _
               "relancez cette procédure.", vbCritical, "Génération du formulaire"
        Exit Sub
    End If

    On Error GoTo Erreur
    nbCtrl = ConstruirePression(vbProj)

    MsgBox "Le formulaire " & NOM_FORM_PRESSION & " a été généré." & vbCrLf & vbCrLf & _
           nbCtrl & " contrôles." & vbCrLf & vbCrLf & _
           "Il s'ouvre par OuvrirSaisiePression, ou par le bouton de la feuille " & _
           FEUILLE_SAISIE & ".", vbInformation, "Génération du formulaire"
    Exit Sub

Erreur:
    MsgBox "La génération a échoué :" & vbCrLf & vbCrLf & _
           Err.Number & " - " & Err.Description, vbCritical, "Génération du formulaire"
End Sub

'==============================================================================
' LE FORMULAIRE
'==============================================================================
Private Function ConstruirePression(vbProj As Object) As Long
    Dim vbComp As Object, dsg As Object

    Set vbComp = PreparerForm(vbProj, NOM_FORM_PRESSION)
    Set dsg = vbComp.Designer

    Prop vbComp, "Caption", "Pression Rena " & ChrW(8212) & " saisie d'une mesure"
    Prop vbComp, "Width", P_LARGEUR
    Prop vbComp, "Height", P_HAUTEUR + P_RESERVE_TITRE
    Prop vbComp, "BackColor", PCOUL_FOND
    Prop vbComp, "SpecialEffect", MSF_SpecialEffectFlat
    Prop vbComp, "StartUpPosition", 1
    Prop vbComp, "ShowModal", False
    PoserPolice dsg, PPOLICE, 8, False

    ConstruireBandeau dsg
    ConstruireFiche dsg
    ConstruireDernieres dsg
    ConstruireBoutons dsg

    vbComp.CodeModule.AddFromString CodePression()
    ConstruirePression = dsg.Controls.Count
End Function

'------------------------------------------------------------------------------
' Le bandeau : le titre à gauche, l'état du classeur à droite.
'------------------------------------------------------------------------------
Private Sub ConstruireBandeau(dsg As Object)
    Dim c As Object, cadre As Object

    Set cadre = Aj(dsg, "Forms.Frame.1", "fraPBandeau", _
                   P_MARGE, P_Z1_TOP, P_CARTE_LARG, P_Z1_HAUT)
    Carte cadre
    cadre.BackColor = PCOUL_BANDEAU
    cadre.BorderColor = PCOUL_BANDEAU

    Set c = Aj(cadre, "Forms.Label.1", "lblPTitre", 14, 8, 340, 20)
    Texte c, "Pression art" & ChrW(233) & "rielle", PSTitre(), MSF_TextAlignLeft

    Set c = Aj(cadre, "Forms.Label.1", "lblPEtatClasseur", P_CARTE_LARG - 364, 12, 350, 24)
    Texte c, vbNullString, PSSousTitre(), MSF_TextAlignRight
    c.WordWrap = False
End Sub

'------------------------------------------------------------------------------
' La fiche de saisie : deux blocs de quatre lignes.
'
' Bloc 1 — QUAND : la date et la demi-journée désignent à elles deux la ligne
' de Datas à remplir, par le calcul de LigneDe. Les changer recharge la fiche.
' Bloc 2 — LA MESURE.
'------------------------------------------------------------------------------
Private Sub ConstruireFiche(dsg As Object)
    Dim c As Object, cadre As Object, x As Single, y As Single, larg As Single

    Set cadre = Aj(dsg, "Forms.Frame.1", "fraPFiche", _
                   P_MARGE, P_Z2_TOP, P_CARTE_LARG, P_Z2_HAUT)
    Carte cadre

    Set c = Aj(cadre, "Forms.Label.1", "lblPSection1", P_GR_X, 5, 240, 12)
    Texte c, UCase$("Quand"), PSSection(), MSF_TextAlignLeft
    Set c = Aj(cadre, "Forms.Label.1", "lblPSection2", PGrilleX(2), 5, 240, 12)
    Texte c, UCase$("La mesure"), PSSection(), MSF_TextAlignLeft

    '--- bloc 1 ---------------------------------------------------------------
    x = PGrilleX(1)

    y = PGrilleY(1)
    Set c = Aj(cadre, "Forms.Label.1", "lblPDate", x, y, P_GR_BLOC, P_LBL_HAUT)
    Texte c, "Date", PSLibelle(), MSF_TextAlignLeft
    Set c = Aj(cadre, "Forms.TextBox.1", "txtPDate", x, y + P_LBL_HAUT, 96, P_CTL_HAUT)
    Zone c, False
    c.ControlTipText = "Jour de la mesure, par exemple 14.09.26"
    Set c = Aj(cadre, "Forms.CommandButton.1", "btnPAujourdhui", _
               x + 104, y + P_LBL_HAUT, P_GR_BLOC - 104, P_CTL_HAUT)
    Bouton c, "Aujourd'hui", "J", PCOUL_AUJOURD, 2
    c.ControlTipText = "Date, demi-journée et heure de maintenant"

    y = PGrilleY(2)
    Set c = Aj(cadre, "Forms.Label.1", "lblPDemi", x, y, P_GR_BLOC, P_LBL_HAUT)
    Texte c, "Demi-journ" & ChrW(233) & "e", PSLibelle(), MSF_TextAlignLeft
    Set c = Aj(cadre, "Forms.ComboBox.1", "cboPDemi", x, y + P_LBL_HAUT, 130, P_CTL_HAUT)
    Liste c, True
    c.ControlTipText = "C'est le RANG de la ligne dans sa journée, pas l'heure : " & _
                       "deux mesures d'un même matin tiennent dans les deux."
    Set c = Aj(cadre, "Forms.Label.1", "lblPHeureCap", x + 138, y, 130, P_LBL_HAUT)
    Texte c, "Heure", PSLibelle(), MSF_TextAlignLeft
    Set c = Aj(cadre, "Forms.TextBox.1", "txtPHeure", x + 138, y + P_LBL_HAUT, 130, P_CTL_HAUT)
    Zone c, False
    c.ControlTipText = "7:42, ou 742"

    y = PGrilleY(3)
    Set c = Aj(cadre, "Forms.Label.1", "lblPActiviteCap", x, y, P_GR_BLOC, P_LBL_HAUT)
    Texte c, "Activit" & ChrW(233), PSLibelle(), MSF_TextAlignLeft
    Set c = Aj(cadre, "Forms.ComboBox.1", "cboPActivite", x, y + P_LBL_HAUT, _
               P_GR_BLOC, P_CTL_HAUT)
    Liste c, False
    c.ControlTipText = "Les activités déjà employées sont proposées ; une nouvelle " & _
                       "peut être tapée."

    y = PGrilleY(4)
    Set c = Aj(cadre, "Forms.Label.1", "lblPLigneCap", x, y, P_GR_BLOC, P_LBL_HAUT)
    Texte c, "Ligne vis" & ChrW(233) & "e dans " & FEUILLE_DATAS, PSLibelle(), MSF_TextAlignLeft
    Set c = Aj(cadre, "Forms.TextBox.1", "txtPLigne", x, y + P_LBL_HAUT, P_GR_BLOC, P_CTL_HAUT)
    Zone c, True
    c.ControlTipText = "Calculée à partir de la date : aucune ligne n'est jamais " & _
                       "insérée ni supprimée."

    '--- bloc 2 ---------------------------------------------------------------
    x = PGrilleX(2)
    larg = (P_GR_BLOC - 2 * P_ENTRE_CHAMPS) / 3

    y = PGrilleY(1)
    Set c = Aj(cadre, "Forms.Label.1", "lblPSysCap", x, y, larg, P_LBL_HAUT)
    Texte c, "Systole", PSLibelle(), MSF_TextAlignLeft
    Set c = Aj(cadre, "Forms.TextBox.1", "txtPSys", x, y + P_LBL_HAUT, larg, P_CTL_HAUT)
    Zone c, False
    Set c = Aj(cadre, "Forms.Label.1", "lblPDiaCap", x + larg + P_ENTRE_CHAMPS, y, _
               larg, P_LBL_HAUT)
    Texte c, "Diastole", PSLibelle(), MSF_TextAlignLeft
    Set c = Aj(cadre, "Forms.TextBox.1", "txtPDia", x + larg + P_ENTRE_CHAMPS, _
               y + P_LBL_HAUT, larg, P_CTL_HAUT)
    Zone c, False
    Set c = Aj(cadre, "Forms.Label.1", "lblPPoulsCap", x + 2 * (larg + P_ENTRE_CHAMPS), y, _
               larg, P_LBL_HAUT)
    Texte c, "Pouls", PSLibelle(), MSF_TextAlignLeft
    Set c = Aj(cadre, "Forms.TextBox.1", "txtPPouls", x + 2 * (larg + P_ENTRE_CHAMPS), _
               y + P_LBL_HAUT, larg, P_CTL_HAUT)
    Zone c, False

    y = PGrilleY(2)
    larg = (P_GR_BLOC - P_ENTRE_CHAMPS) / 2
    Set c = Aj(cadre, "Forms.Label.1", "lblPPoidsCap", x, y, larg, P_LBL_HAUT)
    Texte c, "Poids (kg)", PSLibelle(), MSF_TextAlignLeft
    Set c = Aj(cadre, "Forms.TextBox.1", "txtPPoids", x, y + P_LBL_HAUT, larg, P_CTL_HAUT)
    Zone c, False
    c.ControlTipText = "Relevé le matin. La colonne n'existait pas dans l'ancien fichier."
    Set c = Aj(cadre, "Forms.Label.1", "lblPChangeCap", x + larg + P_ENTRE_CHAMPS, y, _
               larg, P_LBL_HAUT)
    Texte c, "Change", PSLibelle(), MSF_TextAlignLeft
    Set c = Aj(cadre, "Forms.ComboBox.1", "cboPChange", x + larg + P_ENTRE_CHAMPS, _
               y + P_LBL_HAUT, larg, P_CTL_HAUT)
    Liste c, True
    c.ControlTipText = "Repère de changement de situation, porté à 150 ou 160 pour " & _
                       "être visible sur le graphique."

    y = PGrilleY(3)
    Set c = Aj(cadre, "Forms.Label.1", "lblPCommentCap", x, y, P_GR_BLOC, P_LBL_HAUT)
    Texte c, "Commentaire", PSLibelle(), MSF_TextAlignLeft
    Set c = Aj(cadre, "Forms.TextBox.1", "txtPComment", x, y + P_LBL_HAUT, _
               P_GR_BLOC, P_CTL_HAUT)
    Zone c, False

    y = PGrilleY(4)
    Set c = Aj(cadre, "Forms.Label.1", "lblPEtat", x, y + 2, P_GR_BLOC, P_LBL_HAUT + P_CTL_HAUT)
    Texte c, vbNullString, PSEtat(), MSF_TextAlignLeft
    c.WordWrap = True
End Sub

'------------------------------------------------------------------------------
' Le tableau des dernières mesures.
'
' Une grille de libellés, et non une ListBox : c'est ce qui permet de choisir
' librement les largeurs de colonnes, d'aligner les nombres à droite et le
' texte à gauche, et de colorer une ligne. Chaque ligne reçoit d'abord une
' BANDE de fond sur toute la largeur, puis ses cases, transparentes, posées
' dessus — sans quoi les points entre deux cases laisseraient voir le blanc de
' la carte et la ligne choisie paraîtrait rayée.
'------------------------------------------------------------------------------
Private Sub ConstruireDernieres(dsg As Object)
    Dim c As Object, cadre As Object, larg As Variant, lib As Variant, ali As Variant
    Dim i As Long, r As Long, nbCol As Long, x As Single, y As Single, largeur As Single

    Set cadre = Aj(dsg, "Forms.Frame.1", "fraPDernieres", _
                   P_MARGE, P_Z3_TOP, P_CARTE_LARG, P_Z3_HAUT)
    Carte cadre

    Set c = Aj(cadre, "Forms.Label.1", "lblPSection3", AuPixel(P_PAD_X + 2), AuPixel(3), _
               300, 12)
    Texte c, UCase$("Les derni" & ChrW(232) & "res mesures"), PSSection(), MSF_TextAlignLeft

    larg = PDernieresLargeurs()
    lib = PDernieresLibelles()
    ali = PDernieresAlignements()
    nbCol = UBound(larg) - LBound(larg) + 1
    largeur = P_CARTE_LARG - 2

    Set c = Aj(cadre, "Forms.Label.1", "lblPEntFond", AuPixel(1), PGrilleEnteteY(), _
               largeur, AuPixel(P_ENTETE_HAUT))
    Fond c, PCOUL_ENTETE_TBL, PCOUL_ENTETE_TBL

    x = AuPixel(1) + P_PAD_X
    For i = 0 To nbCol - 1
        Set c = Aj(cadre, "Forms.Label.1", "lblPEnt_" & CStr(i + 1), _
                   AuPixel(x), AuPixel(PGrilleEnteteY() + 4), _
                   AuPixel(CSng(larg(i)) - 2 * P_PAD_X), AuPixel(P_LIGNE_H))
        Texte c, CStr(lib(i)), PSEntete(), CLng(ali(i))
        x = x + CSng(larg(i))
    Next i

    For r = 1 To P_NB_DERNIERES
        y = PGrilleLignesY() + (r - 1) * P_LIGNE_H

        Set c = Aj(cadre, "Forms.Label.1", "lblPL_" & CStr(r), _
                   AuPixel(1), AuPixel(y), AuPixel(largeur), AuPixel(P_LIGNE_H))
        With c
            .Caption = vbNullString
            .SpecialEffect = MSF_SpecialEffectFlat
            .BorderStyle = MSF_BorderStyleNone
            .BackStyle = MSF_BackStyleOpaque
            .BackColor = PCOUL_CARTE
        End With
        PoserStyle c, PSCase()

        x = AuPixel(1) + P_PAD_X
        For i = 0 To nbCol - 1
            Set c = Aj(cadre, "Forms.Label.1", _
                       "lblP_" & CStr(r) & "_" & CStr(i + 1), _
                       AuPixel(x), AuPixel(y), _
                       AuPixel(CSng(larg(i)) - 2 * P_PAD_X), AuPixel(P_LIGNE_H))
            Texte c, vbNullString, PSCase(), CLng(ali(i))
            x = x + CSng(larg(i))
        Next i
    Next r
End Sub

'------------------------------------------------------------------------------
' La rangée de boutons, sur le formulaire lui-même.
'------------------------------------------------------------------------------
Private Sub ConstruireBoutons(dsg As Object)
    Dim c As Object, x As Single

    Set c = Aj(dsg, "Forms.Label.1", "lblPMessage", P_MARGE, P_BT_TOP + 8, 240, 14)
    Texte c, vbNullString, PSLibelle(), MSF_TextAlignLeft

    x = P_LARGEUR - P_MARGE - P_BT_LARG
    Set c = Aj(dsg, "Forms.CommandButton.1", "btnPQuitter", x, P_BT_TOP, P_BT_LARG, P_BT_HAUT)
    Bouton c, "Quitter", "Q", PCOUL_QUITTER, 92

    x = x - P_BT_LARG - P_BT_GOUTTIERE
    Set c = Aj(dsg, "Forms.CommandButton.1", "btnPEffacer", x, P_BT_TOP, P_BT_LARG, P_BT_HAUT)
    Bouton c, "Effacer", "E", PCOUL_EFFACER, 91

    x = x - P_BT_LARG - P_BT_GOUTTIERE
    Set c = Aj(dsg, "Forms.CommandButton.1", "btnPEnregistrer", x, P_BT_TOP, P_BT_LARG, P_BT_HAUT)
    Bouton c, "Enregistrer", "N", PCOUL_ENREG, 90
End Sub

'==============================================================================
' PRÉPARATION DU COMPOSANT
'==============================================================================
'------------------------------------------------------------------------------
' Prépare un composant UserForm : créé s'il manque, vidé s'il existe déjà.
'------------------------------------------------------------------------------
Private Function PreparerForm(vbProj As Object, ByVal nom As String) As Object
    Dim vbComp As Object, dsg As Object, i As Long, occupe As Boolean

    On Error Resume Next
    Set vbComp = vbProj.VBComponents(nom)
    On Error GoTo 0

    If vbComp Is Nothing Then
        Set vbComp = vbProj.VBComponents.Add(CT_MSFORM)

        ' LE NOM PEUT ÊTRE ENCORE PRIS. VBA ne rend le nom d'un composant
        ' supprimé qu'au prochain chargement du classeur : un formulaire effacé
        ' à la main dans l'éditeur retient donc le sien jusque-là, et
        ' l'affectation ci-dessous échoue — ce qui produisait une erreur 75
        ' incompréhensible. Le formulaire vide qu'on vient d'ajouter est alors
        ' retiré, sans quoi chaque tentative laisserait un UserForm1,
        ' UserForm2… derrière elle. Le nom est relu après coup : selon les
        ' versions d'Excel, l'affectation échoue tantôt bruyamment, tantôt en
        ' silence.
        On Error Resume Next
        vbComp.Name = nom
        occupe = (Err.Number <> 0)
        Err.Clear
        If Not occupe Then occupe = (StrComp(vbComp.Name, nom, vbTextCompare) <> 0)
        If occupe Then vbProj.VBComponents.Remove vbComp
        Err.Clear
        On Error GoTo 0

        If occupe Then
            Err.Raise ERR_NOM_OCCUPE, "PreparerForm", _
                "Le nom " & nom & " n'est pas encore libre." & vbCrLf & vbCrLf & _
                "Un formulaire portant ce nom a été supprimé pendant cette " & _
                "session. VBA ne rend son nom au projet qu'au prochain " & _
                "chargement du classeur." & vbCrLf & vbCrLf & _
                "À FAIRE : enregistrez, fermez puis rouvrez le classeur, et " & _
                "relancez la génération." & vbCrLf & vbCrLf & _
                "Il n'est jamais utile de supprimer le formulaire au " & _
                "préalable : la génération le vide et le reconstruit sur place."
        End If

        Set PreparerForm = vbComp
        Exit Function
    End If

    ' une instance restée chargée empêcherait la modification
    On Error Resume Next
    For i = UserForms.Count - 1 To 0 Step -1
        If StrComp(UserForms(i).Name, nom, vbTextCompare) = 0 Then Unload UserForms(i)
    Next i
    On Error GoTo 0

    Set dsg = vbComp.Designer
    ViderDesigner dsg, nom

    With vbComp.CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
    End With

    Set PreparerForm = vbComp
End Function

'------------------------------------------------------------------------------
' VIDE un formulaire existant de tous ses contrôles.
'
' CE QUE LE CONCEPTEUR ACCEPTE, ET CE QU'IL REFUSE. La collection Controls du
' formulaire est PLATE : elle montre aussi les contrôles posés dans un cadre.
' Mais elle ne les RETIRE pas — un enfant ne se supprime que depuis la
' collection de son cadre — et un cadre qui a encore des enfants ne part pas
' non plus. Un vidage qui ne s'adresse qu'au formulaire ne retire donc RIEN
' d'un formulaire à cadres.
'------------------------------------------------------------------------------
Private Sub ViderDesigner(dsg As Object, ByVal nom As String)
    RetirerTout dsg
    If dsg.Controls.Count = 0 Then Exit Sub

    Err.Raise ERR_FORM_NON_VIDE, "ViderDesigner", _
        "Le formulaire " & nom & " n'a pas pu être vidé : " & dsg.Controls.Count & _
        " contrôle(s) refusent d'être supprimés." & vbCrLf & vbCrLf & _
        "La génération s'arrête ici : la poursuivre sur un formulaire à moitié " & _
        "plein produirait, bien plus loin, une erreur sans rapport avec la cause." & _
        vbCrLf & vbCrLf & "À FAIRE : enregistrez, fermez puis rouvrez le classeur, " & _
        "et relancez la génération."
End Sub

'------------------------------------------------------------------------------
' Retire d'un conteneur — formulaire ou cadre — tout ce qu'on peut en retirer.
'
' Un CADRE est vidé avant d'être enlevé, par un appel sur lui-même : c'est la
' seule collection qui accepte de lâcher ses enfants. On parcourt à l'envers,
' les index qui précèdent ne bougeant pas, et on recommence tant qu'un tour a
' retiré quelque chose.
'------------------------------------------------------------------------------
Private Sub RetirerTout(cont As Object)
    Dim i As Long, c As Object, nomCtrl As String, avant As Long, nb As Long

    Do
        avant = cont.Controls.Count
        If avant = 0 Then Exit Sub

        For i = avant - 1 To 0 Step -1
            If i <= cont.Controls.Count - 1 Then
                Set c = Nothing
                nomCtrl = vbNullString
                On Error Resume Next
                Set c = cont.Controls(i)
                nomCtrl = c.Name
                On Error GoTo 0

                If Not c Is Nothing Then
                    If EstConteneur(c) Then RetirerTout c
                    nb = cont.Controls.Count
                    On Error Resume Next
                    cont.Controls.Remove nomCtrl
                    If cont.Controls.Count = nb Then cont.Controls.Remove i
                    On Error GoTo 0
                End If
            End If
        Next i

        If cont.Controls.Count = avant Then Exit Sub
    Loop
End Sub

Private Function EstConteneur(c As Object) As Boolean
    Dim n As Long
    On Error Resume Next
    n = c.Controls.Count
    EstConteneur = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

'==============================================================================
' HABILLAGE
'==============================================================================

' Le premier argument est un CONTENEUR : un cadre expose la même collection
' Controls, et les coordonnées comptent alors depuis son coin.
Private Function Aj(conteneur As Object, ByVal progId As String, ByVal nom As String, _
                    ByVal gauche As Single, ByVal haut As Single, _
                    ByVal largeur As Single, ByVal hauteur As Single) As Object
    Dim c As Object

    Set c = conteneur.Controls.Add(progId, nom, True)
    c.Left = gauche
    c.Top = haut
    c.Width = largeur
    c.Height = hauteur
    Set Aj = c
End Function

' SpecialEffect avant BorderStyle : MSForms refuse une bordure simple tant que
' le contrôle est en relief, et un cadre est gravé par défaut.
Private Sub Carte(c As Object)
    c.Caption = vbNullString
    c.BackColor = PCOUL_CARTE
    c.SpecialEffect = MSF_SpecialEffectFlat
    c.BorderStyle = MSF_BorderStyleSingle
    c.BorderColor = PCOUL_BORDURE
    c.ScrollBars = MSF_ScrollBarsNone
End Sub

Private Sub Fond(c As Object, ByVal fondCoul As Long, ByVal bordure As Long)
    c.Caption = vbNullString
    c.BackStyle = MSF_BackStyleOpaque
    c.BackColor = fondCoul
    c.SpecialEffect = MSF_SpecialEffectFlat
    c.BorderStyle = MSF_BorderStyleSingle
    c.BorderColor = bordure
End Sub

Private Sub Texte(c As Object, ByVal contenu As String, ByRef st As StyleTexte, _
                  ByVal alignement As Long)
    c.Caption = contenu
    c.BackStyle = MSF_BackStyleTransparent
    c.SpecialEffect = MSF_SpecialEffectFlat
    c.BorderStyle = MSF_BorderStyleNone
    c.TextAlign = alignement
    c.WordWrap = False
    c.AutoSize = False
    PoserStyle c, st
End Sub

Private Sub PoserStyle(c As Object, ByRef st As StyleTexte)
    c.ForeColor = st.Couleur
    PoserPolice c, st.Police, st.Taille, st.Gras
End Sub

'------------------------------------------------------------------------------
' Zone de saisie. Locked plutôt qu'Enabled = False pour un champ géré par le
' programme : il reste lisible et son contenu copiable.
'
' AUCUNE VARIABLE DE CE MODULE NE DOIT S'APPELER « zone ». Une locale de ce nom
' MASQUE cette procédure : VBA ne dit rien à la compilation, il transforme
' « Zone c, x » en appel tardif sur l'objet, et MSForms répond à l'exécution
' « propriété ou méthode non gérée par cet objet » — une erreur 438 sur une
' ligne d'apparence irréprochable. Les cadres s'appellent donc cadre.
'------------------------------------------------------------------------------
Private Sub Zone(c As Object, ByVal verrouille As Boolean)
    PoserPolice c, PPOLICE, 9.5, False
    c.SpecialEffect = MSF_SpecialEffectFlat
    c.BorderStyle = MSF_BorderStyleSingle
    c.BorderColor = PCOUL_CHAMP_BORD
    c.TextAlign = MSF_TextAlignLeft
    c.EnterKeyBehavior = False
    If verrouille Then
        c.Locked = True
        c.BackColor = PCOUL_VERROU_FOND
        c.ForeColor = PCOUL_VERROU_TXT
        c.TabStop = False
    Else
        c.BackColor = PCOUL_CHAMP_FOND
        c.ForeColor = PCOUL_TEXTE
    End If
End Sub

Private Sub Liste(c As Object, ByVal ferme As Boolean)
    PoserPolice c, PPOLICE, 9.5, False
    c.SpecialEffect = MSF_SpecialEffectFlat
    c.BorderStyle = MSF_BorderStyleSingle
    c.BorderColor = PCOUL_CHAMP_BORD
    c.BackColor = PCOUL_CHAMP_FOND
    c.ForeColor = PCOUL_TEXTE
    c.ListRows = 12
    c.Style = IIf(ferme, MSF_StyleDropDownList, MSF_StyleDropDownCombo)
End Sub

Private Sub Bouton(c As Object, ByVal contenu As String, ByVal raccourci As String, _
                   ByVal couleur As Long, ByVal ordre As Long)
    PoserPolice c, PPOLICE, 9.5, True
    c.Caption = contenu
    c.Accelerator = raccourci
    c.BackColor = couleur
    c.ForeColor = PCOUL_BOUTON_TXT
    c.TabIndex = ordre
End Sub

'------------------------------------------------------------------------------
' L'ORDRE DES QUATRE ASSIGNATIONS EST LE SEUL QUI FONCTIONNE.
'
'   1. la FAMILLE d'abord : en changer remet le reste aux valeurs par défaut
'      de la nouvelle police ;
'   2. la GRAISSE ensuite, jamais après la taille ;
'   3. le POIDS, qui est la même chose vue de plus bas — 400 maigre, 700 gras.
'      Certaines versions ne reprennent que celui-là ; on pose les deux ;
'   4. la TAILLE en DERNIER.
'
' Pourquoi la taille en dernier : MSForms recrée la police en pixels entiers au
' moment où on lui donne un corps, et cette recréation emporte la graisse posée
' APRÈS elle. C'est ce qui laissait les zones de saisie en gras alors que
' Font.Bold = False était bien exécuté.
'
' Le poids est posé sous On Error Resume Next : toutes les versions de MSForms
' n'exposent pas Weight, et son absence ne doit pas faire échouer la génération
' — la graisse aura déjà été posée à la ligne précédente.
'------------------------------------------------------------------------------
Private Sub PoserPolice(c As Object, ByVal police As String, _
                        ByVal taille As Single, ByVal gras As Boolean)
    c.Font.Name = police
    c.Font.Bold = gras
    On Error Resume Next
    c.Font.Weight = IIf(gras, 700, 400)
    On Error GoTo 0
    c.Font.Size = taille
End Sub

Private Sub Prop(vbComp As Object, ByVal nom As String, ByVal valeur As Variant)
    On Error Resume Next
    vbComp.Properties(nom) = valeur
    On Error GoTo 0
End Sub

'==============================================================================
' MODULE DE CODE DE UF_Pression
'------------------------------------------------------------------------------
' Chaque procédure événementielle appelle modPression_Formulaire : régénérer le
' formulaire n'écrase donc jamais de comportement.
'==============================================================================
Private Function CodePression() As String
    Dim larg As Variant, r As Long, i As Long

    mCode = vbNullString
    Lig "'=============================================================================="
    Lig "' " & NOM_FORM_PRESSION & " - MODULE GÉNÉRÉ"
    Lig "'------------------------------------------------------------------------------"
    Lig "' Produit par modPression_Generateur. Toute modification faite ici sera perdue à"
    Lig "' la prochaine génération : le comportement s'écrit dans modPression_Formulaire."
    Lig "'=============================================================================="
    Lig "Option Explicit"
    Lig ""

    Proc "UserForm_Initialize()", "Pression_Initialiser Me"

    ' la date et la demi-journée désignent la ligne : les changer la recharge
    Proc "txtPDate_Change()", "Pression_Cibler Me"
    Proc "cboPDemi_Change()", "Pression_Cibler Me"

    ' l'état se rejuge à chaque frappe des deux valeurs qu'il regarde
    Proc "txtPSys_Change()", "Pression_Juger Me"
    Proc "txtPDia_Change()", "Pression_Juger Me"

    Proc "btnPAujourdhui_Click()", "Pression_Aujourdhui Me"
    Proc "btnPEnregistrer_Click()", "Pression_Enregistrer Me"
    Proc "btnPEffacer_Click()", "Pression_Effacer Me"
    Proc "btnPQuitter_Click()", "Pression_Quitter Me"

    ' la grille est faite de libellés : c'est la case cliquée qui reçoit
    ' l'événement, et elle ne sait dire que sa ligne
    larg = PDernieresLargeurs()
    For r = 1 To P_NB_DERNIERES
        Proc "lblPL_" & CStr(r) & "_Click()", "Pression_Reprendre Me, " & CStr(r)
        For i = LBound(larg) To UBound(larg)
            Proc "lblP_" & CStr(r) & "_" & CStr(i + 1) & "_Click()", _
                 "Pression_Reprendre Me, " & CStr(r)
        Next i
    Next r

    CodePression = mCode
End Function

Private Sub Proc(ByVal entete As String, ByVal corps As String)
    Lig "Private Sub " & entete
    Lig "    " & corps
    Lig "End Sub"
    Lig ""
End Sub

Private Sub Lig(ByVal ligne As String)
    mCode = mCode & ligne & vbNewLine
End Sub
