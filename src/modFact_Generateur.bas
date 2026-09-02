Attribute VB_Name = "modFact_Generateur"
Option Explicit
'==============================================================================
' modFact_Generateur
'------------------------------------------------------------------------------
' Génère de toutes pièces le formulaire UF_Facture, appelé par le bouton
' « Facturer » du formulaire des interventions.
'
'   >>> Procédure à lancer : GenererFormulaireFacture
'
' PRÉREQUIS : Fichier > Options > Centre de gestion de la confidentialité >
'             Paramètres du Centre de gestion de la confidentialité >
'             Paramètres des macros > cocher
'             « Accès approuvé au modèle d'objet du projet VBA ».
'
' IL NE FAUT PAS SUPPRIMER LE FORMULAIRE avant de régénérer : c'est inutile, la
' génération le vide et le reconstruit sur place. VBA ne rend le nom d'un
' composant supprimé qu'au prochain chargement du classeur.
'
' Deux utilitaires viennent de modInterv_Generateur plutôt que d'être réécrits :
' PreparerForm, qui sait reconnaître un nom encore retenu, et PoserPolice, dont
' l'ordre des assignations est la seule chose qui empêche la taille d'emporter
' la graisse.
'==============================================================================

Private Const CT_MSFORM_F As Long = 3
Private Const SIG_SOURIS_F As String = "(ByVal Button As Integer, ByVal Shift As Integer, " & _
                                       "ByVal X As Single, ByVal Y As Single)"

Private mCodeF As String

'==============================================================================
' POINT D'ENTRÉE
'==============================================================================
Public Sub GenererFormulaireFacture()
    Dim vbProj As Object, nbCtrl As Long

    If TableInterventions() Is Nothing Then
        MsgBox "Le tableau " & NOM_TABLE_INTERVENTIONS & " est introuvable dans ce classeur." & _
               vbCrLf & "Le formulaire de facturation ne peut pas être généré.", _
               vbCritical, "Génération du formulaire"
        Exit Sub
    End If

    On Error Resume Next
    Set vbProj = ThisWorkbook.VBProject
    On Error GoTo 0
    If vbProj Is Nothing Then
        MsgBox "Excel refuse l'accès au projet VBA." & vbCrLf & vbCrLf & _
               "Activez d'abord l'option :" & vbCrLf & _
               "Fichier > Options > Centre de gestion de la confidentialité >" & vbCrLf & _
               "Paramètres du Centre de gestion de la confidentialité >" & vbCrLf & _
               "Paramètres des macros > " & Chr$(34) & "Accès approuvé au modèle d'objet du " & _
               "projet VBA" & Chr$(34) & vbCrLf & vbCrLf & _
               "Fermez puis rouvrez le classeur, et relancez cette procédure.", _
               vbCritical, "Génération du formulaire"
        Exit Sub
    End If

    On Error GoTo Erreur
    nbCtrl = ConstruireFacture(vbProj)

    MsgBox "Le formulaire " & NOM_FORM_FACTURE & " a été généré." & vbCrLf & vbCrLf & _
           nbCtrl & " contrôles." & vbCrLf & vbCrLf & _
           "Il s'ouvre par le bouton " & Chr$(34) & "Facturer" & Chr$(34) & " du formulaire " & _
           "des interventions.", vbInformation, "Génération du formulaire"
    Exit Sub

Erreur:
    MsgBox "La génération a échoué :" & vbCrLf & vbCrLf & _
           Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
           "Si le problème persiste : fermez puis rouvrez le classeur, lancez " & _
           "NettoyerFormulairesOrphelins, et relancez la génération.", _
           vbCritical, "Génération du formulaire"
End Sub

'==============================================================================
' LE FORMULAIRE
'==============================================================================
Private Function ConstruireFacture(vbProj As Object) As Long
    Dim vbComp As Object, dsg As Object

    Set vbComp = PreparerForm(vbProj, NOM_FORM_FACTURE)
    Set dsg = vbComp.Designer

    PropF vbComp, "Caption", "Facturation"
    PropF vbComp, "Width", FA_LARGEUR
    PropF vbComp, "Height", FA_HAUTEUR
    PropF vbComp, "BackColor", COUL_FOND
    PropF vbComp, "SpecialEffect", MSF_SpecialEffectFlat
    PropF vbComp, "StartUpPosition", 1
    PropF vbComp, "ShowModal", True
    PoserPolice dsg, POLICE, 8, False

    ConstruireZone1 dsg
    ConstruireZone2 dsg
    ConstruireZone3 dsg
    ConstruireZone4 dsg
    ConstruireZone5 dsg
    ConstruireBoutonsF dsg

    vbComp.CodeModule.AddFromString CodeFacture()
    ConstruireFacture = dsg.Controls.Count
End Function

'------------------------------------------------------------------------------
' Zone 1 : l'intitulé — titre à gauche, année à droite.
'------------------------------------------------------------------------------
Private Sub ConstruireZone1(dsg As Object)
    Dim c As Object, zone As Object

    Set zone = AjF(dsg, "Forms.Frame.1", "fraFCarte1", _
                   FA_MARGE, FA_Z1_TOP, FA_CARTE_LARG, FA_Z1_HAUT)
    CadreF zone

    Set c = AjF(zone, "Forms.Label.1", "lblFTitre", 14, 12, 620, 21)
    TexteF c, "Liste des clients non encore facturés", ZF1Titre(), MSF_TextAlignLeft

    Set c = AjF(zone, "Forms.Label.1", "lblFAnnee", FA_CARTE_LARG - 244, 8, 230, 30)
    TexteF c, vbNullString, ZF1Annee(), MSF_TextAlignRight
End Sub

'------------------------------------------------------------------------------
' Zone 2 : le tableau des clients qu'il reste à facturer.
'------------------------------------------------------------------------------
Private Sub ConstruireZone2(dsg As Object)
    Dim zone As Object

    Set zone = AjF(dsg, "Forms.Frame.1", "fraFCarte2", _
                   FA_MARGE, FA_Z2_TOP, FA_CARTE_LARG, FA_Z2_HAUT)
    CadreF zone

    BandeauF zone, "lblFSection2", "Liste des clients aux travaux non facturés"
    ConstruireGrilleF zone, "C", FClientsLibelles(), FClientsLargeurs(), _
                      FClientsAlignements(), FA_Z2_LIGNES, "sbFClients"
End Sub

'------------------------------------------------------------------------------
' Zone 3 : le filtre des mois, « toutes les opérations », le nouveau numéro.
'------------------------------------------------------------------------------
Private Sub ConstruireZone3(dsg As Object)
    Dim c As Object, zone As Object, y As Single

    Set zone = AjF(dsg, "Forms.Frame.1", "fraFCarte3", _
                   FA_MARGE, FA_Z3_TOP, FA_CARTE_LARG, FA_Z3_HAUT)
    CadreF zone
    y = AuPixel((FA_Z3_HAUT - FA_CTL_HAUT) / 2)

    Set c = AjF(zone, "Forms.Label.1", "lblFMoisCap", FA_MOIS_LBL_X, y + 2, 34, 14)
    TexteF c, "Mois", ZF3Libelle(), MSF_TextAlignLeft

    Set c = AjF(zone, "Forms.ComboBox.1", "cboFMois", FA_MOIS_X, y, FA_MOIS_L, FA_CTL_HAUT)
    ListeF c
    c.ControlTipText = "N'afficher que les travaux de ce mois"

    Set c = AjF(zone, "Forms.CheckBox.1", "chkFToutes", FA_TOUTES_X, y + 1, _
                FA_TOUTES_L, FA_CTL_HAUT)
    CaseF c, "Toutes les opérations"
    c.ControlTipText = "Montrer aussi les interventions déjà facturées"

    Set c = AjF(zone, "Forms.Label.1", "lblFNumCap", FA_NUM_LBL_X, y + 2, 94, 14)
    TexteF c, "Nouv. Numéro", ZF3Libelle(), MSF_TextAlignRight

    Set c = AjF(zone, "Forms.TextBox.1", "txtFNum", FA_NUM_X, y, FA_NUM_L, FA_CTL_HAUT)
    ZoneF c
    c.ControlTipText = "Numéro de facture à attribuer aux lignes cochées"
End Sub

'------------------------------------------------------------------------------
' Zone 4 : le tableau des travaux, ses deux totaux et le bouton Enregistrer.
'------------------------------------------------------------------------------
Private Sub ConstruireZone4(dsg As Object)
    Dim c As Object, zone As Object, larg As Variant, y As Single, x As Single, i As Long

    Set zone = AjF(dsg, "Forms.Frame.1", "fraFCarte4", _
                   FA_MARGE, FA_Z4_TOP, FA_CARTE_LARG, FA_Z4_HAUT)
    CadreF zone

    BandeauF zone, "lblFSection4", "Liste des travaux non facturés"
    ConstruireGrilleF zone, "T", FTravauxLibelles(), FTravauxLargeurs(), _
                      FTravauxAlignements(), FA_Z4_LIGNES, "sbFTravaux"

    ' --- la ligne des totaux, sous les colonnes qu'elle additionne ------------
    larg = FTravauxLargeurs()
    y = AuPixel(FGrilleLignesY() + FA_Z4_LIGNES * FA_LIGNE_H + 2)
    x = AuPixel(FA_PAD_X + 1)

    For i = LBound(larg) To UBound(larg)
        If i - LBound(larg) + 1 = FIndexTravaux(IC_HEURES) Then
            Set c = AjF(zone, "Forms.Label.1", "lblFTotHeures", AuPixel(x), y, _
                        AuPixel(CSng(larg(i)) - 2 * FA_PAD_X), FA_TOTAUX_HAUT)
            TexteF c, vbNullString, ZFTotal(), MSF_TextAlignRight
            c.ControlTipText = "Total des heures affichées"
        ElseIf i - LBound(larg) + 1 = FIndexTravaux(IC_CA) Then
            Set c = AjF(zone, "Forms.Label.1", "lblFTotCA", AuPixel(x), y, _
                        AuPixel(CSng(larg(i)) - 2 * FA_PAD_X), FA_TOTAUX_HAUT)
            TexteF c, vbNullString, ZFTotal(), MSF_TextAlignRight
            c.ControlTipText = "Total du chiffre d'affaires affiché"
        End If
        x = x + CSng(larg(i))
    Next i

    Set c = AjF(zone, "Forms.CommandButton.1", "btnFEnregistrer", _
                FA_CARTE_LARG - FA_BT_LARG - 8, y - 2, FA_BT_LARG, FA_BT_HAUT - 6)
    BoutonF c, "Enregistrer", "E", COUL_AJOUTER, 40
End Sub

'------------------------------------------------------------------------------
' Zone 5 : le texte de facture et les commentaires de la ligne choisie.
'
' Trois lignes chacun, en lecture : ces deux zones montrent le contenu du
' record sélectionné, elles ne le modifient pas.
'------------------------------------------------------------------------------
Private Sub ConstruireZone5(dsg As Object)
    Dim c As Object, zone As Object, demi As Single

    Set zone = AjF(dsg, "Forms.Frame.1", "fraFCarte5", _
                   FA_MARGE, FA_Z5_TOP, FA_CARTE_LARG, FA_Z5_HAUT)
    CadreF zone
    demi = AuPixel((FA_CARTE_LARG - 3 * FA_PADDING) / 2)

    Set c = AjF(zone, "Forms.Label.1", "lblFTexteCap", FA_PADDING, 6, demi, 11)
    TexteF c, "TEXTE DE FACTURE", ZF5Libelle(), MSF_TextAlignLeft
    Set c = AjF(zone, "Forms.TextBox.1", "txtFTexte", FA_PADDING, 19, demi, 48)
    ZoneF c
    MultiLigneF c

    Set c = AjF(zone, "Forms.Label.1", "lblFCommCap", 2 * FA_PADDING + demi, 6, demi, 11)
    TexteF c, "COMMENTAIRES", ZF5Libelle(), MSF_TextAlignLeft
    Set c = AjF(zone, "Forms.TextBox.1", "txtFComm", 2 * FA_PADDING + demi, 19, demi, 48)
    ZoneF c
    MultiLigneF c
End Sub

'------------------------------------------------------------------------------
' La rangée de boutons, sur le formulaire lui-même.
'------------------------------------------------------------------------------
Private Sub ConstruireBoutonsF(dsg As Object)
    Dim c As Object

    Set c = AjF(dsg, "Forms.CommandButton.1", "btnFQuitter", _
                FA_LARGEUR - FA_MARGE - FA_BT_LARG, FA_BT_TOP, FA_BT_LARG, FA_BT_HAUT)
    BoutonF c, "Quitter", "Q", COUL_QUITTER, 50
End Sub

'==============================================================================
' UNE GRILLE DE LIBELLÉS
'------------------------------------------------------------------------------
' Les deux tableaux se construisent par cette seule procédure : même bandeau,
' même en-tête, mêmes bandes de fond, mêmes cases. Seuls changent le préfixe
' des noms — C pour les clients, T pour les travaux — et les quatre tableaux
' qui décrivent les colonnes.
'
' Chaque ligne reçoit d'abord une BANDE de fond sur toute la largeur, puis ses
' cases, transparentes, posées dessus : c'est la bande qui porte la couleur,
' sans quoi les points qui séparent deux cases laisseraient voir le blanc de la
' carte et la ligne choisie paraîtrait rayée.
'
' TOUTE coordonnée passe par AuPixel. L'abscisse d'une colonne est la somme des
' largeurs qui la précèdent, et rien ne garantit qu'elle tombe sur un pixel :
' une case posée à cheval rend son texte décalé et plus épais. Les largeurs,
' elles, restent exactes — seul l'affichage est calé.
'==============================================================================
Private Sub ConstruireGrilleF(zone As Object, ByVal prefixe As String, _
                              ByVal lib As Variant, ByVal larg As Variant, _
                              ByVal ali As Variant, ByVal nbLignes As Long, _
                              ByVal nomBarre As String)
    Dim c As Object, i As Long, r As Long, nbCol As Long
    Dim x As Single, y As Single, gauche As Single, largeur As Single

    gauche = AuPixel(1)
    largeur = FGrilleLargeur()
    nbCol = UBound(larg) - LBound(larg) + 1

    ' --- fond de l'en-tête ----------------------------------------------------
    Set c = AjF(zone, "Forms.Label.1", "lblFEntFond" & prefixe, _
                gauche, FGrilleEnteteY(), largeur, AuPixel(FA_ENTETE_HAUT))
    FondF c, COUL_ENTETE_TBL, COUL_ENTETE_TBL

    ' --- titres de colonnes ---------------------------------------------------
    x = gauche + FA_PAD_X
    For i = 0 To nbCol - 1
        Set c = AjF(zone, "Forms.Label.1", "lblFEnt" & prefixe & "_" & CStr(i + 1), _
                    AuPixel(x), AuPixel(FGrilleEnteteY() + 4), _
                    AuPixel(CSng(larg(i)) - 2 * FA_PAD_X), AuPixel(FA_LIGNE_H))
        TexteF c, CStr(lib(i)), ZFEntete(), CLng(ali(i))
        x = x + CSng(larg(i))
    Next i

    ' --- les lignes et leurs cases --------------------------------------------
    For r = 1 To nbLignes
        y = FGrilleLignesY() + (r - 1) * FA_LIGNE_H

        Set c = AjF(zone, "Forms.Label.1", "lblF" & prefixe & "L_" & CStr(r), _
                    gauche, AuPixel(y), AuPixel(largeur - FA_BARRE_L), AuPixel(FA_LIGNE_H))
        With c
            .Caption = vbNullString
            .SpecialEffect = MSF_SpecialEffectFlat
            .BorderStyle = MSF_BorderStyleNone
            .BackStyle = MSF_BackStyleOpaque
            .BackColor = COUL_CARTE
        End With
        PoserStyleF c, ZFCase()

        x = gauche + FA_PAD_X
        For i = 0 To nbCol - 1
            Set c = AjF(zone, "Forms.Label.1", _
                        "lblF" & prefixe & "_" & CStr(r) & "_" & CStr(i + 1), _
                        AuPixel(x), AuPixel(y), _
                        AuPixel(CSng(larg(i)) - 2 * FA_PAD_X), AuPixel(FA_LIGNE_H))
            TexteF c, vbNullString, ZFCase(), CLng(ali(i))
            x = x + CSng(larg(i))
        Next i
    Next r

    ' --- barre de défilement --------------------------------------------------
    ' Curseur de TAILLE FIXE : avec ProportionalThumb, MSForms le dimensionne à
    ' LargeChange sur la plage, et il occupe alors presque toute la glissière.
    Set c = AjF(zone, "Forms.ScrollBar.1", nomBarre, _
                AuPixel(gauche + largeur - FA_BARRE_L), AuPixel(FGrilleLignesY()), _
                AuPixel(FA_BARRE_L), AuPixel(nbLignes * FA_LIGNE_H))
    With c
        .Min = 0
        .Max = 0
        .SmallChange = 1
        .LargeChange = nbLignes
        .Value = 0
        .ProportionalThumb = False
    End With
End Sub

'------------------------------------------------------------------------------
' Le bandeau de titre d'un tableau, en haut de son cadre.
'------------------------------------------------------------------------------
Private Sub BandeauF(zone As Object, ByVal nom As String, ByVal titre As String)
    Dim c As Object

    Set c = AjF(zone, "Forms.Label.1", nom, AuPixel(FA_PAD_X + 1), AuPixel(4), _
                FGrilleLargeur() - 2 * FA_PAD_X, 13)
    TexteF c, UCase$(titre), ZFSection(), MSF_TextAlignLeft
End Sub

'==============================================================================
' HABILLAGE
'------------------------------------------------------------------------------
' Ces procédures ne font que POSER DES PROPRIÉTÉS, jamais de mise en page : les
' coordonnées restent dans les Construire*, et la police passe toujours par
' PoserPolice, de modInterv_Generateur.
'==============================================================================

' Le premier argument est un CONTENEUR, pas forcément le formulaire : un cadre
' expose la même collection Controls, et les coordonnées comptent alors depuis
' son coin.
Private Function AjF(conteneur As Object, ByVal progId As String, ByVal nom As String, _
                     ByVal gauche As Single, ByVal haut As Single, _
                     ByVal largeur As Single, ByVal hauteur As Single) As Object
    Dim c As Object

    Set c = conteneur.Controls.Add(progId, nom, True)
    c.Left = gauche
    c.Top = haut
    c.Width = largeur
    c.Height = hauteur
    Set AjF = c
End Function

' SpecialEffect avant BorderStyle : MSForms refuse une bordure simple tant que
' le contrôle est en relief, et un cadre est gravé par défaut. Caption vidée,
' sans quoi MSForms réserve une bande de titre en haut du cadre.
Private Sub CadreF(c As Object)
    c.Caption = vbNullString
    c.BackColor = COUL_CARTE
    c.SpecialEffect = MSF_SpecialEffectFlat
    c.BorderStyle = MSF_BorderStyleSingle
    c.BorderColor = COUL_BORDURE
    c.ScrollBars = MSF_ScrollBarsNone
End Sub

Private Sub FondF(c As Object, ByVal fond As Long, ByVal bordure As Long)
    c.Caption = vbNullString
    c.BackStyle = MSF_BackStyleOpaque
    c.BackColor = fond
    c.SpecialEffect = MSF_SpecialEffectFlat
    c.BorderStyle = MSF_BorderStyleSingle
    c.BorderColor = bordure
End Sub

Private Sub TexteF(c As Object, ByVal texte As String, ByRef st As StyleTexte, _
                   ByVal alignement As Long)
    c.Caption = texte
    c.BackStyle = MSF_BackStyleTransparent
    c.SpecialEffect = MSF_SpecialEffectFlat
    c.BorderStyle = MSF_BorderStyleNone
    c.TextAlign = alignement
    c.WordWrap = False
    c.AutoSize = False
    PoserStyleF c, st
End Sub

Private Sub PoserStyleF(c As Object, ByRef st As StyleTexte)
    c.ForeColor = st.Couleur
    PoserPolice c, st.Police, st.Taille, st.Gras
End Sub

Private Sub ZoneF(c As Object)
    PoserPolice c, POLICE, 9, False
    c.SpecialEffect = MSF_SpecialEffectFlat
    c.BorderStyle = MSF_BorderStyleSingle
    c.BorderColor = COUL_CHAMP_BORD
    c.BackColor = COUL_CHAMP_FOND
    c.ForeColor = COUL_TEXTE
    c.TextAlign = MSF_TextAlignLeft
    c.EnterKeyBehavior = False
End Sub

' Une zone haute de trois lignes accueille du texte long : le retour à la ligne
' et l'ascenseur y deviennent utiles.
Private Sub MultiLigneF(c As Object)
    c.MultiLine = True
    c.ScrollBars = MSF_ScrollBarsVertical
    c.Locked = True
    c.BackColor = COUL_VERROU_FOND
    c.ForeColor = COUL_VERROU_TXT
    c.TabStop = False
End Sub

Private Sub ListeF(c As Object)
    PoserPolice c, POLICE, 9, False
    c.SpecialEffect = MSF_SpecialEffectFlat
    c.BorderStyle = MSF_BorderStyleSingle
    c.BorderColor = COUL_CHAMP_BORD
    c.BackColor = COUL_CHAMP_FOND
    c.ForeColor = COUL_TEXTE
    c.ListRows = 13
    c.Style = MSF_StyleDropDownList
End Sub

' Le relief de la case elle-même est conservé : à plat, elle deviendrait
' difficile à distinguer d'un simple libellé.
Private Sub CaseF(c As Object, ByVal libelle As String)
    PoserPolice c, POLICE, 9, False
    c.Caption = libelle
    c.BackStyle = MSF_BackStyleTransparent
    c.ForeColor = COUL_TEXTE
    c.WordWrap = False
End Sub

Private Sub BoutonF(c As Object, ByVal libelle As String, ByVal raccourci As String, _
                    ByVal couleur As Long, ByVal ordre As Long)
    PoserPolice c, POLICE, 9.75, True
    c.Caption = libelle
    c.Accelerator = raccourci
    c.BackColor = couleur
    c.ForeColor = COUL_BOUTON_TXT
    c.TabIndex = ordre
End Sub

Private Sub PropF(vbComp As Object, ByVal nom As String, ByVal valeur As Variant)
    On Error Resume Next
    vbComp.Properties(nom) = valeur
    On Error GoTo 0
End Sub

'==============================================================================
' MODULE DE CODE DE UF_Facture
'------------------------------------------------------------------------------
' Chaque procédure événementielle se contente d'appeler modFact_Formulaire : le
' code métier n'est jamais dupliqué ici, et régénérer le formulaire ne peut
' rien écraser d'écrit à la main.
'==============================================================================
Private Function CodeFacture() As String
    Dim larg As Variant, r As Long, i As Long

    mCodeF = vbNullString
    LigF "'=============================================================================="
    LigF "' " & NOM_FORM_FACTURE & " - MODULE GÉNÉRÉ"
    LigF "'------------------------------------------------------------------------------"
    LigF "' Produit par modFact_Generateur. Toute modification faite ici sera perdue à"
    LigF "' la prochaine génération : le comportement s'écrit dans modFact_Formulaire."
    LigF "'=============================================================================="
    LigF "Option Explicit"
    LigF ""

    ProcF "UserForm_Initialize()", "Fact_Initialiser Me"
    ProcF "UserForm_Activate()", "Fact_Activer Me"

    ' --- tableau des clients --------------------------------------------------
    ' La grille est faite de libellés : c'est la case cliquée qui reçoit
    ' l'événement, et elle ne sait dire que sa ligne — c'est tout ce qui compte,
    ' une ligne entière se sélectionnant d'un bloc.
    larg = FClientsLargeurs()
    For r = 1 To FA_Z2_LIGNES
        For i = LBound(larg) To UBound(larg)
            ProcF "lblFC_" & CStr(r) & "_" & CStr(i + 1) & "_Click()", _
                  "Fact_ClientClic Me, " & CStr(r)
        Next i
        ProcF "lblFCL_" & CStr(r) & "_Click()", "Fact_ClientClic Me, " & CStr(r)
    Next r
    ProcF "sbFClients_Change()", "Fact_DefilerClients Me"

    ' --- tableau des travaux --------------------------------------------------
    ' Le numéro de COLONNE est transmis en plus de la ligne : c'est lui qui
    ' distingue un clic ordinaire, qui choisit la ligne, d'un clic dans la
    ' colonne Select., qui bascule la case.
    larg = FTravauxLargeurs()
    For r = 1 To FA_Z4_LIGNES
        For i = LBound(larg) To UBound(larg)
            ProcF "lblFT_" & CStr(r) & "_" & CStr(i + 1) & "_Click()", _
                  "Fact_TravailClic Me, " & CStr(r) & ", " & CStr(i + 1)
        Next i
        ProcF "lblFTL_" & CStr(r) & "_Click()", _
              "Fact_TravailClic Me, " & CStr(r) & ", 0"
    Next r
    ProcF "sbFTravaux_Change()", "Fact_DefilerTravaux Me"

    ' --- filtres et boutons ---------------------------------------------------
    ProcF "cboFMois_Change()", "Fact_RafraichirTravaux Me"
    ProcF "chkFToutes_Click()", "Fact_RafraichirTravaux Me"
    ProcF "btnFEnregistrer_Click()", "Fact_Enregistrer_Clic Me"
    ProcF "btnFQuitter_Click()", "Fact_Quitter Me"

    CodeFacture = mCodeF
End Function

'------------------------------------------------------------------------------
' Une procédure événementielle d'une seule ligne.
'------------------------------------------------------------------------------
Private Sub ProcF(ByVal entete As String, ByVal corps As String)
    LigF "Private Sub " & entete
    LigF "    " & corps
    LigF "End Sub"
    LigF ""
End Sub

Private Sub LigF(ByVal ligne As String)
    mCodeF = mCodeF & ligne & vbNewLine
End Sub
