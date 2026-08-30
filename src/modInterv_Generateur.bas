Attribute VB_Name = "modInterv_Generateur"
Option Explicit
'==============================================================================
' modInterv_Generateur
'------------------------------------------------------------------------------
' Génère de toutes pièces les deux formulaires des interventions :
'   UF_Interventions  le formulaire principal, quatre fiches empilées ;
'   UF_Calendrier     le sélecteur de date.
'
'   >>> Procédure à lancer : GenererFormulaireInterventions
'
' PRÉREQUIS : Fichier > Options > Centre de gestion de la confidentialité >
'             Paramètres du Centre de gestion de la confidentialité >
'             Paramètres des macros > cocher
'             « Accès approuvé au modèle d'objet du projet VBA ».
'
' Un formulaire déjà présent est VIDÉ puis reconstruit sur place, jamais
' supprimé puis recréé : VBA ne libère le nom d'un composant supprimé qu'au
' retour à Excel, et l'affectation du nom échouerait.
'==============================================================================

Private Const CT_MSFORM As Long = 3         ' vbext_ct_MSForm
Private Const SIG_SOURIS As String = "(ByVal Button As Integer, ByVal Shift As Integer, " & _
                                     "ByVal X As Single, ByVal Y As Single)"

Private mCode As String

'==============================================================================
' POINT D'ENTRÉE
'==============================================================================
Public Sub GenererFormulaireInterventions()
    Dim vbProj As Object, lo As ListObject, manquantes As String
    Dim nbCtrl As Long, nbCal As Long

    ' --- le tableau source doit exister ---------------------------------------
    Set lo = TableInterventions()
    If lo Is Nothing Then
        MsgBox "Le tableau " & NOM_TABLE_INTERVENTIONS & " est introuvable dans ce classeur." & _
               vbCrLf & "Le formulaire ne peut pas être généré.", vbCritical, _
               "Génération du formulaire"
        Exit Sub
    End If

    manquantes = ColonnesNonCouvertesInterv(lo)
    If Len(manquantes) > 0 Then
        If MsgBox("Ces colonnes de " & NOM_TABLE_INTERVENTIONS & " ne figurent pas dans le " & _
                  "schéma du formulaire et ne seront donc pas saisissables :" & vbCrLf & vbCrLf & _
                  manquantes & vbCrLf & vbCrLf & "Poursuivre la génération ?", _
                  vbQuestion + vbYesNo, "Génération du formulaire") <> vbYes Then Exit Sub
    End If

    ' --- accès au projet VBA --------------------------------------------------
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
    nbCtrl = ConstruireFormulairePrincipal(vbProj)
    nbCal = ConstruireCalendrier(vbProj)

    MsgBox "Les formulaires ont été générés." & vbCrLf & vbCrLf & _
           NOM_FORM_INTERV & " : " & nbCtrl & " contrôles pour " & NB_CHAMPS_INTERV & _
           " champs du tableau " & NOM_TABLE_INTERVENTIONS & "." & vbCrLf & _
           NOM_FORM_CALENDRIER & " : " & nbCal & " contrôles." & vbCrLf & vbCrLf & _
           "Lancez maintenant OuvrirGestionInterventions pour l'afficher.", _
           vbInformation, "Génération du formulaire"
    Exit Sub

Erreur:
    MsgBox "La génération a échoué :" & vbCrLf & vbCrLf & _
           Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
           "Si le problème persiste : fermez puis rouvrez le classeur, lancez " & _
           "NettoyerFormulairesOrphelins, et relancez la génération.", _
           vbCritical, "Génération du formulaire"
End Sub

'------------------------------------------------------------------------------
' Colonnes de TblInterv absentes du schéma : ce sont exactement celles qui ne
' seront ni affichées ni saisissables.
'------------------------------------------------------------------------------
Private Function ColonnesNonCouvertesInterv(ByVal lo As ListObject) As String
    Dim lc As ListColumn, ch() As ChampInterv, i As Long, trouve As Boolean, res As String

    ch = ObtenirChampsInterv()
    For Each lc In lo.ListColumns
        trouve = False
        For i = LBound(ch) To UBound(ch)
            If StrComp(ch(i).Colonne, lc.Name, vbTextCompare) = 0 Then
                trouve = True
                Exit For
            End If
        Next i
        If Not trouve Then res = res & IIf(Len(res) > 0, ", ", "") & lc.Name
    Next lc
    ColonnesNonCouvertesInterv = res
End Function

'==============================================================================
' Prépare un composant UserForm : créé s'il manque, vidé s'il existe déjà.
'   renvoie : le composant, prêt à être garni
'==============================================================================
Private Function PreparerForm(vbProj As Object, ByVal nom As String) As Object
    Dim vbComp As Object, dsg As Object, i As Long

    On Error Resume Next
    Set vbComp = vbProj.VBComponents(nom)
    On Error GoTo 0

    If vbComp Is Nothing Then
        Set vbComp = vbProj.VBComponents.Add(CT_MSFORM)
        vbComp.Name = nom
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
    For i = dsg.Controls.Count - 1 To 0 Step -1
        dsg.Controls.Remove dsg.Controls(i).Name
    Next i
    With vbComp.CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
    End With

    Set PreparerForm = vbComp
End Function

'------------------------------------------------------------------------------
' Pose une propriété du formulaire.
' Les erreurs sont ignorées volontairement : si une version d'Excel n'expose pas
' l'une de ces propriétés, le formulaire se génère quand même, avec la valeur par
' défaut pour celle-là.
'------------------------------------------------------------------------------
Private Sub PropForm(vbComp As Object, ByVal nom As String, ByVal valeur As Variant)
    On Error Resume Next
    vbComp.Properties(nom) = valeur
    On Error GoTo 0
End Sub

'==============================================================================
' FORMULAIRE PRINCIPAL
'==============================================================================
Private Function ConstruireFormulairePrincipal(vbProj As Object) As Long
    Dim vbComp As Object, dsg As Object

    Set vbComp = PreparerForm(vbProj, NOM_FORM_INTERV)
    Set dsg = vbComp.Designer

    ' La barre de titre Windows est conservée : la fiche 1 tient lieu d'intitulé
    ' à l'intérieur du formulaire. La hauteur est corrigée à l'ouverture, dans
    ' Interv_Activer, pour que la surface utile mesure bien I_HAUTEUR.
    PropForm vbComp, "Caption", "Interventions"
    PropForm vbComp, "Width", I_LARGEUR
    PropForm vbComp, "Height", I_HAUTEUR
    PropForm vbComp, "BackColor", COUL_FOND
    PropForm vbComp, "SpecialEffect", MSF_SpecialEffectFlat
    PropForm vbComp, "StartUpPosition", 1
    PropForm vbComp, "ShowModal", True

    ConstruireFiche1 dsg
    ConstruireFiche2 dsg
    ConstruireFiche3 dsg
    ConstruireFiltre dsg
    ConstruireFiche4 dsg
    ConstruireBoutonsInterv dsg
    ReculerFonds dsg, Array("lblEnteteTableI", "lblCarte4", "lblCarteFiltreI", _
                            "lblCarte3", "lblCarte2", "lblCarte1")

    vbComp.CodeModule.AddFromString CodeFormulairePrincipal()
    ConstruireFormulairePrincipal = dsg.Controls.Count
End Function

'------------------------------------------------------------------------------
' Fiche 1 : intitulé global — titre à gauche, année à droite, ligne d'état.
'------------------------------------------------------------------------------
Private Sub ConstruireFiche1(dsg As Object)
    Dim c As Object

    Set c = AjCtrl(dsg, "Forms.Label.1", "lblCarte1", I_MARGE, F1_TOP, I_CARTE_LARG, F1_HAUT)
    CarteI c

    Set c = AjCtrl(dsg, "Forms.Label.1", "lblTitreGlobal", 30, F1_TOP + 7, 560, 21)
    Texte c, vbNullString, 14, True, COUL_BANDEAU, MSF_TextAlignLeft

    Set c = AjCtrl(dsg, "Forms.Label.1", "lblEtatI", 30, F1_TOP + 28, 620, 13)
    Texte c, vbNullString, TAILLE_SOUSTITRE, False, COUL_TEXTE_DOUX, MSF_TextAlignLeft

    Set c = AjCtrl(dsg, "Forms.Label.1", "lblAnnee", 700, F1_TOP + 8, 230, 30)
    Texte c, vbNullString, 20, True, COUL_MODIFIER, MSF_TextAlignRight
End Sub

'------------------------------------------------------------------------------
' Fiche 2 : statistiques — graphique à gauche, six tuiles à droite.
'
' Chaque tuile est une image de fond surmontée de deux libellés à fond blanc :
' l'image porte les coins arrondis, les libellés portent le texte. Si l'image
' est absente, un aplat blanc prend sa place et la tuile reste lisible.
'------------------------------------------------------------------------------
Private Sub ConstruireFiche2(dsg As Object)
    Dim c As Object, tuiles As Variant, i As Long, x As Single, y As Single
    Dim col As Long, lig As Long

    Set c = AjCtrl(dsg, "Forms.Label.1", "lblCarte2", I_MARGE, F2_TOP, I_CARTE_LARG, F2_HAUT)
    CarteI c

    ' large : ce libellé accueille aussi, le cas échéant, la raison pour laquelle
    ' le graphique ne s'affiche pas
    Set c = AjCtrl(dsg, "Forms.Label.1", "lblSectionStats", 30, F2_TOP + 6, 700, 13)
    Texte c, "STATISTIQUES", TAILLE_SECTION, True, COUL_SECTION, MSF_TextAlignLeft

    Set c = AjCtrl(dsg, "Forms.Image.1", "imgGraphique", 30, F2_TOP + 24, F2_GRAPH_LARG, 118)
    With c
        .BorderStyle = MSF_BorderStyleNone
        .SpecialEffect = MSF_SpecialEffectFlat
        .BackStyle = MSF_BackStyleTransparent
        .PictureSizeMode = 3            ' fmPictureSizeModeZoom : garde les proportions
        .PictureAlignment = 2           ' fmPictureAlignmentCenter
    End With

    tuiles = TuilesStatistiques()
    For i = LBound(tuiles) To UBound(tuiles)
        col = i Mod F2_TUILES_COL
        lig = i \ F2_TUILES_COL
        x = F2TuilesX() + col * (F2_TUILE_LARG + F2_TUILE_GX)
        y = F2_TOP + 24 + lig * (F2_TUILE_HAUT + F2_TUILE_GY)

        Set c = AjCtrl(dsg, "Forms.Image.1", "imgTuile_" & CStr(i + 1), x, y, _
                       F2_TUILE_LARG, F2_TUILE_HAUT)
        With c
            .BorderStyle = MSF_BorderStyleNone
            .SpecialEffect = MSF_SpecialEffectFlat
            .BackStyle = MSF_BackStyleOpaque
            .BackColor = COUL_CARTE
            .PictureSizeMode = 1        ' fmPictureSizeModeStretch : remplit la tuile
        End With

        Set c = AjCtrl(dsg, "Forms.Label.1", "lblStatCap_" & CStr(i + 1), _
                       x + F2_TUILE_INSET, y + 8, F2_TUILE_LARG - 2 * F2_TUILE_INSET, 11)
        Texte c, UCase$(CStr(tuiles(i)(0))), TAILLE_LIBELLE, True, COUL_TEXTE_DOUX, MSF_TextAlignLeft
        c.BackStyle = MSF_BackStyleOpaque
        c.BackColor = COUL_CARTE

        Set c = AjCtrl(dsg, "Forms.Label.1", "lblStatVal_" & CStr(i + 1), _
                       x + F2_TUILE_INSET, y + 22, F2_TUILE_LARG - 2 * F2_TUILE_INSET, 22)
        Texte c, vbNullString, TAILLE_STAT, True, COUL_BANDEAU, MSF_TextAlignLeft
        c.BackStyle = MSF_BackStyleOpaque
        c.BackColor = COUL_CARTE
    Next i
End Sub

'------------------------------------------------------------------------------
' Fiche 3 : les quinze champs de saisie, posés d'après le schéma.
'------------------------------------------------------------------------------
Private Sub ConstruireFiche3(dsg As Object)
    Dim c As Object, ch() As ChampInterv, i As Long
    Dim x As Single, y As Single, larg As Single, ordre As Long

    Set c = AjCtrl(dsg, "Forms.Label.1", "lblCarte3", I_MARGE, F3_TOP, I_CARTE_LARG, F3_HAUT)
    CarteI c

    Set c = AjCtrl(dsg, "Forms.Label.1", "lblSectionSaisieI", 30, F3_TOP + 6, 300, 13)
    Texte c, "FICHE INTERVENTION", TAILLE_SECTION, True, COUL_SECTION, MSF_TextAlignLeft

    ch = ObtenirChampsInterv()
    ordre = 1

    For i = LBound(ch) To UBound(ch)
        x = IGrilleX(ch(i).Col)
        y = IGrilleY(ch(i).Ligne)
        larg = ILargeurBlocs(ch(i).Blocs)

        ' deux champs peuvent se partager un bloc : les cases TVA et Forfait
        If ch(i).Moitie = 1 Then
            larg = (IG_BLOC - 8) / 2
        ElseIf ch(i).Moitie = 2 Then
            larg = (IG_BLOC - 8) / 2
            x = x + larg + 8
        End If

        If ch(i).TypeCtrl = ITYPE_CASE Then
            Set c = AjCtrl(dsg, "Forms.CheckBox.1", INomControle(ch(i)), _
                           x, y + ICH_LBL_HAUT + 1, larg, ICH_CTL_HAUT)
            Case_ c, ch(i).Libelle
        Else
            Set c = AjCtrl(dsg, "Forms.Label.1", INomLibelle(ch(i)), x, y, larg, ICH_LBL_HAUT)
            Texte c, UCase$(ch(i).Libelle), TAILLE_LIBELLE, True, COUL_TEXTE_DOUX, MSF_TextAlignLeft

            Select Case ch(i).TypeCtrl
                Case ITYPE_LISTE, ITYPE_AUTO
                    Set c = AjCtrl(dsg, "Forms.ComboBox.1", INomControle(ch(i)), _
                                   x, y + ICH_LBL_HAUT + 1, larg, ICH_CTL_HAUT)
                    Liste c, (ch(i).TypeCtrl = ITYPE_AUTO)

                Case ITYPE_DATE
                    ' la zone de date laisse la place au bouton du calendrier
                    Set c = AjCtrl(dsg, "Forms.TextBox.1", INomControle(ch(i)), _
                                   x, y + ICH_LBL_HAUT + 1, larg - 20, ICH_CTL_HAUT)
                    Zone c, False
                    Set c = AjCtrl(dsg, "Forms.Label.1", "lblCalendrier", _
                                   x + larg - 18, y + ICH_LBL_HAUT + 1, 18, ICH_CTL_HAUT)
                    Texte c, ChrW(9662), 9, True, COUL_BOUTON_TXT, MSF_TextAlignCenter
                    c.BackStyle = MSF_BackStyleOpaque
                    c.BackColor = COUL_MODIFIER
                    c.ControlTipText = "Ouvrir le calendrier"
                    Set c = dsg.Controls(INomControle(ch(i)))

                Case Else
                    Set c = AjCtrl(dsg, "Forms.TextBox.1", INomControle(ch(i)), _
                                   x, y + ICH_LBL_HAUT + 1, larg, ICH_CTL_HAUT)
                    Zone c, ch(i).Verrouille
            End Select
        End If

        c.ControlTipText = ch(i).Aide
        If Not ch(i).Verrouille Then
            c.TabIndex = ordre
            ordre = ordre + 1
        End If
    Next i

    ' libellé commun aux deux cases à cocher, qui n'en ont pas d'individuel
    Set c = AjCtrl(dsg, "Forms.Label.1", "lblI_Facturation", IGrilleX(3), IGrilleY(3), _
                   IG_BLOC, ICH_LBL_HAUT)
    Texte c, "FACTURATION", TAILLE_LIBELLE, True, COUL_TEXTE_DOUX, MSF_TextAlignLeft
End Sub

'------------------------------------------------------------------------------
' Barre de filtrage, entre la saisie et le tableau.
'------------------------------------------------------------------------------
Private Sub ConstruireFiltre(dsg As Object)
    Dim c As Object, y As Single

    Set c = AjCtrl(dsg, "Forms.Label.1", "lblCarteFiltreI", I_MARGE, IF_TOP, I_CARTE_LARG, IF_HAUT)
    CarteI c

    y = IF_TOP + 10

    Set c = AjCtrl(dsg, "Forms.Label.1", "lblFiltreTitreI", 30, y + 2, 62, 14)
    Texte c, "Filtrer sur", TAILLE_FILTRE, True, COUL_ENTETE_TXT, MSF_TextAlignLeft

    Set c = AjCtrl(dsg, "Forms.ComboBox.1", "cboChampFiltreI", 100, y, 124, ICH_CTL_HAUT)
    Liste c, False
    c.Style = MSF_StyleDropDownList
    c.ControlTipText = "Colonne du tableau sur laquelle porte le filtre"

    Set c = AjCtrl(dsg, "Forms.TextBox.1", "txtFiltreI", 232, y, 300, ICH_CTL_HAUT)
    Zone c, False
    c.ControlTipText = "Texte à rechercher (accents et majuscules sont ignorés)"

    Set c = AjCtrl(dsg, "Forms.Label.1", "lblResetFiltreI", 544, y + 2, 80, 14)
    Texte c, "Réinitialiser", TAILLE_FILTRE, False, COUL_LIEN, MSF_TextAlignLeft
    c.ControlTipText = "Effacer le filtre et réafficher toutes les interventions"

    Set c = AjCtrl(dsg, "Forms.Label.1", "lblCompteurI", 630, y + 2, 300, 14)
    Texte c, vbNullString, TAILLE_FILTRE, False, COUL_TEXTE_DOUX, MSF_TextAlignRight
End Sub

'------------------------------------------------------------------------------
' Fiche 4 : le tableau des enregistrements.
'
' Une ListBox MSForms ne sait pas afficher d'en-têtes autrement qu'en étant liée
' à une plage de cellules : ce sont donc des libellés posés au-dessus, alignés
' sur les mêmes largeurs — ce qui permet en prime de les rendre cliquables.
'------------------------------------------------------------------------------
Private Sub ConstruireFiche4(dsg As Object)
    Dim c As Object, larg As Variant, lib As Variant, i As Long
    Dim x As Single, gauche As Single, largeur As Single, colw As String, haut As Single

    gauche = I_MARGE + 1
    largeur = I_CARTE_LARG - 2

    Set c = AjCtrl(dsg, "Forms.Label.1", "lblCarte4", I_MARGE, IT_TOP, I_CARTE_LARG, IT_HAUT)
    CarteI c

    Set c = AjCtrl(dsg, "Forms.Label.1", "lblEnteteTableI", gauche, IT_TOP + 1, largeur, IT_ENTETE - 1)
    Fond c, COUL_ENTETE_TBL, COUL_ENTETE_TBL

    larg = ILargeursListe()
    lib = ILibellesListe()
    x = gauche + 3

    For i = LBound(larg) To UBound(larg)
        Set c = AjCtrl(dsg, "Forms.Label.1", "lblEntI_" & CStr(i + 1), _
                       x, IT_TOP + 5, CSng(larg(i)) - 3, 13)
        Texte c, CStr(lib(i)), TAILLE_ENTETE, True, COUL_ENTETE_TXT, MSF_TextAlignLeft
        c.ControlTipText = "Cliquez pour trier sur cette colonne"
        x = x + CSng(larg(i))
        colw = colw & IIf(Len(colw) > 0, ";", "") & CStr(larg(i)) & " pt"
    Next i

    haut = IT_TOP + IT_HAUT - 2 - (IT_TOP + IT_ENTETE + 1)
    Set c = AjCtrl(dsg, "Forms.ListBox.1", "lstInterv", _
                   gauche, IT_TOP + IT_ENTETE + 1, largeur, haut)
    With c
        .Font.Name = POLICE
        .Font.Size = TAILLE_LISTE
        .BackColor = COUL_CARTE
        .ForeColor = COUL_TEXTE
        .SpecialEffect = MSF_SpecialEffectFlat
        .BorderStyle = MSF_BorderStyleNone
        .ColumnCount = UBound(larg) - LBound(larg) + 1
        .ColumnWidths = colw
        .ColumnHeads = False
        .MultiSelect = MSF_MultiSelectSingle
        .ListStyle = MSF_ListStylePlain
        .BoundColumn = 1
        .TabIndex = 90
    End With
    On Error Resume Next
    c.IntegralHeight = False
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Barre de boutons : les quatre commandes de saisie à gauche, Facturer et Info
' détachés au milieu, Quitter renvoyé à droite pour ne pas être cliqué par
' mégarde.
'------------------------------------------------------------------------------
Private Sub ConstruireBoutonsInterv(dsg As Object)
    Dim c As Object, x As Single

    x = I_MARGE
    Set c = AjCtrl(dsg, "Forms.CommandButton.1", "btnIAjouter", x, IB_TOP, IB_LARG, IB_HAUT)
    Bouton_ c, "Ajouter", "A", COUL_AJOUTER, 91
    c.ControlTipText = "Créer une intervention à partir des champs saisis"

    x = x + IB_LARG + IB_GOUTTIERE
    Set c = AjCtrl(dsg, "Forms.CommandButton.1", "btnIModifier", x, IB_TOP, IB_LARG, IB_HAUT)
    Bouton_ c, "Modifier", "M", COUL_MODIFIER, 92
    c.ControlTipText = "Enregistrer les modifications sur l'intervention sélectionnée"

    x = x + IB_LARG + IB_GOUTTIERE
    Set c = AjCtrl(dsg, "Forms.CommandButton.1", "btnISupprimer", x, IB_TOP, IB_LARG, IB_HAUT)
    Bouton_ c, "Supprimer", "S", COUL_SUPPRIMER, 93
    c.ControlTipText = "Supprimer l'intervention sélectionnée"

    x = x + IB_LARG + IB_GOUTTIERE
    Set c = AjCtrl(dsg, "Forms.CommandButton.1", "btnIEffacer", x, IB_TOP, IB_LARG, IB_HAUT)
    Bouton_ c, "Effacer", "E", COUL_EFFACER, 94
    c.ControlTipText = "Vider les zones de saisie sans toucher au tableau"

    x = x + IB_LARG + IB_ECART_GROUPE
    Set c = AjCtrl(dsg, "Forms.CommandButton.1", "btnFacturer", x, IB_TOP, IB_LARG, IB_HAUT)
    Bouton_ c, "Facturer", "F", COUL_FACTURER, 95
    c.ControlTipText = "Ouvrir le formulaire de facturation"

    x = x + IB_LARG + IB_GOUTTIERE
    Set c = AjCtrl(dsg, "Forms.CommandButton.1", "btnInfo", x, IB_TOP, IB_LARG, IB_HAUT)
    Bouton_ c, "Info", "I", COUL_INFO, 96
    c.ControlTipText = "Ouvrir le formulaire d'informations"

    Set c = AjCtrl(dsg, "Forms.CommandButton.1", "btnIQuitter", _
                   I_LARGEUR - I_MARGE - IB_LARG, IB_TOP, IB_LARG, IB_HAUT)
    Bouton_ c, "Quitter", "Q", COUL_QUITTER, 97
    c.Cancel = True
    c.ControlTipText = "Fermer le formulaire (touche Échap)"
End Sub

'==============================================================================
' CALENDRIER
'==============================================================================
Private Function ConstruireCalendrier(vbProj As Object) As Long
    Dim vbComp As Object, dsg As Object, c As Object, i As Long
    Dim col As Long, lig As Long

    Set vbComp = PreparerForm(vbProj, NOM_FORM_CALENDRIER)
    Set dsg = vbComp.Designer

    PropForm vbComp, "Caption", "Date"
    PropForm vbComp, "Width", CAL_LARGEUR
    PropForm vbComp, "Height", CAL_HAUTEUR
    PropForm vbComp, "BackColor", COUL_CARTE
    PropForm vbComp, "SpecialEffect", MSF_SpecialEffectFlat
    PropForm vbComp, "BorderStyle", MSF_BorderStyleNone
    PropForm vbComp, "StartUpPosition", 0            ' position posée par le code
    PropForm vbComp, "ShowModal", True

    ' --- bandeau --------------------------------------------------------------
    Set c = AjCtrl(dsg, "Forms.Label.1", "lblCalBandeau", 0, 0, CAL_LARGEUR, CAL_BANDEAU)
    Fond c, COUL_BANDEAU, COUL_BANDEAU

    Set c = AjCtrl(dsg, "Forms.Label.1", "lblCalPrec", 8, 7, 22, 20)
    Texte c, ChrW(8249), 14, True, COUL_BANDEAU_SOUS, MSF_TextAlignCenter
    c.ControlTipText = "Mois précédent"

    Set c = AjCtrl(dsg, "Forms.Label.1", "lblCalMois", 34, 9, 156, 17)
    Texte c, vbNullString, 10, True, COUL_BANDEAU_TXT, MSF_TextAlignCenter

    Set c = AjCtrl(dsg, "Forms.Label.1", "lblCalSuiv", 194, 7, 22, 20)
    Texte c, ChrW(8250), 14, True, COUL_BANDEAU_SOUS, MSF_TextAlignCenter
    c.ControlTipText = "Mois suivant"

    ' --- en-têtes des jours ---------------------------------------------------
    For i = 0 To 6
        Set c = AjCtrl(dsg, "Forms.Label.1", "lblJS_" & CStr(i + 1), _
                       CAL_GRILLE_X + i * CAL_JOUR_LARG, 42, CAL_JOUR_LARG, 14)
        Texte c, vbNullString, TAILLE_LIBELLE, True, COUL_TEXTE_DOUX, MSF_TextAlignCenter
    Next i

    ' --- grille des 42 cases --------------------------------------------------
    For i = 1 To 42
        col = (i - 1) Mod 7
        lig = (i - 1) \ 7
        Set c = AjCtrl(dsg, "Forms.Label.1", "lblJ_" & CStr(i), _
                       CAL_GRILLE_X + col * CAL_JOUR_LARG, _
                       CAL_GRILLE_Y + lig * CAL_JOUR_HAUT, CAL_JOUR_LARG, CAL_JOUR_HAUT)
        Texte c, vbNullString, TAILLE_CHAMP, False, COUL_TEXTE, MSF_TextAlignCenter
    Next i

    ' --- pied -----------------------------------------------------------------
    Set c = AjCtrl(dsg, "Forms.Label.1", "lblCalAujourdhui", 10, 196, 100, 14)
    Texte c, "Aujourd'hui", TAILLE_FILTRE, False, COUL_LIEN, MSF_TextAlignLeft

    Set c = AjCtrl(dsg, "Forms.Label.1", "lblCalAnnuler", 140, 196, 74, 14)
    Texte c, "Annuler", TAILLE_FILTRE, False, COUL_TEXTE_DOUX, MSF_TextAlignRight

    ReculerFonds dsg, Array("lblCalBandeau")

    vbComp.CodeModule.AddFromString CodeCalendrier()
    ConstruireCalendrier = dsg.Controls.Count
End Function

'==============================================================================
' MISE EN FORME
'==============================================================================
Private Function AjCtrl(dsg As Object, ByVal progId As String, ByVal nom As String, _
                        ByVal gauche As Single, ByVal haut As Single, _
                        ByVal largeur As Single, ByVal hauteur As Single) As Object
    Dim c As Object
    Set c = dsg.Controls.Add(progId, nom, True)
    c.Left = gauche
    c.Top = haut
    c.Width = largeur
    c.Height = hauteur
    Set AjCtrl = c
End Function

'------------------------------------------------------------------------------
' Fond de carte : rectangle blanc à filet fin.
' SpecialEffect doit être posé avant BorderStyle, MSForms refusant une bordure
' simple tant que le contrôle est en relief. Même règle pour tout ce qui suit.
'------------------------------------------------------------------------------
Private Sub CarteI(c As Object)
    c.Caption = vbNullString
    c.BackStyle = MSF_BackStyleOpaque
    c.BackColor = COUL_CARTE
    c.SpecialEffect = MSF_SpecialEffectFlat
    c.BorderStyle = MSF_BorderStyleSingle
    c.BorderColor = COUL_BORDURE
End Sub

'------------------------------------------------------------------------------
' Rectangle de couleur unie : bandeau du calendrier, bande d'en-tête du tableau.
'------------------------------------------------------------------------------
Private Sub Fond(c As Object, ByVal fond As Long, ByVal bordure As Long)
    c.Caption = vbNullString
    c.BackStyle = MSF_BackStyleOpaque
    c.BackColor = fond
    c.SpecialEffect = MSF_SpecialEffectFlat
    c.BorderStyle = MSF_BorderStyleSingle
    c.BorderColor = bordure
End Sub

'------------------------------------------------------------------------------
' Habille un libellé : texte, taille, graisse, couleur, alignement.
' Le fond est transparent pour laisser voir la carte au-dessous.
'------------------------------------------------------------------------------
Private Sub Texte(c As Object, ByVal texte As String, ByVal taille As Single, _
                  ByVal gras As Boolean, ByVal couleur As Long, ByVal alignement As Long)
    c.Caption = texte
    c.BackStyle = MSF_BackStyleTransparent
    c.SpecialEffect = MSF_SpecialEffectFlat
    c.BorderStyle = MSF_BorderStyleNone
    c.ForeColor = couleur
    c.TextAlign = alignement
    c.WordWrap = False
    c.AutoSize = False
    c.Font.Name = POLICE
    c.Font.Size = taille
    c.Font.Bold = gras
End Sub

'------------------------------------------------------------------------------
' Zone de saisie. Locked plutôt qu'Enabled = False pour un champ géré par le
' programme : il reste lisible et son contenu copiable.
'------------------------------------------------------------------------------
Private Sub Zone(c As Object, ByVal verrouille As Boolean)
    c.Font.Name = POLICE
    c.Font.Size = TAILLE_CHAMP
    c.SpecialEffect = MSF_SpecialEffectFlat
    c.BorderStyle = MSF_BorderStyleSingle
    c.BorderColor = COUL_CHAMP_BORD
    c.TextAlign = MSF_TextAlignLeft
    c.EnterKeyBehavior = False
    If verrouille Then
        c.Locked = True
        c.TabStop = False
        c.BackColor = COUL_VERROU_FOND
        c.ForeColor = COUL_VERROU_TXT
    Else
        c.BackColor = COUL_CHAMP_FOND
        c.ForeColor = COUL_TEXTE
    End If
End Sub

'------------------------------------------------------------------------------
' Menu déroulant.
'   filtree : True pour une liste qui se filtre au fil de la frappe — elle doit
'             rester ouverte à la saisie libre, le contenu tapé pouvant ne
'             correspondre à aucun client existant.
'------------------------------------------------------------------------------
Private Sub Liste(c As Object, ByVal filtree As Boolean)
    c.Font.Name = POLICE
    c.Font.Size = TAILLE_CHAMP
    c.SpecialEffect = MSF_SpecialEffectFlat
    c.BorderStyle = MSF_BorderStyleSingle
    c.BorderColor = COUL_CHAMP_BORD
    c.BackColor = COUL_CHAMP_FOND
    c.ForeColor = COUL_TEXTE
    c.ListRows = 12
    c.Style = MSF_StyleDropDownCombo
    ' la complétion automatique gênerait le filtrage, qui réécrit la liste
    c.MatchEntry = IIf(filtree, 2, MSF_MatchEntryComplete)   ' 2 = fmMatchEntryNone
End Sub

'------------------------------------------------------------------------------
' Habille une case à cocher.
' Le relief de la case elle-même est conservé : à plat, elle deviendrait difficile
' à distinguer d'un simple libellé.
'------------------------------------------------------------------------------
Private Sub Case_(c As Object, ByVal libelle As String)
    c.Caption = libelle
    c.BackStyle = MSF_BackStyleTransparent
    c.ForeColor = COUL_TEXTE
    c.Font.Name = POLICE
    c.Font.Size = TAILLE_FILTRE
    c.Font.Bold = False
    c.WordWrap = False
End Sub

'------------------------------------------------------------------------------
' Habille un bouton d'action : couleur de fond, texte blanc, lettre de raccourci
' soulignée et rang de tabulation.
'------------------------------------------------------------------------------
Private Sub Bouton_(c As Object, ByVal libelle As String, ByVal raccourci As String, _
                    ByVal couleur As Long, ByVal ordre As Long)
    c.Caption = libelle
    c.Accelerator = raccourci
    c.BackColor = couleur
    c.ForeColor = COUL_BOUTON_TXT
    c.Font.Name = POLICE
    c.Font.Size = TAILLE_BOUTON
    c.Font.Bold = True
    c.TabIndex = ordre
End Sub

'------------------------------------------------------------------------------
' Renvoie les fonds derrière leur contenu.
' MSForms place chaque contrôle nouvellement créé au premier plan : les cartes,
' créées avant leur contenu, se retrouveraient devant lui.
'------------------------------------------------------------------------------
Private Sub ReculerFonds(dsg As Object, ByVal noms As Variant)
    Dim i As Long, c As Object
    For i = LBound(noms) To UBound(noms)
        On Error Resume Next
        Set c = dsg.Controls(CStr(noms(i)))
        If Not c Is Nothing Then c.ZOrder MSF_ZOrderBack
        Set c = Nothing
        On Error GoTo 0
    Next i
End Sub

'==============================================================================
' MODULE DE CODE DE UF_Interventions
'------------------------------------------------------------------------------
' Chaque procédure événementielle se contente d'appeler modInterv_Formulaire :
' le code métier n'est jamais dupliqué ici, et régénérer le formulaire ne peut
' rien écraser d'écrit à la main.
'
' Les gestionnaires des champs sont produits en parcourant le schéma : ajouter
' un champ crée automatiquement ses événements.
'==============================================================================
Private Function CodeFormulairePrincipal() As String
    Dim ch() As ChampInterv, i As Long, n As String, sortie As String
    Dim boutons As Variant, actions As Variant, larg As Variant

    mCode = vbNullString
    EnteteGenere NOM_FORM_INTERV, "modInterv_Formulaire"

    ProcI "UserForm_Initialize()", "Interv_Initialiser Me"
    ProcI "UserForm_Activate()", "Interv_Activer Me"
    ProcI "UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)", "Interv_Fermeture Me"
    ProcI "UserForm_MouseMove" & SIG_SOURIS, "Interv_Survol Me, " & Guill("")

    ' --- boutons d'action -----------------------------------------------------
    boutons = Array("btnIAjouter", "btnIModifier", "btnISupprimer", "btnIEffacer", _
                    "btnFacturer", "btnInfo", "btnIQuitter")
    actions = Array("Interv_Ajouter", "Interv_Modifier", "Interv_Supprimer", "Interv_Effacer", _
                    "Interv_Facturer", "Interv_Info", "Interv_Quitter")
    For i = LBound(boutons) To UBound(boutons)
        n = CStr(boutons(i))
        ProcI n & "_Click()", CStr(actions(i)) & " Me"
        ProcI n & "_MouseMove" & SIG_SOURIS, "Interv_Survol Me, " & Guill(n)
    Next i

    ' --- filtrage -------------------------------------------------------------
    ProcI "cboChampFiltreI_Change()", "Interv_AppliquerFiltre Me"
    ProcI "txtFiltreI_Change()", "Interv_AppliquerFiltre Me"
    ProcI "lblResetFiltreI_Click()", "Interv_ReinitialiserFiltre Me"
    ProcI "lblResetFiltreI_MouseMove" & SIG_SOURIS, "Interv_Survol Me, " & Guill("lblResetFiltreI")

    ' --- tableau --------------------------------------------------------------
    ProcI "lstInterv_Click()", "Interv_ChargerSelection Me"
    ProcI "lstInterv_DblClick(ByVal Cancel As MSForms.ReturnBoolean)", "Interv_ChargerSelection Me"

    larg = ILargeursListe()
    For i = LBound(larg) To UBound(larg)
        ProcI "lblEntI_" & CStr(i + 1) & "_Click()", "Interv_TrierColonne Me, " & CStr(i + 1)
    Next i

    ' --- sélecteur de date ----------------------------------------------------
    ProcI "lblCalendrier_Click()", "Interv_OuvrirCalendrier Me"

    ' --- zones de saisie ------------------------------------------------------
    ch = ObtenirChampsInterv()
    For i = LBound(ch) To UBound(ch)
        n = INomControle(ch(i))

        If ch(i).TypeCtrl = ITYPE_CASE Then
            ' cocher Forfait change le mode de calcul du chiffre d'affaires
            If StrComp(ch(i).Colonne, IC_FORFAIT, vbTextCompare) = 0 Then
                ProcI n & "_Click()", "Interv_MajCA Me"
            End If

        ElseIf Not ch(i).Verrouille Then
            ProcI n & "_Enter()", "Interv_FocusChamp Me, " & Guill(n) & ", True"

            ' la sortie de champ enchaîne parfois deux traitements
            sortie = "Interv_FocusChamp Me, " & Guill(n) & ", False"
            If ch(i).TypeCtrl = ITYPE_AUTO Then
                sortie = sortie & vbNewLine & "    Interv_ClientChoisi Me, " & Guill(ch(i).Colonne)
            ElseIf ch(i).Saisie = ISAISIE_HEURES Then
                sortie = sortie & vbNewLine & "    Interv_NormaliserHeures Me"
            End If
            ProcI n & "_Exit(ByVal Cancel As MSForms.ReturnBoolean)", sortie

            If ch(i).TypeCtrl = ITYPE_AUTO Then
                ProcI n & "_Change()", "Interv_FiltrerListe Me, " & Guill(ch(i).Colonne)
                ProcI n & "_Click()", "Interv_ClientChoisi Me, " & Guill(ch(i).Colonne)
            End If

            If ch(i).Saisie <> ISAISIE_LIBRE Then
                ProcI n & "_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)", _
                     "Interv_ToucheSaisie KeyAscii, " & CStr(ch(i).Saisie)
            End If

            ' le nombre de personnes entre dans le calcul du chiffre d'affaires
            If StrComp(ch(i).Colonne, IC_PERS, vbTextCompare) = 0 Then
                ProcI n & "_Change()", "Interv_MajCA Me"
            End If
        End If
    Next i

    CodeFormulairePrincipal = mCode
End Function

'==============================================================================
' MODULE DE CODE DE UF_Calendrier
'==============================================================================
Private Function CodeCalendrier() As String
    Dim i As Long

    mCode = vbNullString
    EnteteGenere NOM_FORM_CALENDRIER, "modInterv_Calendrier"

    ProcI "UserForm_Initialize()", "Cal_Initialiser Me"
    ProcI "UserForm_MouseMove" & SIG_SOURIS, _
         "Cal_SurvolLien Me, " & Guill("lblCalPrec") & ", False" & vbNewLine & _
         "    Cal_SurvolLien Me, " & Guill("lblCalSuiv") & ", False"

    ProcI "lblCalPrec_Click()", "Cal_MoisPrecedent Me"
    ProcI "lblCalPrec_MouseMove" & SIG_SOURIS, "Cal_SurvolLien Me, " & Guill("lblCalPrec") & ", True"
    ProcI "lblCalSuiv_Click()", "Cal_MoisSuivant Me"
    ProcI "lblCalSuiv_MouseMove" & SIG_SOURIS, "Cal_SurvolLien Me, " & Guill("lblCalSuiv") & ", True"

    For i = 1 To 42
        ProcI "lblJ_" & CStr(i) & "_Click()", "Cal_ChoisirJour Me, " & CStr(i)
    Next i

    ProcI "lblCalAujourdhui_Click()", "Cal_Aujourdhui Me"
    ProcI "lblCalAnnuler_Click()", "Cal_Annuler Me"

    CodeCalendrier = mCode
End Function

'==============================================================================
' Écriture du code généré
'==============================================================================
Private Sub EnteteGenere(ByVal nomForm As String, ByVal module As String)
    AjLigne "Option Explicit"
    AjLigne "'=============================================================================="
    AjLigne "' Module de code du formulaire " & nomForm
    AjLigne "'------------------------------------------------------------------------------"
    AjLigne "' GÉNÉRÉ AUTOMATIQUEMENT par modInterv_Generateur.GenererFormulaireInterventions."
    AjLigne "' Ne rien écrire ici : toute la logique se trouve dans " & module & ","
    AjLigne "' et ce module est écrasé à chaque régénération du formulaire."
    AjLigne "'=============================================================================="
    AjLigne ""
End Sub

'------------------------------------------------------------------------------
' Écrit une procédure événementielle complète.
'   entete : signature, sans le mot-clef Private Sub
'   corps  : le ou les appels à y placer
'
' Trois lignes distinctes, jamais une seule ligne à deux-points : VBA n'accepte
' pas qu'une déclaration de procédure et son End Sub partagent une ligne.
'------------------------------------------------------------------------------
Private Sub ProcI(ByVal entete As String, ByVal corps As String)
    AjLigne "Private Sub " & entete
    AjLigne "    " & corps
    AjLigne "End Sub"
    AjLigne ""
End Sub

'------------------------------------------------------------------------------
' Ajoute une ligne au code en cours d'écriture.
'------------------------------------------------------------------------------
Private Sub AjLigne(ByVal ligne As String)
    mCode = mCode & ligne & vbNewLine
End Sub

'------------------------------------------------------------------------------
' Entoure un texte de guillemets doubles : évite de les doubler dans le code
' générateur, où ils deviendraient vite illisibles.
'------------------------------------------------------------------------------
Private Function Guill(ByVal s As String) As String
    Guill = Chr$(34) & s & Chr$(34)
End Function
