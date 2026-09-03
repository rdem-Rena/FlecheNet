Attribute VB_Name = "modStat_Generateur"
Option Explicit
'==============================================================================
' modStat_Generateur
'------------------------------------------------------------------------------
' Génère de toutes pièces le formulaire UF_Statistiques, appelé par le bouton
' « Info » du formulaire des interventions.
'
'   >>> Procédure à lancer : GenererFormulaireStatistiques
'
' PRÉREQUIS : « Accès approuvé au modèle d'objet du projet VBA » — voir
' modInterv_Generateur pour le chemin complet dans les options d'Excel.
'
' IL NE FAUT PAS SUPPRIMER LE FORMULAIRE avant de régénérer : la génération le
' vide et le reconstruit sur place, et VBA ne rend le nom d'un composant
' supprimé qu'au prochain chargement du classeur.
'
' PreparerForm et PoserPolice viennent de modInterv_Generateur : la première
' sait reconnaître un nom encore retenu, la seconde connaît l'ordre des
' assignations sans lequel la taille emporte la graisse.
'==============================================================================

Private mCodeS As String

'==============================================================================
' POINT D'ENTRÉE
'==============================================================================
Public Sub GenererFormulaireStatistiques()
    Dim vbProj As Object, nbCtrl As Long

    If TableInterventions() Is Nothing Then
        MsgBox "Le tableau " & NOM_TABLE_INTERVENTIONS & " est introuvable dans ce classeur." & _
               vbCrLf & "Le formulaire des statistiques ne peut pas être généré.", _
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
    nbCtrl = ConstruireStat(vbProj)

    MsgBox "Le formulaire " & NOM_FORM_STAT & " a été généré." & vbCrLf & vbCrLf & _
           nbCtrl & " contrôles." & vbCrLf & vbCrLf & _
           "Il s'ouvre par le bouton " & Chr$(34) & "Info" & Chr$(34) & " du formulaire " & _
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
Private Function ConstruireStat(vbProj As Object) As Long
    Dim vbComp As Object, dsg As Object

    Set vbComp = PreparerForm(vbProj, NOM_FORM_STAT)
    Set dsg = vbComp.Designer

    PropS vbComp, "Caption", "Statistiques"
    PropS vbComp, "Width", ST_LARGEUR
    PropS vbComp, "Height", ST_HAUTEUR + ST_RESERVE_TITRE
    PropS vbComp, "BackColor", COUL_FOND
    PropS vbComp, "SpecialEffect", MSF_SpecialEffectFlat
    PropS vbComp, "StartUpPosition", 1
    PropS vbComp, "ShowModal", True
    PoserPolice dsg, POLICE, 8, False

    ConstruireSZone1 dsg
    ConstruireSZone2 dsg
    ConstruireSZone3 dsg
    ConstruireSZone4 dsg
    ConstruireSZone5 dsg
    ConstruireSBoutons dsg

    vbComp.CodeModule.AddFromString CodeStat()
    ConstruireStat = dsg.Controls.Count
End Function

'------------------------------------------------------------------------------
' Zone 1 : l'intitulé — titre à gauche, année à droite.
'------------------------------------------------------------------------------
Private Sub ConstruireSZone1(dsg As Object)
    Dim c As Object, zone As Object

    Set zone = AjS(dsg, "Forms.Frame.1", "fraSCarte1", _
                   ST_MARGE, ST_Z1_TOP, ST_CARTE_LARG, ST_Z1_HAUT)
    CadreS zone

    Set c = AjS(zone, "Forms.Label.1", "lblSTitre", 14, 12, 620, 21)
    TexteS c, "Statistiques", ZS1Titre(), MSF_TextAlignLeft

    Set c = AjS(zone, "Forms.Label.1", "lblSAnnee", ST_CARTE_LARG - 244, 8, 230, 30)
    TexteS c, vbNullString, ZS1Annee(), MSF_TextAlignRight
End Sub

'------------------------------------------------------------------------------
' Zone 2 : les cinq tuiles, sur une seule ligne.
'
' Même construction que celles du formulaire des interventions : une image de
' fond porte les coins arrondis, deux libellés à fond blanc portent le texte.
' Si l'image manque, un aplat blanc prend sa place et la tuile reste lisible.
'------------------------------------------------------------------------------
Private Sub ConstruireSZone2(dsg As Object)
    Dim c As Object, zone As Object, tu As Variant, i As Long, x As Single, y As Single
    Dim yCap As Single

    Set zone = AjS(dsg, "Forms.Frame.1", "fraSCarte2", _
                   ST_MARGE, ST_Z2_TOP, ST_CARTE_LARG, ST_Z2_HAUT)
    CadreS zone

    tu = STuiles()
    y = AuPixel((ST_Z2_HAUT - F2_TUILE_HAUT) / 2)

    For i = LBound(tu) To UBound(tu)
        x = AuPixel(STuilesX() + (i - LBound(tu)) * (F2_TUILE_LARG + F2_TUILE_GX))

        Set c = AjS(zone, "Forms.Image.1", "imgSTuile_" & CStr(i + 1), x, y, _
                    F2_TUILE_LARG, F2_TUILE_HAUT)
        With c
            .SpecialEffect = MSF_SpecialEffectFlat
            .BorderStyle = MSF_BorderStyleNone
            .BorderColor = COUL_CARTE
            .BackStyle = MSF_BackStyleOpaque
            .BackColor = COUL_CARTE
            .PictureSizeMode = 1        ' fmPictureSizeModeStretch
        End With

        yCap = y + (F2_TUILE_HAUT - (F2_TUILE_CAP_HAUT + F2_TUILE_ECART + _
                                     F2_TUILE_VAL_HAUT)) / 2

        Set c = AjS(zone, "Forms.Label.1", "lblSTuileCap_" & CStr(i + 1), _
                    x + F2_TUILE_INSET, yCap, _
                    F2_TUILE_LARG - 2 * F2_TUILE_INSET, F2_TUILE_CAP_HAUT)
        TexteS c, UCase$(CStr(tu(i)(0))), ZSTuileCap(), MSF_TextAlignCenter
        FondTuileS c

        Set c = AjS(zone, "Forms.Label.1", "lblSTuileVal_" & CStr(i + 1), _
                    x + F2_TUILE_INSET, yCap + F2_TUILE_CAP_HAUT + F2_TUILE_ECART, _
                    F2_TUILE_LARG - 2 * F2_TUILE_INSET, F2_TUILE_VAL_HAUT)
        TexteS c, vbNullString, ZSTuileVal(), MSF_TextAlignCenter
        FondTuileS c
    Next i
End Sub

'------------------------------------------------------------------------------
' Zone 3 : le graphique mensuel, et la barre d'objectif à sa droite.
'
' Le décor est posé ici, les barres sont placées à l'exécution par
' modStat_Formulaire : leur hauteur dépend des données filtrées. Les deux
' repartent de ST_GR_X et ST_GR_Y, la même origine — un tracé qui repartirait
' d'ailleurs se retrouverait décalé de son propre décor.
'------------------------------------------------------------------------------
Private Sub ConstruireSZone3(dsg As Object)
    Dim c As Object, zone As Object, i As Long, y As Single, largeur As Single

    Set zone = AjS(dsg, "Forms.Frame.1", "fraSCarte3", _
                   ST_MARGE, ST_Z3_TOP, ST_CARTE_LARG, ST_Z3_HAUT)
    CadreS zone

    largeur = F2_GRAPH_LARG - GR_MARGE_G

    Set c = AjS(zone, "Forms.Label.1", "lblSGrLegende", ST_GR_X, ST_GR_Y, _
                F2_GRAPH_LARG, GR_LEGENDE_HAUT)
    TexteS c, "Chiffre d'affaires par mois", ZSGraphTitre(), MSF_TextAlignLeft

    Set c = AjS(zone, "Forms.Label.1", "lblSGrMax", ST_GR_X, _
                ST_GR_Y + GR_TRACE_TOP - 5, GR_MARGE_G - 8, 11)
    TexteS c, vbNullString, ZSGraphAxe(), MSF_TextAlignRight

    ' trois repères : le sommet de l'échelle, la moitié, la ligne de base
    For i = 1 To 3
        y = ST_GR_Y + GR_TRACE_TOP + (i - 1) * (GR_TRACE_HAUT / 2)
        Set c = AjS(zone, "Forms.Label.1", "lblSGrGrille_" & CStr(i), _
                    ST_GR_X + GR_MARGE_G, y, largeur, 1)
        FondS c, IIf(i = 3, COUL_GR_BASE, COUL_GR_GRILLE), _
                 IIf(i = 3, COUL_GR_BASE, COUL_GR_GRILLE)
    Next i

    For i = 1 To GR_NB_MOIS
        Set c = AjS(zone, "Forms.Label.1", "lblSGrBarre_" & CStr(i), _
                    ST_GR_X + GR_MARGE_G, ST_GR_Y + GR_TRACE_TOP, GR_BARRE_LARG, 1)
        FondS c, COUL_GR_BARRE, COUL_GR_BARRE

        Set c = AjS(zone, "Forms.Label.1", "lblSGrMois_" & CStr(i), _
                    ST_GR_X + GR_MARGE_G, _
                    ST_GR_Y + GR_TRACE_TOP + GR_TRACE_HAUT + 3, _
                    GR_BARRE_LARG, GR_MOIS_HAUT)
        TexteS c, vbNullString, ZSGraphAxe(), MSF_TextAlignCenter
    Next i

    ' --- la barre d'objectif, verticale ---------------------------------------
    ' La cuve est posée ici une fois pour toutes ; le remplissage est un second
    ' libellé dont le haut et la hauteur bougent à l'exécution. Il se remplit du
    ' BAS vers le haut, comme un thermomètre : c'est son Top qui descend.
    Set c = AjS(zone, "Forms.Label.1", "lblSObjTitre", ST_OBJ_X, ST_GR_Y, _
                ST_OBJ_TXT_L, GR_LEGENDE_HAUT)
    TexteS c, "Objectif annuel", ZSGraphTitre(), MSF_TextAlignLeft

    Set c = AjS(zone, "Forms.Label.1", "lblSObjCuve", ST_OBJ_X, ST_OBJ_TOP, _
                ST_OBJ_LARG, ST_OBJ_HAUT)
    FondS c, COUL_GR_GRILLE, COUL_GR_BASE

    Set c = AjS(zone, "Forms.Label.1", "lblSObjPlein", ST_OBJ_X + 1, _
                ST_OBJ_TOP + ST_OBJ_HAUT - 1, ST_OBJ_LARG - 2, 1)
    FondS c, COUL_GR_BARRE, COUL_GR_BARRE

    Set c = AjS(zone, "Forms.Label.1", "lblSObjPct", ST_OBJ_X + ST_OBJ_LARG + 12, _
                ST_OBJ_TOP + 4, ST_OBJ_TXT_L, 22)
    TexteS c, vbNullString, ZSTuileVal(), MSF_TextAlignLeft

    Set c = AjS(zone, "Forms.Label.1", "lblSObjDetail", ST_OBJ_X + ST_OBJ_LARG + 12, _
                ST_OBJ_TOP + 28, ST_OBJ_TXT_L, 26)
    TexteS c, vbNullString, ZSGraphAxe(), MSF_TextAlignLeft
    c.WordWrap = True
End Sub

'------------------------------------------------------------------------------
' Zone 4 : les sept filtres, sur une seule ligne.
'
' TVA, Forfait et Facturé sont des cases à TROIS états : grisée elle ne filtre
' pas, cochée elle ne garde que les lignes vraies, décochée que les fausses.
' Une case à deux états ne saurait pas exprimer « seulement les non facturées »,
' qui est justement le cas utile.
'------------------------------------------------------------------------------
Private Sub ConstruireSZone4(dsg As Object)
    Dim c As Object, zone As Object, y As Single

    Set zone = AjS(dsg, "Forms.Frame.1", "fraSCarte4", _
                   ST_MARGE, ST_Z4_TOP, ST_CARTE_LARG, ST_Z4_HAUT)
    CadreS zone
    y = AuPixel((ST_Z4_HAUT - ST_CTL_HAUT) / 2)

    Set c = AjS(zone, "Forms.Label.1", "lblSFMoisCap", ST_F_MOIS_LBL, y + 2, 32, 14)
    TexteS c, "Mois", ZSFiltre(), MSF_TextAlignLeft
    Set c = AjS(zone, "Forms.ComboBox.1", "cboSMois", ST_F_MOIS, y, ST_F_MOIS_L, ST_CTL_HAUT)
    ListeS c

    Set c = AjS(zone, "Forms.Label.1", "lblSFEntCap", ST_F_ENT_LBL, y + 2, 62, 14)
    TexteS c, "Entreprise", ZSFiltre(), MSF_TextAlignLeft
    Set c = AjS(zone, "Forms.TextBox.1", "txtSEntreprise", ST_F_ENT, y, ST_F_ENT_L, ST_CTL_HAUT)
    ZoneS c
    c.ControlTipText = "Ne garder que les entreprises contenant ce texte"

    Set c = AjS(zone, "Forms.Label.1", "lblSFNomCap", ST_F_NOM_LBL, y + 2, 34, 14)
    TexteS c, "Nom", ZSFiltre(), MSF_TextAlignLeft
    Set c = AjS(zone, "Forms.TextBox.1", "txtSNom", ST_F_NOM, y, ST_F_NOM_L, ST_CTL_HAUT)
    ZoneS c
    c.ControlTipText = "Ne garder que les noms contenant ce texte"

    Set c = AjS(zone, "Forms.CheckBox.1", "chkSTVA", ST_F_TVA, y + 1, ST_F_TVA_L, ST_CTL_HAUT)
    CaseS c, "TVA"
    Set c = AjS(zone, "Forms.CheckBox.1", "chkSForfait", ST_F_FORF, y + 1, ST_F_FORF_L, ST_CTL_HAUT)
    CaseS c, "Forfait"
    Set c = AjS(zone, "Forms.CheckBox.1", "chkSFacture", ST_F_FACT, y + 1, ST_F_FACT_L, ST_CTL_HAUT)
    CaseS c, "Facturé"
    c.ControlTipText = "Grisée : toutes. Cochée : les facturées. Décochée : celles qui " & _
                       "n'ont pas de numéro."

    Set c = AjS(zone, "Forms.Label.1", "lblSFNumCap", ST_F_NUM_LBL, y + 2, 74, 14)
    TexteS c, "N" & ChrW(176) & " facture", ZSFiltre(), MSF_TextAlignRight
    Set c = AjS(zone, "Forms.TextBox.1", "txtSNoFacture", ST_F_NUM, y, ST_F_NUM_L, ST_CTL_HAUT)
    ZoneS c
    c.ControlTipText = "Ne garder que les numéros de facture contenant ce texte"
End Sub

'------------------------------------------------------------------------------
' Zone 5 : le tableau des travaux retenus par les filtres.
'------------------------------------------------------------------------------
Private Sub ConstruireSZone5(dsg As Object)
    Dim c As Object, zone As Object, larg As Variant, lib As Variant, ali As Variant
    Dim i As Long, r As Long, nbCol As Long
    Dim x As Single, y As Single, gauche As Single, largeur As Single

    Set zone = AjS(dsg, "Forms.Frame.1", "fraSCarte5", _
                   ST_MARGE, ST_Z5_TOP, ST_CARTE_LARG, ST_Z5_HAUT)
    CadreS zone

    Set c = AjS(zone, "Forms.Label.1", "lblSSection", AuPixel(IGR_PAD_X + 1), AuPixel(4), _
                SGrilleLargeur() - 2 * IGR_PAD_X, 13)
    TexteS c, UCase$("Liste travaux " & ChrW(8211) & " Clients"), ZSSection(), MSF_TextAlignLeft

    larg = SLargeurs()
    lib = SLibelles()
    ali = SAlignements()
    nbCol = UBound(larg) - LBound(larg) + 1
    gauche = AuPixel(1)
    largeur = SGrilleLargeur()

    Set c = AjS(zone, "Forms.Label.1", "lblSEntFond", gauche, SGrilleEnteteY(), _
                largeur, AuPixel(ST_ENTETE_HAUT))
    FondS c, COUL_ENTETE_TBL, COUL_ENTETE_TBL

    x = gauche + IGR_PAD_X
    For i = 0 To nbCol - 1
        Set c = AjS(zone, "Forms.Label.1", "lblSEnt_" & CStr(i + 1), _
                    AuPixel(x), AuPixel(SGrilleEnteteY() + 4), _
                    AuPixel(CSng(larg(i)) - 2 * IGR_PAD_X), AuPixel(IGR_LIGNE_H))
        TexteS c, CStr(lib(i)), ZSEntete(), CLng(ali(i))
        x = x + CSng(larg(i))
    Next i

    ' Chaque ligne reçoit d'abord une BANDE de fond sur toute la largeur, puis
    ' ses cases, transparentes, posées dessus : c'est la bande qui porte la
    ' couleur, sans quoi les points entre deux cases laisseraient voir le blanc
    ' de la carte et la ligne choisie paraîtrait rayée.
    For r = 1 To ST_Z5_LIGNES
        y = SGrilleLignesY() + (r - 1) * IGR_LIGNE_H

        Set c = AjS(zone, "Forms.Label.1", "lblSL_" & CStr(r), _
                    gauche, AuPixel(y), AuPixel(largeur - IGR_BARRE_L), AuPixel(IGR_LIGNE_H))
        With c
            .Caption = vbNullString
            .SpecialEffect = MSF_SpecialEffectFlat
            .BorderStyle = MSF_BorderStyleNone
            .BackStyle = MSF_BackStyleOpaque
            .BackColor = COUL_CARTE
        End With
        PoserStyleS c, ZSCase()

        x = gauche + IGR_PAD_X
        For i = 0 To nbCol - 1
            Set c = AjS(zone, "Forms.Label.1", _
                        "lblS_" & CStr(r) & "_" & CStr(i + 1), _
                        AuPixel(x), AuPixel(y), _
                        AuPixel(CSng(larg(i)) - 2 * IGR_PAD_X), AuPixel(IGR_LIGNE_H))
            TexteS c, vbNullString, ZSCase(), CLng(ali(i))
            x = x + CSng(larg(i))
        Next i
    Next r

    Set c = AjS(zone, "Forms.ScrollBar.1", "sbStat", _
                AuPixel(gauche + largeur - IGR_BARRE_L), AuPixel(SGrilleLignesY()), _
                AuPixel(IGR_BARRE_L), AuPixel(ST_Z5_LIGNES * IGR_LIGNE_H))
    With c
        .Min = 0
        .Max = 0
        .SmallChange = 1
        .LargeChange = ST_Z5_LIGNES
        .Value = 0
        .ProportionalThumb = False
    End With
End Sub

'------------------------------------------------------------------------------
' La rangée de boutons, sur le formulaire lui-même.
'------------------------------------------------------------------------------
Private Sub ConstruireSBoutons(dsg As Object)
    Dim c As Object

    Set c = AjS(dsg, "Forms.Label.1", "lblSCompteur", ST_MARGE, ST_BT_TOP + 8, 320, 14)
    TexteS c, vbNullString, ZSFiltre(), MSF_TextAlignLeft

    Set c = AjS(dsg, "Forms.CommandButton.1", "btnSQuitter", _
                ST_LARGEUR - ST_MARGE - ST_BT_LARG, ST_BT_TOP, ST_BT_LARG, ST_BT_HAUT)
    BoutonS c, "Quitter", "Q", COUL_QUITTER, 50
End Sub

'==============================================================================
' HABILLAGE
'------------------------------------------------------------------------------
' Ces procédures ne posent que des propriétés, jamais de mise en page. La
' police passe toujours par PoserPolice, de modInterv_Generateur : l'ordre de
' ses assignations est la seule chose qui empêche la taille d'emporter la
' graisse.
'==============================================================================

' Le premier argument est un CONTENEUR : un cadre expose la même collection
' Controls, et les coordonnées comptent alors depuis son coin.
Private Function AjS(conteneur As Object, ByVal progId As String, ByVal nom As String, _
                     ByVal gauche As Single, ByVal haut As Single, _
                     ByVal largeur As Single, ByVal hauteur As Single) As Object
    Dim c As Object

    Set c = conteneur.Controls.Add(progId, nom, True)
    c.Left = gauche
    c.Top = haut
    c.Width = largeur
    c.Height = hauteur
    Set AjS = c
End Function

' SpecialEffect avant BorderStyle : MSForms refuse une bordure simple tant que
' le contrôle est en relief, et un cadre est gravé par défaut.
Private Sub CadreS(c As Object)
    c.Caption = vbNullString
    c.BackColor = COUL_CARTE
    c.SpecialEffect = MSF_SpecialEffectFlat
    c.BorderStyle = MSF_BorderStyleSingle
    c.BorderColor = COUL_BORDURE
    c.ScrollBars = MSF_ScrollBarsNone
End Sub

Private Sub FondS(c As Object, ByVal fond As Long, ByVal bordure As Long)
    c.Caption = vbNullString
    c.BackStyle = MSF_BackStyleOpaque
    c.BackColor = fond
    c.SpecialEffect = MSF_SpecialEffectFlat
    c.BorderStyle = MSF_BorderStyleSingle
    c.BorderColor = bordure
End Sub

' Fond blanc opaque et sans filet : le libellé d'une tuile se pose SUR l'image
' de fond et doit la masquer là où il écrit.
Private Sub FondTuileS(c As Object)
    c.BackStyle = MSF_BackStyleOpaque
    c.BackColor = COUL_CARTE
    c.BorderStyle = MSF_BorderStyleNone
    c.BorderColor = COUL_CARTE
End Sub

Private Sub TexteS(c As Object, ByVal texte As String, ByRef st As StyleTexte, _
                   ByVal alignement As Long)
    c.Caption = texte
    c.BackStyle = MSF_BackStyleTransparent
    c.SpecialEffect = MSF_SpecialEffectFlat
    c.BorderStyle = MSF_BorderStyleNone
    c.TextAlign = alignement
    c.WordWrap = False
    c.AutoSize = False
    PoserStyleS c, st
End Sub

Private Sub PoserStyleS(c As Object, ByRef st As StyleTexte)
    c.ForeColor = st.Couleur
    PoserPolice c, st.Police, st.Taille, st.Gras
End Sub

Private Sub ZoneS(c As Object)
    PoserPolice c, POLICE, 9, False
    c.SpecialEffect = MSF_SpecialEffectFlat
    c.BorderStyle = MSF_BorderStyleSingle
    c.BorderColor = COUL_CHAMP_BORD
    c.BackColor = COUL_CHAMP_FOND
    c.ForeColor = COUL_TEXTE
    c.TextAlign = MSF_TextAlignLeft
    c.EnterKeyBehavior = False
End Sub

Private Sub ListeS(c As Object)
    PoserPolice c, POLICE, 9, False
    c.SpecialEffect = MSF_SpecialEffectFlat
    c.BorderStyle = MSF_BorderStyleSingle
    c.BorderColor = COUL_CHAMP_BORD
    c.BackColor = COUL_CHAMP_FOND
    c.ForeColor = COUL_TEXTE
    c.ListRows = 13
    c.Style = MSF_StyleDropDownList
End Sub

' TripleState : la case a un troisième état, grisé, qui ne filtre rien. Value
' vaut alors Null, ce que modStat_Formulaire traduit en TRI_INDIFFERENT.
Private Sub CaseS(c As Object, ByVal libelle As String)
    PoserPolice c, POLICE, 9, False
    c.Caption = libelle
    c.BackStyle = MSF_BackStyleTransparent
    c.ForeColor = COUL_TEXTE
    c.WordWrap = False
    c.TripleState = True
    c.Value = Null
End Sub

Private Sub BoutonS(c As Object, ByVal libelle As String, ByVal raccourci As String, _
                    ByVal couleur As Long, ByVal ordre As Long)
    PoserPolice c, POLICE, 9.75, True
    c.Caption = libelle
    c.Accelerator = raccourci
    c.BackColor = couleur
    c.ForeColor = COUL_BOUTON_TXT
    c.TabIndex = ordre
End Sub

Private Sub PropS(vbComp As Object, ByVal nom As String, ByVal valeur As Variant)
    On Error Resume Next
    vbComp.Properties(nom) = valeur
    On Error GoTo 0
End Sub

'==============================================================================
' MODULE DE CODE DE UF_Statistiques
'------------------------------------------------------------------------------
' Chaque procédure événementielle appelle modStat_Formulaire : régénérer le
' formulaire n'écrase donc jamais de comportement.
'==============================================================================
Private Function CodeStat() As String
    Dim larg As Variant, r As Long, i As Long

    mCodeS = vbNullString
    LigS "'=============================================================================="
    LigS "' " & NOM_FORM_STAT & " - MODULE GÉNÉRÉ"
    LigS "'------------------------------------------------------------------------------"
    LigS "' Produit par modStat_Generateur. Toute modification faite ici sera perdue à la"
    LigS "' prochaine génération : le comportement s'écrit dans modStat_Formulaire."
    LigS "'=============================================================================="
    LigS "Option Explicit"
    LigS ""

    ProcS "UserForm_Initialize()", "Stat_Initialiser Me"
    ProcS "UserForm_Activate()", "Stat_Activer Me"

    ' la grille est faite de libellés : c'est la case cliquée qui reçoit
    ' l'événement, et elle ne sait dire que sa ligne
    larg = SLargeurs()
    For r = 1 To ST_Z5_LIGNES
        For i = LBound(larg) To UBound(larg)
            ProcS "lblS_" & CStr(r) & "_" & CStr(i + 1) & "_Click()", _
                  "Stat_GrilleClic Me, " & CStr(r)
        Next i
        ProcS "lblSL_" & CStr(r) & "_Click()", "Stat_GrilleClic Me, " & CStr(r)
    Next r
    ProcS "sbStat_Change()", "Stat_Defiler Me"

    ' les sept filtres rejouent tous le même recalcul
    ProcS "cboSMois_Change()", "Stat_Rafraichir Me"
    ProcS "txtSEntreprise_Change()", "Stat_Rafraichir Me"
    ProcS "txtSNom_Change()", "Stat_Rafraichir Me"
    ProcS "chkSTVA_Click()", "Stat_Rafraichir Me"
    ProcS "chkSForfait_Click()", "Stat_Rafraichir Me"
    ProcS "chkSFacture_Click()", "Stat_Rafraichir Me"
    ProcS "txtSNoFacture_Change()", "Stat_Rafraichir Me"

    ProcS "btnSQuitter_Click()", "Stat_Quitter Me"

    CodeStat = mCodeS
End Function

Private Sub ProcS(ByVal entete As String, ByVal corps As String)
    LigS "Private Sub " & entete
    LigS "    " & corps
    LigS "End Sub"
    LigS ""
End Sub

Private Sub LigS(ByVal ligne As String)
    mCodeS = mCodeS & ligne & vbNewLine
End Sub
