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
' prochain chargement du classeur, et l'affectation du nom échouerait.
'
' D'où : IL NE FAUT PAS SUPPRIMER LE FORMULAIRE avant de régénérer, c'est
' inutile et le nom reste pris jusqu'à la fin de la session. Si c'est déjà
' fait, enregistrez, fermez puis rouvrez le classeur avant de relancer la
' génération — le générateur le dit maintenant de lui-même.
'==============================================================================

Private Const CT_MSFORM As Long = 3         ' vbext_ct_MSForm

' Levée quand le nom du formulaire est encore retenu par un composant supprimé
' à la main pendant la session : ce n'est pas une panne, c'est une manoeuvre à
' faire, et le message doit le dire au lieu d'afficher une erreur VBA brute.
Private Const ERR_NOM_OCCUPE As Long = vbObjectError + 513
Private Const SIG_SOURIS As String = "(ByVal Button As Integer, ByVal Shift As Integer, " & _
                                     "ByVal X As Single, ByVal Y As Single)"

Private mCode As String

' L'ÉTAPE EN COURS, pour que le message d'erreur dise OÙ la génération a
' échoué. Une erreur MSForms ne nomme jamais le contrôle fautif : sans cela,
' « propriété non gérée par cet objet » laisse chercher dans mille lignes.
Private mEtape As String

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
    ' Le nom retenu par un formulaire supprimé à la main n'est pas une panne :
    ' la marche à suivre est dans le message, inutile de l'habiller en erreur.
    If Err.Number = ERR_NOM_OCCUPE Then
        MsgBox Err.Description, vbExclamation, "Génération du formulaire"
    Else
        MsgBox "La génération a échoué :" & vbCrLf & vbCrLf & _
               Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
               IIf(Len(mEtape) > 0, "Étape : " & mEtape & vbCrLf & vbCrLf, "") & _
               "Si le problème persiste : fermez puis rouvrez le classeur, lancez " & _
               "NettoyerFormulairesOrphelins, et relancez la génération.", _
               vbCritical, "Génération du formulaire"
    End If
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
' PUBLIQUE : le générateur de la facturation s'en sert aussi. Elle porte la
' détection du nom encore retenu par un formulaire supprimé à la main, qu'on ne
' veut pas voir réécrite ailleurs.
Public Function PreparerForm(vbProj As Object, ByVal nom As String) As Object
    Dim vbComp As Object, dsg As Object, i As Long, occupe As Boolean

    On Error Resume Next
    Set vbComp = vbProj.VBComponents(nom)
    On Error GoTo 0

    If vbComp Is Nothing Then
        Set vbComp = vbProj.VBComponents.Add(CT_MSFORM)

        ' LE NOM PEUT ÊTRE ENCORE PRIS. VBA ne rend le nom d'un composant
        ' supprimé qu'au prochain chargement du classeur : un formulaire
        ' effacé à la main dans le VBE retient donc le sien jusque-là, et
        ' l'affectation ci-dessous échoue — c'est ce qui produisait une
        ' erreur 75 incompréhensible.
        '
        ' Le formulaire vide qu'on vient d'ajouter est alors retiré : sans
        ' cela chaque tentative laisserait un UserForm1, UserForm2… derrière
        ' elle. Le nom est relu après coup, car selon les versions d'Excel
        ' l'affectation échoue tantôt bruyamment, tantôt en silence.
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

    ' VIDAGE. Un parcours par index à l'envers ne convient plus depuis qu'une
    ' zone peut être un CADRE : ses enfants partent avec lui, la collection
    ' rétrécit de plusieurs éléments d'un coup, et les index suivants ne
    ' désignent plus ce qu'on croit. On retire donc toujours le PREMIER, tant
    ' qu'il en reste, et on s'arrête dès qu'un tour ne retire plus rien —
    ' sinon un contrôle impossible à supprimer ferait tourner sans fin.
    Set dsg = vbComp.Designer
    On Error Resume Next
    Do While dsg.Controls.Count > 0
        i = dsg.Controls.Count
        dsg.Controls.Remove dsg.Controls(0).Name
        If dsg.Controls.Count >= i Then Exit Do
    Loop
    On Error GoTo 0
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
    PoliceParDefaut dsg

    mEtape = "fiche 1 (intitulé)"
    ConstruireFiche1 dsg
    mEtape = "fiche 2 (statistiques)"
    ConstruireFiche2 dsg
    mEtape = "fiche 3 (saisie)"
    ConstruireFiche3 dsg
    mEtape = "barre de filtrage"
    ConstruireFiltre dsg
    mEtape = "fiche 4 (tableau)"
    ConstruireFiche4 dsg
    mEtape = "boutons"
    ConstruireBoutonsInterv dsg
    mEtape = "ordre de plan"
    ' Un cadre n'a rien à faire ici : ses enfants sont devant lui par
    ' construction. Les deux premières cartes y figurent tout de même, sous le
    ' nom que leur donne le thème — inoffensif pour un cadre, indispensable
    ' pour un libellé de fond si l'essai est désactivé.
    ReculerFonds dsg, Array("lblEnteteTableI", "lblCarte4", NomCarteFiltre(), _
                            NomCarte3(), NomCarte2(), NomCarte1())

    mEtape = "module de code"
    vbComp.CodeModule.AddFromString CodeFormulairePrincipal()
    mEtape = vbNullString
    ConstruireFormulairePrincipal = dsg.Controls.Count
End Function

'------------------------------------------------------------------------------
' Fiche 1 : intitulé global — titre à gauche, année à droite, ligne d'état.
'
' ESSAI DU CADRE, commandé par F1_EN_CADRE (modInterv_Theme).
'
' Avec un cadre, les trois libellés sont DANS la carte au lieu d'être posés
' dessus. Leurs coordonnées comptent alors depuis le coin du cadre, d'où les
' deux décalages ox et oy retranchés plus bas : les positions écrites restent
' celles du formulaire, lisibles à côté du reste de la mise en page.
'
' Sans cadre, zone vaut le formulaire lui-même et les décalages sont nuls :
' une seule série de lignes sert aux deux chemins.
'------------------------------------------------------------------------------
Private Sub ConstruireFiche1(dsg As Object)
    Dim c As Object, zone As Object, ox As Single, oy As Single

    If F1_EN_CADRE Then
        Set zone = AjCtrl(dsg, "Forms.Frame.1", NomCarte1(), _
                          I_MARGE, F1_TOP, I_CARTE_LARG, F1_HAUT)
        CadreI zone
        EnBandeauI zone
        ox = I_MARGE
        oy = F1_TOP
    Else
        Set c = AjCtrl(dsg, "Forms.Label.1", NomCarte1(), _
                       I_MARGE, F1_TOP, I_CARTE_LARG, F1_HAUT)
        CarteI c
        EnBandeauI c
        Set zone = dsg
    End If

    Set c = AjCtrl(zone, "Forms.Label.1", "lblTitreGlobal", _
                   30 - ox, F1_TOP + 7 - oy, 560, 21)
    Texte c, vbNullString, Z1Titre(), MSF_TextAlignLeft

    Set c = AjCtrl(zone, "Forms.Label.1", "lblEtatI", _
                   30 - ox, F1_TOP + 28 - oy, 620, 13)
    Texte c, vbNullString, Z1Etat(), MSF_TextAlignLeft

    Set c = AjCtrl(zone, "Forms.Label.1", "lblAnnee", _
                   700 - ox, F1_TOP + 8 - oy, 230, 30)
    Texte c, vbNullString, Z1Annee(), MSF_TextAlignRight
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
    Dim col As Long, lig As Long, yCap As Single
    Dim zone As Object, ox As Single, oy As Single

    If F2_EN_CADRE Then
        Set zone = AjCtrl(dsg, "Forms.Frame.1", NomCarte2(), _
                          I_MARGE, F2_TOP, I_CARTE_LARG, F2_HAUT)
        CadreI zone
        ox = I_MARGE
        oy = F2_TOP
    Else
        Set c = AjCtrl(dsg, "Forms.Label.1", NomCarte2(), _
                       I_MARGE, F2_TOP, I_CARTE_LARG, F2_HAUT)
        CarteI c
        Set zone = dsg
    End If

    ' large : ce libellé accueille aussi, le cas échéant, la raison pour laquelle
    ' le graphique ne s'affiche pas
    Set c = AjCtrl(zone, "Forms.Label.1", "lblSectionStats", _
                   30 - ox, F2_TOP + 6 - oy, 700, 13)
    Texte c, "STATISTIQUES", Z2Section(), MSF_TextAlignLeft

    ' GrOrigineX / GrOrigineY tiennent compte du cadre : c'est la MÊME origine
    ' que celle dont Graph_Tracer repart pour placer les barres.
    ConstruireGraphique zone, GrOrigineX(), GrOrigineY()

    tuiles = TuilesStatistiques()
    For i = LBound(tuiles) To UBound(tuiles)
        col = i Mod F2_TUILES_COL
        lig = i \ F2_TUILES_COL
        x = F2TuilesX() - ox + col * (F2_TUILE_LARG + F2_TUILE_GX)
        y = F2_TOP + 24 - oy + lig * (F2_TUILE_HAUT + F2_TUILE_GY)

        Set c = AjCtrl(zone, "Forms.Image.1", "imgTuile_" & CStr(i + 1), x, y, _
                       F2_TUILE_LARG, F2_TUILE_HAUT)
        With c
            .SpecialEffect = MSF_SpecialEffectFlat
            .BorderStyle = MSF_BorderStyleNone
            .BorderColor = COUL_CARTE   ' blanc : invisible même si un filet réapparaît
            .BackStyle = MSF_BackStyleOpaque
            .BackColor = COUL_CARTE
            .PictureSizeMode = 1        ' fmPictureSizeModeStretch : remplit la tuile
        End With

        ' bloc « libellé + montant » centré verticalement dans la tuile
        yCap = y + (F2_TUILE_HAUT - (F2_TUILE_CAP_HAUT + F2_TUILE_ECART + F2_TUILE_VAL_HAUT)) / 2

        Set c = AjCtrl(zone, "Forms.Label.1", "lblStatCap_" & CStr(i + 1), _
                       x + F2_TUILE_INSET, yCap, _
                       F2_TUILE_LARG - 2 * F2_TUILE_INSET, F2_TUILE_CAP_HAUT)
        Texte c, UCase$(CStr(tuiles(i)(0))), Z2TuileCap(), MSF_TextAlignCenter
        FondTuile c

        Set c = AjCtrl(zone, "Forms.Label.1", "lblStatVal_" & CStr(i + 1), _
                       x + F2_TUILE_INSET, yCap + F2_TUILE_CAP_HAUT + F2_TUILE_ECART, _
                       F2_TUILE_LARG - 2 * F2_TUILE_INSET, F2_TUILE_VAL_HAUT)
        Texte c, vbNullString, Z2TuileVal(), MSF_TextAlignCenter
        FondTuile c
    Next i
End Sub

'------------------------------------------------------------------------------
' Fond d'un texte de tuile.
'
' Opaque et blanc : le texte doit se détacher de l'image de fond de la tuile.
' Le filet est retiré ET peint en blanc — MSForms redessine parfois une bordure
' quand un contrôle est déplacé dans le concepteur, et une bordure blanche sur
' fond blanc ne se verra pas.
'------------------------------------------------------------------------------
Private Sub FondTuile(c As Object)
    c.BackStyle = MSF_BackStyleOpaque
    c.BackColor = COUL_CARTE
    c.BorderStyle = MSF_BorderStyleNone
    c.BorderColor = COUL_CARTE
End Sub

'------------------------------------------------------------------------------
' Contrôles du graphique du chiffre d'affaires.
'   ox, oy : coin haut-gauche de la zone qui lui est réservée
'
' Rien n'est positionné définitivement ici sauf le décor : la légende, l'échelle
' et les lignes de repère. Les douze barres et les noms de mois sont créés puis
' placés à l'exécution par Graph_Tracer, qui seul connaît les valeurs.
'
' Une seule série, donc une seule teinte et aucune légende de couleurs : le
' titre nomme ce qui est représenté. La grille reste volontairement discrète, et
' seul le sommet de l'échelle est écrit — les autres montants se lisent au
' survol, plutôt que d'imprimer un nombre sur chaque barre.
'------------------------------------------------------------------------------
Private Sub ConstruireGraphique(zone As Object, ByVal ox As Single, ByVal oy As Single)
    Dim c As Object, i As Long, y As Single, largeur As Single

    largeur = F2_GRAPH_LARG - GR_MARGE_G

    Set c = AjCtrl(zone, "Forms.Label.1", "lblGrLegende", ox, oy, F2_GRAPH_LARG, GR_LEGENDE_HAUT)
    Texte c, "Chiffre d'affaires par mois", Z2GraphTitre(), MSF_TextAlignLeft

    Set c = AjCtrl(zone, "Forms.Label.1", "lblGrMax", ox, oy + GR_TRACE_TOP - 5, GR_MARGE_G - 8, 11)
    Texte c, vbNullString, Z2GraphAxe(), MSF_TextAlignRight

    ' trois repères : le sommet de l'échelle, la moitié, la ligne de base
    For i = 1 To 3
        y = oy + GR_TRACE_TOP + (i - 1) * (GR_TRACE_HAUT / 2)
        Set c = AjCtrl(zone, "Forms.Label.1", "lblGrGrille_" & CStr(i), _
                       ox + GR_MARGE_G, y, largeur, 1)
        Fond c, IIf(i = 3, COUL_GR_BASE, COUL_GR_GRILLE), _
                IIf(i = 3, COUL_GR_BASE, COUL_GR_GRILLE)
    Next i

    ' barres et noms de mois : créés ici, placés par Graph_Tracer
    For i = 1 To GR_NB_MOIS
        Set c = AjCtrl(zone, "Forms.Label.1", "lblGrBarre_" & CStr(i), _
                       ox + GR_MARGE_G, oy + GR_TRACE_TOP, GR_BARRE_LARG, 1)
        Fond c, COUL_GR_BARRE, COUL_GR_BARRE

        Set c = AjCtrl(zone, "Forms.Label.1", "lblGrMois_" & CStr(i), _
                       ox + GR_MARGE_G, oy + GR_TRACE_TOP + GR_TRACE_HAUT + 3, _
                       GR_BARRE_LARG, GR_MOIS_HAUT)
        Texte c, vbNullString, Z2GraphAxe(), MSF_TextAlignCenter
    Next i
End Sub

'------------------------------------------------------------------------------
' Fiche 3 : les quinze champs de saisie, posés d'après le schéma.
'------------------------------------------------------------------------------
Private Sub ConstruireFiche3(dsg As Object)
    Dim c As Object, ch() As ChampInterv, i As Long
    Dim x As Single, y As Single, larg As Single, haut As Single, ordre As Long
    Dim zone As Object, ox As Single, oy As Single, cZoneDate As Object

    If F3_EN_CADRE Then
        Set zone = AjCtrl(dsg, "Forms.Frame.1", NomCarte3(), _
                          I_MARGE, F3_TOP, I_CARTE_LARG, F3_HAUT)
        CadreI zone
        ox = I_MARGE
        oy = F3_TOP
    Else
        Set c = AjCtrl(dsg, "Forms.Label.1", NomCarte3(), _
                       I_MARGE, F3_TOP, I_CARTE_LARG, F3_HAUT)
        CarteI c
        Set zone = dsg
    End If

    Set c = AjCtrl(zone, "Forms.Label.1", "lblSectionSaisieI", _
                   30 - ox, F3_TOP + 6 - oy, 300, 13)
    Texte c, "FICHE INTERVENTION", Z3Section(), MSF_TextAlignLeft

    ch = ObtenirChampsInterv()
    ordre = 1

    For i = LBound(ch) To UBound(ch)
        ' l'étape nomme le champ : une erreur MSForms ne dit jamais sur quel
        ' contrôle elle s'est produite
        mEtape = "fiche 3 (saisie), champ " & ch(i).Colonne
        x = IGrilleX(ch(i).Col) - ox
        y = IGrilleY(ch(i).Ligne) - oy
        larg = ILargeurBlocs(ch(i).Blocs)
        haut = IHauteurLignes(ch(i).NbLignes)

        ' deux champs peuvent se partager une colonne : les cases TVA et Forfait
        If ch(i).Moitie = 1 Then
            larg = (IG_BLOC - 8) / 2
        ElseIf ch(i).Moitie = 2 Then
            larg = (IG_BLOC - 8) / 2
            x = x + larg + 8
        End If

        If ch(i).TypeCtrl = ITYPE_CASE Then
            Set c = AjCtrl(zone, "Forms.CheckBox.1", INomControle(ch(i)), _
                           x, y + ICH_LBL_HAUT + 1, larg, ICH_CTL_HAUT)
            Case_ c, ch(i).Libelle
        Else
            Set c = AjCtrl(zone, "Forms.Label.1", INomLibelle(ch(i)), x, y, larg, ICH_LBL_HAUT)
            Texte c, UCase$(ch(i).Libelle), Z3Libelle(), MSF_TextAlignLeft

            Select Case ch(i).TypeCtrl
                Case ITYPE_LISTE, ITYPE_AUTO
                    Set c = AjCtrl(zone, "Forms.ComboBox.1", INomControle(ch(i)), _
                                   x, y + ICH_LBL_HAUT + 1, larg, haut)
                    Liste c, (ch(i).TypeCtrl = ITYPE_AUTO)

                Case ITYPE_DATE
                    ' La zone de date laisse la place au bouton du calendrier.
                    '
                    ' Elle est MISE DE CÔTÉ dans cZoneDate, et non rappelée par son
                    ' nom après coup : zone.Controls("...") interrogeait la
                    ' collection d'un CADRE, ce que le concepteur ne gère pas
                    ' toujours — d'où une erreur 438 à la génération. Garder
                    ' l'objet évite la question, et c'était de toute façon un
                    ' détour : on venait de le créer.
                    Set cZoneDate = AjCtrl(zone, "Forms.TextBox.1", INomControle(ch(i)), _
                                       x, y + ICH_LBL_HAUT + 1, larg - 20, haut)
                    Zone cZoneDate, False

                    Set c = AjCtrl(zone, "Forms.Label.1", "lblCalendrier", _
                                   x + larg - 18, y + ICH_LBL_HAUT + 1, 18, haut)
                    Texte c, ChrW(9662), Z3Chevron(), MSF_TextAlignCenter
                    c.BackStyle = MSF_BackStyleOpaque
                    c.BackColor = COUL_MODIFIER
                    c.ControlTipText = "Ouvrir le calendrier"

                    ' la suite du tour porte sur la zone de date, pas sur le
                    ' chevron : c'est elle qui reçoit l'aide et le tabulateur
                    Set c = cZoneDate

                Case Else
                    Set c = AjCtrl(zone, "Forms.TextBox.1", INomControle(ch(i)), _
                                   x, y + ICH_LBL_HAUT + 1, larg, haut)
                    Zone c, ch(i).Verrouille
                    ' une zone haute de plusieurs lignes accueille du texte long :
                    ' le retour à la ligne et l'ascenseur y deviennent utiles.
                    ' Zone a laissé EnterKeyBehavior à False : Entrée valide la
                    ' fiche, Ctrl+Entrée va à la ligne.
                    If ch(i).NbLignes > 1 Then
                        c.MultiLine = True
                        c.ScrollBars = MSF_ScrollBarsVertical
                    End If
            End Select
        End If

        c.ControlTipText = ch(i).Aide
        If Not ch(i).Verrouille Then
            c.TabIndex = ordre
            ordre = ordre + 1
        End If
    Next i

    ' libellé commun aux deux cases à cocher, qui n'en ont pas d'individuel :
    ' il occupe la ligne de libellés que TVA et Forfait laissent vide
    Set c = AjCtrl(zone, "Forms.Label.1", "lblI_Facturation", _
                   IGrilleX(3) - ox, IGrilleY(4) - oy, IG_BLOC, ICH_LBL_HAUT)
    Texte c, "FACTURATION", Z3Libelle(), MSF_TextAlignLeft
End Sub

'------------------------------------------------------------------------------
' Barre de filtrage, entre la saisie et le tableau.
'------------------------------------------------------------------------------
Private Sub ConstruireFiltre(dsg As Object)
    Dim c As Object, y As Single
    Dim zone As Object, ox As Single, oy As Single

    If F4_EN_CADRE Then
        Set zone = AjCtrl(dsg, "Forms.Frame.1", NomCarteFiltre(), _
                          I_MARGE, IF_TOP, I_CARTE_LARG, IF_HAUT)
        CadreI zone
        ox = I_MARGE
        oy = IF_TOP
    Else
        Set c = AjCtrl(dsg, "Forms.Label.1", NomCarteFiltre(), _
                       I_MARGE, IF_TOP, I_CARTE_LARG, IF_HAUT)
        CarteI c
        Set zone = dsg
    End If

    ' toutes les abscisses de la barre passent par ici : une seule soustraction
    y = IF_TOP + 10 - oy

    Set c = AjCtrl(zone, "Forms.Label.1", "lblFiltreTitreI", IFB_TITRE_X - ox, y + 2, 62, 14)
    Texte c, "Filtrer sur", Z4Libelle(), MSF_TextAlignLeft

    Set c = AjCtrl(zone, "Forms.ComboBox.1", "cboChampFiltreI", IFB_CHAMP_X - ox, y, _
                   IFB_CHAMP_L, ICH_CTL_HAUT)
    Liste c, False
    c.Style = MSF_StyleDropDownList
    c.ControlTipText = "Colonne du tableau sur laquelle porte le filtre"

    Set c = AjCtrl(zone, "Forms.TextBox.1", "txtFiltreI", IFB_TEXTE_X - ox, y, _
                   IFB_TEXTE_L, ICH_CTL_HAUT)
    Zone c, False
    c.ControlTipText = "Texte à rechercher (accents et majuscules sont ignorés)"

    ' --- second filtre, sur le mois de la date --------------------------------
    Set c = AjCtrl(zone, "Forms.Label.1", "lblMoisFiltreI", IFB_MOIS_LBL_X - ox, y + 2, 28, 14)
    Texte c, "Mois", Z4Libelle(), MSF_TextAlignLeft

    Set c = AjCtrl(zone, "Forms.ComboBox.1", "cboMoisI", IFB_MOIS_X - ox, y, _
                   IFB_MOIS_L, ICH_CTL_HAUT)
    Liste c, False
    c.Style = MSF_StyleDropDownList
    c.ControlTipText = "N'afficher que les interventions d'un mois"

    Set c = AjCtrl(zone, "Forms.Label.1", "lblResetFiltreI", IFB_RESET_X - ox, y + 2, 80, 14)
    Texte c, "Réinitialiser", Z4Lien(), MSF_TextAlignLeft
    c.ControlTipText = "Effacer les deux filtres et réafficher toutes les interventions"

    Set c = AjCtrl(zone, "Forms.Label.1", "lblCompteurI", _
                   I_MARGE + I_CARTE_LARG - F2_PADDING - IFB_COMPTEUR_L - ox, y + 2, _
                   IFB_COMPTEUR_L, 14)
    Texte c, vbNullString, Z4Compteur(), MSF_TextAlignRight
End Sub

'------------------------------------------------------------------------------
' Fiche 4 : le tableau des enregistrements.
'
' Une ListBox MSForms ne sait pas afficher d'en-têtes autrement qu'en étant liée
' à une plage de cellules : ce sont donc des libellés posés au-dessus, alignés
' sur les mêmes largeurs — ce qui permet en prime de les rendre cliquables.
'------------------------------------------------------------------------------
Private Sub ConstruireFiche4(dsg As Object)
    Dim c As Object, larg As Variant, lib As Variant, ali As Variant
    Dim i As Long, r As Long, nbCol As Long
    Dim x As Single, y As Single, gauche As Single, largeur As Single, base As Single

    gauche = AuPixel(I_MARGE + 1)
    largeur = I_CARTE_LARG - 2

    Set c = AjCtrl(dsg, "Forms.Label.1", "lblCarte4", I_MARGE, IT_TOP, I_CARTE_LARG, IT_HAUT)
    CarteI c

    Set c = AjCtrl(dsg, "Forms.Label.1", "lblEnteteTableI", gauche, AuPixel(IT_TOP + 1), _
                   largeur, AuPixel(IT_ENTETE - 1))
    Fond c, COUL_ENTETE_TBL, COUL_ENTETE_TBL

    larg = ILargeursListe()
    lib = ILibellesListe()
    ali = IAlignementsListe()
    nbCol = UBound(larg) - LBound(larg) + 1
    base = IT_TOP + IT_ENTETE + 1

    ' --- en-têtes -------------------------------------------------------------
    ' Demi-gras SANS Bold : c'est la famille Segoe UI Semibold qui porte la
    ' graisse, MSForms ne connaissant que gras ou pas gras.
    x = gauche + IGR_PAD_X
    For i = 0 To nbCol - 1
        Set c = AjCtrl(dsg, "Forms.Label.1", "lblEntI_" & CStr(i + 1), _
                       AuPixel(x), AuPixel(IT_TOP + 5), _
                       AuPixel(CSng(larg(i)) - 2 * IGR_PAD_X), AuPixel(IGR_LIGNE_H))
        Texte c, CStr(lib(i)), Z5Entete(), CLng(ali(i))
        c.ControlTipText = "Cliquez pour trier sur cette colonne"
        x = x + CSng(larg(i))
    Next i

    ' --- les lignes et leurs cases --------------------------------------------
    ' Chaque ligne reçoit d'abord une BANDE de fond sur toute la largeur, puis
    ' ses cases, transparentes, posées dessus. C'est la bande qui porte la
    ' couleur : sans elle, les points qui séparent deux cases laisseraient voir
    ' le blanc de la carte et la ligne choisie paraîtrait rayée.
    '
    ' Les contrôles s'ajoutent PAR-DEVANT : la bande créée en premier se retrouve
    ' donc bien derrière ses cases.
    '
    ' TOUTE coordonnée passe par AuPixel. L'abscisse d'une colonne est la somme
    ' des largeurs qui la précèdent, et rien ne garantit qu'elle tombe sur un
    ' pixel : une case posée à cheval rend son texte décalé et plus épais, d'une
    ' colonne à l'autre sans régularité. Les largeurs, elles, restent exactes —
    ' seul l'affichage est calé.
    For r = 1 To IGR_NB_LIGNES
        y = base + (r - 1) * IGR_LIGNE_H

        ' Fond plat, SANS filet : un libellé bordé dessine un cadre d'un pixel,
        ' et deux bandes voisines feraient une ligne double entre elles.
        Set c = AjCtrl(dsg, "Forms.Label.1", "lblGL_" & CStr(r), _
                       gauche, AuPixel(y), AuPixel(largeur - IGR_BARRE_L), _
                       AuPixel(IGR_LIGNE_H))
        With c
            .Caption = vbNullString
            .SpecialEffect = MSF_SpecialEffectFlat
            .BorderStyle = MSF_BorderStyleNone
            .BackStyle = MSF_BackStyleOpaque
            .BackColor = COUL_CARTE
        End With
        ' La bande ne montre aucun texte, mais elle fait partie du tableau :
        ' elle porte donc le style de la zone 6, et non le gras du formulaire.
        PoserStyle c, Z6Case()

        x = gauche + IGR_PAD_X
        For i = 0 To nbCol - 1
            Set c = AjCtrl(dsg, "Forms.Label.1", _
                           "lblG_" & CStr(r) & "_" & CStr(i + 1), _
                           AuPixel(x), AuPixel(y), _
                           AuPixel(CSng(larg(i)) - 2 * IGR_PAD_X), AuPixel(IGR_LIGNE_H))
            Texte c, vbNullString, Z6Case(), CLng(ali(i))
            x = x + CSng(larg(i))
        Next i
    Next r

    ' --- barre de défilement --------------------------------------------------
    Set c = AjCtrl(dsg, "Forms.ScrollBar.1", "sbGrille", _
                   AuPixel(gauche + largeur - IGR_BARRE_L), AuPixel(base), _
                   AuPixel(IGR_BARRE_L), AuPixel(IGR_NB_LIGNES * IGR_LIGNE_H))
    With c
        .Min = 0
        .Max = 0
        .Value = 0
        .SmallChange = 1
        .LargeChange = IGR_NB_LIGNES     ' un clic dans le vide fait une page

        ' CURSEUR DE TAILLE FIXE. Avec ProportionalThumb, MSForms dimensionne le
        ' curseur à LargeChange / (plage + LargeChange). Ici la plage vaut le
        ' nombre de lignes en trop — douze pour vingt-neuf interventions — et
        ' LargeChange une page entière : le curseur occupait presque toute la
        ' glissière et ne se distinguait plus du fond.
        .ProportionalThumb = False
        .TabIndex = 90
    End With
End Sub

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

    mEtape = "calendrier"
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

    Set c = AjCtrl(dsg, "Forms.Label.1", "lblCalPrec", 10, 10, 24, 22)
    Texte c, ChrW(8249), Z7Fleche(), MSF_TextAlignCenter
    c.ControlTipText = "Mois précédent"

    Set c = AjCtrl(dsg, "Forms.Label.1", "lblCalMois", 40, 11, 180, 19)
    Texte c, vbNullString, Z7Mois(), MSF_TextAlignCenter

    Set c = AjCtrl(dsg, "Forms.Label.1", "lblCalSuiv", CAL_LARGEUR - 34, 10, 24, 22)
    Texte c, ChrW(8250), Z7Fleche(), MSF_TextAlignCenter
    c.ControlTipText = "Mois suivant"

    ' --- en-têtes des jours ---------------------------------------------------
    For i = 0 To 6
        Set c = AjCtrl(dsg, "Forms.Label.1", "lblJS_" & CStr(i + 1), _
                       CAL_GRILLE_X + i * CAL_JOUR_LARG, CAL_ENTETES_Y, CAL_JOUR_LARG, 13)
        Texte c, vbNullString, Z7JourSem(), MSF_TextAlignCenter
    Next i

    ' --- colonnes du week-end, teintées derrière la grille --------------------
    For i = 1 To 2
        Set c = AjCtrl(dsg, "Forms.Label.1", "lblCalWE_" & CStr(i), _
                       CAL_GRILLE_X + (4 + i) * CAL_JOUR_LARG, CAL_GRILLE_Y, _
                       CAL_JOUR_LARG, 6 * CAL_JOUR_HAUT)
        Fond c, COUL_CAL_WEEKEND, COUL_CAL_WEEKEND
    Next i

    ' --- grille des 42 cases --------------------------------------------------
    For i = 1 To 42
        col = (i - 1) Mod 7
        lig = (i - 1) \ 7
        Set c = AjCtrl(dsg, "Forms.Label.1", "lblJ_" & CStr(i), _
                       CAL_GRILLE_X + col * CAL_JOUR_LARG, _
                       CAL_GRILLE_Y + lig * CAL_JOUR_HAUT, CAL_JOUR_LARG, CAL_JOUR_HAUT)
        Texte c, vbNullString, Z7Jour(), MSF_TextAlignCenter
    Next i

    ' --- raccourcis du bas ----------------------------------------------------
    Set c = AjCtrl(dsg, "Forms.Label.1", "lblCalAujourdhui", CAL_GRILLE_X, CAL_PIED_Y, 92, 14)
    Texte c, "Aujourd'hui", Z7Lien(), MSF_TextAlignLeft
    c.ControlTipText = "Revenir au mois en cours"

    Set c = AjCtrl(dsg, "Forms.Label.1", "lblCalFinMois", CAL_GRILLE_X + 97, CAL_PIED_Y, 80, 14)
    Texte c, "Fin de mois", Z7Lien(), MSF_TextAlignLeft
    c.ControlTipText = "Choisir le dernier jour du mois affiché"

    Set c = AjCtrl(dsg, "Forms.Label.1", "lblCalAnnuler", CAL_LARGEUR - CAL_GRILLE_X - 57, _
                   CAL_PIED_Y, 57, 14)
    Texte c, "Annuler", Z7Annuler(), MSF_TextAlignRight

    ReculerFonds dsg, Array("lblCalWE_1", "lblCalWE_2", "lblCalBandeau")

    vbComp.CodeModule.AddFromString CodeCalendrier()
    ConstruireCalendrier = dsg.Controls.Count
End Function

'==============================================================================
' MISE EN FORME
'==============================================================================
'------------------------------------------------------------------------------
' Police par défaut du FORMULAIRE — à poser avant le premier contrôle.
'
' Un contrôle MSForms recopie la police du formulaire au moment où il est créé.
' Tant que le formulaire n'en a pas, cette police est MS Sans Serif 8 gras, la
' valeur d'usine. Posée ici une fois pour toutes, elle devient Segoe UI 8 gras :
' c'est le défaut du formulaire, dont chaque contrôle s'écarte ensuite s'il le
' veut — l'intérieur du tableau, notamment, qui est en neuf points maigres.
'
' L'ORDRE COMPTE DEUX FOIS. Le nom de famille avant la taille et la graisse,
' sans quoi changer de police les remet aux valeurs par défaut de la nouvelle.
' Et l'appel avant le premier contrôle : après coup il ne servirait à rien, les
' contrôles déjà nés ayant chacun leur propre copie de la police.
'------------------------------------------------------------------------------
Private Sub PoliceParDefaut(dsg As Object)
    On Error Resume Next
    PoserPolice dsg, POLICE, TAILLE_DEFAUT, GRAS_DEFAUT
    On Error GoTo 0
End Sub

' Le premier argument est un CONTENEUR, pas forcément le formulaire : un cadre
' expose la même collection Controls, et y ajouter un contrôle l'y place. Les
' coordonnées comptent alors depuis le coin du conteneur.
Private Function AjCtrl(conteneur As Object, ByVal progId As String, ByVal nom As String, _
                        ByVal gauche As Single, ByVal haut As Single, _
                        ByVal largeur As Single, ByVal hauteur As Single) As Object
    Dim c As Object
    Set c = conteneur.Controls.Add(progId, nom, True)
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
'------------------------------------------------------------------------------
' Fond de zone en CADRE : même aspect que CarteI — rectangle blanc à filet fin
' — mais c'est un conteneur, pas un libellé.
'
' SpecialEffect AVANT BorderStyle, comme partout ailleurs : MSForms refuse une
' bordure simple tant que le contrôle est en relief, et un cadre est en relief
' gravé par défaut.
'
' Caption vidée : MSForms réserve sinon une bande en haut du cadre pour le
' titre. ScrollBars éteintes explicitement : un cadre est un conteneur
' défilable, et on ne veut ici qu'un fond.
'------------------------------------------------------------------------------
Private Sub CadreI(c As Object)
    c.Caption = vbNullString
    c.BackColor = COUL_CARTE
    c.SpecialEffect = MSF_SpecialEffectFlat
    c.BorderStyle = MSF_BorderStyleSingle
    c.BorderColor = COUL_BORDURE
    c.ScrollBars = MSF_ScrollBarsNone
End Sub

'------------------------------------------------------------------------------
' Le BANDEAU SUPÉRIEUR : la carte de la zone 1, en aplat bleu foncé.
'
' NE POSE QUE LA COULEUR, et n'appelle aucun habilleur. Elle en appelait un,
' celui des CADRES, y compris quand la carte était un libellé : ce dernier n'a
' pas de barres de défilement, et MSForms répondait « propriété non gérée par
' cet objet ». L'appelant choisit désormais son habilleur, puis recolore.
'
' Les quatre formulaires du classeur portent le même bandeau, celui de la
' gestion des clients ; le titre, la ligne d'état et l'année prennent leurs
' styles du bandeau commun, dans modInterv_Theme.
'------------------------------------------------------------------------------
Private Sub EnBandeauI(c As Object)
    c.BackColor = COUL_BANDEAU
    c.BorderColor = COUL_BANDEAU
End Sub

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
Private Sub Texte(c As Object, ByVal texte As String, ByRef st As StyleTexte, _
                  ByVal alignement As Long)
    c.Caption = texte
    c.BackStyle = MSF_BackStyleTransparent
    c.SpecialEffect = MSF_SpecialEffectFlat
    c.BorderStyle = MSF_BorderStyleNone
    c.TextAlign = alignement
    c.WordWrap = False
    c.AutoSize = False
    PoserStyle c, st
End Sub

'------------------------------------------------------------------------------
' Pose les quatre caractéristiques d'un style sur un contrôle.
'
' Le nom de FAMILLE en premier : la changer après coup remettrait la taille et
' la graisse aux valeurs par défaut de la nouvelle police.
'
' Séparée de Texte pour servir aussi aux contrôles sans légende — les bandes de
' fond de la grille — qui portent la police de leur zone sans rien afficher.
'------------------------------------------------------------------------------
Private Sub PoserStyle(c As Object, ByRef st As StyleTexte)
    c.ForeColor = st.Couleur
    PoserPolice c, st.Police, st.Taille, st.Gras
End Sub

'------------------------------------------------------------------------------
' Pose une police sur un contrôle. LE SEUL ENDROIT qui écrive dans c.Font.
'
' L'ORDRE EST LA RAISON D'ÊTRE DE CETTE PROCÉDURE, et il n'est pas celui qu'on
' écrirait spontanément.
'
'   1. la FAMILLE d'abord : en changer remet le reste aux valeurs par défaut
'      de la nouvelle police ;
'   2. la GRAISSE ensuite, jamais après la taille ;
'   3. le POIDS, qui est la même chose vue de plus bas — 400 maigre, 700 gras.
'      Certaines versions ne reprennent que celui-là ; on pose les deux ;
'   4. la TAILLE en DERNIER.
'
' Pourquoi la taille en dernier : MSForms recrée la police en pixels entiers au
' moment où on lui donne un corps. Demander 8 points en rend 8,25, soit onze
' pixels tout ronds — on le voit dans le diagnostic. Cette recréation emporte
' la graisse posée APRÈS elle : c'est pour cela que les zones de saisie
' restaient en gras alors que Font.Bold = False était bien exécuté, et que ni
' l'override ni le changement du défaut du formulaire n'y ont rien changé.
'
' Le poids est posé sous On Error Resume Next : toutes les versions de MSForms
' n'exposent pas Weight, et son absence ne doit pas faire échouer la
' génération — la graisse aura déjà été posée à la ligne précédente.
'------------------------------------------------------------------------------
' PUBLIQUE : tout générateur doit poser ses polices par ici, l'ordre des
' assignations étant la seule chose qui empêche la taille d'emporter la graisse.
Public Sub PoserPolice(c As Object, ByVal police As String, _
                       ByVal taille As Single, ByVal gras As Boolean)
    c.Font.Name = police
    c.Font.Bold = gras
    On Error Resume Next
    c.Font.Weight = IIf(gras, 700, 400)
    On Error GoTo 0
    c.Font.Size = taille
End Sub

'------------------------------------------------------------------------------
' Zone de saisie. Locked plutôt qu'Enabled = False pour un champ géré par le
' programme : il reste lisible et son contenu copiable.
'------------------------------------------------------------------------------
Private Sub Zone(c As Object, ByVal verrouille As Boolean)
    PoserPolice c, POLICE, TAILLE_CHAMP, False
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
    PoserPolice c, POLICE, TAILLE_CHAMP, False
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
    PoserPolice c, POLICE, TAILLE_FILTRE, False
    c.Caption = libelle
    c.BackStyle = MSF_BackStyleTransparent
    c.ForeColor = COUL_TEXTE
    c.WordWrap = False
End Sub

'------------------------------------------------------------------------------
' Habille un bouton d'action : couleur de fond, texte blanc, lettre de raccourci
' soulignée et rang de tabulation.
'------------------------------------------------------------------------------
Private Sub Bouton_(c As Object, ByVal libelle As String, ByVal raccourci As String, _
                    ByVal couleur As Long, ByVal ordre As Long)
    PoserPolice c, POLICE, TAILLE_BOUTON, True
    c.Caption = libelle
    c.Accelerator = raccourci
    c.BackColor = couleur
    c.ForeColor = COUL_BOUTON_TXT
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
    Dim ch() As ChampInterv, i As Long, r As Long, n As String, sortie As String
    Dim boutons As Variant, actions As Variant, larg As Variant

    mCode = vbNullString
    EnteteGenere NOM_FORM_INTERV, "modInterv_Formulaire"

    ProcI "UserForm_Initialize()", "Interv_Initialiser Me"
    ProcI "UserForm_Activate()", "Interv_Activer Me"
    ProcI "UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)", "Interv_Fermeture Me"
    ' quitter tout contrôle actif ramène boutons et graphique au repos
    ProcI "UserForm_MouseMove" & SIG_SOURIS, _
         "Interv_Survol Me, " & Guill("") & vbNewLine & _
         "    Graph_Survol Me, 0"

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

    ' --- graphique ------------------------------------------------------------
    ' survoler une barre écrit son mois et son montant dans la légende ; le fond
    ' de la carte, lui, rend la main au titre (voir UserForm_MouseMove ci-dessus)
    For i = 1 To GR_NB_MOIS
        ProcI "lblGrBarre_" & CStr(i) & "_MouseMove" & SIG_SOURIS, _
             "Graph_Survol Me, " & CStr(i)
    Next i

    ' La carte recouvre toute la zone du graphique : c'est elle, et non le fond
    ' du formulaire, que la souris atteint en quittant une barre.
    ProcI NomCarte2() & "_MouseMove" & SIG_SOURIS, _
         "Interv_Survol Me, " & Guill("") & vbNewLine & _
         "    Graph_Survol Me, 0"

    ' --- filtrage -------------------------------------------------------------
    ProcI "cboChampFiltreI_Change()", "Interv_AppliquerFiltre Me"
    ProcI "txtFiltreI_Change()", "Interv_AppliquerFiltre Me"
    ProcI "cboMoisI_Change()", "Interv_AppliquerFiltre Me"
    ProcI "lblResetFiltreI_Click()", "Interv_ReinitialiserFiltre Me"
    ProcI "lblResetFiltreI_MouseMove" & SIG_SOURIS, "Interv_Survol Me, " & Guill("lblResetFiltreI")

    ' --- tableau --------------------------------------------------------------
    ' La grille est faite de libellés : c'est la case cliquée qui reçoit
    ' l'événement, et elle ne sait dire que sa LIGNE — c'est tout ce qui compte,
    ' une ligne entière se sélectionnant d'un bloc. Le survol suit le même
    ' chemin ; Interv_GrilleSurvol ne repeint que la ligne quittée et la ligne
    ' prise, jamais les dix-sept.
    larg = ILargeursListe()
    For r = 1 To IGR_NB_LIGNES
        For i = LBound(larg) To UBound(larg)
            n = "lblG_" & CStr(r) & "_" & CStr(i + 1)
            ProcI n & "_Click()", "Interv_GrilleClic Me, " & CStr(r)
            ProcI n & "_MouseMove" & SIG_SOURIS, "Interv_GrilleSurvol Me, " & CStr(r)
        Next i

        ' la bande dépasse des cases entre deux colonnes : elle répond aussi
        n = "lblGL_" & CStr(r)
        ProcI n & "_Click()", "Interv_GrilleClic Me, " & CStr(r)
        ProcI n & "_MouseMove" & SIG_SOURIS, "Interv_GrilleSurvol Me, " & CStr(r)
    Next r

    ' quitter la grille éteint le survol : la carte est le seul fond qu'on
    ' atteigne en sortant d'une case
    ProcI "lblCarte4_MouseMove" & SIG_SOURIS, "Interv_GrilleSurvol Me, 0"
    ProcI "lblEnteteTableI_MouseMove" & SIG_SOURIS, "Interv_GrilleSurvol Me, 0"

    ' Change SEULEMENT, jamais Scroll. Scroll se déclenche à chaque pixel de
    ' déplacement du curseur ; repeindre dix-sept lignes à cette cadence entrait
    ' en conflit avec le dessin de la barre, qui clignotait entre gris et noir
    ' dès qu'on l'avait saisie. Change se déclenche aux flèches, aux clics dans
    ' la glissière et au relâchement du curseur : le tableau suit donc le
    ' glissement à la fin plutôt qu'en continu, et rien ne scintille.
    ProcI "sbGrille_Change()", "Interv_Defiler Me"

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

            ' le nombre de personnes et le taux entrent dans le calcul du
            ' chiffre d'affaires, qui se met à jour à chaque frappe
            If StrComp(ch(i).Colonne, IC_PERS, vbTextCompare) = 0 _
               Or StrComp(ch(i).Colonne, IC_TAUX, vbTextCompare) = 0 Then
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

    ' un seul point de remise au repos : la souris qui sort d'une case ou d'un
    ' lien retombe forcément sur le fond du formulaire
    ProcI "UserForm_MouseMove" & SIG_SOURIS, "Cal_SurvolRepos Me"

    ' Le bandeau et les en-têtes de jours recouvrent le fond : sans eux, sortir
    ' d'une flèche ou d'une case ne déclencherait aucun événement.
    ProcI "lblCalBandeau_MouseMove" & SIG_SOURIS, "Cal_SurvolRepos Me"
    ProcI "lblCalMois_MouseMove" & SIG_SOURIS, "Cal_SurvolRepos Me"
    For i = 1 To 7
        ProcI "lblJS_" & CStr(i) & "_MouseMove" & SIG_SOURIS, "Cal_SurvolRepos Me"
    Next i

    ' --- bandeau --------------------------------------------------------------
    ProcI "lblCalPrec_Click()", "Cal_MoisPrecedent Me"
    ProcI "lblCalPrec_MouseMove" & SIG_SOURIS, "Cal_SurvolLien Me, " & Guill("lblCalPrec") & ", True"
    ProcI "lblCalSuiv_Click()", "Cal_MoisSuivant Me"
    ProcI "lblCalSuiv_MouseMove" & SIG_SOURIS, "Cal_SurvolLien Me, " & Guill("lblCalSuiv") & ", True"

    ' --- grille : un clic choisit, un survol éclaire --------------------------
    For i = 1 To 42
        ProcI "lblJ_" & CStr(i) & "_Click()", "Cal_ChoisirJour Me, " & CStr(i)
        ProcI "lblJ_" & CStr(i) & "_MouseMove" & SIG_SOURIS, "Cal_SurvolJour Me, " & CStr(i)
    Next i

    ' --- raccourcis du pied ---------------------------------------------------
    ProcI "lblCalAujourdhui_Click()", "Cal_Aujourdhui Me"
    ProcI "lblCalAujourdhui_MouseMove" & SIG_SOURIS, _
         "Cal_SurvolPied Me, " & Guill("lblCalAujourdhui") & ", True, False"
    ProcI "lblCalFinMois_Click()", "Cal_FinMois Me"
    ProcI "lblCalFinMois_MouseMove" & SIG_SOURIS, _
         "Cal_SurvolPied Me, " & Guill("lblCalFinMois") & ", True, False"
    ProcI "lblCalAnnuler_Click()", "Cal_Annuler Me"
    ProcI "lblCalAnnuler_MouseMove" & SIG_SOURIS, _
         "Cal_SurvolPied Me, " & Guill("lblCalAnnuler") & ", True, True"

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
