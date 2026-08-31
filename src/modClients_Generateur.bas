Attribute VB_Name = "modClients_Generateur"
Option Explicit
'==============================================================================
' modClients_Generateur
'------------------------------------------------------------------------------
' Génère de toutes pièces le UserForm "UF_Clients" à partir du schéma décrit
' dans modClients_Schéma : création du formulaire, de ses contrôles, de leur
' mise en forme, puis injection de son module de code.
'
'   >>> Procédure à lancer : GenererFormulaireClients
'
' PREREQUIS : Fichier > Options > Centre de gestion de la confidentialité >
'             Parametres du Centre de gestion de la confidentialité >
'             Parametres des macros > cocher
'             "Accès approuve au modèle d'objet du projet VBA".
'
' Ce module n'est nécessaire qu'au moment de la génération : une fois le
' formulaire créé, il peut rester en place (il permet de le régénérer après
' toute modification du schéma ou de la charte graphique).
'==============================================================================

Private Const CT_MSFORM As Long = 3         ' vbext_ct_MSForm
Private Const SIG_SOURIS As String = "(ByVal Button As Integer, ByVal Shift As Integer, " & _
                                     "ByVal X As Single, ByVal Y As Single)"

Private mCode As String

'==============================================================================
' POINT D'ENTREE
'==============================================================================
'------------------------------------------------------------------------------
' Crée le formulaire UF_Clients de toutes pièces.
'
' Déroulement :
'   1. vérifie que TblClients existe et signale ses colonnes absentes du schéma ;
'   2. demande l'accès au projet VBA, refusé par Excel tant que l'option de
'      confiance n'est pas cochée ;
'   3. remplace, après confirmation, une version précédente du formulaire ;
'   4. crée le UserForm et pose ses propriétés ;
'   5. construit les cinq zones, du fond vers le contenu ;
'   6. écrit le module de code du formulaire.
'
' À relancer après toute modification de modClients_Schema ou modClients_Theme.
'------------------------------------------------------------------------------
Public Sub GenererFormulaireClients()
    Dim vbProj As Object, vbComp As Object, dsg As Object
    Dim lo As ListObject, manquantes As String
    Dim cree As Boolean          ' True si le composant vient d'être créé ici

    ' --- le tableau source doit exister ---------------------------------------
    Set lo = TableClients()
    If lo Is Nothing Then
        MsgBox "Le tableau " & NOM_TABLE_CLIENTS & " est introuvable dans ce classeur." & vbCrLf & _
               "Le formulaire ne peut pas être généré.", vbCritical, "Génération du formulaire"
        Exit Sub
    End If

    manquantes = ColonnesNonCouvertes(lo)
    If Len(manquantes) > 0 Then
        If MsgBox("Ces colonnes de " & NOM_TABLE_CLIENTS & " ne figurent pas dans le schéma du " & _
                  "formulaire et ne seront donc pas saisissables :" & vbCrLf & vbCrLf & manquantes & _
                  vbCrLf & vbCrLf & "Poursuivre la génération ?", _
                  vbQuestion + vbYesNo, "Génération du formulaire") <> vbYes Then Exit Sub
    End If

    ' --- accès au projet VBA --------------------------------------------------
    ' Sans l'option de confiance, cette seule ligne déclenche l'erreur 1004.
    ' On la neutralise pour pouvoir afficher un message compréhensible plutôt
    ' que de laisser remonter l'erreur brute.
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

    ' --- création ou remise à neuf du formulaire ------------------------------
    On Error Resume Next
    Set vbComp = vbProj.VBComponents(NOM_FORMULAIRE)
    On Error GoTo 0

    If Not vbComp Is Nothing Then
        If MsgBox("Le formulaire " & NOM_FORMULAIRE & " existe déjà." & vbCrLf & _
                  "Le reconstruire à neuf ?", vbQuestion + vbYesNo, _
                  "Génération du formulaire") <> vbYes Then Exit Sub
    End If

    On Error GoTo Erreur
    PreparerComposant vbProj, vbComp, cree
    Set dsg = vbComp.Designer

    PropFormulaire vbComp, "Caption", "Gestion des clients"
    PropFormulaire vbComp, "Width", F_LARGEUR
    PropFormulaire vbComp, "Height", F_HAUTEUR
    PropFormulaire vbComp, "BackColor", COUL_FOND
    PropFormulaire vbComp, "SpecialEffect", MSF_SpecialEffectFlat
    PropFormulaire vbComp, "BorderStyle", MSF_BorderStyleNone
    PropFormulaire vbComp, "StartUpPosition", 1          ' centre sur Excel
    PropFormulaire vbComp, "ShowModal", True

    ConstruireBandeau dsg
    ConstruireCarteSaisie dsg
    ConstruireCarteFiltre dsg
    ConstruireCarteTableau dsg
    ConstruireBoutons dsg
    ReculerArrierePlans dsg

    ' --- module de code du formulaire ----------------------------------------
    vbComp.CodeModule.AddFromString CodeDuFormulaire()

    MsgBox "Le formulaire " & NOM_FORMULAIRE & _
           IIf(cree, " a été généré.", " a été reconstruit.") & vbCrLf & vbCrLf & _
           dsg.Controls.Count & " contrôles créés pour " & NB_CHAMPS & " champs du tableau " & _
           NOM_TABLE_CLIENTS & "." & vbCrLf & vbCrLf & _
           "Lancez maintenant la procédure OuvrirGestionClients pour l'afficher.", _
           vbInformation, "Génération du formulaire"
    Exit Sub

Erreur:
    ' Un composant tout juste créé mais inachevé est retiré : sans cela, chaque
    ' tentative laisserait un UserForm1, UserForm2 … orphelin dans le projet.
    If cree And Not vbComp Is Nothing Then
        On Error Resume Next
        vbProj.VBComponents.Remove vbComp
        On Error GoTo 0
    End If
    MsgBox "La génération a échoué :" & vbCrLf & vbCrLf & _
           Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
           "Si le problème persiste : fermez puis rouvrez le classeur, lancez " & _
           "NettoyerFormulairesOrphelins, et relancez la génération.", _
           vbCritical, "Génération du formulaire"
End Sub

'------------------------------------------------------------------------------
' Pose une propriété du formulaire.
' Les erreurs sont ignorées volontairement : si une version d'Excel n'expose pas
' l'une de ces propriétés, le formulaire se génère quand même, avec la valeur par
' défaut pour celle-là.
'------------------------------------------------------------------------------
Private Sub PropFormulaire(vbComp As Object, ByVal nom As String, ByVal valeur As Variant)
    On Error Resume Next
    vbComp.Properties(nom) = valeur
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Rend un composant UF_Clients prêt à être reconstruit.
'   vbComp : en entrée le composant déjà présent, ou Nothing ; en sortie celui
'            à garnir. Renseigné dès la création, pour que l'appelant puisse le
'            retirer même si la suite échoue.
'   cree   : passé à True si le composant vient d'être créé ici
'
' Trois cas, du plus sûr au plus rustique :
'
'   1. le formulaire n'existe pas    -> on l'ajoute et on le nomme ;
'   2. il existe                     -> on le VIDE et on le réutilise ;
'   3. le vidage échoue              -> on renomme l'ancien, on le supprime,
'                                       puis on ajoute un composant neuf.
'
' Ce qu'on ne fait jamais, c'est supprimer puis recréer sous le même nom dans la
' même exécution : VBA ne libère le nom d'un composant supprimé qu'au retour à
' Excel, et l'affectation du nom échouerait — erreur 75, « objet spécifié
' introuvable ». Le renommage du cas 3 libère le nom immédiatement, lui.
'------------------------------------------------------------------------------
Private Sub PreparerComposant(vbProj As Object, ByRef vbComp As Object, ByRef cree As Boolean)
    cree = False

    ' cas 1 : rien à reprendre, on crée
    If vbComp Is Nothing Then
        Set vbComp = vbProj.VBComponents.Add(CT_MSFORM)
        cree = True
        vbComp.Name = NOM_FORMULAIRE
        Exit Sub
    End If

    ' cas 2 : remise à neuf sur place
    On Error GoTo Repli
    ViderFormulaire vbComp
    On Error GoTo 0
    Exit Sub

    ' cas 3 : le renommage libère le nom tout de suite, la suppression peut
    ' bien attendre le retour à Excel
Repli:
    On Error GoTo 0
    vbComp.Name = NomLibre(vbProj, "UF_Clients_remplace")
    vbProj.VBComponents.Remove vbComp
    Set vbComp = Nothing
    Set vbComp = vbProj.VBComponents.Add(CT_MSFORM)
    cree = True
    vbComp.Name = NOM_FORMULAIRE
End Sub

'------------------------------------------------------------------------------
' Vide un formulaire existant : tous ses contrôles, puis tout son module de code.
'------------------------------------------------------------------------------
Private Sub ViderFormulaire(vbComp As Object)
    Dim dsg As Object, i As Long

    ' une instance restée chargée en mémoire empêcherait la modification
    On Error Resume Next
    For i = UserForms.Count - 1 To 0 Step -1
        If StrComp(UserForms(i).Name, NOM_FORMULAIRE, vbTextCompare) = 0 Then
            Unload UserForms(i)
        End If
    Next i
    On Error GoTo 0

    Set dsg = vbComp.Designer
    For i = dsg.Controls.Count - 1 To 0 Step -1
        dsg.Controls.Remove dsg.Controls(i).Name
    Next i

    With vbComp.CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
    End With
End Sub

'------------------------------------------------------------------------------
' Nom de composant encore disponible dans le projet, dérivé d'une base.
'   renvoie : base, ou base suivie d'un numéro
'------------------------------------------------------------------------------
Private Function NomLibre(vbProj As Object, ByVal base As String) As String
    Dim n As Long, essai As String, pris As Object

    essai = base
    Do
        Set pris = Nothing
        On Error Resume Next
        Set pris = vbProj.VBComponents(essai)
        On Error GoTo 0
        If pris Is Nothing Then
            NomLibre = essai
            Exit Function
        End If
        n = n + 1
        essai = base & CStr(n)
    Loop While n < 200
    NomLibre = base & Format$(Now, "hhnnss")
End Function

'------------------------------------------------------------------------------
' Retire UF_Clients du projet. Utile pour repartir de zéro, ou pour livrer le
' classeur sans le formulaire.
'------------------------------------------------------------------------------
Public Sub SupprimerFormulaireClients()
    Dim vbProj As Object, vbComp As Object

    On Error Resume Next
    Set vbProj = ThisWorkbook.VBProject
    If vbProj Is Nothing Then Exit Sub
    Set vbComp = vbProj.VBComponents(NOM_FORMULAIRE)
    On Error GoTo 0
    If vbComp Is Nothing Then
        MsgBox "Le formulaire " & NOM_FORMULAIRE & " n'existe pas.", vbInformation, "Suppression"
        Exit Sub
    End If
    ' Le composant est d'abord renommé : VBA ne libère le nom d'un composant
    ' supprimé qu'au retour à Excel, et une génération lancée dans la foulée
    ' échouerait sinon sur un nom encore pris.
    On Error Resume Next
    vbComp.Name = NomLibre(vbProj, "UF_Clients_supprime")
    On Error GoTo 0
    vbProj.VBComponents.Remove vbComp
    MsgBox "Formulaire " & NOM_FORMULAIRE & " supprimé.", vbInformation, "Suppression"
End Sub

'------------------------------------------------------------------------------
' Retire les UserForm1, UserForm2 … qu'une génération interrompue a pu laisser
' dans le projet. Ne sont concernés que les composants portant un nom
' automatique, sans aucun contrôle ni la moindre ligne de code : un formulaire
' construit à la main ne peut donc pas être supprimé par erreur.
'------------------------------------------------------------------------------
Public Sub NettoyerFormulairesOrphelins()
    Dim vbProj As Object, vbComp As Object, liste As String, n As Long
    Dim aRetirer As Collection, i As Long

    On Error Resume Next
    Set vbProj = ThisWorkbook.VBProject
    On Error GoTo 0
    If vbProj Is Nothing Then
        MsgBox "Excel refuse l'accès au projet VBA.", vbExclamation, "Nettoyage"
        Exit Sub
    End If

    ' les noms sont relevés d'abord : on ne modifie pas une collection
    ' pendant qu'on la parcourt
    Set aRetirer = New Collection
    For Each vbComp In vbProj.VBComponents
        If EstFormulaireOrphelin(vbComp) Then
            n = n + 1
            liste = liste & vbCrLf & "   " & vbComp.Name
            aRetirer.Add vbComp.Name
        End If
    Next vbComp

    If n = 0 Then
        MsgBox "Aucun formulaire orphelin dans ce projet.", vbInformation, "Nettoyage"
        Exit Sub
    End If

    If MsgBox(n & " formulaire(s) vide(s) trouvé(s) :" & vbCrLf & liste & vbCrLf & vbCrLf & _
              "Les supprimer ?", vbQuestion + vbYesNo, "Nettoyage") <> vbYes Then Exit Sub

    For i = 1 To aRetirer.Count
        On Error Resume Next
        vbProj.VBComponents.Remove vbProj.VBComponents(aRetirer(i))
        On Error GoTo 0
    Next i
    MsgBox n & " formulaire(s) supprimé(s).", vbInformation, "Nettoyage"
End Sub

'------------------------------------------------------------------------------
' True si le composant est un UserForm au nom automatique, sans contrôle ni code.
'------------------------------------------------------------------------------
Private Function EstFormulaireOrphelin(vbComp As Object) As Boolean
    Dim nom As String, suffixe As String, i As Long

    On Error GoTo Fin
    If vbComp.Type <> CT_MSFORM Then Exit Function

    nom = vbComp.Name
    If Len(nom) <= 8 Then Exit Function
    If StrComp(Left$(nom, 8), "UserForm", vbTextCompare) <> 0 Then Exit Function

    suffixe = Mid$(nom, 9)
    For i = 1 To Len(suffixe)
        If Mid$(suffixe, i, 1) < "0" Or Mid$(suffixe, i, 1) > "9" Then Exit Function
    Next i

    If vbComp.Designer.Controls.Count > 0 Then Exit Function
    If vbComp.CodeModule.CountOfLines > 0 Then Exit Function
    EstFormulaireOrphelin = True
Fin:
End Function

'------------------------------------------------------------------------------
' Colonnes de TblClients absentes du schéma du formulaire.
'   renvoie : leurs noms séparés par des virgules, chaîne vide si tout est couvert
' Ce sont exactement les colonnes qui ne seront ni affichées ni saisissables.
'------------------------------------------------------------------------------
Private Function ColonnesNonCouvertes(ByVal lo As ListObject) As String
    Dim lc As ListColumn, ch() As ChampClient, i As Long, trouve As Boolean, res As String

    ch = ObtenirChamps()
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
    ColonnesNonCouvertes = res
End Function

'==============================================================================
' CONSTRUCTION DES CONTROLES
'==============================================================================
'------------------------------------------------------------------------------
' Zone 1 : bandeau de titre, sous-titre d'état et croix de fermeture.
'------------------------------------------------------------------------------
Private Sub ConstruireBandeau(dsg As Object)
    Dim c As Object

    Set c = AjouterControle(dsg, "Forms.Label.1", "lblEntete", 0, 0, F_LARGEUR, BAND_HAUT)
    FondPlein c, COUL_BANDEAU, COUL_BANDEAU

    Set c = AjouterControle(dsg, "Forms.Label.1", "lblTitre", 24, 9, 520, 19)
    TexteLabel c, "Gestion des clients", TAILLE_TITRE, True, COUL_BANDEAU_TXT, MSF_TextAlignLeft

    Set c = AjouterControle(dsg, "Forms.Label.1", "lblSousTitre", 24, 28, 620, 13)
    TexteLabel c, "Tableau " & NOM_TABLE_CLIENTS & " — feuille Clients", _
               TAILLE_SOUSTITRE, False, COUL_BANDEAU_SOUS, MSF_TextAlignLeft

    Set c = AjouterControle(dsg, "Forms.Label.1", "lblFermer", F_LARGEUR - 40, 11, 26, 26)
    TexteLabel c, GlypheFermer(), 14, True, COUL_BANDEAU_TXT, MSF_TextAlignCenter
    c.BackColor = COUL_FERMER_H
    c.ControlTipText = "Fermer le formulaire"
End Sub

'------------------------------------------------------------------------------
' Zone 2 : la fiche client.
'
' Parcourt le schéma et pose, pour chaque champ, son libellé puis son contrôle
' aux coordonnées calculées par GrilleX et GrilleY. L'ordre de tabulation suit
' l'ordre du schéma, les champs verrouillés étant sautés.
'------------------------------------------------------------------------------
Private Sub ConstruireCarteSaisie(dsg As Object)
    Dim c As Object, ch() As ChampClient, i As Long
    Dim x As Single, y As Single, larg As Single, ordreTab As Long

    Set c = AjouterControle(dsg, "Forms.Label.1", "lblCarteSaisie", MARGE, CS_TOP, CARTE_LARG, CS_HAUT)
    Carte c

    Set c = AjouterControle(dsg, "Forms.Label.1", "lblSectionSaisie", GR_X, CS_TOP + 8, 400, 14)
    TexteLabel c, "FICHE CLIENT", TAILLE_SECTION, True, COUL_SECTION, MSF_TextAlignLeft

    ch = ObtenirChamps()
    ordreTab = 1

    For i = LBound(ch) To UBound(ch)
        x = GrilleX(ch(i).Col)
        y = GrilleY(ch(i).Ligne)
        larg = GR_BLOC

        ' Deux champs peuvent partager un bloc — les cases TVA et Forfait :
        ' chacun prend la moitié de la largeur, moins 8 points de gouttière.
        If ch(i).Moitie = 1 Then
            larg = (GR_BLOC - 8) / 2
        ElseIf ch(i).Moitie = 2 Then
            larg = (GR_BLOC - 8) / 2
            x = x + larg + 8
        End If

        If ch(i).TypeCtrl = TYPE_CASE Then
            ' une case à cocher porte son propre libellé
            Set c = AjouterControle(dsg, "Forms.CheckBox.1", NomControle(ch(i)), _
                                    x, y + CH_LBL_HAUT + 1, larg, CH_CTL_HAUT)
            CaseACocher c, ch(i).Libelle
        Else
            Set c = AjouterControle(dsg, "Forms.Label.1", NomLibelle(ch(i)), x, y, larg, CH_LBL_HAUT)
            TexteLabel c, UCase$(ch(i).Libelle), TAILLE_LIBELLE, True, COUL_TEXTE_DOUX, MSF_TextAlignLeft

            If ch(i).TypeCtrl = TYPE_LISTE Then
                Set c = AjouterControle(dsg, "Forms.ComboBox.1", NomControle(ch(i)), _
                                        x, y + CH_LBL_HAUT + 1, larg, CH_CTL_HAUT)
                ZoneListe c, (StrComp(ch(i).Colonne, "Titre", vbTextCompare) = 0)
            Else
                Set c = AjouterControle(dsg, "Forms.TextBox.1", NomControle(ch(i)), _
                                        x, y + CH_LBL_HAUT + 1, larg, CH_CTL_HAUT)
                ZoneTexte c, ch(i).Verrouille
            End If
        End If

        c.ControlTipText = ch(i).Aide
        If Not ch(i).Verrouille Then
            c.TabIndex = ordreTab
            ordreTab = ordreTab + 1
        End If
    Next i

    ' les deux cases à cocher partagent un même libellé de bloc
    Set c = AjouterControle(dsg, "Forms.Label.1", "lblChamp_Facturation", _
                            GrilleX(2), GrilleY(5), GR_BLOC, CH_LBL_HAUT)
    TexteLabel c, "FACTURATION", TAILLE_LIBELLE, True, COUL_TEXTE_DOUX, MSF_TextAlignLeft
End Sub

'------------------------------------------------------------------------------
' Zone 3 : barre de filtrage — libellé, choix de la colonne, zone de
' recherche, lien de réinitialisation et compteur de fiches.
'------------------------------------------------------------------------------
Private Sub ConstruireCarteFiltre(dsg As Object)
    Dim c As Object, y As Single

    Set c = AjouterControle(dsg, "Forms.Label.1", "lblCarteFiltre", MARGE, CF_TOP, CARTE_LARG, CF_HAUT)
    Carte c

    y = CF_TOP + 11

    Set c = AjouterControle(dsg, "Forms.Label.1", "lblFiltreTitre", GR_X, y + 2, 62, 14)
    TexteLabel c, "Filtrer sur", TAILLE_FILTRE, True, COUL_ENTETE_TXT, MSF_TextAlignLeft

    Set c = AjouterControle(dsg, "Forms.ComboBox.1", "cboChampFiltre", GR_X + 72, y, 124, CH_CTL_HAUT)
    ZoneListe c, True
    c.ControlTipText = "Colonne du tableau sur laquelle porte le filtre"

    Set c = AjouterControle(dsg, "Forms.TextBox.1", "txtFiltre", GR_X + 204, y, 284, CH_CTL_HAUT)
    ZoneTexte c, False
    c.ControlTipText = "Texte à rechercher (accents et majuscules sont ignorés)"

    Set c = AjouterControle(dsg, "Forms.Label.1", "lblResetFiltre", GR_X + 496, y + 2, 80, 14)
    TexteLabel c, "Réinitialiser", TAILLE_FILTRE, False, COUL_LIEN, MSF_TextAlignLeft
    c.ControlTipText = "Effacer le filtre et réafficher toutes les fiches"

    Set c = AjouterControle(dsg, "Forms.Label.1", "lblCompteur", GR_X + 584, y + 2, 200, 14)
    TexteLabel c, vbNullString, TAILLE_FILTRE, False, COUL_TEXTE_DOUX, MSF_TextAlignRight
End Sub

'------------------------------------------------------------------------------
' Zone 4 : tableau des enregistrements.
'
' La ListBox MSForms ne sait pas afficher d'en-têtes de colonnes autrement qu'en
' étant liée à une plage de cellules. Les en-têtes sont donc de simples libellés
' posés au-dessus d'elle, alignés sur les mêmes largeurs — ce qui permet en prime
' de les rendre cliquables pour le tri.
'------------------------------------------------------------------------------
Private Sub ConstruireCarteTableau(dsg As Object)
    Dim c As Object, larg As Variant, lib As Variant
    Dim i As Long, x As Single, gauche As Single, largeur As Single
    Dim colw As String, hautListe As Single

    gauche = MARGE + 1
    largeur = CARTE_LARG - 2

    Set c = AjouterControle(dsg, "Forms.Label.1", "lblCarteTable", MARGE, CT_TOP, CARTE_LARG, CT_HAUT)
    Carte c

    Set c = AjouterControle(dsg, "Forms.Label.1", "lblEnteteTable", gauche, CT_TOP + 1, largeur, CT_ENTETE - 1)
    FondPlein c, COUL_ENTETE_TBL, COUL_ENTETE_TBL

    larg = LargeursListe()
    lib = LibellesListe()
    colw = vbNullString
    x = gauche + 3

    For i = LBound(larg) To UBound(larg)
        Set c = AjouterControle(dsg, "Forms.Label.1", "lblEnt_" & CStr(i + 1), _
                                x, CT_TOP + 5, CSng(larg(i)) - 3, 13)
        TexteLabel c, CStr(lib(i)), TAILLE_ENTETE, True, COUL_ENTETE_TXT, MSF_TextAlignLeft
        c.ControlTipText = "Cliquez pour trier sur cette colonne"
        ' Chaque en-tête est posé à l'abscisse cumulée des colonnes qui le
        ' précèdent : ils restent alignés sur la ListBox quelles que soient les
        ' largeurs choisies dans LargeursListe.
        x = x + CSng(larg(i))
        colw = colw & IIf(Len(colw) > 0, ";", "") & CStr(larg(i)) & " pt"
    Next i

    hautListe = CT_TOP + CT_HAUT - 2 - (CT_TOP + CT_ENTETE + 1)
    Set c = AjouterControle(dsg, "Forms.ListBox.1", "lstClients", _
                            gauche, CT_TOP + CT_ENTETE + 1, largeur, hautListe)
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
    ' IntegralHeight n'est modifiable qu'à la conception ; l'affectation est
    ' protégée au cas où une version d'Excel la refuserait ici.
    On Error Resume Next
    c.IntegralHeight = False
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Zone 5 : les cinq boutons d'action.
' Les quatre premiers sont alignés à gauche, Quitter est renvoyé à droite pour
' qu'on ne le clique pas par mégarde.
'------------------------------------------------------------------------------
Private Sub ConstruireBoutons(dsg As Object)
    Dim c As Object, x As Single

    x = MARGE
    Set c = AjouterControle(dsg, "Forms.CommandButton.1", "btnAjouter", x, BT_TOP, BT_LARG, BT_HAUT)
    Bouton c, "Ajouter", "A", COUL_AJOUTER, 91
    c.ControlTipText = "Créer une nouvelle fiche à partir des champs saisis"

    x = x + BT_LARG + BT_GOUTTIERE
    Set c = AjouterControle(dsg, "Forms.CommandButton.1", "btnModifier", x, BT_TOP, BT_LARG, BT_HAUT)
    Bouton c, "Modifier", "M", COUL_MODIFIER, 92
    c.ControlTipText = "Enregistrer les modifications sur la fiche sélectionnée"

    x = x + BT_LARG + BT_GOUTTIERE
    Set c = AjouterControle(dsg, "Forms.CommandButton.1", "btnSupprimer", x, BT_TOP, BT_LARG, BT_HAUT)
    Bouton c, "Supprimer", "S", COUL_SUPPRIMER, 93
    c.ControlTipText = "Supprimer la fiche sélectionnée du tableau"

    x = x + BT_LARG + BT_GOUTTIERE
    Set c = AjouterControle(dsg, "Forms.CommandButton.1", "btnEffacer", x, BT_TOP, BT_LARG, BT_HAUT)
    Bouton c, "Effacer", "E", COUL_EFFACER, 94
    c.ControlTipText = "Vider les zones de saisie sans toucher au tableau"

    Set c = AjouterControle(dsg, "Forms.CommandButton.1", "btnQuitter", _
                            F_LARGEUR - MARGE - BT_LARG, BT_TOP, BT_LARG, BT_HAUT)
    Bouton c, "Quitter", "Q", COUL_QUITTER, 95
    c.Cancel = True
    c.ControlTipText = "Fermer le formulaire (touche Échap)"
End Sub

'------------------------------------------------------------------------------
' Renvoie les fonds de cartes derrière leur contenu.
'
' MSForms place chaque contrôle nouvellement créé au premier plan : les cartes,
' créées avant leur contenu, se retrouveraient donc devant lui. Elles sont
' reculées ici dans l'ordre inverse de leur profondeur.
'------------------------------------------------------------------------------
Private Sub ReculerArrierePlans(dsg As Object)
    Dim noms As Variant, i As Long, c As Object

    noms = Array("lblEnteteTable", "lblCarteTable", "lblCarteFiltre", "lblCarteSaisie", "lblEntete")
    For i = LBound(noms) To UBound(noms)
        On Error Resume Next
        Set c = dsg.Controls(CStr(noms(i)))
        If Not c Is Nothing Then c.ZOrder MSF_ZOrderBack
        Set c = Nothing
        On Error GoTo 0
    Next i
End Sub

'==============================================================================
' MISE EN FORME
'==============================================================================
'------------------------------------------------------------------------------
' Crée un contrôle et le positionne.
'   progId  : Forms.Label.1, Forms.TextBox.1, Forms.ComboBox.1, Forms.CheckBox.1,
'             Forms.ListBox.1 ou Forms.CommandButton.1
'   renvoie : le contrôle créé, pour la suite de sa mise en forme
'------------------------------------------------------------------------------
Private Function AjouterControle(dsg As Object, ByVal progId As String, ByVal nom As String, _
                                 ByVal gauche As Single, ByVal haut As Single, _
                                 ByVal largeur As Single, ByVal hauteur As Single) As Object
    Dim c As Object
    Set c = dsg.Controls.Add(progId, nom, True)
    c.Left = gauche
    c.Top = haut
    c.Width = largeur
    c.Height = hauteur
    Set AjouterControle = c
End Function

'------------------------------------------------------------------------------
' Fond de carte : rectangle blanc à filet fin.
'
' SpecialEffect doit être posé avant BorderStyle : MSForms refuse une bordure
' simple tant que le contrôle est encore en relief. La même règle vaut pour tous
' les contrôles habillés plus bas.
'------------------------------------------------------------------------------
Private Sub Carte(c As Object)
    c.Caption = vbNullString
    c.BackStyle = MSF_BackStyleOpaque
    c.BackColor = COUL_CARTE
    c.SpecialEffect = MSF_SpecialEffectFlat
    c.BorderStyle = MSF_BorderStyleSingle
    c.BorderColor = COUL_BORDURE
End Sub

'------------------------------------------------------------------------------
' Rectangle de couleur unie : bandeau de titre, bande d'en-tête du tableau.
'------------------------------------------------------------------------------
Private Sub FondPlein(c As Object, ByVal fond As Long, ByVal bordure As Long)
    c.Caption = vbNullString
    c.BackStyle = MSF_BackStyleOpaque
    c.BackColor = fond
    c.SpecialEffect = MSF_SpecialEffectFlat
    c.BorderStyle = MSF_BorderStyleSingle
    c.BorderColor = bordure
End Sub

'------------------------------------------------------------------------------
' Habille un libellé : texte, taille, graisse, couleur, alignement.
' Le fond est transparent pour laisser voir la carte ou le bandeau au-dessous.
'------------------------------------------------------------------------------
Private Sub TexteLabel(c As Object, ByVal texte As String, ByVal taille As Single, _
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
' Habille une zone de saisie : filet fin, fond très clair.
'   verrouille : True pour un champ géré par le programme — fond gris, texte
'                atténué, exclu de la tabulation
'
' Locked plutôt qu'Enabled = False : le champ reste lisible et son contenu
' copiable, alors qu'un contrôle désactivé s'affiche en gris illisible.
'------------------------------------------------------------------------------
Private Sub ZoneTexte(c As Object, ByVal verrouille As Boolean)
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
' Habille un menu déroulant.
'   choixImpose : True pour une liste fermée (Titre, colonne de filtrage),
'                 False pour une liste ouverte où la saisie libre reste possible
'                 (Adresse, NPA, Texte de facture), toutes les valeurs n'étant
'                 pas forcément déjà répertoriées.
'------------------------------------------------------------------------------
Private Sub ZoneListe(c As Object, ByVal choixImpose As Boolean)
    c.Font.Name = POLICE
    c.Font.Size = TAILLE_CHAMP
    c.SpecialEffect = MSF_SpecialEffectFlat
    c.BorderStyle = MSF_BorderStyleSingle
    c.BorderColor = COUL_CHAMP_BORD
    c.BackColor = COUL_CHAMP_FOND
    c.ForeColor = COUL_TEXTE
    c.ListRows = 12
    c.MatchEntry = MSF_MatchEntryComplete
    c.Style = IIf(choixImpose, MSF_StyleDropDownList, MSF_StyleDropDownCombo)
End Sub

'------------------------------------------------------------------------------
' Habille une case à cocher.
' Le relief de la case elle-même est conservé : à plat, elle deviendrait
' difficile à distinguer d'un simple libellé.
'------------------------------------------------------------------------------
Private Sub CaseACocher(c As Object, ByVal libelle As String)
    c.Caption = libelle
    c.BackStyle = MSF_BackStyleTransparent
    c.ForeColor = COUL_TEXTE
    c.Font.Name = POLICE
    c.Font.Size = TAILLE_FILTRE
    c.Font.Bold = False
    c.WordWrap = False
End Sub

'------------------------------------------------------------------------------
' Habille un bouton d'action : couleur de fond, texte blanc, lettre de
' raccourci soulignée et rang de tabulation.
'------------------------------------------------------------------------------
Private Sub Bouton(c As Object, ByVal libelle As String, ByVal raccourci As String, _
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

'==============================================================================
' MODULE DE CODE DU FORMULAIRE
'------------------------------------------------------------------------------
' Chaque procédure événementielle se contente d'appeler modClients_Formulaire.
'==============================================================================
'------------------------------------------------------------------------------
' Écrit le module de code du formulaire et le renvoie sous forme de texte.
'
' Chaque procédure événementielle tient en une ligne et se contente d'appeler
' modClients_Formulaire. Le code métier n'est donc jamais dupliqué ici, et
' régénérer le formulaire ne peut rien écraser d'écrit à la main.
'
' Les gestionnaires des champs sont produits en parcourant le schéma : ajouter un
' champ crée automatiquement ses événements Enter, Exit et, s'il est numérique,
' son filtre de frappe.
'------------------------------------------------------------------------------
Private Function CodeDuFormulaire() As String
    Dim ch() As ChampClient, i As Long, n As String
    Dim boutons As Variant, larg As Variant

    mCode = vbNullString

    L "Option Explicit"
    L "'=============================================================================="
    L "' Module de code du formulaire " & NOM_FORMULAIRE
    L "'------------------------------------------------------------------------------"
    L "' GÉNÉRÉ AUTOMATIQUEMENT par modClients_Generateur.GenererFormulaireClients."
    L "' Ne rien écrire ici : toute la logique se trouve dans modClients_Formulaire,"
    L "' et ce module est écrasé à chaque régénération du formulaire."
    L "'=============================================================================="
    L ""

    ' --- formulaire -----------------------------------------------------------
    Proc "UserForm_Initialize()", "Clients_Initialiser Me"
    Proc "UserForm_Activate()", "Clients_Activer Me"
    Proc "UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)", "Clients_Fermeture Me"
    Proc "UserForm_MouseMove" & SIG_SOURIS, "Clients_Survol Me, " & Q("")

    ' --- bandeau : déplacement de la fenêtre ---------------------------------
    EmettreZoneDeplacement "lblEntete"
    EmettreZoneDeplacement "lblTitre"
    EmettreZoneDeplacement "lblSousTitre"

    Proc "lblFermer_Click()", "Clients_Quitter Me"
    Proc "lblFermer_MouseMove" & SIG_SOURIS, "Clients_Survol Me, " & Q("lblFermer")

    ' --- boutons d'action -----------------------------------------------------
    boutons = Array("Ajouter", "Modifier", "Supprimer", "Effacer", "Quitter")
    For i = LBound(boutons) To UBound(boutons)
        n = "btn" & CStr(boutons(i))
        Proc n & "_Click()", "Clients_" & CStr(boutons(i)) & " Me"
        Proc n & "_MouseMove" & SIG_SOURIS, "Clients_Survol Me, " & Q(n)
    Next i

    ' --- filtrage -------------------------------------------------------------
    Proc "cboChampFiltre_Change()", "Clients_AppliquerFiltre Me"
    Proc "txtFiltre_Change()", "Clients_AppliquerFiltre Me"
    Proc "lblResetFiltre_Click()", "Clients_ReinitialiserFiltre Me"
    Proc "lblResetFiltre_MouseMove" & SIG_SOURIS, "Clients_Survol Me, " & Q("lblResetFiltre")

    ' --- tableau des enregistrements ------------------------------------------
    Proc "lstClients_Click()", "Clients_ChargerSelection Me"
    Proc "lstClients_DblClick(ByVal Cancel As MSForms.ReturnBoolean)", "Clients_ChargerSelection Me"

    larg = LargeursListe()
    For i = LBound(larg) To UBound(larg)
        Proc "lblEnt_" & CStr(i + 1) & "_Click()", "Clients_TrierColonne Me, " & CStr(i + 1)
    Next i

    ' --- zones de saisie ------------------------------------------------------
    ch = ObtenirChamps()
    For i = LBound(ch) To UBound(ch)
        If Not ch(i).Verrouille And ch(i).TypeCtrl <> TYPE_CASE Then
            n = NomControle(ch(i))
            Proc n & "_Enter()", "Clients_FocusChamp Me, " & Q(n) & ", True"
            Proc n & "_Exit(ByVal Cancel As MSForms.ReturnBoolean)", _
                 "Clients_FocusChamp Me, " & Q(n) & ", False"

            If StrComp(ch(i).Colonne, COL_ADRESSE, vbTextCompare) = 0 Then
                Proc n & "_Change()", "Clients_AdresseChoisie Me"
            ElseIf StrComp(ch(i).Colonne, COL_NPA, vbTextCompare) = 0 Then
                Proc n & "_Change()", "Clients_NpaChoisi Me"
            End If

            If ch(i).Numerique <> NUM_NON Then
                Proc n & "_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)", _
                     "Clients_ToucheNumerique KeyAscii, " & IIf(ch(i).Numerique = NUM_DECIMAL, "True", "False")
            End If
        End If
    Next i

    CodeDuFormulaire = mCode
End Function

'------------------------------------------------------------------------------
' Écrit les trois gestionnaires qui rendent une zone du bandeau saisissable
' pour déplacer la fenêtre.
'------------------------------------------------------------------------------
Private Sub EmettreZoneDeplacement(ByVal nom As String)
    Proc nom & "_MouseDown" & SIG_SOURIS, "Clients_DebutDeplacement X, Y"
    Proc nom & "_MouseMove" & SIG_SOURIS, "Clients_Deplacer Me, Button, X, Y" & vbNewLine & _
         "    Clients_Survol Me, " & Q("")
    Proc nom & "_MouseUp" & SIG_SOURIS, "Clients_FinDeplacement"
End Sub

'------------------------------------------------------------------------------
' Écrit une procédure événementielle complète.
'   entete : signature, sans le mot-clef Private Sub
'   corps  : le ou les appels à placer dedans
'
' Trois lignes séparées, jamais une seule ligne à deux-points : VBA n'accepte pas
' qu'une déclaration de procédure et son End Sub partagent une ligne.
'------------------------------------------------------------------------------
Private Sub Proc(ByVal entete As String, ByVal corps As String)
    L "Private Sub " & entete
    L "    " & corps
    L "End Sub"
    L ""
End Sub

'------------------------------------------------------------------------------
' Ajoute une ligne au code en cours d'écriture.
'------------------------------------------------------------------------------
Private Sub L(ByVal ligne As String)
    mCode = mCode & ligne & vbNewLine
End Sub

'------------------------------------------------------------------------------
' Entoure un texte de guillemets doubles.
' Évite d'avoir à doubler les guillemets dans le code générateur, où ils
' deviendraient vite illisibles.
'------------------------------------------------------------------------------
Private Function Q(ByVal s As String) As String
    Q = Chr$(34) & s & Chr$(34)
End Function
