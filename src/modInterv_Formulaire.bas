Attribute VB_Name = "modInterv_Formulaire"
Option Explicit
'==============================================================================
' modInterv_Formulaire
'------------------------------------------------------------------------------
' Comportement du formulaire UF_Interventions.
'
' Comme pour le formulaire des clients, le module de code du UserForm ne
' contient que des procédures événementielles d'une ligne qui appellent les
' procédures publiques ci-dessous : la logique reste dans un module standard, le
' projet compile avant même que le formulaire existe, et le régénérer n'écrase
' jamais de code métier.
'
' Le formulaire est toujours reçu en Object — liaison tardive — ce qui évite
' toute dépendance de compilation au UserForm.
'==============================================================================

Private mNoCourant As String        ' NoInterv de la fiche affichée ("" = nouvelle)
Private mLignesVis() As Long        ' lignes du cache visibles dans le tableau
Private mNbVis As Long

' État de la grille. mLigneSel et mSurvolGrille comptent dans mLignesVis, de 1 à
' mNbVis, et valent 0 quand rien n'est choisi ni survolé ; mPremiereLigne est le
' décalage du défilement, 0 quand la grille est en haut.
Private mPremiereLigne As Long
Private mLigneSel As Long
Private mSurvolGrille As Long
Private mTriColonne As Long
Private mTriDecroissant As Boolean
Private mChargement As Boolean      ' True pendant un remplissage programmé
Private mAjuste As Boolean          ' True une fois la fenêtre redimensionnée

'==============================================================================
' INITIALISATION
'==============================================================================
'------------------------------------------------------------------------------
' Prépare le formulaire : données, listes déroulantes, en-têtes, statistiques,
' puis affichage du tableau. Appelée une fois, depuis UserForm_Initialize.
'------------------------------------------------------------------------------
Public Sub Interv_Initialiser(f As Object)
    Dim i As Long, lib As Variant

    On Error GoTo Erreur
    mChargement = True
    mNoCourant = vbNullString
    mTriColonne = 0
    mTriDecroissant = False

    Interv_ToutRecharger
    Interv_Charger

    ' --- fiche 1 : intitulé ---------------------------------------------------
    ICtl(f, "lblTitreGlobal").Caption = EnTexte(Interv_CelluleNommee(CEL_TITRE))
    ICtl(f, "lblAnnee").Caption = AnneeAffichee()

    ' --- fiche 2 : statistiques -----------------------------------------------
    Interv_MajStatistiques f

    ' --- listes déroulantes ---------------------------------------------------
    RemplirCombo ICtl(f, INomControleColonne(IC_ENTREPRISE)), Interv_ListeClients("Entreprise")
    RemplirCombo ICtl(f, INomControleColonne(IC_NOM)), Interv_ListeClients("Nom")
    RemplirCombo ICtl(f, INomControleColonne(IC_TITRE)), Interv_ListeClients("Titre")
    RemplirCombo ICtl(f, INomControleColonne(IC_TEXTE)), Adresses_TextesFacture()

    With ICtl(f, "cboChampFiltreI")
        .Clear
        lib = IChampsFiltrables()
        For i = LBound(lib) To UBound(lib)
            .AddItem lib(i)
        Next i
        .ListIndex = 0
    End With

    ' Douze mois précédés de « Tous les mois » : le premier élément ne filtre
    ' rien, ce qui donne au menu un état de repos plutôt qu'une case vide.
    With ICtl(f, "cboMoisI")
        .Clear
        .AddItem TOUS_LES_MOIS
        For i = 1 To 12
            .AddItem NomDuMois(i)
        Next i
        .ListIndex = 0
    End With

    ' --- fiche 4 : en-têtes du tableau ----------------------------------------
    MajEntetesInterv f
    mPremiereLigne = 0
    mLigneSel = 0
    mSurvolGrille = 0

    mChargement = False
    ViderChampsInterv f
    Interv_RafraichirListe f
    MajBoutonsInterv f
    MajEtatInterv f
    Exit Sub

Erreur:
    mChargement = False
    MsgBox "Initialisation impossible :" & vbCrLf & vbCrLf & Err.Description, _
           vbExclamation, "Interventions"
End Sub

'------------------------------------------------------------------------------
' Ajuste la fenêtre à la première activation pour que sa surface utile mesure
' exactement I_LARGEUR x I_HAUTEUR.
'
' Width et Height d'un UserForm désignent les dimensions EXTÉRIEURES, barre de
' titre comprise : sans cette correction, le bas du formulaire — la rangée de
' boutons — se retrouverait rogné d'une vingtaine de points.
'------------------------------------------------------------------------------
Public Sub Interv_Activer(f As Object)
    If mAjuste Then Exit Sub
    mAjuste = True
    On Error Resume Next
    f.Width = f.Width + (I_LARGEUR - f.InsideWidth)
    f.Height = f.Height + (I_HAUTEUR - f.InsideHeight)
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Année de référence, lue dans la cellule nommée AnneeEnCours et affichée sur
' quatre chiffres.
'------------------------------------------------------------------------------
Private Function AnneeAffichee() As String
    Dim v As Variant
    v = Interv_CelluleNommee(CEL_ANNEE)
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If VarType(v) = vbDate Then
        AnneeAffichee = Format$(v, "yyyy")
    ElseIf IsNumeric(v) Then
        ' un nombre à quatre chiffres est déjà une année ; au-delà c'est une
        ' date au format numérique d'Excel
        If CDbl(v) > 3000 Then
            AnneeAffichee = Format$(CDate(CDbl(v)), "yyyy")
        Else
            AnneeAffichee = Format$(CLng(v), "0000")
        End If
    Else
        AnneeAffichee = CStr(v)
    End If
End Function

'------------------------------------------------------------------------------
' Remplit un menu déroulant à partir d'un tableau de valeurs.
' Ne fait rien si le contrôle ou le tableau est absent : le formulaire reste
' utilisable même si une des feuilles sources manque.
'------------------------------------------------------------------------------
Private Sub RemplirCombo(cbo As Object, ByVal valeurs As Variant)
    Dim i As Long
    If cbo Is Nothing Then Exit Sub
    cbo.Clear
    If Not IsArray(valeurs) Then Exit Sub
    On Error Resume Next
    For i = LBound(valeurs) To UBound(valeurs)
        cbo.AddItem valeurs(i)
    Next i
    On Error GoTo 0
End Sub

'==============================================================================
' FICHE 2 : STATISTIQUES
'==============================================================================
'------------------------------------------------------------------------------
' Rafraîchit le graphique et les six tuiles de chiffres.
'
' Chaque tuile reçoit son image de fond — celle qui lui donne ses coins
' arrondis — puis le montant lu dans sa cellule nommée. L'image n'est chargée
' qu'une fois et partagée par les six contrôles.
'
' Le graphique, lui, n'est plus une image : modInterv_Graphique le dessine en
' contrôles, à partir des données de la feuille Statistiques. Il reste donc net
' quelle que soit sa taille et ne dépend plus ni d'un export ni d'un fichier
' temporaire.
'------------------------------------------------------------------------------
Public Sub Interv_MajStatistiques(f As Object)
    Dim chemin As String, img As Object, lb As Object
    Dim tuiles As Variant, i As Long, fond As Object

    tuiles = TuilesStatistiques()
    chemin = Interv_CheminImageTuile()

    ' --- tuiles : image de fond, puis montant ---------------------------------
    If Len(chemin) > 0 Then
        On Error Resume Next
        Set fond = LoadPicture(chemin)
        On Error GoTo 0
    End If

    For i = LBound(tuiles) To UBound(tuiles)
        If Not fond Is Nothing Then
            Set img = ICtl(f, "imgTuile_" & CStr(i + 1))
            If Not img Is Nothing Then
                On Error Resume Next
                Set img.Picture = fond
                On Error GoTo 0
            End If
        End If
        Set lb = ICtl(f, "lblStatVal_" & CStr(i + 1))
        If Not lb Is Nothing Then lb.Caption = Interv_MontantNomme(CStr(tuiles(i)(1)))
    Next i

    ' --- graphique ------------------------------------------------------------
    ' Graph_Tracer se charge de tout, y compris de dire dans sa légende quand les
    ' données manquent : rien à surveiller ici.
    Graph_Tracer f
End Sub

'------------------------------------------------------------------------------
' Les six indicateurs affichés : libellé, puis nom de la cellule Excel.
'
' Attention au nom de la sixième : la demande mentionnait CATotalMmoisActuel,
' déjà employé par « CA mois actuel ». La cellule réellement définie dans le
' classeur pour le mois actuel non facturé est CATotalMoisActuelNonFacture,
' c'est donc celle-là qui est lue.
'------------------------------------------------------------------------------
Public Function TuilesStatistiques() As Variant
    ' Libellés courts : les tuiles sont étroites, et la grille se lit d'elle-même
    ' — la première ligne donne les totaux, la seconde les mois.
    TuilesStatistiques = Array( _
        Array("CA total", "CATotal"), _
        Array("Facturé", "CATotalFacture"), _
        Array("Non facturé", "CATotalNonFacture"), _
        Array("Mois précédent", "CATotalMmoisPrecedent"), _
        Array("Mois actuel", "CATotalMmoisActuel"), _
        Array("Mois non facturé", "CATotalMoisActuelNonFacture"))
End Function

'==============================================================================
' FICHE 4 : TABLEAU DES ENREGISTREMENTS
'==============================================================================
Private Sub MajEntetesInterv(f As Object)
    Dim lib As Variant, i As Long, lb As Object, txt As String

    lib = ILibellesListe()
    For i = LBound(lib) To UBound(lib)
        Set lb = ICtl(f, "lblEntI_" & CStr(i + 1))
        If Not lb Is Nothing Then
            txt = CStr(lib(i))
            If mTriColonne = i + 1 Then
                txt = txt & IIf(mTriDecroissant, GlypheTriDecroissant(), GlypheTriCroissant())
            End If
            lb.Caption = txt
        End If
    Next i
End Sub

'------------------------------------------------------------------------------
' Reconstruit le tableau : filtre, tri, puis affichage.
'
' Deux tableaux sont tenus en parallèle : ce que voit l'utilisateur, et
' mLignesVis, qui garde pour chaque ligne affichée son numéro dans le cache.
' C'est ce second tableau qui sert à retrouver la fiche sélectionnée, si bien
' que les colonnes affichées peuvent changer sans rien casser.
'------------------------------------------------------------------------------
Public Sub Interv_RafraichirListe(f As Object)
    Dim nb As Long, i As Long, champ As String, cible As String, mois As Long
    Dim garde As Boolean

    On Error GoTo Erreur
    nb = Interv_NbLignes()
    ReDim mLignesVis(1 To IIf(nb > 0, nb, 1))
    mNbVis = 0

    champ = EnTexte(ICtl(f, "cboChampFiltreI").Value)
    If Len(champ) = 0 Then champ = IC_ENTREPRISE
    cible = Normaliser(Trim$(EnTexte(ICtl(f, "txtFiltreI").Text)))
    mois = MoisFiltre(f)

    For i = 1 To nb
        If Len(cible) = 0 Then
            garde = True
        Else
            garde = (InStr(1, Normaliser(Interv_ValeurAffichee(i, champ)), cible, vbBinaryCompare) > 0)
        End If
        ' les deux filtres se cumulent : le texte ET le mois
        If garde And mois > 0 Then garde = (MoisDeLaLigne(i) = mois)
        If garde Then
            mNbVis = mNbVis + 1
            mLignesVis(mNbVis) = i
        End If
    Next i

    If mTriColonne > 0 And mNbVis > 1 Then TrierVisibles

    ' le filtre a changé : on repart du haut, et rien n'est plus sélectionné
    If mPremiereLigne > MaxDefilement() Then mPremiereLigne = MaxDefilement()
    mLigneSel = 0
    mSurvolGrille = 0
    MajBarreGrille f
    Interv_PeindreGrille f

    MajCompteurInterv f
    If Len(mNoCourant) > 0 Then SelectionnerNo f, mNoCourant
    Exit Sub

Erreur:
    MsgBox "Affichage de la liste impossible :" & vbCrLf & vbCrLf & Err.Description, _
           vbExclamation, "Interventions"
End Sub

'------------------------------------------------------------------------------
' Relance le filtrage. Appelée à chaque frappe dans la zone de filtre et à
' chaque changement de colonne.
' Ne fait rien pendant un remplissage programmé, pour ne pas se déclencher quand
' le code écrit lui-même dans les contrôles.
'------------------------------------------------------------------------------
Public Sub Interv_AppliquerFiltre(f As Object)
    If mChargement Then Exit Sub
    Interv_RafraichirListe f
End Sub

'------------------------------------------------------------------------------
' Vide la zone de filtre et revient à la première colonne filtrable.
' Déclenchée par le lien Réinitialiser de la barre de filtrage.
'------------------------------------------------------------------------------
Public Sub Interv_ReinitialiserFiltre(f As Object)
    Dim garde As Boolean
    garde = mChargement
    mChargement = True
    ICtl(f, "txtFiltreI").Text = vbNullString
    ICtl(f, "cboChampFiltreI").ListIndex = 0
    ICtl(f, "cboMoisI").ListIndex = 0
    mChargement = garde
    Interv_RafraichirListe f
End Sub

'------------------------------------------------------------------------------
' Trie le tableau sur une colonne.
'   colonne : 1 à 10, position dans IColonnesListe
' Un clic sur la colonne déjà triée inverse simplement le sens.
'------------------------------------------------------------------------------
Public Sub Interv_TrierColonne(f As Object, ByVal colonne As Long)
    If mTriColonne = colonne Then
        mTriDecroissant = Not mTriDecroissant
    Else
        mTriColonne = colonne
        mTriDecroissant = False
    End If
    MajEntetesInterv f
    Interv_RafraichirListe f
End Sub

'------------------------------------------------------------------------------
' Trie mLignesVis sur la colonne active. Seul l'ordre des numéros de ligne
' change : ni le cache ni le tableau Excel ne sont réordonnés.
'------------------------------------------------------------------------------
Private Sub TrierVisibles()
    Dim cols As Variant
    cols = IColonnesListe()
    If mTriColonne < 1 Or mTriColonne > (UBound(cols) - LBound(cols) + 1) Then Exit Sub
    TriRapideVis 1, mNbVis, CStr(cols(LBound(cols) + mTriColonne - 1))
End Sub

'------------------------------------------------------------------------------
' Tri rapide sur le tableau des lignes visibles.
'------------------------------------------------------------------------------
Private Sub TriRapideVis(ByVal g As Long, ByVal d As Long, ByVal nomCol As String)
    Dim i As Long, j As Long, pivot As Long, tmp As Long
    i = g: j = d
    pivot = mLignesVis((g + d) \ 2)
    Do While i <= j
        Do While ComparerVis(mLignesVis(i), pivot, nomCol) < 0
            i = i + 1
        Loop
        Do While ComparerVis(mLignesVis(j), pivot, nomCol) > 0
            j = j - 1
        Loop
        If i <= j Then
            tmp = mLignesVis(i): mLignesVis(i) = mLignesVis(j): mLignesVis(j) = tmp
            i = i + 1: j = j - 1
        End If
    Loop
    If g < j Then TriRapideVis g, j, nomCol
    If i < d Then TriRapideVis i, d, nomCol
End Sub

'------------------------------------------------------------------------------
' Compare deux interventions sur une colonne.
'   renvoie : -1, 0 ou 1, sens de tri déjà appliqué
'
' Les dates sont comparées chronologiquement et les nombres numériquement : un
' tri alphabétique placerait 100 avant 20. Les valeurs manquantes sont
' renvoyées en fin de liste dans les deux sens.
'------------------------------------------------------------------------------
Private Function ComparerVis(ByVal a As Long, ByVal b As Long, ByVal nomCol As String) As Long
    Dim va As Variant, vb As Variant, r As Long, sa As String, sb As String

    ' Une case qui réunit deux colonnes n'a pas de valeur brute : on la compare
    ' sur le texte affiché, « Aiello Rosalba », donc par nom puis par prénom.
    If IColonneComposee(nomCol) Then
        sa = Normaliser(Interv_ValeurListe(a, nomCol))
        sb = Normaliser(Interv_ValeurListe(b, nomCol))
        If Len(sa) = 0 And Len(sb) > 0 Then
            r = 1
        ElseIf Len(sb) = 0 And Len(sa) > 0 Then
            r = -1
        ElseIf sa < sb Then
            r = -1
        ElseIf sa > sb Then
            r = 1
        End If
        If mTriDecroissant Then r = -r
        ComparerVis = r
        Exit Function
    End If

    va = Interv_Valeur(a, nomCol)
    vb = Interv_Valeur(b, nomCol)

    If VarType(va) = vbDate And VarType(vb) = vbDate Then
        If CDbl(CDate(va)) < CDbl(CDate(vb)) Then
            r = -1
        ElseIf CDbl(CDate(va)) > CDbl(CDate(vb)) Then
            r = 1
        End If
    ElseIf IsNumeric(va) And IsNumeric(vb) Then
        If CDbl(va) < CDbl(vb) Then
            r = -1
        ElseIf CDbl(va) > CDbl(vb) Then
            r = 1
        End If
    Else
        sa = Normaliser(Interv_ValeurAffichee(a, nomCol))
        sb = Normaliser(Interv_ValeurAffichee(b, nomCol))
        If Len(sa) = 0 And Len(sb) > 0 Then
            r = 1
        ElseIf Len(sb) = 0 And Len(sa) > 0 Then
            r = -1
        ElseIf sa < sb Then
            r = -1
        ElseIf sa > sb Then
            r = 1
        End If
    End If

    If mTriDecroissant Then r = -r
    ComparerVis = r
End Function

'==============================================================================
' GRILLE DES ENREGISTREMENTS
'------------------------------------------------------------------------------
' Le tableau est fait de libellés, un par cellule, et seules IGR_NB_LIGNES
' lignes existent. Défiler ne déplace aucun contrôle : cela change la première
' ligne affichée, puis repeint les mêmes cases. Un tableau de milliers de lignes
' tient donc dans deux cents contrôles, et l'affichage coûte le même prix quelle
' que soit la taille du tableau.
'==============================================================================
Public Sub Interv_PeindreGrille(f As Object)
    Dim r As Long
    For r = 1 To IGR_NB_LIGNES
        PeindreLigneGrille f, r
    Next r
End Sub

'------------------------------------------------------------------------------
' Repeint une ligne de la grille.
'   r : rang à l'écran, de 1 à IGR_NB_LIGNES
'
' Quatre aspects, dans cet ordre de priorité : la ligne choisie, la ligne
' survolée, une ligne sur deux teintée, le fond ordinaire. Le zébrage suit le
' rang RÉEL et non le rang à l'écran, sans quoi les bandes sauteraient d'une
' ligne à chaque cran de défilement.
'------------------------------------------------------------------------------
Private Sub PeindreLigneGrille(f As Object, ByVal r As Long)
    Dim cols As Variant, c As Long, nbCol As Long, ligne As Long
    Dim lb As Object, fond As Long, encre As Long, vide As Boolean

    If r < 1 Or r > IGR_NB_LIGNES Then Exit Sub

    cols = IColonnesListe()
    nbCol = UBound(cols) - LBound(cols) + 1
    ligne = mPremiereLigne + r
    vide = (ligne < 1 Or ligne > mNbVis)

    If vide Then
        fond = COUL_CARTE
        encre = COUL_CARTE
    ElseIf ligne = mLigneSel Then
        fond = COUL_MODIFIER
        encre = COUL_BOUTON_TXT
    ElseIf ligne = mSurvolGrille Then
        fond = COUL_ENTETE_TBL
        encre = COUL_TEXTE
    ElseIf (ligne Mod 2) = 0 Then
        fond = COUL_GRILLE_ZEBRE
        encre = COUL_TEXTE
    Else
        fond = COUL_CARTE
        encre = COUL_TEXTE
    End If

    ' la bande donne sa couleur à toute la ligne, filet compris
    Set lb = ICtl(f, "lblGL_" & CStr(r))
    If Not lb Is Nothing Then
        lb.BackColor = fond
        lb.BorderColor = fond
    End If

    For c = 1 To nbCol
        Set lb = ICtl(f, "lblG_" & CStr(r) & "_" & CStr(c))
        If Not lb Is Nothing Then
            lb.ForeColor = encre
            If vide Then
                lb.Caption = vbNullString
                lb.ControlTipText = vbNullString
            Else
                lb.Caption = Interv_ValeurListe(mLignesVis(ligne), _
                                                CStr(cols(LBound(cols) + c - 1)))
                ' une colonne étroite tronque : l'info-bulle rend le texte entier
                lb.ControlTipText = lb.Caption
            End If
        End If
    Next c
End Sub

'==============================================================================
' Clic sur une case : c'est toute la ligne qui est choisie.
'   r : rang à l'écran de la case cliquée
'==============================================================================
Public Sub Interv_GrilleClic(f As Object, ByVal r As Long)
    Dim ligne As Long, ancienne As Long

    If mChargement Then Exit Sub
    ligne = mPremiereLigne + r
    If ligne < 1 Or ligne > mNbVis Then Exit Sub
    If ligne = mLigneSel Then Exit Sub

    ancienne = mLigneSel
    mLigneSel = ligne

    ' deux lignes repeintes au plus : celle qu'on quitte, celle qu'on prend
    If ancienne > mPremiereLigne And ancienne <= mPremiereLigne + IGR_NB_LIGNES Then
        PeindreLigneGrille f, ancienne - mPremiereLigne
    End If
    PeindreLigneGrille f, r

    Interv_ChargerSelection f
End Sub

'==============================================================================
' Survol d'une case.
'   r : rang à l'écran, ou 0 pour n'en survoler aucun
'==============================================================================
Public Sub Interv_GrilleSurvol(f As Object, ByVal r As Long)
    Dim ligne As Long, ancienne As Long

    ligne = 0
    If r >= 1 Then
        ligne = mPremiereLigne + r
        If ligne > mNbVis Then ligne = 0
    End If
    If ligne = mSurvolGrille Then Exit Sub

    ancienne = mSurvolGrille
    mSurvolGrille = ligne

    If ancienne > mPremiereLigne And ancienne <= mPremiereLigne + IGR_NB_LIGNES Then
        PeindreLigneGrille f, ancienne - mPremiereLigne
    End If
    If ligne > 0 Then PeindreLigneGrille f, r
End Sub

'==============================================================================
' Défilement : la barre donne la première ligne à afficher.
'==============================================================================
Public Sub Interv_Defiler(f As Object)
    Dim c As Object, v As Long

    If mChargement Then Exit Sub
    Set c = ICtl(f, "sbGrille")
    If c Is Nothing Then Exit Sub

    v = c.Value
    If v < 0 Then v = 0
    If v > MaxDefilement() Then v = MaxDefilement()
    If v = mPremiereLigne Then Exit Sub

    mPremiereLigne = v
    mSurvolGrille = 0
    Interv_PeindreGrille f
End Sub

'------------------------------------------------------------------------------
' Plus grand décalage possible : au-delà, la grille montrerait du vide sous la
' dernière ligne.
'------------------------------------------------------------------------------
Private Function MaxDefilement() As Long
    MaxDefilement = mNbVis - IGR_NB_LIGNES
    If MaxDefilement < 0 Then MaxDefilement = 0
End Function

'------------------------------------------------------------------------------
' Accorde la barre de défilement au nombre de lignes retenues par le filtre.
' Elle est désactivée quand tout tient à l'écran, plutôt que laissée active et
' sans effet.
'------------------------------------------------------------------------------
Private Sub MajBarreGrille(f As Object)
    Dim c As Object, garde As Boolean, maxi As Long

    Set c = ICtl(f, "sbGrille")
    If c Is Nothing Then Exit Sub

    maxi = MaxDefilement()
    garde = mChargement
    mChargement = True
    On Error Resume Next
    c.Max = maxi
    c.Value = mPremiereLigne
    c.Enabled = (maxi > 0)
    On Error GoTo 0
    mChargement = garde
End Sub

'==============================================================================
' SÉLECTION
'==============================================================================
Public Sub Interv_ChargerSelection(f As Object)
    If mChargement Then Exit Sub
    If mLigneSel < 1 Or mLigneSel > mNbVis Then Exit Sub

    EcrireChampsInterv f, mLignesVis(mLigneSel)
    mNoCourant = Interv_ValeurAffichee(mLignesVis(mLigneSel), IC_NO)
    MajBoutonsInterv f
    MajEtatInterv f
End Sub

'------------------------------------------------------------------------------
' Mois retenu par le menu déroulant.
'   renvoie : 1 à 12, ou 0 si aucun mois n'est demandé
'
' La position dans la liste suffit : l'élément 0 est « Tous les mois », donc
' l'indice vaut directement le numéro du mois. Aucun besoin de retrouver le mois
' à partir de son nom.
'------------------------------------------------------------------------------
Private Function MoisFiltre(f As Object) As Long
    Dim c As Object
    Set c = ICtl(f, "cboMoisI")
    If c Is Nothing Then Exit Function
    If c.ListIndex > 0 Then MoisFiltre = c.ListIndex
End Function

'------------------------------------------------------------------------------
' Mois de la date d'une intervention, 0 si la date manque ou n'en est pas une.
'------------------------------------------------------------------------------
Private Function MoisDeLaLigne(ByVal ligne As Long) As Long
    Dim v As Variant
    v = Interv_Valeur(ligne, IC_DATE)
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If Not IsDate(v) Then Exit Function
    MoisDeLaLigne = Month(CDate(v))
End Function

'------------------------------------------------------------------------------
' Sélectionne dans le tableau la ligne portant un numéro d'intervention donné.
' Le drapeau de chargement est levé pendant l'opération : écrire ListIndex
' déclencherait sinon l'événement Click, donc un rechargement inutile des champs.
'------------------------------------------------------------------------------
Private Sub SelectionnerNo(f As Object, ByVal noInterv As String)
    Dim i As Long, garde As Boolean

    garde = mChargement
    mChargement = True
    mLigneSel = 0
    For i = 1 To mNbVis
        If StrComp(Interv_ValeurAffichee(mLignesVis(i), IC_NO), noInterv, vbTextCompare) = 0 Then
            mLigneSel = i
            Exit For
        End If
    Next i

    ' amener la ligne choisie dans la partie visible si elle n'y est pas
    If mLigneSel > 0 Then
        If mLigneSel <= mPremiereLigne Then
            mPremiereLigne = mLigneSel - 1
        ElseIf mLigneSel > mPremiereLigne + IGR_NB_LIGNES Then
            mPremiereLigne = mLigneSel - IGR_NB_LIGNES
        End If
        If mPremiereLigne < 0 Then mPremiereLigne = 0
    End If

    MajBarreGrille f
    Interv_PeindreGrille f
    mChargement = garde
End Sub

'==============================================================================
' LECTURE ET ÉCRITURE DES ZONES DE SAISIE
'==============================================================================
'------------------------------------------------------------------------------
' Recopie une intervention du cache vers les zones de saisie.
' Le drapeau de chargement empêche les listes filtrées et le calcul du CA de se
' déclencher pendant le remplissage.
'------------------------------------------------------------------------------
Private Sub EcrireChampsInterv(f As Object, ByVal ligne As Long)
    Dim ch() As ChampInterv, i As Long, c As Object, garde As Boolean

    garde = mChargement
    mChargement = True
    ch = ObtenirChampsInterv()

    For i = LBound(ch) To UBound(ch)
        Set c = ICtl(f, INomControle(ch(i)))
        If Not c Is Nothing Then
            If ch(i).TypeCtrl = ITYPE_CASE Then
                c.Value = EnBooleenI(Interv_Valeur(ligne, ch(i).Colonne))
            Else
                c.Value = Interv_ValeurAffichee(ligne, ch(i).Colonne)
            End If
        End If
    Next i

    mChargement = garde
End Sub

'------------------------------------------------------------------------------
' Interprète une valeur de cellule comme une case à cocher.
' Accepte le booléen d'Excel, un nombre — 0 valant non coché — et les textes
' usuels, au cas où la colonne aurait été saisie à la main.
'------------------------------------------------------------------------------
Private Function EnBooleenI(ByVal v As Variant) As Boolean
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If VarType(v) = vbBoolean Then
        EnBooleenI = CBool(v)
    ElseIf IsNumeric(v) Then
        EnBooleenI = (CDbl(v) <> 0)
    Else
        Select Case LCase$(Trim$(CStr(v)))
            Case "vrai", "true", "oui", "x", "1": EnBooleenI = True
        End Select
    End If
End Function

'------------------------------------------------------------------------------
' Contenu des zones de saisie, typé pour l'écriture dans la feuille.
'   renvoie : un Dictionary nom de colonne -> valeur
'
' Ce qui est écarté, ce ne sont PAS les champs verrouillés : un champ peut être
' hors de portée de la frappe et devoir malgré tout être enregistré — c'est le
' cas du n° de client, rempli par le choix du client. Seules les colonnes de
' IColonnesNonEcrites et les colonnes calculées restent en dehors.
'
' Un champ vide donne Empty, et non une chaîne vide, pour que la cellule Excel
' reste réellement vide.
'------------------------------------------------------------------------------
Private Function LireChampsInterv(f As Object) As Object
    Dim ch() As ChampInterv, i As Long, c As Object
    Dim d As Object, s As String, ok As Boolean

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    ch = ObtenirChampsInterv()

    For i = LBound(ch) To UBound(ch)
        If Not IColonneNonEcrite(ch(i).Colonne) And Not Interv_EstCalculee(ch(i).Colonne) Then
            Set c = ICtl(f, INomControle(ch(i)))
            If Not c Is Nothing Then
                Select Case ch(i).TypeCtrl
                    Case ITYPE_CASE
                        d(ch(i).Colonne) = CBool(c.Value)

                    Case ITYPE_DATE
                        s = Trim$(EnTexte(c.Value))
                        If Len(s) = 0 Then
                            d(ch(i).Colonne) = Empty
                        ElseIf IsDate(s) Then
                            d(ch(i).Colonne) = CDate(s)
                        Else
                            d(ch(i).Colonne) = s
                        End If

                    Case Else
                        s = Trim$(EnTexte(c.Value))
                        If Len(s) = 0 Then
                            d(ch(i).Colonne) = Empty
                        ElseIf ch(i).Saisie = ISAISIE_HEURES Then
                            ok = False
                            d(ch(i).Colonne) = Interv_TexteVersHeures(s, ok)
                            If Not ok Then d(ch(i).Colonne) = Empty
                        ElseIf ch(i).Saisie = ISAISIE_ENTIER Then
                            If QueDesChiffres(s) Then
                                d(ch(i).Colonne) = CDbl(s)
                            Else
                                d(ch(i).Colonne) = s
                            End If
                        ElseIf ch(i).Saisie = ISAISIE_MONTANT Then
                            ' « 13'260.00 » doit atterrir en NOMBRE dans la
                            ' cellule ; écrit tel quel, ce serait du texte, et
                            ' aucune somme ne le compterait
                            ok = False
                            d(ch(i).Colonne) = Interv_TexteVersMontant(s, ok)
                            If Not ok Then d(ch(i).Colonne) = Empty
                        Else
                            d(ch(i).Colonne) = s
                        End If
                End Select
            End If
        End If
    Next i

    Set LireChampsInterv = d
End Function

'------------------------------------------------------------------------------
' Vide les zones de saisie et prépare une nouvelle intervention : le numéro
' affiche « (automatique) » et la date, celle du jour.
'------------------------------------------------------------------------------
Private Sub ViderChampsInterv(f As Object)
    Dim ch() As ChampInterv, i As Long, c As Object, garde As Boolean

    garde = mChargement
    mChargement = True
    ch = ObtenirChampsInterv()

    For i = LBound(ch) To UBound(ch)
        Set c = ICtl(f, INomControle(ch(i)))
        If Not c Is Nothing Then
            If ch(i).TypeCtrl = ITYPE_CASE Then
                c.Value = False
            Else
                c.Value = vbNullString
            End If
        End If
    Next i

    ICtl(f, INomControleColonne(IC_NO)).Value = "(automatique)"
    ICtl(f, INomControleColonne(IC_DATE)).Value = Format$(Date, "dd/mm/yyyy")
    ICtl(f, INomControleColonne(IC_PERS)).Value = "1"

    mChargement = garde
End Sub

'==============================================================================
' SAISIE ASSISTÉE : ENTREPRISE ET NOM
'==============================================================================
'------------------------------------------------------------------------------
' Filtre la liste déroulante au fil de la frappe.
'   nomColonne : Entreprise ou Nom
'
' Vider puis regarnir la liste efface le texte saisi : il est donc relevé avant,
' puis rétabli avec la position du curseur, sans quoi la frappe repartirait du
' début à chaque lettre.
'------------------------------------------------------------------------------
Public Sub Interv_FiltrerListe(f As Object, ByVal nomColonne As String)
    Dim cbo As Object, saisie As String, cible As String, pos As Long
    Dim liste As Variant, i As Long, n As Long, garde As Boolean

    If mChargement Then Exit Sub
    Set cbo = ICtl(f, INomControleColonne(nomColonne))
    If cbo Is Nothing Then Exit Sub

    saisie = EnTexte(cbo.Text)
    pos = Len(saisie)
    On Error Resume Next
    pos = cbo.SelStart
    On Error GoTo 0
    cible = Normaliser(Trim$(saisie))

    garde = mChargement
    mChargement = True

    liste = Interv_ListeClients(nomColonne)
    cbo.Clear
    If IsArray(liste) Then
        For i = LBound(liste) To UBound(liste)
            If Len(cible) = 0 Then
                cbo.AddItem liste(i)
                n = n + 1
            ElseIf InStr(1, Normaliser(CStr(liste(i))), cible, vbBinaryCompare) > 0 Then
                cbo.AddItem liste(i)
                n = n + 1
            End If
        Next i
    End If

    cbo.Text = saisie
    On Error Resume Next
    cbo.SelStart = pos
    On Error GoTo 0
    mChargement = garde

    ' la liste ne se déroule que s'il y a matière à choisir
    If n > 0 And Len(cible) > 0 And n < 40 Then
        On Error Resume Next
        cbo.DropDown
        On Error GoTo 0
    End If
End Sub

'------------------------------------------------------------------------------
' Reporte les informations du client dans la fiche.
'   nomColonne : la colonne sur laquelle le client a été identifié
'
' Les correspondances sont décrites une fois pour toutes dans
' IReportsDepuisClients : colonne source dans TblClients, colonne cible ici.
'------------------------------------------------------------------------------
Public Sub Interv_ClientChoisi(f As Object, ByVal nomColonne As String)
    Dim cbo As Object, ligne As Long, reports As Variant, i As Long
    Dim c As Object, v As Variant, cible As String, garde As Boolean

    If mChargement Then Exit Sub
    Set cbo = ICtl(f, INomControleColonne(nomColonne))
    If cbo Is Nothing Then Exit Sub

    ligne = Interv_TrouverClient(nomColonne, EnTexte(cbo.Value))
    If ligne = 0 Then Exit Sub

    garde = mChargement
    mChargement = True

    reports = IReportsDepuisClients()
    For i = LBound(reports) To UBound(reports)
        cible = CStr(reports(i)(1))
        Set c = ICtl(f, INomControleColonne(cible))
        If Not c Is Nothing Then
            v = Interv_ValeurClient(ligne, CStr(reports(i)(0)))
            If StrComp(cible, IC_TVA, vbTextCompare) = 0 _
               Or StrComp(cible, IC_FORFAIT, vbTextCompare) = 0 Then
                c.Value = EnBooleenI(v)
            ElseIf StrComp(cible, IC_TAUX, vbTextCompare) = 0 Then
                ' le taux s'affiche comme un montant, ici comme dans le tableau
                If IsNumeric(v) Then
                    c.Value = Format$(CDbl(v), "#,##0.00")
                Else
                    c.Value = EnTexte(v)
                End If
            Else
                c.Value = EnTexte(v)
            End If
        End If
    Next i

    mChargement = garde
    Interv_MajCA f
End Sub

'==============================================================================
' CHIFFRE D'AFFAIRES
'------------------------------------------------------------------------------
' Le montant est calculé ici, au fil de la saisie, puis enregistré tel quel
' dans la colonne CA : heures x personnes x taux, ou le taux seul au forfait.
'
' Il se recalcule dès qu'une des quatre valeurs change — heures, personnes,
' taux, case Forfait — et aussi quand le choix d'un client rapporte son taux.
'==============================================================================
Public Sub Interv_MajCA(f As Object)
    Dim c As Object, heures As Double, pers As Double, taux As Double
    Dim forfait As Boolean, ok As Boolean, montant As Double, garde As Boolean

    Set c = ICtl(f, INomControleColonne(IC_CA))
    If c Is Nothing Then Exit Sub

    taux = MontantSaisi(f, IC_TAUX)
    If taux = 0 Then
        ' sans taux il n'y a rien à calculer : mieux vaut une case vide qu'un
        ' zéro, qui se lirait comme une prestation gratuite
        garde = mChargement
        mChargement = True
        c.Value = vbNullString
        mChargement = garde
        Exit Sub
    End If

    ok = False
    heures = Interv_TexteVersHeures(EnTexte(ICtl(f, INomControleColonne(IC_HEURES)).Value), ok)
    If Not ok Then heures = 0

    pers = MontantSaisi(f, IC_PERS)
    If pers < 1 Then pers = 1

    forfait = CBool(ICtl(f, INomControleColonne(IC_FORFAIT)).Value)
    montant = Interv_EstimerCA(taux, forfait, pers, heures)

    garde = mChargement
    mChargement = True
    c.Value = Format$(montant, "#,##0.00")
    mChargement = garde
End Sub

'------------------------------------------------------------------------------
' Valeur numérique d'une zone de saisie, 0 si elle est vide ou illisible.
'------------------------------------------------------------------------------
Private Function MontantSaisi(f As Object, ByVal nomColonne As String) As Double
    Dim c As Object, ok As Boolean, v As Double

    Set c = ICtl(f, INomControleColonne(nomColonne))
    If c Is Nothing Then Exit Function

    ok = False
    v = Interv_TexteVersMontant(EnTexte(c.Value), ok)
    If ok Then MontantSaisi = v
End Function

'==============================================================================
' SÉLECTEUR DE DATE
'==============================================================================
'------------------------------------------------------------------------------
' Ouvre le calendrier sous la zone de date et y reporte le choix.
'------------------------------------------------------------------------------
Public Sub Interv_OuvrirCalendrier(f As Object)
    Dim c As Object, depart As Date, choisie As Date, ok As Boolean
    Dim gauche As Single, haut As Single, garde As Boolean

    Set c = ICtl(f, INomControleColonne(IC_DATE))
    If c Is Nothing Then Exit Sub

    depart = Date
    If IsDate(EnTexte(c.Value)) Then depart = CDate(EnTexte(c.Value))

    ' le calendrier s'ouvre juste sous la zone de date
    gauche = f.Left + c.Left
    haut = f.Top + c.Top + c.Height + 24
    choisie = Calendrier_Choisir(depart, gauche, haut, ok)
    If Not ok Then Exit Sub

    garde = mChargement
    mChargement = True
    c.Value = Format$(choisie, "dd/mm/yyyy")
    mChargement = garde
End Sub

'==============================================================================
' CONTRÔLE DES FRAPPES
'==============================================================================
'------------------------------------------------------------------------------
' N'autorise que les chiffres, et le deux-points pour une durée.
'   KeyAscii : reçu en Object pour éviter toute dépendance à MSForms ici ;
'              mettre sa valeur à 0 annule la frappe
'------------------------------------------------------------------------------
Public Sub Interv_ToucheSaisie(ByVal KeyAscii As Object, ByVal contrainte As Long)
    Dim c As String
    If KeyAscii.Value = 8 Then Exit Sub          ' retour arrière
    If KeyAscii.Value < 32 Then Exit Sub         ' touches de contrôle
    c = Chr$(KeyAscii.Value)
    If c >= "0" And c <= "9" Then Exit Sub
    If contrainte = ISAISIE_HEURES And c = ":" Then Exit Sub
    If contrainte = ISAISIE_MONTANT And (c = "." Or c = ",") Then Exit Sub
    KeyAscii.Value = 0
End Sub

'------------------------------------------------------------------------------
' Remet en forme une durée quittée : 420 devient 420:00.
'------------------------------------------------------------------------------
Public Sub Interv_NormaliserHeures(f As Object)
    Dim c As Object, ok As Boolean, v As Double, garde As Boolean

    If mChargement Then Exit Sub
    Set c = ICtl(f, INomControleColonne(IC_HEURES))
    If c Is Nothing Then Exit Sub
    If Len(Trim$(EnTexte(c.Value))) = 0 Then Exit Sub

    ok = False
    v = Interv_TexteVersHeures(EnTexte(c.Value), ok)
    If Not ok Then Exit Sub

    garde = mChargement
    mChargement = True
    c.Value = Interv_HeuresVersTexte(v)
    mChargement = garde
    Interv_MajCA f
End Sub

'==============================================================================
' BOUTONS
'==============================================================================
Public Sub Interv_Ajouter(f As Object)
    Dim msg As String, vals As Object, no As String

    On Error GoTo Erreur
    If Not ValiderInterv(f, msg) Then
        MsgBox msg, vbExclamation, "Ajout impossible"
        Exit Sub
    End If

    Set vals = LireChampsInterv(f)
    no = IntervBD_Ajouter(vals)
    mNoCourant = no
    Interv_RafraichirListe f
    SelectionnerNo f, no
    Interv_MajStatistiques f
    MajBoutonsInterv f
    MajEtatInterv f, "Intervention " & no & " ajoutée au tableau."
    Exit Sub

Erreur:
    MsgBox "Ajout impossible :" & vbCrLf & vbCrLf & Err.Description, vbCritical, "Interventions"
End Sub

'------------------------------------------------------------------------------
' Bouton Modifier : enregistre les zones de saisie sur l'intervention
' sélectionnée. Le numéro et le chiffre d'affaires ne sont pas touchés.
'------------------------------------------------------------------------------
Public Sub Interv_Modifier(f As Object)
    Dim msg As String, vals As Object

    On Error GoTo Erreur
    If Len(mNoCourant) = 0 Then
        MsgBox "Sélectionnez d'abord une intervention dans le tableau.", _
               vbInformation, "Modification"
        Exit Sub
    End If
    If Not ValiderInterv(f, msg) Then
        MsgBox msg, vbExclamation, "Modification impossible"
        Exit Sub
    End If

    Set vals = LireChampsInterv(f)
    If IntervBD_Modifier(mNoCourant, vals) Then
        Interv_RafraichirListe f
        Interv_MajStatistiques f
        MajEtatInterv f, "Intervention " & mNoCourant & " mise à jour."
    Else
        MsgBox "L'intervention " & mNoCourant & " est introuvable dans le tableau.", _
               vbExclamation, "Modification"
    End If
    Exit Sub

Erreur:
    MsgBox "Modification impossible :" & vbCrLf & vbCrLf & Err.Description, _
           vbCritical, "Interventions"
End Sub

'------------------------------------------------------------------------------
' Bouton Supprimer : demande confirmation, en signalant la facture sur laquelle
' l'intervention figure éventuellement, puis supprime la ligne du tableau.
'------------------------------------------------------------------------------
Public Sub Interv_Supprimer(f As Object)
    Dim ligne As Long, no As String, question As String, facture As String

    On Error GoTo Erreur
    If Len(mNoCourant) = 0 Then
        MsgBox "Sélectionnez d'abord une intervention dans le tableau.", _
               vbInformation, "Suppression"
        Exit Sub
    End If

    no = mNoCourant
    ligne = Interv_TrouverLigne(no)
    If ligne = 0 Then
        MsgBox "L'intervention " & no & " est introuvable dans le tableau.", _
               vbExclamation, "Suppression"
        Exit Sub
    End If

    question = "Supprimer définitivement l'intervention " & no & " ?" & vbCrLf & vbCrLf & _
               Interv_ValeurAffichee(ligne, IC_DATE) & "   " & _
               Trim$(Interv_ValeurAffichee(ligne, IC_ENTREPRISE) & " " & _
                     Interv_ValeurAffichee(ligne, IC_NOM) & " " & _
                     Interv_ValeurAffichee(ligne, IC_PRENOM))

    facture = Trim$(Interv_ValeurAffichee(ligne, IC_FACTURE))
    If Len(facture) > 0 Then
        question = question & vbCrLf & vbCrLf & _
                   "Attention : cette intervention figure sur la facture " & facture & "."
    End If

    If MsgBox(question, vbQuestion + vbYesNo + vbDefaultButton2, "Suppression") <> vbYes Then Exit Sub

    If IntervBD_Supprimer(no) Then
        mNoCourant = vbNullString
        ViderChampsInterv f
        Interv_RafraichirListe f
        Interv_MajStatistiques f
        MajBoutonsInterv f
        MajEtatInterv f, "Intervention " & no & " supprimée."
    End If
    Exit Sub

Erreur:
    MsgBox "Suppression impossible :" & vbCrLf & vbCrLf & Err.Description, _
           vbCritical, "Interventions"
End Sub

'------------------------------------------------------------------------------
' Bouton Effacer : vide les zones de saisie et repasse en mode nouvelle
' intervention. Le tableau et la feuille ne sont pas touchés, et le filtre reste
' en place.
'------------------------------------------------------------------------------
Public Sub Interv_Effacer(f As Object)
    Dim garde As Boolean

    mNoCourant = vbNullString
    ViderChampsInterv f

    garde = mChargement
    mChargement = True
    mLigneSel = 0
    Interv_PeindreGrille f
    mChargement = garde

    MajBoutonsInterv f
    MajEtatInterv f
    On Error Resume Next
    ICtl(f, INomControleColonne(IC_ENTREPRISE)).SetFocus
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Bouton Quitter et touche Échap : ferme le formulaire.
'------------------------------------------------------------------------------
Public Sub Interv_Quitter(f As Object)
    Unload f
End Sub

'------------------------------------------------------------------------------
' Remet l'état à zéro à la fermeture, pour que la prochaine ouverture reparte
' proprement — y compris le drapeau d'ajustement de la fenêtre.
'------------------------------------------------------------------------------
Public Sub Interv_Fermeture(f As Object)
    mNoCourant = vbNullString
    mNbVis = 0
    mAjuste = False
End Sub

'------------------------------------------------------------------------------
' Boutons Facturer et Info : les formulaires qu'ils ouvriront restent à définir.
' Le point d'entrée est en place, il n'y aura qu'à remplacer le message par
' l'appel voulu.
'------------------------------------------------------------------------------
Public Sub Interv_Facturer(f As Object)
    MsgBox "Le formulaire de facturation reste à définir." & vbCrLf & vbCrLf & _
           "Point de branchement : modInterv_Formulaire.Interv_Facturer." & _
           IIf(Len(mNoCourant) > 0, vbCrLf & vbCrLf & "Intervention sélectionnée : " & mNoCourant, ""), _
           vbInformation, "Facturer"
End Sub

'------------------------------------------------------------------------------
' Bouton Info : le formulaire d'informations reste à définir.
'------------------------------------------------------------------------------
Public Sub Interv_Info(f As Object)
    MsgBox "Le formulaire d'informations reste à définir." & vbCrLf & vbCrLf & _
           "Point de branchement : modInterv_Formulaire.Interv_Info.", _
           vbInformation, "Info"
End Sub

'==============================================================================
' CONTRÔLE DE LA SAISIE
'==============================================================================
'------------------------------------------------------------------------------
' Contrôles avant écriture. Volontairement souples : ils empêchent les fautes de
' frappe, pas l'enregistrement d'une fiche encore incomplète.
'------------------------------------------------------------------------------
Private Function ValiderInterv(f As Object, ByRef msg As String) As Boolean
    Dim s As String, ok As Boolean

    msg = vbNullString

    s = Trim$(EnTexte(ICtl(f, INomControleColonne(IC_DATE)).Value))
    If Len(s) = 0 Then
        msg = "Indiquez la date de l'intervention."
        Exit Function
    End If
    If Not IsDate(s) Then
        msg = "La date ne semble pas valide : " & s
        Exit Function
    End If

    If Len(Trim$(EnTexte(ICtl(f, INomControleColonne(IC_NOM)).Value))) = 0 _
       And Len(Trim$(EnTexte(ICtl(f, INomControleColonne(IC_ENTREPRISE)).Value))) = 0 Then
        msg = "Renseignez au moins le nom du client ou son entreprise."
        Exit Function
    End If

    s = Trim$(EnTexte(ICtl(f, INomControleColonne(IC_HEURES)).Value))
    If Len(s) > 0 Then
        ok = False
        Interv_TexteVersHeures s, ok
        If Not ok Then
            msg = "La durée doit s'écrire heures:minutes, par exemple 420:00."
            Exit Function
        End If
    End If

    s = Trim$(EnTexte(ICtl(f, INomControleColonne(IC_PERS)).Value))
    If Len(s) > 0 Then
        If Not QueDesChiffres(s) Or Len(s) > 2 Or CDbl(s) < 1 Then
            msg = "Le nombre de personnes doit être un entier de 1 à 99."
            Exit Function
        End If
    End If

    ValiderInterv = True
End Function

'==============================================================================
' HABILLAGE
'==============================================================================
Public Sub Interv_Survol(f As Object, ByVal nomControle As String)
    Dim boutons As Variant, i As Long, b As Object, lb As Object

    boutons = Array("btnIAjouter", "btnIModifier", "btnISupprimer", "btnIEffacer", _
                    "btnFacturer", "btnInfo", "btnIQuitter")
    For i = LBound(boutons) To UBound(boutons)
        Set b = ICtl(f, CStr(boutons(i)))
        If Not b Is Nothing Then
            If b.Enabled Then
                CouleurFondSi b, ICouleurBouton(CStr(boutons(i)), CStr(boutons(i)) = nomControle)
            End If
        End If
    Next i

    Set lb = ICtl(f, "lblResetFiltreI")
    If Not lb Is Nothing Then
        If lb.ForeColor <> IIf(nomControle = "lblResetFiltreI", COUL_LIEN_H, COUL_LIEN) Then
            lb.ForeColor = IIf(nomControle = "lblResetFiltreI", COUL_LIEN_H, COUL_LIEN)
        End If
    End If
End Sub

'------------------------------------------------------------------------------
' Change la couleur de fond seulement si elle diffère : réécrire la même valeur
' à chaque MouseMove ferait scintiller le contrôle.
'------------------------------------------------------------------------------
Private Sub CouleurFondSi(c As Object, ByVal couleur As Long)
    If c.BackColor <> couleur Then c.BackColor = couleur
End Sub

'------------------------------------------------------------------------------
' Souligne le champ actif : filet bleu à l'entrée, filet gris à la sortie.
'------------------------------------------------------------------------------
Public Sub Interv_FocusChamp(f As Object, ByVal nomControle As String, ByVal actif As Boolean)
    Dim c As Object
    Set c = ICtl(f, nomControle)
    If c Is Nothing Then Exit Sub
    On Error Resume Next
    c.BorderColor = IIf(actif, COUL_CHAMP_FOCUS, COUL_CHAMP_BORD)
    On Error GoTo 0
End Sub

'==============================================================================
' BANDEAU ET BOUTONS : MISE À JOUR DE L'ÉTAT
'==============================================================================
Private Sub MajEtatInterv(f As Object, Optional ByVal message As String = vbNullString)
    Dim lb As Object, txt As String

    Set lb = ICtl(f, "lblEtatI")
    If lb Is Nothing Then Exit Sub

    If Len(message) > 0 Then
        txt = message
    ElseIf Len(mNoCourant) = 0 Then
        txt = "Nouvelle intervention — remplissez les champs puis cliquez sur Ajouter."
    Else
        txt = "Intervention " & mNoCourant & " sélectionnée — modifiez puis cliquez sur Modifier."
    End If
    lb.Caption = txt
End Sub

'------------------------------------------------------------------------------
' Met à jour le compteur de la barre de filtrage.
'------------------------------------------------------------------------------
Private Sub MajCompteurInterv(f As Object)
    Dim lb As Object
    Set lb = ICtl(f, "lblCompteurI")
    If lb Is Nothing Then Exit Sub
    lb.Caption = mNbVis & " intervention(s) affichée(s) sur " & Interv_NbLignes()
End Sub

'------------------------------------------------------------------------------
' Active ou grise Modifier, Supprimer et Facturer selon qu'une intervention est
' sélectionnée.
'------------------------------------------------------------------------------
Private Sub MajBoutonsInterv(f As Object)
    Dim actif As Boolean
    actif = (Len(mNoCourant) > 0)
    ActiverBoutonI f, "btnIModifier", actif
    ActiverBoutonI f, "btnISupprimer", actif
    ActiverBoutonI f, "btnFacturer", actif
End Sub

'------------------------------------------------------------------------------
' Active ou grise un bouton, couleur comprise.
'------------------------------------------------------------------------------
Private Sub ActiverBoutonI(f As Object, ByVal nom As String, ByVal actif As Boolean)
    Dim b As Object
    Set b = ICtl(f, nom)
    If b Is Nothing Then Exit Sub
    b.Enabled = actif
    b.BackColor = IIf(actif, ICouleurBouton(nom, False), COUL_BOUTON_OFF)
End Sub

'==============================================================================
' Accès sûr aux contrôles du formulaire
'------------------------------------------------------------------------------
' Le formulaire est reçu en Object, jamais en UF_Interventions : c'est ce qui
' permet à ce module de compiler avant que le formulaire ait été généré, et de
' survivre à sa régénération.
'==============================================================================
Private Function ICtl(f As Object, ByVal nom As String) As Object
    On Error Resume Next
    Set ICtl = f.Controls(nom)
    On Error GoTo 0
End Function
