Attribute VB_Name = "modClients_Formulaire"
Option Explicit
'==============================================================================
' modClients_Formulaire
'------------------------------------------------------------------------------
' Comportement du formulaire UF_Clients.
'
' Le module de code du UserForm ne contient que des procédures événementielles
' d'une ligne qui appellent les procédures publiques ci-dessous. Toute la
' logique reste donc dans un module standard : elle est lisible, modifiable et
' le projet compile même si le UserForm n'a pas encore été généré.
'
' Le formulaire est toujours reçu en tant qu'Object (liaison tardive), ce qui
' évite toute dépendance de compilation au UserForm.
'==============================================================================

Private mClefCourante As String     ' Clef_BD de la fiche affichée ("" = nouvelle fiche)
Private mLignesVis() As Long        ' lignes du cache visibles dans le tableau
Private mNbVis As Long
Private mTriColonne As Long         ' index (1 à 10) de la colonne de tri, 0 = aucun
Private mTriDecroissant As Boolean
Private mChargement As Boolean      ' True pendant un remplissage programme
Private mDeplacement As Boolean
Private mDepX As Single
Private mDepY As Single

'==============================================================================
' INITIALISATION
'==============================================================================
'------------------------------------------------------------------------------
' Prépare le formulaire : appelée une fois, depuis UserForm_Initialize.
' Charge les données, remplit les menus déroulants, écrit les en-têtes du
' tableau, affiche la liste et place le formulaire en mode « nouvelle fiche ».
'------------------------------------------------------------------------------
Public Sub Clients_Initialiser(f As Object)
    Dim lst As Object, i As Long, lib As Variant

    On Error GoTo Erreur
    mChargement = True
    mClefCourante = vbNullString
    mTriColonne = 0
    mTriDecroissant = False

    Donnees_Charger
    Adresses_Recharger

    Ctl(f, "lblTitre").Caption = "Gestion des clients"

    ' --- menus déroulants -----------------------------------------------------
    RemplirTitres f
    RemplirListeCombo Ctl(f, NomControleColonne(COL_ADRESSE)), Adresses_ListeRues()
    RemplirListeCombo Ctl(f, NomControleColonne(COL_NPA)), Adresses_ListeNpa()
    RemplirListeCombo Ctl(f, NomControleColonne(COL_TEXTE_FACTURE)), Adresses_TextesFacture()

    With Ctl(f, "cboChampFiltre")
        .Clear
        lib = ChampsFiltrables()
        For i = LBound(lib) To UBound(lib)
            .AddItem lib(i)
        Next i
        .ListIndex = 1                      ' "Nom" par défaut
    End With

    ' --- en-têtes du tableau ---------------------------------------------------
    MajEntetesTableau f

    Set lst = Ctl(f, "lstClients")
    lst.Clear

    mChargement = False
    ViderChamps f
    Clients_RafraichirListe f
    MajBoutons f
    MajEntete f
    Exit Sub

Erreur:
    mChargement = False
    MsgBox "Initialisation impossible :" & vbCrLf & vbCrLf & Err.Description, _
           vbExclamation, "Gestion des clients"
End Sub

'------------------------------------------------------------------------------
' Remplit le menu déroulant Titre : d'abord les civilités proposées par le
' schéma, puis celles déjà présentes dans le tableau qui n'y figurent pas.
' Sans ce complément, une fiche « Famille » ou « Docteur » s'afficherait vide,
' puisque le menu est en liste fermée.
'------------------------------------------------------------------------------
Private Sub RemplirTitres(f As Object)
    Dim cbo As Object, prop As Variant, sup As Variant, i As Long, d As Object

    Set cbo = Ctl(f, NomControleColonne("Titre"))
    If cbo Is Nothing Then Exit Sub

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1

    cbo.Clear
    prop = TitresProposes()
    For i = LBound(prop) To UBound(prop)
        cbo.AddItem prop(i)
        d.Add CStr(prop(i)), 1
    Next i

    ' les civilités déjà présentes dans le tableau restent sélectionnables
    sup = TrierChaines(Donnees_ValeursDistinctes("Titre"))
    If IsArray(sup) Then
        For i = LBound(sup) To UBound(sup)
            If Not d.Exists(CStr(sup(i))) Then
                cbo.AddItem sup(i)
                d.Add CStr(sup(i)), 1
            End If
        Next i
    End If
End Sub

'------------------------------------------------------------------------------
' Remplit un menu déroulant à partir d'un tableau de valeurs.
' Ne fait rien si le contrôle ou le tableau est absent : le formulaire reste
' utilisable même si une des feuilles sources manque.
'------------------------------------------------------------------------------
Private Sub RemplirListeCombo(cbo As Object, ByVal valeurs As Variant)
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

'------------------------------------------------------------------------------
' Écrit les en-têtes de colonnes du tableau et ajoute le triangle de tri sur
' la colonne active.
'------------------------------------------------------------------------------
Private Sub MajEntetesTableau(f As Object)
    Dim lib As Variant, i As Long, lb As Object, txt As String

    lib = LibellesListe()
    For i = LBound(lib) To UBound(lib)
        Set lb = Ctl(f, "lblEnt_" & CStr(i + 1))
        If Not lb Is Nothing Then
            txt = CStr(lib(i))
            If mTriColonne = i + 1 Then
                txt = txt & IIf(mTriDecroissant, GlypheTriDecroissant(), GlypheTriCroissant())
            End If
            lb.Caption = txt
        End If
    Next i
End Sub

'==============================================================================
' TABLEAU DES ENREGISTREMENTS : filtrage, tri, affichage
'==============================================================================
'------------------------------------------------------------------------------
' Reconstruit le tableau des enregistrements : filtre, tri, puis affichage.
'
' Deux tableaux sont tenus en parallèle : ce que voit l'utilisateur, et
' mLignesVis, qui garde pour chaque ligne affichée son numéro dans le cache.
' C'est ce second tableau qui sert à retrouver la fiche sélectionnée, si bien que
' les colonnes affichées peuvent être changées librement sans rien casser.
'------------------------------------------------------------------------------
Public Sub Clients_RafraichirListe(f As Object)
    Dim nb As Long, i As Long, champ As String, cible As String
    Dim cols As Variant, arr() As Variant, lst As Object
    Dim ic As Long, garde As Boolean

    On Error GoTo Erreur
    nb = Donnees_NbLignes()
    ' Dimensionné au pire cas — toutes les fiches visibles — puis rempli sur
    ' mNbVis éléments. La borne minimale de 1 évite un ReDim(1 To 0), refusé
    ' par VBA quand le tableau est vide.
    ReDim mLignesVis(1 To IIf(nb > 0, nb, 1))
    mNbVis = 0

    champ = EnTexte(Ctl(f, "cboChampFiltre").Value)
    If Len(champ) = 0 Then champ = "Nom"
    cible = Normaliser(Trim$(EnTexte(Ctl(f, "txtFiltre").Text)))

    For i = 1 To nb
        If Len(cible) = 0 Then
            garde = True
        Else
            garde = (InStr(1, Normaliser(Donnees_ValeurAffichee(i, champ)), cible, vbBinaryCompare) > 0)
        End If
        If garde Then
            mNbVis = mNbVis + 1
            mLignesVis(mNbVis) = i
        End If
    Next i

    If mTriColonne > 0 And mNbVis > 1 Then TrierLignesVisibles

    Set lst = Ctl(f, "lstClients")
    cols = ColonnesListe()

    If mNbVis = 0 Then
        lst.Clear
    Else
        ' La ListBox est alimentée en une seule affectation : ajouter les
        ' lignes une à une avec AddItem serait nettement plus lent.
        ReDim arr(1 To mNbVis, 1 To UBound(cols) - LBound(cols) + 1)
        For i = 1 To mNbVis
            For ic = LBound(cols) To UBound(cols)
                arr(i, ic + 1) = Donnees_ValeurAffichee(mLignesVis(i), CStr(cols(ic)))
            Next ic
        Next i
        lst.List = arr
    End If

    MajCompteur f
    If Len(mClefCourante) > 0 Then SelectionnerClef f, mClefCourante
    Exit Sub

Erreur:
    MsgBox "Affichage de la liste impossible :" & vbCrLf & vbCrLf & Err.Description, _
           vbExclamation, "Gestion des clients"
End Sub

'------------------------------------------------------------------------------
' Relance le filtrage. Appelée à chaque frappe dans la zone de filtre et à
' chaque changement de colonne.
' Ne fait rien pendant un remplissage programmé, pour ne pas se déclencher quand
' le code écrit lui-même dans les contrôles.
'------------------------------------------------------------------------------
Public Sub Clients_AppliquerFiltre(f As Object)
    If mChargement Then Exit Sub
    Clients_RafraichirListe f
End Sub

'------------------------------------------------------------------------------
' Vide la zone de filtre et revient à la colonne Nom.
' Déclenchée par le lien Réinitialiser de la barre de filtrage.
'------------------------------------------------------------------------------
Public Sub Clients_ReinitialiserFiltre(f As Object)
    Dim garde As Boolean
    garde = mChargement
    mChargement = True
    Ctl(f, "txtFiltre").Text = vbNullString
    Ctl(f, "cboChampFiltre").ListIndex = 1
    mChargement = garde
    Clients_RafraichirListe f
End Sub

'------------------------------------------------------------------------------
' Trie le tableau sur une colonne.
'   colonne : 1 à 10, position dans ColonnesListe
' Un clic sur la colonne déjà triée inverse simplement le sens.
'------------------------------------------------------------------------------
Public Sub Clients_TrierColonne(f As Object, ByVal colonne As Long)
    If mTriColonne = colonne Then
        mTriDecroissant = Not mTriDecroissant
    Else
        mTriColonne = colonne
        mTriDecroissant = False
    End If
    MajEntetesTableau f
    Clients_RafraichirListe f
End Sub

'------------------------------------------------------------------------------
' Trie mLignesVis sur la colonne active. Seul l'ordre des numéros de ligne
' change : le cache lui-même n'est jamais réordonné, et le tableau Excel non plus.
'------------------------------------------------------------------------------
Private Sub TrierLignesVisibles()
    Dim cols As Variant, nomCol As String
    cols = ColonnesListe()
    If mTriColonne < 1 Or mTriColonne > (UBound(cols) - LBound(cols) + 1) Then Exit Sub
    nomCol = CStr(cols(LBound(cols) + mTriColonne - 1))
    TriRapideLignes 1, mNbVis, nomCol
End Sub

'------------------------------------------------------------------------------
' Tri rapide sur le tableau des lignes visibles.
'------------------------------------------------------------------------------
Private Sub TriRapideLignes(ByVal g As Long, ByVal d As Long, ByVal nomCol As String)
    Dim i As Long, j As Long, pivot As Long, tmp As Long
    i = g: j = d
    pivot = mLignesVis((g + d) \ 2)
    Do While i <= j
        Do While Comparer(mLignesVis(i), pivot, nomCol) < 0
            i = i + 1
        Loop
        Do While Comparer(mLignesVis(j), pivot, nomCol) > 0
            j = j - 1
        Loop
        If i <= j Then
            tmp = mLignesVis(i): mLignesVis(i) = mLignesVis(j): mLignesVis(j) = tmp
            i = i + 1: j = j - 1
        End If
    Loop
    If g < j Then TriRapideLignes g, j, nomCol
    If i < d Then TriRapideLignes i, d, nomCol
End Sub

'------------------------------------------------------------------------------
' Compare deux fiches sur une colonne.
'   renvoie : -1, 0 ou 1, sens de tri déjà appliqué
'
' Les dates sont comparées chronologiquement et les nombres numériquement : un
' tri purement alphabétique placerait le 02/01 avant le 14/08 mais aussi 100
' avant 20. Les fiches sans valeur sont renvoyées en fin de liste, dans les deux
' sens de tri.
'------------------------------------------------------------------------------
Private Function Comparer(ByVal ligneA As Long, ByVal ligneB As Long, ByVal nomCol As String) As Long
    Dim va As Variant, vb As Variant, r As Long
    Dim sa As String, sb As String

    va = Donnees_Valeur(ligneA, nomCol)
    vb = Donnees_Valeur(ligneB, nomCol)

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
        sa = Normaliser(Donnees_ValeurAffichee(ligneA, nomCol))
        sb = Normaliser(Donnees_ValeurAffichee(ligneB, nomCol))
        ' les fiches sans valeur sont renvoyées en fin de liste
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
    Comparer = r
End Function

'==============================================================================
' SELECTION D'UNE FICHE
'==============================================================================
'------------------------------------------------------------------------------
' Charge la fiche sélectionnée dans les zones de saisie et passe le formulaire
' en mode modification.
'------------------------------------------------------------------------------
Public Sub Clients_ChargerSelection(f As Object)
    Dim idx As Long
    If mChargement Then Exit Sub
    idx = Ctl(f, "lstClients").ListIndex
    If idx < 0 Or idx + 1 > mNbVis Then Exit Sub

    EcrireChamps f, mLignesVis(idx + 1)
    mClefCourante = Donnees_ValeurAffichee(mLignesVis(idx + 1), COL_CLEF)
    MajBoutons f
    MajEntete f
End Sub

'------------------------------------------------------------------------------
' Sélectionne dans le tableau la ligne portant une Clef_BD donnée.
' Le drapeau de chargement est levé pendant l'opération : écrire ListIndex
' déclencherait sinon l'événement Click, donc un rechargement inutile des champs.
'------------------------------------------------------------------------------
Private Sub SelectionnerClef(f As Object, ByVal clef As String)
    Dim i As Long, lst As Object, garde As Boolean

    Set lst = Ctl(f, "lstClients")
    garde = mChargement
    mChargement = True
    On Error Resume Next
    lst.ListIndex = -1
    For i = 1 To mNbVis
        If StrComp(Donnees_ValeurAffichee(mLignesVis(i), COL_CLEF), clef, vbTextCompare) = 0 Then
            lst.ListIndex = i - 1
            Exit For
        End If
    Next i
    On Error GoTo 0
    mChargement = garde
End Sub

'==============================================================================
' LECTURE / ECRITURE DES ZONES DE SAISIE
'==============================================================================
'------------------------------------------------------------------------------
' Recopie une fiche du cache vers les zones de saisie.
'   ligne : numéro de fiche dans le cache
' Le drapeau de chargement empêche les événements Change des menus déroulants
' Adresse et NPA de se déclencher et d'écraser la ville ou le canton lus.
'------------------------------------------------------------------------------
Private Sub EcrireChamps(f As Object, ByVal ligne As Long)
    Dim ch() As ChampClient, i As Long, c As Object, v As Variant
    Dim garde As Boolean

    garde = mChargement
    mChargement = True
    ch = ObtenirChamps()

    For i = LBound(ch) To UBound(ch)
        Set c = Ctl(f, NomControle(ch(i)))
        If Not c Is Nothing Then
            v = Donnees_Valeur(ligne, ch(i).Colonne)
            Select Case ch(i).TypeCtrl
                Case TYPE_CASE
                    c.Value = EnBooleen(v)
                Case Else
                    c.Value = ValeurPourSaisie(ligne, ch(i))
            End Select
        End If
    Next i

    mChargement = garde
End Sub

'------------------------------------------------------------------------------
' Interprète une valeur de cellule comme une case à cocher.
' Accepte le booléen d'Excel, un nombre (0 = non coché) et les textes usuels,
' au cas où la colonne aurait été saisie à la main.
'------------------------------------------------------------------------------
Private Function EnBooleen(ByVal v As Variant) As Boolean
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If VarType(v) = vbBoolean Then
        EnBooleen = CBool(v)
    ElseIf IsNumeric(v) Then
        EnBooleen = (CDbl(v) <> 0)
    Else
        Select Case LCase$(Trim$(CStr(v)))
            Case "vrai", "true", "oui", "x", "1": EnBooleen = True
        End Select
    End If
End Function

'------------------------------------------------------------------------------
' Valeur d'un champ mise en forme pour sa zone de saisie : date en jj/mm/aaaa,
' montant à deux décimales, tout le reste tel quel.
'------------------------------------------------------------------------------
Private Function ValeurPourSaisie(ByVal ligne As Long, ByRef ch As ChampClient) As String
    Dim v As Variant

    v = Donnees_Valeur(ligne, ch.Colonne)
    If IsEmpty(v) Or IsNull(v) Then Exit Function

    If StrComp(ch.Colonne, COL_DATE, vbTextCompare) = 0 Then
        ValeurPourSaisie = Donnees_ValeurAffichee(ligne, ch.Colonne)
    ElseIf ch.Numerique = NUM_DECIMAL And IsNumeric(v) Then
        ValeurPourSaisie = Format$(CDbl(v), "0.00")
    Else
        ValeurPourSaisie = CStr(v)
    End If
End Function

'------------------------------------------------------------------------------
' Relève le contenu des zones de saisie.
'   renvoie : un Dictionary nom de colonne -> valeur typée
'
' Le typage se fait ici, une bonne fois : case à cocher en booléen, champ
' numérique en nombre, champ vide en Empty — et non en chaîne vide, pour que la
' cellule Excel reste réellement vide. Les champs verrouillés sont ignorés.
'
' Le NPA fait exception et reste du texte : c'est le type de la colonne dans le
' tableau, et le convertir en nombre ferait perdre un éventuel zéro initial.
'------------------------------------------------------------------------------
Private Function LireChamps(f As Object) As Object
    Dim ch() As ChampClient, i As Long, c As Object
    Dim d As Object, s As String, ok As Boolean

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    ch = ObtenirChamps()

    For i = LBound(ch) To UBound(ch)
        If Not ch(i).Verrouille Then
            Set c = Ctl(f, NomControle(ch(i)))
            If Not c Is Nothing Then
                Select Case ch(i).TypeCtrl
                    Case TYPE_CASE
                        d(ch(i).Colonne) = CBool(c.Value)
                    Case Else
                        s = Trim$(EnTexte(c.Value))
                        If Len(s) = 0 Then
                            ' Empty, et non "" : une chaîne vide écrite dans
                            ' une cellule la rend non vide pour ESTVIDE et
                            ' pour les filtres d'Excel.
                            d(ch(i).Colonne) = Empty
                        ElseIf ch(i).Numerique <> NUM_NON _
                               And StrComp(ch(i).Colonne, COL_NPA, vbTextCompare) <> 0 Then
                            ok = False
                            d(ch(i).Colonne) = EnNombre(s, ok)
                            If Not ok Then d(ch(i).Colonne) = s
                        Else
                            d(ch(i).Colonne) = s
                        End If
                End Select
            End If
        End If
    Next i

    Set LireChamps = d
End Function

'------------------------------------------------------------------------------
' Vide toutes les zones de saisie et prépare une nouvelle fiche : la clef
' affiche « (automatique) » et la date, celle du jour.
'------------------------------------------------------------------------------
Private Sub ViderChamps(f As Object)
    Dim ch() As ChampClient, i As Long, c As Object, garde As Boolean

    garde = mChargement
    mChargement = True
    ch = ObtenirChamps()

    For i = LBound(ch) To UBound(ch)
        Set c = Ctl(f, NomControle(ch(i)))
        If Not c Is Nothing Then
            If ch(i).TypeCtrl = TYPE_CASE Then
                c.Value = False
            Else
                c.Value = vbNullString
            End If
        End If
    Next i

    ' valeurs proposées par défaut pour une nouvelle fiche
    Ctl(f, NomControleColonne(COL_CLEF)).Value = "(automatique)"
    Ctl(f, NomControleColonne(COL_DATE)).Value = Format$(Date, "dd/mm/yyyy")

    mChargement = garde
End Sub

'==============================================================================
' BOUTON "AJOUTER"
'==============================================================================
'------------------------------------------------------------------------------
' Bouton Ajouter : contrôle la saisie, prévient en cas d'ID Crésus en doublon,
' écrit la nouvelle fiche puis la sélectionne dans le tableau.
'------------------------------------------------------------------------------
Public Sub Clients_Ajouter(f As Object)
    Dim msg As String, vals As Object, clef As String, doublon As String

    On Error GoTo Erreur
    If Not Valider(f, msg) Then
        MsgBox msg, vbExclamation, "Ajout impossible"
        Exit Sub
    End If

    Set vals = LireChamps(f)

    doublon = Donnees_ClefAvecMemeIdCresus(EnTexte(vals(COL_ID_CRESUS)), vbNullString)
    If Len(doublon) > 0 Then
        If MsgBox("L'ID Crésus saisi est déjà utilisé par la fiche " & doublon & "." & vbCrLf & _
                  "Ajouter tout de même cette fiche ?", vbQuestion + vbYesNo + vbDefaultButton2, _
                  "Doublon d'ID Crésus") <> vbYes Then Exit Sub
    End If

    clef = Donnees_Ajouter(vals)
    mClefCourante = clef
    Clients_RafraichirListe f
    SelectionnerClef f, clef
    MajBoutons f
    MajEntete f, "Fiche " & clef & " ajoutée au tableau."
    Exit Sub

Erreur:
    MsgBox "Ajout impossible :" & vbCrLf & vbCrLf & Err.Description, vbCritical, "Gestion des clients"
End Sub

'==============================================================================
' BOUTON "MODIFIER"
'==============================================================================
'------------------------------------------------------------------------------
' Bouton Modifier : enregistre les zones de saisie sur la fiche sélectionnée.
' La clef et la date de création ne sont pas touchées.
'------------------------------------------------------------------------------
Public Sub Clients_Modifier(f As Object)
    Dim msg As String, vals As Object, doublon As String

    On Error GoTo Erreur
    If Len(mClefCourante) = 0 Then
        MsgBox "Sélectionnez d'abord une fiche dans le tableau.", vbInformation, "Modification"
        Exit Sub
    End If
    If Not Valider(f, msg) Then
        MsgBox msg, vbExclamation, "Modification impossible"
        Exit Sub
    End If

    Set vals = LireChamps(f)

    doublon = Donnees_ClefAvecMemeIdCresus(EnTexte(vals(COL_ID_CRESUS)), mClefCourante)
    If Len(doublon) > 0 Then
        If MsgBox("L'ID Crésus saisi est déjà utilisé par la fiche " & doublon & "." & vbCrLf & _
                  "Enregistrer tout de même ?", vbQuestion + vbYesNo + vbDefaultButton2, _
                  "Doublon d'ID Crésus") <> vbYes Then Exit Sub
    End If

    If Donnees_Modifier(mClefCourante, vals) Then
        Clients_RafraichirListe f
        MajEntete f, "Fiche " & mClefCourante & " mise à jour."
    Else
        MsgBox "La fiche " & mClefCourante & " est introuvable dans le tableau.", _
               vbExclamation, "Modification"
    End If
    Exit Sub

Erreur:
    MsgBox "Modification impossible :" & vbCrLf & vbCrLf & Err.Description, vbCritical, "Gestion des clients"
End Sub

'==============================================================================
' BOUTON "SUPPRIMER"
'==============================================================================
'------------------------------------------------------------------------------
' Bouton Supprimer : demande confirmation, en signalant le nombre
' d'interventions rattachées au client, puis supprime la ligne du tableau.
'------------------------------------------------------------------------------
Public Sub Clients_Supprimer(f As Object)
    Dim ligne As Long, nom As String, nbInterv As Long, question As String
    Dim clef As String

    On Error GoTo Erreur
    If Len(mClefCourante) = 0 Then
        MsgBox "Sélectionnez d'abord une fiche dans le tableau.", vbInformation, "Suppression"
        Exit Sub
    End If

    clef = mClefCourante
    ligne = Donnees_TrouverLigne(clef)
    If ligne = 0 Then
        MsgBox "La fiche " & clef & " est introuvable dans le tableau.", vbExclamation, "Suppression"
        Exit Sub
    End If

    nom = Trim$(Donnees_ValeurAffichee(ligne, "Entreprise") & " " & _
                Donnees_ValeurAffichee(ligne, "Nom") & " " & _
                Donnees_ValeurAffichee(ligne, "Prenom"))

    question = "Supprimer définitivement la fiche " & clef & " ?" & vbCrLf & vbCrLf & nom
    nbInterv = Donnees_NbInterventions(Donnees_ValeurAffichee(ligne, COL_ID_CRESUS))
    If nbInterv > 0 Then
        question = question & vbCrLf & vbCrLf & _
                   "Attention : " & nbInterv & " intervention(s) du tableau " & NOM_TABLE_INTERV & _
                   " font référence à ce client."
    End If

    If MsgBox(question, vbQuestion + vbYesNo + vbDefaultButton2, "Suppression") <> vbYes Then Exit Sub

    If Donnees_Supprimer(clef) Then
        mClefCourante = vbNullString
        ViderChamps f
        Clients_RafraichirListe f
        MajBoutons f
        MajEntete f, "Fiche " & clef & " supprimée."
    End If
    Exit Sub

Erreur:
    MsgBox "Suppression impossible :" & vbCrLf & vbCrLf & Err.Description, vbCritical, "Gestion des clients"
End Sub

'==============================================================================
' BOUTON "EFFACER"
'==============================================================================
'------------------------------------------------------------------------------
' Bouton Effacer : vide les zones de saisie et repasse en mode nouvelle fiche.
' Le tableau et la feuille ne sont pas touchés, et le filtre reste en place.
'------------------------------------------------------------------------------
Public Sub Clients_Effacer(f As Object)
    Dim garde As Boolean

    mClefCourante = vbNullString
    ViderChamps f

    garde = mChargement
    mChargement = True
    On Error Resume Next
    Ctl(f, "lstClients").ListIndex = -1
    On Error GoTo 0
    mChargement = garde

    MajBoutons f
    MajEntete f
    On Error Resume Next
    Ctl(f, NomControleColonne("Entreprise")).SetFocus
    On Error GoTo 0
End Sub

'==============================================================================
' BOUTON "QUITTER"
'==============================================================================
'------------------------------------------------------------------------------
' Bouton Quitter, croix du bandeau et touche Échap : ferme le formulaire.
'------------------------------------------------------------------------------
Public Sub Clients_Quitter(f As Object)
    Unload f
End Sub

'------------------------------------------------------------------------------
' Remet l'état à zéro à la fermeture, pour que la prochaine ouverture reparte
' proprement.
'------------------------------------------------------------------------------
Public Sub Clients_Fermeture(f As Object)
    mClefCourante = vbNullString
    mNbVis = 0
End Sub

'==============================================================================
' CONTROLES DE SAISIE
'==============================================================================
'------------------------------------------------------------------------------
' Contrôle la saisie avant écriture.
'   msg     : renseigné avec le motif du refus
'   renvoie : True si la fiche peut être écrite
'
' Contrôles volontairement souples : ils empêchent les fautes de frappe, pas la
' saisie d'une fiche incomplète, courante en cours de constitution.
'------------------------------------------------------------------------------
Private Function Valider(f As Object, ByRef msg As String) As Boolean
    Dim s As String, npa As String, ok As Boolean

    msg = vbNullString

    If Len(Trim$(EnTexte(Ctl(f, NomControleColonne("Nom")).Value))) = 0 _
       And Len(Trim$(EnTexte(Ctl(f, NomControleColonne("Entreprise")).Value))) = 0 Then
        msg = "Renseignez au moins le nom du client ou le nom de l'entreprise."
        Exit Function
    End If

    npa = Trim$(EnTexte(Ctl(f, NomControleColonne(COL_NPA)).Value))
    If Len(npa) > 0 Then
        If Len(npa) <> 4 Or Not EstNumerique(npa) Then
            msg = "Le NPA doit comporter 4 chiffres (valeur saisie : " & npa & ")."
            Exit Function
        End If
    End If

    s = Trim$(EnTexte(Ctl(f, NomControleColonne(COL_ID_CRESUS)).Value))
    If Len(s) > 0 Then
        If Not EstNumerique(s) Then
            msg = "L'ID Crésus doit être un nombre entier."
            Exit Function
        End If
    End If

    s = Trim$(EnTexte(Ctl(f, NomControleColonne("Tx_hrs_Forf")).Value))
    If Len(s) > 0 Then
        ok = False
        EnNombre s, ok
        If Not ok Then
            msg = "Le taux horaire / forfait doit être un nombre."
            Exit Function
        End If
    End If

    s = Trim$(EnTexte(Ctl(f, NomControleColonne("Email")).Value))
    If Len(s) > 0 Then
        If InStr(s, "@") < 2 Or InStr(InStr(s, "@"), s, ".") = 0 Or Right$(s, 1) = "." Then
            msg = "L'adresse de courriel ne semble pas valide : " & s
            Exit Function
        End If
    End If

    Valider = True
End Function

'------------------------------------------------------------------------------
' True si la chaîne ne contient que des chiffres.
' Plus strict qu'IsNumeric, qui accepterait « 1e5 », « -3 » ou « 1,5 » — inadaptés
' pour un NPA ou un identifiant.
'------------------------------------------------------------------------------
Private Function EstNumerique(ByVal s As String) As Boolean
    Dim i As Long, c As String
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If c < "0" Or c > "9" Then Exit Function
    Next i
    EstNumerique = True
End Function

'==============================================================================
' RECHERCHE D'ADRESSE (onglet Adresses)
'==============================================================================
'------------------------------------------------------------------------------
' Menu déroulant Adresse : renseigne NPA, ville et canton à partir de la rue.
' Déclenchée à chaque frappe ; ne fait rien tant que la rue saisie ne correspond
' à aucune adresse connue.
'------------------------------------------------------------------------------
Public Sub Clients_AdresseChoisie(f As Object)
    Dim rue As String, npa As String, ville As String, canton As String
    Dim n As Long, garde As Boolean

    If mChargement Then Exit Sub
    rue = Trim$(EnTexte(Ctl(f, NomControleColonne(COL_ADRESSE)).Value))
    If Len(rue) = 0 Then Exit Sub

    n = Adresses_ChercherParRue(rue, Trim$(EnTexte(Ctl(f, NomControleColonne(COL_NPA)).Value)), _
                                npa, ville, canton)
    If n = 0 Then Exit Sub

    garde = mChargement
    mChargement = True
    If Len(npa) > 0 Then Ctl(f, NomControleColonne(COL_NPA)).Value = npa
    If Len(ville) > 0 Then Ctl(f, NomControleColonne(COL_VILLE)).Value = ville
    If Len(canton) > 0 Then Ctl(f, NomControleColonne(COL_CANTON)).Value = canton
    mChargement = garde
End Sub

'------------------------------------------------------------------------------
' Menu déroulant NPA : renseigne la ville et le canton.
' N'agit qu'à partir de quatre caractères, pour ne pas réagir aux frappes
' intermédiaires.
'------------------------------------------------------------------------------
Public Sub Clients_NpaChoisi(f As Object)
    Dim npa As String, ville As String, canton As String, garde As Boolean

    If mChargement Then Exit Sub
    npa = Trim$(EnTexte(Ctl(f, NomControleColonne(COL_NPA)).Value))
    If Len(npa) < 4 Then Exit Sub
    If Not Adresses_ChercherParNpa(npa, ville, canton) Then Exit Sub

    garde = mChargement
    mChargement = True
    If Len(ville) > 0 Then Ctl(f, NomControleColonne(COL_VILLE)).Value = ville
    If Len(canton) > 0 Then Ctl(f, NomControleColonne(COL_CANTON)).Value = canton
    mChargement = garde
End Sub

'==============================================================================
' SAISIE NUMERIQUE
'==============================================================================
'------------------------------------------------------------------------------
' Filtre les frappes dans un champ numérique.
'   KeyAscii   : reçu en Object pour éviter toute dépendance à MSForms dans ce
'                module ; mettre sa valeur à 0 annule la frappe
'   decimales  : True pour accepter aussi le point et la virgule
'------------------------------------------------------------------------------
Public Sub Clients_ToucheNumerique(ByVal KeyAscii As Object, ByVal decimales As Boolean)
    Dim c As String
    ' KeyAscii arrive en MSForms.ReturnInteger : un objet, dont la propriété
    ' Value se lit et s'écrit. Y mettre 0 revient à ne jamais avoir tapé.
    If KeyAscii.Value = 8 Then Exit Sub          ' retour arrière
    If KeyAscii.Value < 32 Then Exit Sub         ' touches de contrôle
    c = Chr$(KeyAscii.Value)
    If c >= "0" And c <= "9" Then Exit Sub
    If decimales And (c = "." Or c = ",") Then Exit Sub
    KeyAscii.Value = 0
End Sub

'==============================================================================
' HABILLAGE : survol, focus, déplacement de la fenêtre
'==============================================================================
'------------------------------------------------------------------------------
' Met à jour l'effet de survol.
'   nomControle : contrôle sous la souris, chaîne vide si la souris est ailleurs
'
' Un seul point d'entrée pour tous les contrôles survolables : chacun est remis
' dans son état normal sauf celui qui est nommé. C'est ce qui remet un bouton à
' sa couleur quand la souris le quitte, MSForms ne signalant pas la sortie.
'------------------------------------------------------------------------------
Public Sub Clients_Survol(f As Object, ByVal nomControle As String)
    Dim boutons As Variant, i As Long, b As Object, lb As Object

    boutons = Array("btnAjouter", "btnModifier", "btnSupprimer", "btnEffacer", "btnQuitter")
    For i = LBound(boutons) To UBound(boutons)
        Set b = Ctl(f, CStr(boutons(i)))
        If Not b Is Nothing Then
            If b.Enabled Then
                CouleurSi b, CouleurBouton(CStr(boutons(i)), CStr(boutons(i)) = nomControle)
            End If
        End If
    Next i

    Set lb = Ctl(f, "lblFermer")
    If Not lb Is Nothing Then
        If nomControle = "lblFermer" Then
            If lb.BackStyle <> MSF_BackStyleOpaque Then lb.BackStyle = MSF_BackStyleOpaque
            CouleurSi lb, COUL_FERMER_H
        Else
            If lb.BackStyle <> MSF_BackStyleTransparent Then lb.BackStyle = MSF_BackStyleTransparent
        End If
    End If

    Set lb = Ctl(f, "lblResetFiltre")
    If Not lb Is Nothing Then
        CouleurTexteSi lb, IIf(nomControle = "lblResetFiltre", COUL_LIEN_H, COUL_LIEN)
    End If
End Sub

'------------------------------------------------------------------------------
' Change la couleur de fond seulement si elle diffère : réécrire la même
' valeur à chaque MouseMove ferait scintiller le contrôle.
'------------------------------------------------------------------------------
Private Sub CouleurSi(c As Object, ByVal couleur As Long)
    If c.BackColor <> couleur Then c.BackColor = couleur
End Sub

'------------------------------------------------------------------------------
' Change la couleur du texte seulement si elle diffère.
'------------------------------------------------------------------------------
Private Sub CouleurTexteSi(c As Object, ByVal couleur As Long)
    If c.ForeColor <> couleur Then c.ForeColor = couleur
End Sub

'------------------------------------------------------------------------------
' Souligne le champ actif : filet bleu à l'entrée, filet gris à la sortie.
'------------------------------------------------------------------------------
Public Sub Clients_FocusChamp(f As Object, ByVal nomControle As String, ByVal actif As Boolean)
    Dim c As Object
    Set c = Ctl(f, nomControle)
    If c Is Nothing Then Exit Sub
    On Error Resume Next
    c.BorderColor = IIf(actif, COUL_CHAMP_FOCUS, COUL_CHAMP_BORD)
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Début du déplacement de la fenêtre : mémorise le point de saisie dans le
' bandeau.
'------------------------------------------------------------------------------
Public Sub Clients_DebutDeplacement(ByVal x As Single, ByVal y As Single)
    mDeplacement = True
    mDepX = x
    mDepY = y
End Sub

'------------------------------------------------------------------------------
' Déplace la fenêtre pendant le glissement.
' Le formulaire n'a pas de barre de titre système : c'est ce trio de procédures
' qui en tient lieu. Le déplacement se calcule en écart par rapport au point de
' départ, ce qui fonctionne quel que soit le contrôle du bandeau saisi.
'------------------------------------------------------------------------------
Public Sub Clients_Deplacer(f As Object, ByVal bouton As Integer, ByVal x As Single, ByVal y As Single)
    If Not mDeplacement Then Exit Sub
    If bouton <> 1 Then Exit Sub
    f.Left = f.Left + (x - mDepX)
    f.Top = f.Top + (y - mDepY)
End Sub

'------------------------------------------------------------------------------
' Fin du déplacement, au relâchement du bouton.
'------------------------------------------------------------------------------
Public Sub Clients_FinDeplacement()
    mDeplacement = False
End Sub

'==============================================================================
' BANDEAU ET BOUTONS : mise à jour de l'état
'==============================================================================
'------------------------------------------------------------------------------
' Met à jour la deuxième ligne du bandeau.
'   message : texte à afficher ; si vide, l'état courant est affiché
' Sert de zone d'information discrète : « fiche ajoutée », « fiche mise à jour »,
' là où une boîte de dialogue interromprait la saisie.
'------------------------------------------------------------------------------
Private Sub MajEntete(f As Object, Optional ByVal message As String = vbNullString)
    Dim lb As Object, txt As String

    Set lb = Ctl(f, "lblSousTitre")
    If lb Is Nothing Then Exit Sub

    If Len(message) > 0 Then
        txt = message
    ElseIf Len(mClefCourante) = 0 Then
        txt = "Nouvelle fiche — remplissez les champs puis cliquez sur Ajouter."
    Else
        txt = "Fiche " & mClefCourante & " sélectionnée — modifiez puis cliquez sur Modifier."
    End If
    lb.Caption = txt
End Sub

'------------------------------------------------------------------------------
' Met à jour le compteur de la barre de filtrage.
'------------------------------------------------------------------------------
Private Sub MajCompteur(f As Object)
    Dim lb As Object
    Set lb = Ctl(f, "lblCompteur")
    If lb Is Nothing Then Exit Sub
    lb.Caption = mNbVis & " fiche(s) affichée(s) sur " & Donnees_NbLignes()
End Sub

'------------------------------------------------------------------------------
' Active ou grise Modifier et Supprimer selon qu'une fiche est sélectionnée.
'------------------------------------------------------------------------------
Private Sub MajBoutons(f As Object)
    Dim actif As Boolean
    actif = (Len(mClefCourante) > 0)
    ActiverBouton f, "btnModifier", actif
    ActiverBouton f, "btnSupprimer", actif
End Sub

'------------------------------------------------------------------------------
' Active ou grise un bouton, couleur comprise.
'------------------------------------------------------------------------------
Private Sub ActiverBouton(f As Object, ByVal nom As String, ByVal actif As Boolean)
    Dim b As Object
    Set b = Ctl(f, nom)
    If b Is Nothing Then Exit Sub
    b.Enabled = actif
    b.BackColor = IIf(actif, CouleurBouton(nom, False), COUL_BOUTON_OFF)
End Sub

'==============================================================================
' Accès aux contrôles du formulaire
'==============================================================================
'------------------------------------------------------------------------------
' Renvoie un contrôle du formulaire par son nom, ou Nothing s'il n'existe pas.
'
' Le formulaire est reçu en Object, jamais en UF_Clients : c'est ce qui permet à
' ce module de compiler avant que le formulaire ait été généré, et de survivre à
' sa régénération.
'------------------------------------------------------------------------------
Private Function Ctl(f As Object, ByVal nom As String) As Object
    On Error Resume Next
    Set Ctl = f.Controls(nom)
    On Error GoTo 0
End Function
