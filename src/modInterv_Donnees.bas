Attribute VB_Name = "modInterv_Donnees"
Option Explicit
'==============================================================================
' modInterv_Donnees
'------------------------------------------------------------------------------
' Accès aux données du formulaire des interventions :
'   - lecture et écriture du tableau TblInterv ;
'   - recherches dans TblClients pour la saisie assistée ;
'   - lecture des cellules nommées de la feuille Statistiques ;
'   - export du graphique CAGraphique en image.
'
' Les utilitaires généraux — Normaliser, EnNombre, EnTexte, ObtenirTable,
' IndexColonne, TrierChaines — viennent des modules du formulaire des clients :
' ils sont partagés, pas recopiés.
'==============================================================================

Private mDonnees As Variant         ' cache de TblInterv : (1 To n, 1 To nbCol)
Private mNbLignes As Long
Private mNbColonnes As Long
Private mIdxColonne As Object       ' nom de colonne -> numéro
Private mCalculees As Object        ' numéros des colonnes portant une formule

Private mClients As Variant         ' cache de TblClients
Private mIdxClient As Object
Private mNbClients As Long
Private mClientsCharges As Boolean

'==============================================================================
' CHARGEMENT
'==============================================================================
'------------------------------------------------------------------------------
' Relit TblInterv et le range en mémoire, en relevant au passage les colonnes
' calculées : ce sont celles qu'il ne faudra jamais écrire, sous peine de
' remplacer leur formule par une valeur figée.
'------------------------------------------------------------------------------
Public Sub Interv_Charger()
    Dim lo As ListObject, lc As ListColumn, c As Range

    Set lo = TableInterventions()
    If lo Is Nothing Then
        Err.Raise vbObjectError + 520, "modInterv_Donnees", _
                  "Le tableau " & NOM_TABLE_INTERVENTIONS & " est introuvable dans ce classeur."
    End If

    Set mIdxColonne = CreateObject("Scripting.Dictionary")
    mIdxColonne.CompareMode = 1
    Set mCalculees = CreateObject("Scripting.Dictionary")
    For Each lc In lo.ListColumns
        If Not mIdxColonne.Exists(lc.Name) Then mIdxColonne.Add lc.Name, lc.Index
        If Not lc.DataBodyRange Is Nothing Then
            Set c = lc.DataBodyRange.Cells(1, 1)
            If c.HasFormula Then mCalculees(lc.Index) = True
        End If
    Next lc
    mNbColonnes = lo.ListColumns.Count

    If lo.ListRows.Count = 0 Then
        mNbLignes = 0
        mDonnees = Empty
    Else
        mNbLignes = lo.ListRows.Count
        mDonnees = lo.DataBodyRange.Value
    End If
End Sub

'------------------------------------------------------------------------------
' Charge le cache des interventions au premier besoin. Placée en tête des
' procédures de lecture pour qu'aucune ne dépende de l'ordre des appels.
'------------------------------------------------------------------------------
Private Sub AssurerInterv()
    If mIdxColonne Is Nothing Then Interv_Charger
End Sub

'------------------------------------------------------------------------------
' Relit TblClients : sert aux listes d'entreprises et de noms, et au report des
' informations du client dans la fiche d'intervention.
'------------------------------------------------------------------------------
Private Sub ChargerClients()
    Dim lo As ListObject, lc As ListColumn

    mClientsCharges = True
    mClients = Empty
    mNbClients = 0
    Set mIdxClient = CreateObject("Scripting.Dictionary")
    mIdxClient.CompareMode = 1

    Set lo = ObtenirTable(NOM_TABLE_CLIENTS)
    If lo Is Nothing Then Exit Sub
    For Each lc In lo.ListColumns
        If Not mIdxClient.Exists(lc.Name) Then mIdxClient.Add lc.Name, lc.Index
    Next lc
    If lo.ListRows.Count = 0 Then Exit Sub

    mNbClients = lo.ListRows.Count
    mClients = lo.DataBodyRange.Value
End Sub

'------------------------------------------------------------------------------
' Charge le cache des clients au premier besoin.
'------------------------------------------------------------------------------
Private Sub AssurerClients()
    If Not mClientsCharges Then ChargerClients
End Sub

'------------------------------------------------------------------------------
' Force la relecture des deux caches au prochain accès. Appelée à l'ouverture
' du formulaire, pour prendre en compte ce qui a changé dans les feuilles depuis
' la dernière fois.
'------------------------------------------------------------------------------
Public Sub Interv_ToutRecharger()
    Set mIdxColonne = Nothing
    mClientsCharges = False
End Sub

'==============================================================================
' LECTURE DU TABLEAU
'==============================================================================
Public Function Interv_NbLignes() As Long
    AssurerInterv
    Interv_NbLignes = mNbLignes
End Function

'------------------------------------------------------------------------------
' Position d'une colonne dans le cache.
'   renvoie : le numéro de colonne, ou 0 si elle n'existe pas
'------------------------------------------------------------------------------
Public Function Interv_IndexColonne(ByVal nomColonne As String) As Long
    AssurerInterv
    If mIdxColonne.Exists(nomColonne) Then Interv_IndexColonne = mIdxColonne(nomColonne)
End Function

'------------------------------------------------------------------------------
' Valeur brute d'une cellule. Empty si la ligne ou la colonne n'existe pas,
' plutôt qu'une erreur : les appelants n'ont pas à se protéger.
'------------------------------------------------------------------------------
Public Function Interv_Valeur(ByVal ligne As Long, ByVal nomColonne As String) As Variant
    Dim ic As Long
    AssurerInterv
    ic = Interv_IndexColonne(nomColonne)
    If ic = 0 Or ligne < 1 Or ligne > mNbLignes Then
        Interv_Valeur = Empty
    Else
        Interv_Valeur = mDonnees(ligne, ic)
    End If
End Function

'------------------------------------------------------------------------------
' Valeur mise en forme pour l'affichage : dates en jj/mm/aaaa, durées en
' heures:minutes, montants à deux décimales, booléens en Oui / Non.
'------------------------------------------------------------------------------
Public Function Interv_ValeurAffichee(ByVal ligne As Long, ByVal nomColonne As String) As String
    Dim v As Variant
    v = Interv_Valeur(ligne, nomColonne)

    If IsEmpty(v) Or IsNull(v) Then
        Interv_ValeurAffichee = vbNullString
    ElseIf VarType(v) = vbBoolean Then
        Interv_ValeurAffichee = IIf(v, "Oui", "Non")
    ElseIf StrComp(nomColonne, IC_DATE, vbTextCompare) = 0 Then
        Interv_ValeurAffichee = Interv_DateAffichee(v)
    ElseIf StrComp(nomColonne, IC_HEURES, vbTextCompare) = 0 Then
        Interv_ValeurAffichee = Interv_HeuresVersTexte(v)
    ElseIf StrComp(nomColonne, IC_CA, vbTextCompare) = 0 Then
        If IsNumeric(v) Then Interv_ValeurAffichee = Format$(CDbl(v), "#,##0.00")
    Else
        Interv_ValeurAffichee = CStr(v)
    End If
End Function

'------------------------------------------------------------------------------
' Met une date en forme jj/mm/aaaa, qu'Excel la rende en Date ou en numéro de
' série.
'------------------------------------------------------------------------------
Public Function Interv_DateAffichee(ByVal v As Variant) As String
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If VarType(v) = vbDate Then
        Interv_DateAffichee = Format$(v, "dd/mm/yyyy")
    ElseIf IsNumeric(v) Then
        If CDbl(v) > 0 Then Interv_DateAffichee = Format$(CDate(CDbl(v)), "dd/mm/yyyy")
    ElseIf IsDate(v) Then
        Interv_DateAffichee = Format$(CDate(v), "dd/mm/yyyy")
    End If
End Function

'------------------------------------------------------------------------------
' Numéro de ligne portant un NoInterv donné ; 0 si introuvable.
'------------------------------------------------------------------------------
Public Function Interv_TrouverLigne(ByVal noInterv As String) As Long
    Dim i As Long, ic As Long
    AssurerInterv
    If Len(noInterv) = 0 Then Exit Function
    ic = Interv_IndexColonne(IC_NO)
    If ic = 0 Then Exit Function
    For i = 1 To mNbLignes
        If StrComp(CStr(mDonnees(i, ic) & ""), noInterv, vbTextCompare) = 0 Then
            Interv_TrouverLigne = i
            Exit Function
        End If
    Next i
End Function

'------------------------------------------------------------------------------
' Numéro de la prochaine intervention : IN suivi du plus grand numéro + 1.
' Recalculé à chaque fois plutôt que déduit du nombre de lignes, pour qu'une
' suppression au milieu du tableau ne provoque jamais de doublon.
'------------------------------------------------------------------------------
Public Function Interv_NouveauNumero() As String
    Dim i As Long, ic As Long, maxi As Long, n As Long
    AssurerInterv
    ic = Interv_IndexColonne(IC_NO)
    If ic > 0 Then
        For i = 1 To mNbLignes
            n = ChiffresDe(CStr(mDonnees(i, ic) & ""))
            If n > maxi Then maxi = n
        Next i
    End If
    Interv_NouveauNumero = PREFIXE_NO_INTERV & CStr(maxi + 1)
End Function

'------------------------------------------------------------------------------
' Extrait les chiffres d'une chaîne : IN29 donne 29.
' Renvoie 0 si la chaîne n'en contient aucun, ou plus de neuf — au-delà, la
' conversion en Long déborderait.
'------------------------------------------------------------------------------
Private Function ChiffresDe(ByVal s As String) As Long
    Dim i As Long, c As String, res As String
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If c >= "0" And c <= "9" Then res = res & c
    Next i
    If Len(res) > 0 And Len(res) <= 9 Then ChiffresDe = CLng(res)
End Function

'==============================================================================
' DURÉES
'------------------------------------------------------------------------------
' Excel range Nb_Hres en fraction de jour et l'affiche au format [h]:mm : 0,125
' vaut 3 heures, 17,5 vaut 420 heures. Le formulaire, lui, parle en heures.
'==============================================================================
'------------------------------------------------------------------------------
' Fraction de jour -> "420:00". Le calcul passe par les minutes entières pour
' éviter qu'un arrondi ne produise un absurde 419:60.
'------------------------------------------------------------------------------
Public Function Interv_HeuresVersTexte(ByVal v As Variant) As String
    Dim minutes As Long
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    minutes = CLng(Round(CDbl(v) * 1440#, 0))
    Interv_HeuresVersTexte = CStr(minutes \ 60) & ":" & Format$(minutes Mod 60, "00")
End Function

'------------------------------------------------------------------------------
' "420:00", "420" ou "420:30" -> fraction de jour.
'   ok : passé à True si la saisie a pu être interprétée
'------------------------------------------------------------------------------
Public Function Interv_TexteVersHeures(ByVal s As String, ByRef ok As Boolean) As Double
    Dim p As Long, h As String, m As String, t As String

    ok = False
    t = Trim$(Replace$(s, " ", vbNullString))
    t = Replace$(t, "h", vbNullString)
    t = Replace$(t, "H", vbNullString)
    If Len(t) = 0 Then Exit Function

    p = InStr(t, ":")
    If p = 0 Then
        h = t
        m = "0"
    Else
        h = Left$(t, p - 1)
        m = Mid$(t, p + 1)
        If Len(m) = 0 Then m = "0"
    End If
    If Len(h) = 0 Then h = "0"

    If Not QueDesChiffres(h) Or Not QueDesChiffres(m) Then Exit Function
    If CDbl(m) > 59 Then Exit Function

    ok = True
    Interv_TexteVersHeures = (CDbl(h) * 60# + CDbl(m)) / 1440#
End Function

'------------------------------------------------------------------------------
' True si la chaîne ne contient que des chiffres.
' Plus strict qu'IsNumeric, qui accepterait « 1e5 », « -3 » ou « 1,5 » — inadaptés
' pour un nombre de personnes.
'------------------------------------------------------------------------------
Public Function QueDesChiffres(ByVal s As String) As Boolean
    Dim i As Long, c As String
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If c < "0" Or c > "9" Then Exit Function
    Next i
    QueDesChiffres = True
End Function

'==============================================================================
' ÉCRITURES
'------------------------------------------------------------------------------
' valeurs : Dictionary nom de colonne -> valeur déjà typée.
' Les colonnes calculées et NoInterv sont écartées : la première catégorie
' garde ses formules, la seconde est posée ici.
'==============================================================================
Public Function IntervBD_Ajouter(ByVal valeurs As Object) As String
    Dim lo As ListObject, lr As ListRow, noInterv As String

    AssurerInterv
    Set lo = TableInterventions()
    noInterv = Interv_NouveauNumero()

    On Error GoTo Fin
    Application.EnableEvents = False
    Set lr = lo.ListRows.Add
    EcrireCellules lr.Range, valeurs, noInterv

Fin:
    Application.EnableEvents = True
    If Err.Number <> 0 Then Err.Raise Err.Number, "IntervBD_Ajouter", Err.Description

    Interv_Charger
    IntervBD_Ajouter = noInterv
End Function

'------------------------------------------------------------------------------
' Met à jour une intervention existante.
'   noInterv : numéro de l'intervention à modifier
'   valeurs  : Dictionary nom de colonne -> valeur ; les colonnes absentes du
'              dictionnaire gardent leur valeur actuelle
'   renvoie  : True si l'intervention a été trouvée et écrite
'------------------------------------------------------------------------------
Public Function IntervBD_Modifier(ByVal noInterv As String, ByVal valeurs As Object) As Boolean
    Dim lo As ListObject, i As Long

    AssurerInterv
    i = Interv_TrouverLigne(noInterv)
    If i = 0 Then Exit Function
    Set lo = TableInterventions()

    On Error GoTo Fin
    Application.EnableEvents = False
    EcrireCellules lo.ListRows(i).Range, valeurs, noInterv

Fin:
    Application.EnableEvents = True
    If Err.Number <> 0 Then Err.Raise Err.Number, "IntervBD_Modifier", Err.Description

    Interv_Charger
    IntervBD_Modifier = True
End Function

'------------------------------------------------------------------------------
' Supprime la ligne correspondant à un numéro d'intervention.
'   renvoie : True si l'intervention existait
' La confirmation de l'utilisateur est demandée en amont, par le formulaire.
'------------------------------------------------------------------------------
Public Function IntervBD_Supprimer(ByVal noInterv As String) As Boolean
    Dim lo As ListObject, i As Long

    AssurerInterv
    i = Interv_TrouverLigne(noInterv)
    If i = 0 Then Exit Function
    Set lo = TableInterventions()

    On Error GoTo Fin
    Application.EnableEvents = False
    lo.ListRows(i).Delete

Fin:
    Application.EnableEvents = True
    If Err.Number <> 0 Then Err.Raise Err.Number, "IntervBD_Supprimer", Err.Description

    Interv_Charger
    IntervBD_Supprimer = True
End Function

'------------------------------------------------------------------------------
' Écrit une ligne cellule par cellule, en sautant les colonnes calculées.
'
' Écrire la ligne entière d'un seul coup, comme pour le formulaire des clients,
' est impossible ici : la colonne CA porte une formule, qu'une écriture globale
' remplacerait par une valeur figée. Excel la recalcule tout seul dès que les
' colonnes dont elle dépend sont renseignées.
'------------------------------------------------------------------------------
Private Sub EcrireCellules(ByVal ligne As Range, ByVal valeurs As Object, ByVal noInterv As String)
    Dim k As Variant, ic As Long

    ic = Interv_IndexColonne(IC_NO)
    If ic > 0 Then ligne.Cells(1, ic).Value = noInterv

    For Each k In valeurs.Keys
        If StrComp(CStr(k), IC_NO, vbTextCompare) <> 0 Then
            ic = Interv_IndexColonne(CStr(k))
            If ic > 0 Then
                If Not mCalculees.Exists(ic) Then ligne.Cells(1, ic).Value = valeurs(k)
            End If
        End If
    Next k
End Sub

'------------------------------------------------------------------------------
' True si la colonne porte une formule et ne doit donc pas être écrite.
'------------------------------------------------------------------------------
Public Function Interv_EstCalculee(ByVal nomColonne As String) As Boolean
    Dim ic As Long
    AssurerInterv
    ic = Interv_IndexColonne(nomColonne)
    If ic > 0 Then Interv_EstCalculee = mCalculees.Exists(ic)
End Function

'==============================================================================
' SAISIE ASSISTÉE DEPUIS TblClients
'==============================================================================
'------------------------------------------------------------------------------
' Valeurs distinctes non vides d'une colonne de TblClients, triées.
' Sert à alimenter les listes Entreprise et Nom.
'------------------------------------------------------------------------------
Public Function Interv_ListeClients(ByVal nomColonne As String) As Variant
    Dim d As Object, i As Long, ic As Long, s As String

    AssurerClients
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    If mIdxClient.Exists(nomColonne) And IsArray(mClients) Then
        ic = mIdxClient(nomColonne)
        For i = 1 To mNbClients
            s = Trim$(CStr(mClients(i, ic) & ""))
            If Len(s) > 0 Then
                If Not d.Exists(s) Then d.Add s, 1
            End If
        Next i
    End If
    Interv_ListeClients = TrierChaines(d.Keys)
End Function

'------------------------------------------------------------------------------
' Ligne de TblClients dont une colonne vaut exactement une valeur donnée.
'   renvoie : le numéro de ligne dans le cache des clients, ou 0
'------------------------------------------------------------------------------
Public Function Interv_TrouverClient(ByVal nomColonne As String, ByVal valeur As String) As Long
    Dim i As Long, ic As Long, cible As String

    AssurerClients
    cible = Normaliser(Trim$(valeur))
    If Len(cible) = 0 Then Exit Function
    If Not mIdxClient.Exists(nomColonne) Or Not IsArray(mClients) Then Exit Function

    ic = mIdxClient(nomColonne)
    For i = 1 To mNbClients
        If Normaliser(Trim$(CStr(mClients(i, ic) & ""))) = cible Then
            Interv_TrouverClient = i
            Exit Function
        End If
    Next i
End Function

'------------------------------------------------------------------------------
' Valeur d'une colonne de TblClients pour une ligne du cache des clients.
'------------------------------------------------------------------------------
Public Function Interv_ValeurClient(ByVal ligne As Long, ByVal nomColonne As String) As Variant
    AssurerClients
    If ligne < 1 Or ligne > mNbClients Then Exit Function
    If Not mIdxClient.Exists(nomColonne) Then Exit Function
    Interv_ValeurClient = mClients(ligne, mIdxClient(nomColonne))
End Function

'------------------------------------------------------------------------------
' Taux horaire ou montant forfaitaire d'un client, cherché par son ID Crésus.
' C'est la valeur qu'utilise la formule de la colonne CA.
'------------------------------------------------------------------------------
Public Function Interv_TauxClient(ByVal idCresus As String) As Double
    Dim i As Long, icId As Long, icTx As Long, cible As String, v As Variant

    AssurerClients
    cible = Trim$(idCresus)
    If Len(cible) = 0 Then Exit Function
    If Not mIdxClient.Exists("ID_Cresus") Or Not mIdxClient.Exists("Tx_hrs_Forf") Then Exit Function
    If Not IsArray(mClients) Then Exit Function

    icId = mIdxClient("ID_Cresus")
    icTx = mIdxClient("Tx_hrs_Forf")
    For i = 1 To mNbClients
        If StrComp(Trim$(CStr(mClients(i, icId) & "")), cible, vbTextCompare) = 0 Then
            v = mClients(i, icTx)
            If IsNumeric(v) Then Interv_TauxClient = CDbl(v)
            Exit Function
        End If
    Next i
End Function

'------------------------------------------------------------------------------
' Estimation du chiffre d'affaires, reprise de la formule de la colonne CA :
'   au forfait          -> le montant forfaitaire du client
'   à l'heure           -> taux x personnes x heures
' Elle sert à afficher une valeur pendant la saisie ; la valeur définitive est
' celle que le tableau Excel calcule après écriture.
'------------------------------------------------------------------------------
Public Function Interv_EstimerCA(ByVal idCresus As String, ByVal forfait As Boolean, _
                                 ByVal nbPers As Double, ByVal heures As Double) As Double
    Dim taux As Double
    taux = Interv_TauxClient(idCresus)
    If forfait Then
        Interv_EstimerCA = taux
    Else
        Interv_EstimerCA = taux * nbPers * heures * 24#
    End If
End Function

'==============================================================================
' FEUILLE STATISTIQUES
'==============================================================================
'------------------------------------------------------------------------------
' Valeur d'une cellule nommée du classeur.
'   renvoie : la valeur, ou Empty si le nom n'existe pas
'------------------------------------------------------------------------------
Public Function Interv_CelluleNommee(ByVal nom As String) As Variant
    On Error Resume Next
    Interv_CelluleNommee = ThisWorkbook.Names(nom).RefersToRange.Value
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' Montant d'une cellule nommée, mis en forme pour l'affichage.
'------------------------------------------------------------------------------
Public Function Interv_MontantNomme(ByVal nom As String) As String
    Dim v As Variant
    v = Interv_CelluleNommee(nom)
    If IsNumeric(v) Then
        Interv_MontantNomme = Format$(CDbl(v), "#,##0") & " CHF"
    Else
        Interv_MontantNomme = "-"
    End If
End Function

'------------------------------------------------------------------------------
' Exporte le graphique CAGraphique en image, pour l'afficher dans le formulaire.
'   renvoie : le chemin du fichier créé, ou une chaîne vide en cas d'échec
'
' MSForms ne sait pas afficher un graphique Excel : il faut passer par une
' image. Le fichier est écrit dans le dossier temporaire de Windows et réécrit
' à chaque ouverture, pour refléter les données du moment.
'------------------------------------------------------------------------------
Public Function Interv_ExporterGraphique(Optional ByRef motif As String) As String
    Dim ws As Worksheet, gr As Chart, chemin As String

    motif = vbNullString
    On Error GoTo Erreur

    Set ws = ThisWorkbook.Worksheets(NOM_FEUILLE_STATS)
    Set gr = GraphiqueDeLaFeuille(ws)
    If gr Is Nothing Then
        motif = "graphique " & NOM_GRAPHIQUE & " introuvable sur la feuille " & NOM_FEUILLE_STATS
        Exit Function
    End If

    ' GIF et non PNG : LoadPicture vient de la bibliothèque OLE, qui ne lit que
    ' bmp, ico, wmf, emf, gif et jpg. Un PNG s'exporte sans erreur mais reste
    ' ensuite illisible — le graphique n'apparaissait pas, en silence.
    chemin = Environ$("TEMP") & Application.PathSeparator & NOM_GRAPHIQUE & ".gif"

    ' un fichier resté d'une session précédente ferait échouer l'export
    On Error Resume Next
    If Len(Dir$(chemin)) > 0 Then Kill chemin
    On Error GoTo Erreur

    gr.Export chemin, "GIF"
    If Len(Dir$(chemin)) = 0 Then
        motif = "l'export n'a produit aucun fichier"
        Exit Function
    End If

    Interv_ExporterGraphique = chemin
    Exit Function

Erreur:
    motif = "erreur " & Err.Number & " - " & Err.Description
End Function

'------------------------------------------------------------------------------
' Retrouve le graphique de la feuille des statistiques.
'   renvoie : le Chart, ou Nothing
'
' Trois tentatives : par le nom de l'objet graphique, puis par celui de la forme
' — les deux ne coïncident pas toujours — puis, s'il n'y a qu'un seul graphique
' sur la feuille, celui-là.
'------------------------------------------------------------------------------
Private Function GraphiqueDeLaFeuille(ws As Worksheet) As Chart
    Dim co As ChartObject, sh As Shape

    On Error Resume Next
    Set co = ws.ChartObjects(NOM_GRAPHIQUE)
    On Error GoTo 0
    If Not co Is Nothing Then
        Set GraphiqueDeLaFeuille = co.Chart
        Exit Function
    End If

    For Each sh In ws.Shapes
        If StrComp(sh.Name, NOM_GRAPHIQUE, vbTextCompare) = 0 Then
            On Error Resume Next
            Set GraphiqueDeLaFeuille = sh.Chart
            On Error GoTo 0
            If Not GraphiqueDeLaFeuille Is Nothing Then Exit Function
        End If
    Next sh

    If ws.ChartObjects.Count = 1 Then Set GraphiqueDeLaFeuille = ws.ChartObjects(1).Chart
End Function

'------------------------------------------------------------------------------
' Chemin de l'image de fond des tuiles de statistiques.
'   renvoie : le chemin si le fichier existe, une chaîne vide sinon
'------------------------------------------------------------------------------
Public Function Interv_CheminImageTuile() As String
    Dim chemin As String
    On Error Resume Next
    ' un classeur jamais enregistré n'a pas de dossier : rien à chercher
    If Len(ThisWorkbook.Path) = 0 Then Exit Function
    chemin = ThisWorkbook.Path & Application.PathSeparator & DOSSIER_IMAGES & _
             Application.PathSeparator & IMAGE_TUILE
    If Len(Dir$(chemin)) > 0 Then Interv_CheminImageTuile = chemin
    On Error GoTo 0
End Function
