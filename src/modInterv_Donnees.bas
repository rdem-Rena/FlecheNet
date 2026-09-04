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
' Valeur d'une CASE du tableau du formulaire.
'   cle : nom de colonne, ou plusieurs réunis par ICL_SEPARATEUR
'
' Une case peut réunir deux colonnes — le nom et le prénom — pour tenir dans les
' dix colonnes qu'accepte une ListBox. Les morceaux vides sont sautés, sans quoi
' un prénom absent laisserait une espace en fin de case.
Public Function Interv_ValeurListe(ByVal ligne As Long, ByVal cle As String) As String
    Dim src As Variant, i As Long, bout As String, r As String

    src = IColonnesSources(cle)
    For i = LBound(src) To UBound(src)
        bout = Trim$(Interv_ValeurAffichee(ligne, CStr(src(i))))
        If Len(bout) > 0 Then
            If Len(r) > 0 Then r = r & " "
            r = r & bout
        End If
    Next i
    Interv_ValeurListe = r
End Function

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
    ElseIf StrComp(nomColonne, IC_CA, vbTextCompare) = 0 _
           Or StrComp(nomColonne, IC_TAUX, vbTextCompare) = 0 Then
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
' Lit un montant saisi : « 39 », « 39.50 », « 39,50 », « 1'250.00 ».
'   ok : passé à True si le texte a pu être lu
'
' La virgule vaut le point — on tape « 39,50 » sur un clavier suisse romand —
' et l'apostrophe des milliers est retirée, puisque c'est ainsi que le montant
' est réaffiché. CDbl dépendrait des paramètres régionaux de Windows : la
' conversion est donc faite à la main, chiffre par chiffre.
Public Function Interv_TexteVersMontant(ByVal s As String, ByRef ok As Boolean) As Double
    Dim t As String, i As Long, c As String, ent As String, dec As String
    Dim apres As Boolean

    ok = False
    t = Replace$(Trim$(s), " ", vbNullString)
    t = Replace$(t, ChrW(39), vbNullString)      ' apostrophe des milliers
    t = Replace$(t, ChrW(8217), vbNullString)    ' apostrophe typographique
    t = Replace$(t, ",", ".")
    If Len(t) = 0 Then Exit Function

    For i = 1 To Len(t)
        c = Mid$(t, i, 1)
        If c >= "0" And c <= "9" Then
            If apres Then dec = dec & c Else ent = ent & c
        ElseIf c = "." And Not apres Then
            apres = True
        Else
            Exit Function                        ' caractère inattendu
        End If
    Next i

    If Len(ent) = 0 And Len(dec) = 0 Then Exit Function
    If Len(ent) = 0 Then ent = "0"

    ok = True
    Interv_TexteVersMontant = CDbl(ent)
    If Len(dec) > 0 Then
        Interv_TexteVersMontant = Interv_TexteVersMontant + CDbl(dec) / (10 ^ Len(dec))
    End If
End Function

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
' Chiffre d'affaires d'une intervention.
'   taux    : taux horaire, ou montant forfaitaire si forfait vaut True
'   nbPers  : nombre de personnes intervenues
'   heures  : durée en FRACTION DE JOUR, comme Excel la stocke
'
'   au forfait -> le montant forfaitaire, tel quel
'   à l'heure  -> taux x personnes x heures
'
' Le x 24 convertit la fraction de jour en heures : 0,125 jour vaut 3 heures.
' Oublier ce facteur diviserait tous les montants par vingt-quatre.
'
' Le taux vient désormais de la colonne Taux/Forfait de l'intervention, et non
' plus d'une recherche dans TblClients : le tarif est celui qui s'appliquait au
' moment de la prestation, et changer celui d'un client ne réécrit pas le passé.
'------------------------------------------------------------------------------
Public Function Interv_EstimerCA(ByVal taux As Double, ByVal forfait As Boolean, _
                                 ByVal nbPers As Double, ByVal heures As Double) As Double
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
' Chemin de l'image de fond des tuiles de statistiques.
'
' Le dossier vient de DossierClasseur (modChemins) et NON de ThisWorkbook.Path :
' dans un dossier synchronisé par OneDrive, ce dernier rend une adresse web, que
' ni Dir$ ni LoadPicture ne savent lire. L'image était alors introuvable sans
' que rien ne le dise.
'
'   renvoie : le chemin si le fichier existe, une chaîne vide sinon
'------------------------------------------------------------------------------
Public Function Interv_CheminImageTuile() As String
    Dim dossier As String, chemin As String

    dossier = DossierClasseur()
    If Len(dossier) = 0 Then Exit Function

    chemin = dossier & Application.PathSeparator & DOSSIER_IMAGES & _
             Application.PathSeparator & IMAGE_TUILE
    If FichierExiste(chemin) Then Interv_CheminImageTuile = chemin
End Function
