Attribute VB_Name = "modClients_Adresses"
Option Explicit
'==============================================================================
' modClients_Adresses
'------------------------------------------------------------------------------
' Alimentation et recherche des champs Adresse / NoPost / Ville / Cant à partir
' de l'onglet "Adresses" (tableau Tabl_Adresses).
'
' Quand un NPA n'existe pas dans Tabl_Adresses, la recherche se poursuit dans
' l'onglet "Liste_NPA_Suisse" (tableau Tabl_Villes_CH), qui contient la liste
' officielle des localités suisses.
'==============================================================================

Private Const COL_AD_RUE As String = "Nom_Rue_complet"
Private Const COL_AD_NPA As String = "No Postal"
Private Const COL_AD_VILLE As String = "Ville"
Private Const COL_AD_CANTON As String = "Canton"

Private Const COL_VI_LIEU As String = "Nom_du_Lieu"
Private Const COL_VI_NPA As String = "NPA"
Private Const COL_VI_CANTON As String = "Canton"

Private mAdr As Variant             ' cache du tableau Tabl_Adresses
Private mAdrRue As Long
Private mAdrNpa As Long
Private mAdrVille As Long
Private mAdrCanton As Long
Private mAdrCharge As Boolean

'==============================================================================
' Chargement du cache des adresses
'==============================================================================
'------------------------------------------------------------------------------
' Range le tableau Tabl_Adresses en mémoire et repère ses colonnes.
' Les 500 lignes sont lues d'un coup ; les recherches suivantes ne touchent plus
' la feuille, ce qui les rend assez rapides pour être lancées à chaque frappe.
'------------------------------------------------------------------------------
Private Sub ChargerAdresses()
    Dim lo As ListObject

    mAdrCharge = True
    mAdr = Empty
    Set lo = ObtenirTable(NOM_TABLE_ADRESSES)
    If lo Is Nothing Then Exit Sub
    If lo.ListRows.Count = 0 Then Exit Sub

    mAdrRue = IndexColonne(lo, COL_AD_RUE)
    If mAdrRue = 0 Then mAdrRue = IndexColonne(lo, "Nom_Rue")
    mAdrNpa = IndexColonne(lo, COL_AD_NPA)
    mAdrVille = IndexColonne(lo, COL_AD_VILLE)
    mAdrCanton = IndexColonne(lo, COL_AD_CANTON)

    mAdr = lo.DataBodyRange.Value
End Sub

'------------------------------------------------------------------------------
' Charge le cache des adresses au premier besoin.
'------------------------------------------------------------------------------
Private Sub AssurerAdresses()
    If Not mAdrCharge Then ChargerAdresses
End Sub

'------------------------------------------------------------------------------
' Force la relecture des adresses au prochain accès.
' Appelée à l'ouverture du formulaire : les adresses ajoutées dans la feuille
' depuis la dernière ouverture sont ainsi prises en compte.
'------------------------------------------------------------------------------
Public Sub Adresses_Recharger()
    mAdrCharge = False
End Sub

'==============================================================================
' Listes proposées dans les menus déroulants
'==============================================================================
'------------------------------------------------------------------------------
' Noms de rue proposés par le menu déroulant Adresse.
'   renvoie : les valeurs distinctes de Nom_Rue_complet, triées alphabétiquement
'------------------------------------------------------------------------------
Public Function Adresses_ListeRues() As Variant
    Dim d As Object, i As Long, s As String

    AssurerAdresses
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    If IsArray(mAdr) And mAdrRue > 0 Then
        For i = LBound(mAdr, 1) To UBound(mAdr, 1)
            s = Trim$(CStr(mAdr(i, mAdrRue) & ""))
            If Len(s) > 0 Then
                If Not d.Exists(s) Then d.Add s, 1
            End If
        Next i
    End If
    Adresses_ListeRues = TrierChaines(d.Keys)
End Function

'------------------------------------------------------------------------------
' NPA proposés par le menu déroulant NoPost.
'   renvoie : les NPA distincts de Tabl_Adresses, triés
'
' Volontairement limité aux adresses du classeur : charger les 5743 localités de
' Tabl_Villes_CH rendrait le menu inutilisable. Un NPA absent de cette liste
' reste saisissable à la main et sera cherché dans Tabl_Villes_CH.
'------------------------------------------------------------------------------
Public Function Adresses_ListeNpa() As Variant
    Dim d As Object, i As Long, s As String

    AssurerAdresses
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    If IsArray(mAdr) And mAdrNpa > 0 Then
        For i = LBound(mAdr, 1) To UBound(mAdr, 1)
            s = Trim$(CStr(mAdr(i, mAdrNpa) & ""))
            If Len(s) > 0 Then
                If Not d.Exists(s) Then d.Add s, 1
            End If
        Next i
    End If
    Adresses_ListeNpa = TrierChaines(d.Keys)
End Function

'------------------------------------------------------------------------------
' Textes standards proposés par le menu déroulant Texte de facture.
'   renvoie : le contenu de TblTxtStd, onglet Parametres, trié
' Renvoie une liste vide, sans erreur, si le tableau n'existe pas.
'------------------------------------------------------------------------------
Public Function Adresses_TextesFacture() As Variant
    Dim lo As ListObject, v As Variant, i As Long
    Dim d As Object, s As String

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    Set lo = ObtenirTable(NOM_TABLE_TEXTES)
    If Not lo Is Nothing Then
        If lo.ListRows.Count > 0 Then
            v = lo.DataBodyRange.Value
            If IsArray(v) Then
                For i = LBound(v, 1) To UBound(v, 1)
                    s = Trim$(CStr(v(i, 1) & ""))
                    If Len(s) > 0 Then
                        If Not d.Exists(s) Then d.Add s, 1
                    End If
                Next i
            End If
        End If
    End If
    Adresses_TextesFacture = TrierChaines(d.Keys)
End Function

'==============================================================================
' Recherche par rue
'==============================================================================
'------------------------------------------------------------------------------
' Cherche le NPA, la ville et le canton correspondant à un nom de rue.
'   rue       : nom de rue saisi ou choisi dans la liste
'   npaActuel : NPA déjà présent dans le formulaire, éventuellement vide
'   npa, ville, canton : renseignés en sortie
'   renvoie   : le nombre de localités trouvées pour cette rue
'
' Une même rue existe dans plusieurs communes. Si npaActuel correspond à l'une
' d'elles, c'est celle-là qui l'emporte ; sinon la première trouvée est proposée.
' Dans tous les cas l'utilisateur peut corriger à la main.
'------------------------------------------------------------------------------
Public Function Adresses_ChercherParRue(ByVal rue As String, ByVal npaActuel As String, _
                                        ByRef npa As String, ByRef ville As String, _
                                        ByRef canton As String) As Long
    Dim i As Long, n As Long, cible As String, courant As String
    Dim pNpa As String, pVille As String, pCanton As String

    AssurerAdresses
    npa = vbNullString: ville = vbNullString: canton = vbNullString
    cible = Normaliser(Trim$(rue))
    If Len(cible) = 0 Then Exit Function
    If Not IsArray(mAdr) Or mAdrRue = 0 Then Exit Function

    For i = LBound(mAdr, 1) To UBound(mAdr, 1)
        courant = Normaliser(Trim$(CStr(mAdr(i, mAdrRue) & "")))
        If courant = cible Then
            n = n + 1
            pNpa = ValeurAdr(i, mAdrNpa)
            pVille = ValeurAdr(i, mAdrVille)
            pCanton = ValeurAdr(i, mAdrCanton)
            If n = 1 Then
                npa = pNpa: ville = pVille: canton = pCanton
            End If
            If Len(npaActuel) > 0 Then
                If StrComp(pNpa, Trim$(npaActuel), vbTextCompare) = 0 Then
                    npa = pNpa: ville = pVille: canton = pCanton
                End If
            End If
        End If
    Next i
    Adresses_ChercherParRue = n
End Function

'==============================================================================
' Recherche par NPA
'==============================================================================
'------------------------------------------------------------------------------
' Cherche la ville et le canton d'un NPA.
'   renvoie : True si le NPA a été trouvé
'
' Tabl_Adresses d'abord, car ce sont les localités réellement utilisées par
' l'entreprise ; puis Tabl_Villes_CH, qui couvre toute la Suisse.
'------------------------------------------------------------------------------
Public Function Adresses_ChercherParNpa(ByVal npa As String, ByRef ville As String, _
                                        ByRef canton As String) As Boolean
    Dim i As Long, cible As String

    AssurerAdresses
    ville = vbNullString: canton = vbNullString
    cible = Trim$(npa)
    If Len(cible) = 0 Then Exit Function

    If IsArray(mAdr) And mAdrNpa > 0 Then
        For i = LBound(mAdr, 1) To UBound(mAdr, 1)
            If StrComp(ValeurAdr(i, mAdrNpa), cible, vbTextCompare) = 0 Then
                ville = ValeurAdr(i, mAdrVille)
                canton = ValeurAdr(i, mAdrCanton)
                Adresses_ChercherParNpa = True
                Exit Function
            End If
        Next i
    End If

    Adresses_ChercherParNpa = ChercherNpaSuisse(cible, ville, canton)
End Function

'------------------------------------------------------------------------------
' Repli sur la liste officielle des NPA suisses, onglet Liste_NPA_Suisse.
' Ce tableau n'est pas mis en cache : il est volumineux et n'est consulté que
' pour les NPA absents de Tabl_Adresses, ce qui reste rare.
'------------------------------------------------------------------------------
Private Function ChercherNpaSuisse(ByVal npa As String, ByRef ville As String, _
                                   ByRef canton As String) As Boolean
    Dim lo As ListObject, v As Variant, i As Long
    Dim icNpa As Long, icLieu As Long, icCant As Long

    Set lo = ObtenirTable(NOM_TABLE_VILLES)
    If lo Is Nothing Then Exit Function
    If lo.ListRows.Count = 0 Then Exit Function

    icNpa = IndexColonne(lo, COL_VI_NPA)
    icLieu = IndexColonne(lo, COL_VI_LIEU)
    icCant = IndexColonne(lo, COL_VI_CANTON)
    If icNpa = 0 Then Exit Function

    v = lo.DataBodyRange.Value
    If Not IsArray(v) Then Exit Function

    For i = LBound(v, 1) To UBound(v, 1)
        If StrComp(Trim$(CStr(v(i, icNpa) & "")), npa, vbTextCompare) = 0 Then
            If icLieu > 0 Then ville = Trim$(CStr(v(i, icLieu) & ""))
            If icCant > 0 Then canton = Trim$(CStr(v(i, icCant) & ""))
            ChercherNpaSuisse = True
            Exit Function
        End If
    Next i
End Function

'------------------------------------------------------------------------------
' Lit une cellule du cache des adresses, en texte, et sans erreur si la
' colonne demandée n'existe pas dans ce classeur.
'------------------------------------------------------------------------------
Private Function ValeurAdr(ByVal ligne As Long, ByVal colonne As Long) As String
    If colonne > 0 Then ValeurAdr = Trim$(CStr(mAdr(ligne, colonne) & ""))
End Function

'==============================================================================
' Tri alphabétique
'==============================================================================
'------------------------------------------------------------------------------
' Trie un tableau de chaînes par ordre alphabétique, accents ignorés.
'   renvoie : une copie triée ; le tableau d'origine n'est pas modifié
'------------------------------------------------------------------------------
Public Function TrierChaines(ByVal arr As Variant) As Variant
    Dim v As Variant
    If Not IsArray(arr) Then
        TrierChaines = arr
        Exit Function
    End If
    v = arr
    If UBound(v) > LBound(v) Then TriRapide v, LBound(v), UBound(v)
    TrierChaines = v
End Function

'------------------------------------------------------------------------------
' Tri rapide (quicksort) sur place, entre les bornes g et d.
' Préféré à un tri à bulles : les listes d'adresses dépassent 300 entrées.
'------------------------------------------------------------------------------
Private Sub TriRapide(ByRef v As Variant, ByVal g As Long, ByVal d As Long)
    Dim i As Long, j As Long, pivot As String, tmp As Variant
    i = g: j = d
    pivot = Normaliser(CStr(v((g + d) \ 2)))
    Do While i <= j
        Do While Normaliser(CStr(v(i))) < pivot
            i = i + 1
        Loop
        Do While Normaliser(CStr(v(j))) > pivot
            j = j - 1
        Loop
        If i <= j Then
            tmp = v(i): v(i) = v(j): v(j) = tmp
            i = i + 1: j = j - 1
        End If
    Loop
    If g < j Then TriRapide v, g, j
    If i < d Then TriRapide v, i, d
End Sub
