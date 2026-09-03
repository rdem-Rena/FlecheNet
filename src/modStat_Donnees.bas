Attribute VB_Name = "modStat_Donnees"
Option Explicit
'==============================================================================
' modStat_Donnees
'------------------------------------------------------------------------------
' Ce que le formulaire des statistiques tire de TblInterv.
'
' Rien n'est relu du classeur ici : modInterv_Donnees a déjà chargé le tableau
' et sait le lire. Ce module ne fait que FILTRER et TOTALISER — et tout ce
' qu'il rend porte sur les mêmes lignes, celles que le tableau affiche. Les
' tuiles, le graphique et la barre d'objectif ne peuvent donc pas se contredire.
'
' Les critères voyagent dans un objet FiltreStat plutôt qu'en sept arguments :
' en ajouter un ne fait pas retoucher chaque appel.
'==============================================================================

' Une case à cocher à trois états. Null n'existant pas pour un Boolean, les
' trois valeurs sont explicites : le gris de la case veut dire « ne filtre pas ».
Public Const TRI_INDIFFERENT As Long = 0
Public Const TRI_VRAI As Long = 1
Public Const TRI_FAUX As Long = 2

Public Type FiltreStat
    Mois As Long              ' 1 à 12, ou 0 pour tous
    Entreprise As String
    Nom As String
    TVA As Long               ' TRI_*
    Forfait As Long           ' TRI_*
    Facture As Long           ' TRI_* : la colonne No_Facture est-elle remplie ?
    NoFacture As String
End Type

'==============================================================================
' Les lignes retenues par les filtres.
'------------------------------------------------------------------------------
'   renvoie : un tableau d'index de lignes de TblInterv, vide si rien ne passe
'==============================================================================
Public Function Stat_Lignes(ByRef f As FiltreStat) As Variant
    Dim res() As Long, n As Long, i As Long

    If Interv_NbLignes() = 0 Then
        Stat_Lignes = Array()
        Exit Function
    End If

    ReDim res(1 To Interv_NbLignes())
    For i = 1 To Interv_NbLignes()
        If Stat_LigneRetenue(i, f) Then
            n = n + 1
            res(n) = i
        End If
    Next i

    If n = 0 Then
        Stat_Lignes = Array()
    Else
        ReDim Preserve res(1 To n)
        Stat_Lignes = res
    End If
End Function

'------------------------------------------------------------------------------
' Une ligne passe-t-elle tous les filtres ?
'
' Les deux champs texte cherchent une SOUS-CHAÎNE, sans distinction de casse :
' taper « aeb » retrouve Aebi. Un champ vide ne filtre pas.
'------------------------------------------------------------------------------
Public Function Stat_LigneRetenue(ByVal ligne As Long, ByRef f As FiltreStat) As Boolean
    If f.Mois <> 0 Then
        If Fact_MoisDeLaLigne(ligne) <> f.Mois Then Exit Function
    End If
    If Not STexteContient(Interv_Valeur(ligne, IC_ENTREPRISE), f.Entreprise) Then Exit Function
    If Not STexteContient(Interv_Valeur(ligne, IC_NOM), f.Nom) Then Exit Function
    If Not SBooleenRetenu(Fact_EnBooleen(Interv_Valeur(ligne, IC_TVA)), f.TVA) Then Exit Function
    If Not SBooleenRetenu(Fact_EnBooleen(Interv_Valeur(ligne, IC_FORFAIT)), f.Forfait) Then Exit Function
    If Not SBooleenRetenu(Fact_EstFacturee(ligne), f.Facture) Then Exit Function
    If Not STexteContient(Interv_Valeur(ligne, IC_FACTURE), f.NoFacture) Then Exit Function

    Stat_LigneRetenue = True
End Function

'==============================================================================
' LES CINQ TOTAUX DES TUILES
'------------------------------------------------------------------------------
' Tous portent sur les lignes reçues, donc sur ce que le tableau montre.
'==============================================================================
Public Function Stat_Total(ByVal lignes As Variant, ByVal cle As String) As Double
    Dim i As Long, total As Double, facturee As Boolean

    If Not IsArray(lignes) Then Exit Function
    On Error GoTo Fin
    For i = LBound(lignes) To UBound(lignes)
        Select Case cle
            Case TU_CA
                total = total + Fact_CADeLaLigne(CLng(lignes(i)))
            Case TU_HEURES
                total = total + Fact_EnNombre(Interv_Valeur(CLng(lignes(i)), IC_HEURES))
            Case TU_NB
                total = total + 1
            Case TU_FACTURE, TU_A_FACTURER
                facturee = Fact_EstFacturee(CLng(lignes(i)))
                If facturee = (cle = TU_FACTURE) Then
                    total = total + Fact_CADeLaLigne(CLng(lignes(i)))
                End If
        End Select
    Next i
Fin:
    Stat_Total = total
End Function

'==============================================================================
' Le chiffre d'affaires mois par mois, pour le graphique.
'------------------------------------------------------------------------------
' Douze cases, de janvier à décembre, remplies à partir des lignes REÇUES : le
' graphique suit donc les filtres, contrairement à celui du formulaire des
' interventions, qui lit une feuille de totaux déjà faite.
'==============================================================================
Public Function Stat_CAParMois(ByVal lignes As Variant) As Variant
    Dim ca(1 To 12) As Double, i As Long, m As Long

    If IsArray(lignes) Then
        On Error Resume Next
        For i = LBound(lignes) To UBound(lignes)
            m = Fact_MoisDeLaLigne(CLng(lignes(i)))
            If m >= 1 And m <= 12 Then ca(m) = ca(m) + Fact_CADeLaLigne(CLng(lignes(i)))
        Next i
        On Error GoTo 0
    End If
    Stat_CAParMois = ca
End Function

'------------------------------------------------------------------------------
' L'objectif annuel, lu dans sa cellule nommée ; 0 si elle manque ou n'est pas
' un nombre — la barre de progression reste alors vide et le dit.
'------------------------------------------------------------------------------
Public Function Stat_ObjectifAnnuel() As Double
    Stat_ObjectifAnnuel = Fact_EnNombre(Interv_CelluleNommee(CEL_OBJECTIF))
End Function

'==============================================================================
' COMPARAISONS
'==============================================================================

' True si le motif est vide, ou s'il figure dans la valeur. Sans distinction de
' casse ni d'accents non : StrComp en mode texte suffit ici.
Private Function STexteContient(ByVal v As Variant, ByVal motif As String) As Boolean
    If Len(Trim$(motif)) = 0 Then
        STexteContient = True
    Else
        STexteContient = (InStr(1, EnTexte(v), Trim$(motif), vbTextCompare) > 0)
    End If
End Function

' True si l'état de la ligne satisfait une case à trois états.
Private Function SBooleenRetenu(ByVal etat As Boolean, ByVal critere As Long) As Boolean
    Select Case critere
        Case TRI_VRAI:  SBooleenRetenu = etat
        Case TRI_FAUX:  SBooleenRetenu = Not etat
        Case Else:      SBooleenRetenu = True
    End Select
End Function
