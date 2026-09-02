Attribute VB_Name = "modFact_Donnees"
Option Explicit
'==============================================================================
' modFact_Donnees
'------------------------------------------------------------------------------
' Ce que le formulaire de facturation lit et écrit dans TblInterv.
'
' Rien n'est relu du classeur ici : modInterv_Donnees a déjà chargé le tableau
' et sait le lire, ligne par ligne et colonne par colonne. Ce module ne fait
' que TRIER — quels clients restent à facturer, quels travaux appartiennent à
' l'un d'eux — et ÉCRIRE le numéro de facture par le même chemin que le
' formulaire des interventions.
'
' Les deux fonctions de sélection rendent un tableau d'INDEX de lignes, jamais
' des copies des données : une valeur ne se lit donc qu'au moment de
' l'afficher, et ne peut pas devenir périmée entre-temps.
'==============================================================================

'==============================================================================
' Les clients dont des travaux restent à facturer.
'------------------------------------------------------------------------------
' Une ligne est retenue si son No_Facture est vide. Deux interventions d'un
' même client ne donnent qu'une entrée : la clé du dédoublonnage est le
' n-uplet AFFICHÉ tout entier, et non le seul numéro de client — deux fiches
' au même numéro mais au nom différent restent donc distinctes, ce qui les
' rend visibles plutôt que de les cacher.
'
'   renvoie : un tableau d'index de lignes, vide si rien n'est à facturer
'==============================================================================
Public Function Fact_ClientsNonFactures() As Variant
    Dim cols As Variant, vus As Object, res() As Long, n As Long
    Dim i As Long, j As Long, cle As String

    cols = FClientsColonnes()
    Set vus = CreateObject("Scripting.Dictionary")
    vus.CompareMode = 1
    ReDim res(1 To WorksheetFunctionMax(Interv_NbLignes(), 1))

    For i = 1 To Interv_NbLignes()
        If Not Fact_EstFacturee(i) Then
            cle = vbNullString
            For j = LBound(cols) To UBound(cols)
                cle = cle & EnTexte(Interv_Valeur(i, CStr(cols(j)))) & vbTab
            Next j
            If Not vus.Exists(cle) Then
                vus.Add cle, i
                n = n + 1
                res(n) = i
            End If
        End If
    Next i

    If n = 0 Then
        Fact_ClientsNonFactures = Array()
    Else
        ReDim Preserve res(1 To n)
        Fact_ClientsNonFactures = res
    End If
End Function

'==============================================================================
' Les travaux d'un client.
'------------------------------------------------------------------------------
'   clientNo : le numéro de client, tel qu'il figure dans TblInterv
'   toutes   : True pour montrer aussi les interventions déjà facturées
'   mois     : 1 à 12, ou 0 pour ne pas filtrer sur le mois
'
'   renvoie : un tableau d'index de lignes, vide si le client n'a rien
'==============================================================================
Public Function Fact_TravauxDuClient(ByVal clientNo As String, ByVal toutes As Boolean, _
                                     ByVal mois As Long) As Variant
    Dim res() As Long, n As Long, i As Long

    If Len(clientNo) = 0 Then
        Fact_TravauxDuClient = Array()
        Exit Function
    End If

    ReDim res(1 To WorksheetFunctionMax(Interv_NbLignes(), 1))
    For i = 1 To Interv_NbLignes()
        If StrComp(EnTexte(Interv_Valeur(i, IC_CLIENT)), clientNo, vbTextCompare) = 0 Then
            If toutes Or Not Fact_EstFacturee(i) Then
                If mois = 0 Or Fact_MoisDeLaLigne(i) = mois Then
                    n = n + 1
                    res(n) = i
                End If
            End If
        End If
    Next i

    If n = 0 Then
        Fact_TravauxDuClient = Array()
    Else
        ReDim Preserve res(1 To n)
        Fact_TravauxDuClient = res
    End If
End Function

'------------------------------------------------------------------------------
' True si la ligne porte déjà un numéro de facture.
'
' La colonne est de type TEXTE dans le classeur : un numéro y est une chaîne,
' et « vide » veut dire chaîne vide une fois les espaces retirés.
'------------------------------------------------------------------------------
Public Function Fact_EstFacturee(ByVal ligne As Long) As Boolean
    Fact_EstFacturee = (Len(Trim$(EnTexte(Interv_Valeur(ligne, IC_FACTURE)))) > 0)
End Function

'------------------------------------------------------------------------------
' Mois d'une ligne, 1 à 12 ; 0 si la date est absente ou illisible.
'------------------------------------------------------------------------------
Public Function Fact_MoisDeLaLigne(ByVal ligne As Long) As Long
    Dim v As Variant

    v = Interv_Valeur(ligne, IC_DATE)
    On Error Resume Next
    If IsDate(v) Then
        Fact_MoisDeLaLigne = Month(CDate(v))
    ElseIf IsNumeric(v) Then
        Fact_MoisDeLaLigne = Month(CDate(CDbl(v)))
    End If
    On Error GoTo 0
End Function

'==============================================================================
' Le chiffre d'affaires d'une ligne, CALCULÉ et non lu.
'------------------------------------------------------------------------------
' Heures x Personnes x Taux, ou le Taux seul quand la case Forfait est cochée.
' C'est le calcul que le formulaire des interventions applique déjà, dans
' Interv_EstimerCA : les deux écrans doivent afficher le même montant pour la
' même intervention, ce qui interdit d'en écrire une seconde version ici.
'
' Nb_Hres est une FRACTION DE JOUR dans le classeur — la conversion en heures
' est faite par Interv_EstimerCA.
'==============================================================================
Public Function Fact_CADeLaLigne(ByVal ligne As Long) As Double
    Dim taux As Double, heures As Double, pers As Double, forfait As Boolean

    taux = Fact_EnNombre(Interv_Valeur(ligne, IC_TAUX))
    heures = Fact_EnNombre(Interv_Valeur(ligne, IC_HEURES))
    pers = Fact_EnNombre(Interv_Valeur(ligne, IC_PERS))
    If pers = 0 Then pers = 1
    forfait = Fact_EnBooleen(Interv_Valeur(ligne, IC_FORFAIT))

    Fact_CADeLaLigne = Interv_EstimerCA(taux, forfait, pers, heures)
End Function

'------------------------------------------------------------------------------
' Somme d'une colonne sur les lignes données. Le CA passe par Fact_CADeLaLigne,
' puisqu'il n'est pas stocké.
'------------------------------------------------------------------------------
Public Function Fact_Somme(ByVal lignes As Variant, ByVal colonne As String) As Double
    Dim i As Long, total As Double

    If Not IsArray(lignes) Then Exit Function
    On Error GoTo Fin
    For i = LBound(lignes) To UBound(lignes)
        If colonne = IC_CA Then
            total = total + Fact_CADeLaLigne(CLng(lignes(i)))
        Else
            total = total + Fact_EnNombre(Interv_Valeur(CLng(lignes(i)), colonne))
        End If
    Next i
Fin:
    Fact_Somme = total
End Function

'==============================================================================
' Attribue un numéro de facture aux lignes données.
'------------------------------------------------------------------------------
' L'écriture passe par IntervBD_Modifier, le même chemin que le formulaire des
' interventions : c'est lui qui sait retrouver la ligne dans le classeur et
' respecter le format texte de la colonne.
'
'   renvoie : le nombre de lignes réellement mises à jour
'==============================================================================
Public Function Fact_Enregistrer(ByVal numero As String, ByVal lignes As Variant) As Long
    Dim i As Long, n As Long, valeurs As Object, noInterv As String

    If Len(Trim$(numero)) = 0 Then Exit Function
    If Not IsArray(lignes) Then Exit Function

    On Error GoTo Fin
    For i = LBound(lignes) To UBound(lignes)
        noInterv = EnTexte(Interv_Valeur(CLng(lignes(i)), IC_NO))
        If Len(noInterv) > 0 Then
            Set valeurs = CreateObject("Scripting.Dictionary")
            valeurs.CompareMode = 1
            valeurs.Add IC_FACTURE, Trim$(numero)
            If IntervBD_Modifier(noInterv, valeurs) Then n = n + 1
        End If
    Next i
Fin:
    Fact_Enregistrer = n
End Function

'------------------------------------------------------------------------------
' Conversions tolérantes : une cellule vide, un texte ou une erreur ne doivent
' jamais interrompre un affichage.
'------------------------------------------------------------------------------
Public Function Fact_EnNombre(ByVal v As Variant) As Double
    On Error Resume Next
    If IsNumeric(v) Then Fact_EnNombre = CDbl(v)
    On Error GoTo 0
End Function

Public Function Fact_EnBooleen(ByVal v As Variant) As Boolean
    Dim s As String

    On Error Resume Next
    If VarType(v) = vbBoolean Then
        Fact_EnBooleen = CBool(v)
    Else
        s = UCase$(Trim$(EnTexte(v)))
        Fact_EnBooleen = (s = "VRAI" Or s = "TRUE" Or s = "1" Or s = "OUI" Or s = "X")
    End If
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' Le plus grand de deux entiers. Écrit ici plutôt qu'emprunté à Excel : un
' ReDim doit réussir même si la feuille de calcul est indisponible.
'------------------------------------------------------------------------------
Private Function WorksheetFunctionMax(ByVal a As Long, ByVal b As Long) As Long
    WorksheetFunctionMax = IIf(a > b, a, b)
End Function
