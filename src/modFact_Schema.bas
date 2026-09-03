Attribute VB_Name = "modFact_Schema"
Option Explicit
'==============================================================================
' modFact_Schema
'------------------------------------------------------------------------------
' Ce que le formulaire de facturation affiche : les colonnes de ses deux
' tableaux, leurs titres, leurs largeurs et leurs alignements.
'
' C'est la SOURCE UNIQUE. Les quatre tableaux de chaque grille se lisent
' position par position et doivent rester de même longueur ; le générateur, le
' code d'affichage et le simulateur les relisent tous ici.
'
' Les noms de colonnes eux-mêmes ne sont pas redéfinis : ils viennent de
' modInterv_Schema, qui décrit déjà TblInterv. Une colonne renommée là-bas
' suit donc ici sans rien à toucher.
'==============================================================================

Public Const NOM_FORM_FACTURE As String = "UF_Facture"

' Colonne PROPRE au formulaire : elle n'existe pas dans TblInterv. L'utilisateur
' y coche les interventions à facturer, et c'est elle qui décide des lignes
' qu'Enregistrer met à jour.
Public Const FC_SELECT As String = "Select."

Public Const TOUS_LES_MOIS_F As String = "Tous les mois"

'------------------------------------------------------------------------------
' Ce qu'affiche une case à cocher de la grille : le « vu ».
'
' Une FONCTION et non une constante, parce qu'un Const ne peut porter qu'une
' valeur littérale, et que ce caractère doit s'écrire en ChrW : il n'existe pas
' en Windows-1252, et le laisser tel quel dans le source le ferait perdre à la
' conversion. U+2713 est présent dans Segoe UI, la police des deux grilles.
'------------------------------------------------------------------------------
Public Function FCoche() As String
    FCoche = ChrW(10003)
End Function

'==============================================================================
' TABLEAU DU HAUT — les clients dont des travaux restent à facturer
'------------------------------------------------------------------------------
' Une ligne par client, sans doublon : plusieurs interventions non facturées
' d'un même client n'en donnent qu'une.
'==============================================================================
Public Function FClientsColonnes() As Variant
    FClientsColonnes = Array(IC_CLIENT, IC_ENTREPRISE, IC_TITRE, IC_NOM, IC_PRENOM)
End Function

Public Function FClientsLibelles() As Variant
    FClientsLibelles = Array("N" & ChrW(176) & " client", "Entreprise", "Titre", _
                             "Nom", ChrW(80) & "r" & ChrW(233) & "nom")
End Function

Public Function FClientsLargeurs() As Variant
    FClientsLargeurs = Array(80, 320, 90, 210, 190)
End Function

Public Function FClientsAlignements() As Variant
    FClientsAlignements = Array(MSF_TextAlignLeft, MSF_TextAlignLeft, MSF_TextAlignLeft, _
                                MSF_TextAlignLeft, MSF_TextAlignLeft)
End Function

'==============================================================================
' TABLEAU DU BAS — les travaux du client choisi
'------------------------------------------------------------------------------
' CA est CALCULÉ, jamais lu : voir Fact_CADeLaLigne. Select. n'existe pas dans
' TblInterv et se coche à l'écran.
'==============================================================================
Public Function FTravauxColonnes() As Variant
    FTravauxColonnes = Array(IC_NO, IC_DATE, IC_HEURES, IC_PERS, IC_TAUX, _
                             IC_TVA, IC_FORFAIT, IC_CA, IC_TEXTE, IC_COMMENT, _
                             IC_FACTURE, FC_SELECT)
End Function

Public Function FTravauxLibelles() As Variant
    FTravauxLibelles = Array("N" & ChrW(176), "Date", "Heures", "Pers.", "Taux/Forf.", _
                             "TVA", "Forf.", "CA", "Texte de facture", "Commentaires", _
                             "Fact.", "Select.")
End Function

' 906 points au total, comme la grille des interventions : la carte fait
' I_CARTE_LARG moins deux, moins la barre de défilement.
Public Function FTravauxLargeurs() As Variant
    FTravauxLargeurs = Array(46, 64, 52, 38, 62, 34, 38, 62, 200, 200, 46, 46)
End Function

Public Function FTravauxAlignements() As Variant
    FTravauxAlignements = Array(MSF_TextAlignRight, MSF_TextAlignLeft, MSF_TextAlignRight, _
                                MSF_TextAlignCenter, MSF_TextAlignRight, MSF_TextAlignCenter, _
                                MSF_TextAlignCenter, MSF_TextAlignRight, MSF_TextAlignLeft, _
                                MSF_TextAlignLeft, MSF_TextAlignLeft, MSF_TextAlignCenter)
End Function

'------------------------------------------------------------------------------
' Les colonnes qui s'affichent comme une case à cocher.
'------------------------------------------------------------------------------
Public Function FEstCase(ByVal colonne As String) As Boolean
    FEstCase = (colonne = IC_TVA) Or (colonne = IC_FORFAIT) Or (colonne = FC_SELECT)
End Function

'------------------------------------------------------------------------------
' Position d'une colonne dans le tableau du bas, 1 = la première ; 0 si absente.
'------------------------------------------------------------------------------
Public Function FIndexTravaux(ByVal colonne As String) As Long
    Dim cols As Variant, i As Long

    cols = FTravauxColonnes()
    For i = LBound(cols) To UBound(cols)
        If StrComp(CStr(cols(i)), colonne, vbTextCompare) = 0 Then
            FIndexTravaux = i - LBound(cols) + 1
            Exit Function
        End If
    Next i
End Function
