Attribute VB_Name = "modStat_Schema"
Option Explicit
'==============================================================================
' modStat_Schema
'------------------------------------------------------------------------------
' Ce que le formulaire des statistiques affiche : les colonnes de son tableau
' et les cinq tuiles qui le résument.
'
' Les noms de colonnes viennent de modInterv_Schema, qui décrit déjà TblInterv :
' une colonne renommée là-bas suit ici sans rien à toucher.
'==============================================================================

Public Const NOM_FORM_STAT As String = "UF_Statistiques"

' L'objectif de chiffre d'affaires de l'année, dans une cellule nommée du
' classeur. La barre de progression s'y rapporte ; si la cellule manque, la
' barre reste vide et le dit, plutôt que de diviser par zéro.
Public Const CEL_OBJECTIF As String = "Objectif_annuel_CA"

Public Const TOUS_LES_MOIS_S As String = "Tous les mois"

' Les cinq tuiles, dans l'ordre où elles se posent. Chaque entrée donne le
' libellé et la CLÉ que modStat_Formulaire sait calculer.
Public Const TU_CA As String = "CA"
Public Const TU_HEURES As String = "HEURES"
Public Const TU_NB As String = "NB"
Public Const TU_FACTURE As String = "FACTURE"
Public Const TU_A_FACTURER As String = "AFACTURER"

'==============================================================================
' LES TUILES
'------------------------------------------------------------------------------
' Toutes portent sur les lignes AFFICHÉES : elles suivent donc les filtres, et
' se recalculent à chaque fois que le tableau change.
'==============================================================================
Public Function STuiles() As Variant
    STuiles = Array(Array("Chiffre d'affaires", TU_CA), _
                    Array("Nbr d'heures", TU_HEURES), _
                    Array("Nb d'interventions", TU_NB), _
                    Array("Factur" & ChrW(233) & "es", TU_FACTURE), _
                    Array("A facturer", TU_A_FACTURER))
End Function

'==============================================================================
' LE TABLEAU
'==============================================================================
Public Function SColonnes() As Variant
    SColonnes = Array(IC_DATE, IC_CLIENT, IC_ENTREPRISE, IC_TITRE, IC_NOM, _
                      IC_PRENOM, IC_HEURES, IC_PERS, IC_TAUX, IC_CA, _
                      IC_TVA, IC_FORFAIT, IC_FACTURE)
End Function

Public Function SLibelles() As Variant
    SLibelles = Array("Date", "N" & ChrW(176) & " client", "Entreprise", "Titre", _
                      "Nom", "Pr" & ChrW(233) & "nom", "Heures", "Pers.", _
                      "Taux/Forf.", "CA", "TVA", "Forf.", "Fact.")
End Function

' 900 points, plus la barre de défilement : la carte en offre 926.
Public Function SLargeurs() As Variant
    SLargeurs = Array(64, 60, 160, 60, 120, 100, 52, 38, 62, 62, 34, 38, 50)
End Function

Public Function SAlignements() As Variant
    SAlignements = Array(MSF_TextAlignLeft, MSF_TextAlignLeft, MSF_TextAlignLeft, _
                         MSF_TextAlignLeft, MSF_TextAlignLeft, MSF_TextAlignLeft, _
                         MSF_TextAlignRight, MSF_TextAlignCenter, MSF_TextAlignRight, _
                         MSF_TextAlignRight, MSF_TextAlignCenter, MSF_TextAlignCenter, _
                         MSF_TextAlignLeft)
End Function

'------------------------------------------------------------------------------
' Les colonnes qui s'affichent comme une case à cocher.
'------------------------------------------------------------------------------
Public Function SEstCase(ByVal colonne As String) As Boolean
    SEstCase = (colonne = IC_TVA) Or (colonne = IC_FORFAIT)
End Function
