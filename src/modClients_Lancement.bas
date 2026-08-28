Attribute VB_Name = "modClients_Lancement"
Option Explicit
'==============================================================================
' modClients_Lancement
'------------------------------------------------------------------------------
' Procédures à rattacher aux boutons de la feuille ou à lancer depuis
' Développeur > Macros.
'
'   1. GenererFormulaireClients  (module modClients_Generateur) : créé le
'      UserForm. À lancer une seule fois, puis à chaque fois que le schéma ou
'      la charte graphique changent.
'   2. OuvrirGestionClients      : affiche le formulaire.
'==============================================================================

'------------------------------------------------------------------------------
' Affiche le formulaire de gestion des clients.
'------------------------------------------------------------------------------
Public Sub OuvrirGestionClients()
    Dim f As Object

    If TableClients() Is Nothing Then
        MsgBox "Le tableau " & NOM_TABLE_CLIENTS & " est introuvable dans ce classeur.", _
               vbCritical, "Gestion des clients"
        Exit Sub
    End If

    On Error GoTo Erreur
    Set f = UserForms.Add(NOM_FORMULAIRE)
    f.Show
    Exit Sub

Erreur:
    If Err.Number = 424 Or Err.Number = 5 Then
        MsgBox "Le formulaire " & NOM_FORMULAIRE & " n'existe pas encore dans ce classeur." & _
               vbCrLf & vbCrLf & "Lancez d'abord la procédure GenererFormulaireClients " & _
               "(module modClients_Generateur).", vbExclamation, "Gestion des clients"
    Else
        MsgBox "Ouverture impossible :" & vbCrLf & vbCrLf & _
               Err.Number & " - " & Err.Description, vbCritical, "Gestion des clients"
    End If
End Sub

'------------------------------------------------------------------------------
' Génère le formulaire puis rappelle comment l'ouvrir.
' (La génération et l'affichage ne peuvent pas avoir lieu dans la même
'  exécution : VBA doit d'abord recompiler le projet.)
'------------------------------------------------------------------------------
Public Sub InstallerGestionClients()
    GenererFormulaireClients
End Sub

'------------------------------------------------------------------------------
' Diagnostic : vérifie que le classeur contient tout ce dont le formulaire a
' besoin. Utile lorsque le formulaire ne se comporte pas comme prévu.
'------------------------------------------------------------------------------
Public Sub VerifierClasseur()
    Dim msg As String, lo As ListObject, ch() As ChampClient, i As Long
    Dim manque As String

    msg = "Vérification du classeur" & vbCrLf & String$(46, "-") & vbCrLf

    Set lo = TableClients()
    If lo Is Nothing Then
        msg = msg & "[X] Tableau " & NOM_TABLE_CLIENTS & " : INTROUVABLE" & vbCrLf
    Else
        msg = msg & "[OK] Tableau " & NOM_TABLE_CLIENTS & " : feuille " & lo.Parent.Name & _
              ", " & lo.ListRows.Count & " fiches, " & lo.ListColumns.Count & " colonnes" & vbCrLf

        ch = ObtenirChamps()
        For i = LBound(ch) To UBound(ch)
            If IndexColonne(lo, ch(i).Colonne) = 0 Then
                manque = manque & IIf(Len(manque) > 0, ", ", "") & ch(i).Colonne
            End If
        Next i
        If Len(manque) > 0 Then
            msg = msg & "[X] Colonnes attendues et absentes : " & manque & vbCrLf
        Else
            msg = msg & "[OK] Les " & NB_CHAMPS & " colonnes du schéma sont présentes" & vbCrLf
        End If
    End If

    msg = msg & LigneTable(NOM_TABLE_ADRESSES, "recherche d'adresses")
    msg = msg & LigneTable(NOM_TABLE_VILLES, "recherche de NPA (repli)")
    msg = msg & LigneTable(NOM_TABLE_TEXTES, "textes de facture")
    msg = msg & LigneTable(NOM_TABLE_INTERV, "contrôle avant suppression")

    msg = msg & vbCrLf & "Formulaire " & NOM_FORMULAIRE & " : " & EtatFormulaire()

    MsgBox msg, vbInformation, "Gestion des clients"
End Sub

Private Function LigneTable(ByVal nomTable As String, ByVal usage As String) As String
    Dim lo As ListObject
    Set lo = ObtenirTable(nomTable)
    If lo Is Nothing Then
        LigneTable = "[ ] Tableau " & nomTable & " : absent - " & usage & " indisponible" & vbCrLf
    Else
        LigneTable = "[OK] Tableau " & nomTable & " : " & lo.ListRows.Count & " lignes - " & _
                     usage & vbCrLf
    End If
End Function

Private Function EtatFormulaire() As String
    Dim vbProj As Object, vbComp As Object

    On Error Resume Next
    Set vbProj = ThisWorkbook.VBProject
    On Error GoTo 0
    If vbProj Is Nothing Then
        EtatFormulaire = "état inconnu (accès au modèle d'objet du projet VBA non autorisé)"
        Exit Function
    End If

    On Error Resume Next
    Set vbComp = vbProj.VBComponents(NOM_FORMULAIRE)
    On Error GoTo 0
    EtatFormulaire = IIf(vbComp Is Nothing, "à générer (GenererFormulaireClients)", "présent")
End Function
