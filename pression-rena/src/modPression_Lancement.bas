Attribute VB_Name = "modPression_Lancement"
Option Explicit
'==============================================================================
' modPression_Lancement
'------------------------------------------------------------------------------
' Procédures à rattacher à un bouton de la feuille Saisie, ou à lancer depuis
' Développeur > Macros.
'
'   1. GenererFormulairePression (modPression_Generateur) : crée le UserForm.
'      À lancer une fois, puis à chaque fois que la charte ou la géométrie
'      changent.
'   2. OuvrirSaisiePression : affiche le formulaire.
'   3. VerifierClasseur : dit si le classeur a tout ce qu'il faut.
'==============================================================================

'------------------------------------------------------------------------------
' Affiche le formulaire de saisie.
'
' UserForms.Add crée le formulaire par son NOM, sous forme de texte : ce module
' compile donc même avant la première génération, et un message clair remplace
' l'erreur VBA si le formulaire n'existe pas encore.
'------------------------------------------------------------------------------
Public Sub OuvrirSaisiePression()
    Dim f As Object

    If FeuilleDatas() Is Nothing Then
        MsgBox "La feuille " & FEUILLE_DATAS & " est introuvable dans ce classeur.", _
               vbCritical, "Pression Rena"
        Exit Sub
    End If

    On Error GoTo Erreur
    Set f = UserForms.Add(NOM_FORM_PRESSION)
    f.Show
    Exit Sub

Erreur:
    If Err.Number = 424 Or Err.Number = 5 Then
        MsgBox "Le formulaire " & NOM_FORM_PRESSION & " n'existe pas encore dans ce " & _
               "classeur." & vbCrLf & vbCrLf & "Lancez d'abord la procédure " & _
               "GenererFormulairePression (module modPression_Generateur).", _
               vbExclamation, "Pression Rena"
    Else
        MsgBox "Ouverture impossible :" & vbCrLf & vbCrLf & _
               Err.Number & " - " & Err.Description, vbCritical, "Pression Rena"
    End If
End Sub

'------------------------------------------------------------------------------
' Génère le formulaire. Alias de GenererFormulairePression, pour retrouver
' l'installation dans la liste des macros sans connaître le nom du module.
'
' La génération et l'affichage ne peuvent pas avoir lieu dans la même
' exécution : VBA doit d'abord recompiler le projet.
'------------------------------------------------------------------------------
Public Sub InstallerSaisiePression()
    GenererFormulairePression
End Sub

'------------------------------------------------------------------------------
' Diagnostic : vérifie que le classeur porte tout ce dont le formulaire a
' besoin, et que l'invariant tient. À lancer en premier quand quelque chose ne
' se comporte pas comme prévu.
'------------------------------------------------------------------------------
Public Sub VerifierClasseur()
    Dim msg As String, ws As Worksheet, nom As Variant, controle As Variant
    Dim lignes As Long, manque As String

    msg = "Vérification du classeur Pression Rena" & vbCrLf & String$(46, "-") & vbCrLf

    For Each nom In Array(FEUILLE_SAISIE, FEUILLE_DATAS, FEUILLE_CALCUL, FEUILLE_REGLAGES)
        Set ws = ObtenirFeuille(CStr(nom))
        If ws Is Nothing Then
            msg = msg & "[X] Feuille " & nom & " : INTROUVABLE" & vbCrLf
        Else
            msg = msg & "[OK] Feuille " & nom & vbCrLf
        End If
    Next nom

    Set ws = FeuilleDatas()
    If Not ws Is Nothing Then
        lignes = DerniereLigne() - LIGNE_DEB + 1
        msg = msg & vbCrLf & "Calendrier : " & Format$(PremierJour(), "dd.mm.yyyy") & _
              " " & ChrW(8211) & " " & Format$(DernierJour(), "dd.mm.yyyy") & _
              ", " & lignes & " lignes" & vbCrLf

        ' L'invariant : deux lignes par jour, exactement.
        If lignes Mod 2 <> 0 Then
            msg = msg & "[X] Le nombre de lignes est IMPAIR : une ligne a été insérée " & _
                  "ou supprimée." & vbCrLf
        ElseIf lignes / 2 <> DateDiff("d", PremierJour(), DernierJour()) + 1 Then
            msg = msg & "[X] " & lignes / 2 & " paires de lignes pour " & _
                  DateDiff("d", PremierJour(), DernierJour()) + 1 & " jours : le " & _
                  "calendrier a un trou." & vbCrLf
        Else
            msg = msg & "[OK] Deux lignes par jour, sans trou." & vbCrLf
        End If
    End If

    For Each nom In Array(CEL_CONTROLE, CEL_SEUIL_SYS, CEL_SEUIL_DIA, _
                          "Date_Deb", "Date_Fin", "Nb_Jours", "Systole", "Diastole")
        If IsEmpty(ValeurNommee(CStr(nom))) Then
            manque = manque & IIf(Len(manque) > 0, ", ", "") & nom
        End If
    Next nom
    If Len(manque) > 0 Then
        msg = msg & "[X] Plages nommées absentes : " & manque & vbCrLf
    Else
        msg = msg & "[OK] Les plages nommées des graphiques répondent." & vbCrLf
    End If

    controle = ValeurNommee(CEL_CONTROLE)
    If Not IsEmpty(controle) Then msg = msg & vbCrLf & CStr(controle) & vbCrLf

    msg = msg & vbCrLf & "Formulaire " & NOM_FORM_PRESSION & " : " & EtatFormulaire()

    MsgBox msg, vbInformation, "Pression Rena"
End Sub

'------------------------------------------------------------------------------
' État du formulaire dans le projet.
'   renvoie : présent, à générer, ou état inconnu quand l'accès au projet VBA
'             n'est pas autorisé — auquel cas l'absence n'est pas démontrable
'------------------------------------------------------------------------------
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
    Set vbComp = vbProj.VBComponents(NOM_FORM_PRESSION)
    On Error GoTo 0
    EtatFormulaire = IIf(vbComp Is Nothing, "à générer (GenererFormulairePression)", "présent")
End Function
