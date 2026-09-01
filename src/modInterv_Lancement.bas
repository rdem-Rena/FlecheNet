Attribute VB_Name = "modInterv_Lancement"
Option Explicit
'==============================================================================
' modInterv_Lancement
'------------------------------------------------------------------------------
' Procédures à rattacher à un bouton de feuille ou à lancer par Alt+F8.
'
'   1. GenererFormulaireInterventions  crée UF_Interventions et UF_Calendrier.
'      À lancer une fois, puis à chaque changement du schéma ou de la charte.
'   2. OuvrirGestionInterventions      affiche le formulaire.
'==============================================================================

'------------------------------------------------------------------------------
' Affiche le formulaire des interventions.
'
' UserForms.Add crée le formulaire par son nom, sous forme de texte : ce module
' compile donc avant même la première génération, et un message clair remplace
' l'erreur VBA si le formulaire n'existe pas encore.
'------------------------------------------------------------------------------
Public Sub OuvrirGestionInterventions()
    Dim f As Object

    If TableInterventions() Is Nothing Then
        MsgBox "Le tableau " & NOM_TABLE_INTERVENTIONS & " est introuvable dans ce classeur.", _
               vbCritical, "Interventions"
        Exit Sub
    End If

    On Error GoTo Erreur
    Set f = UserForms.Add(NOM_FORM_INTERV)
    f.Show
    Exit Sub

Erreur:
    If Err.Number = 424 Or Err.Number = 5 Then
        MsgBox "Le formulaire " & NOM_FORM_INTERV & " n'existe pas encore dans ce classeur." & _
               vbCrLf & vbCrLf & "Lancez d'abord GenererFormulaireInterventions " & _
               "(module modInterv_Generateur).", vbExclamation, "Interventions"
    Else
        MsgBox "Ouverture impossible :" & vbCrLf & vbCrLf & _
               Err.Number & " - " & Err.Description, vbCritical, "Interventions"
    End If
End Sub

'------------------------------------------------------------------------------
' Diagnostic : vérifie que le classeur contient tout ce dont le formulaire des
' interventions a besoin. À lancer en premier quand quelque chose ne se comporte
' pas comme prévu.
'------------------------------------------------------------------------------
Public Sub VerifierClasseurInterventions()
    Dim msg As String, lo As ListObject, ch() As ChampInterv, i As Long
    Dim manque As String, tuiles As Variant, v As Variant

    msg = "Vérification — formulaire des interventions" & vbCrLf & String$(48, "-") & vbCrLf

    Set lo = TableInterventions()
    If lo Is Nothing Then
        msg = msg & "[X] Tableau " & NOM_TABLE_INTERVENTIONS & " : INTROUVABLE" & vbCrLf
    Else
        msg = msg & "[OK] Tableau " & NOM_TABLE_INTERVENTIONS & " : feuille " & lo.Parent.Name & _
              ", " & lo.ListRows.Count & " interventions, " & lo.ListColumns.Count & _
              " colonnes" & vbCrLf
        ch = ObtenirChampsInterv()
        For i = LBound(ch) To UBound(ch)
            If IndexColonne(lo, ch(i).Colonne) = 0 Then
                manque = manque & IIf(Len(manque) > 0, ", ", "") & ch(i).Colonne
            End If
        Next i
        If Len(manque) > 0 Then
            msg = msg & "[X] Colonnes attendues et absentes : " & manque & vbCrLf
        Else
            msg = msg & "[OK] Les " & NB_CHAMPS_INTERV & " colonnes du schéma sont présentes" & vbCrLf
        End If
        ' Le formulaire calcule le CA et l'enregistre. Si la colonne porte
        ' encore la formule d'origine, c'est elle qui gagne et le montant écrit
        ' est ignoré : les deux régimes donnent le même résultat, mais mieux
        ' vaut savoir lequel s'applique.
        If Interv_EstCalculee(IC_CA) Then
            msg = msg & "[ ] Colonne " & IC_CA & " : porte une formule — c'est elle qui" & _
                  " calcule, le montant du formulaire n'est pas écrit" & vbCrLf
        Else
            msg = msg & "[OK] Colonne " & IC_CA & " : calculée par le formulaire et enregistrée" & vbCrLf
        End If
        If Interv_IndexColonne(IC_TAUX) = 0 Then
            msg = msg & "[X] Colonne " & IC_TAUX & " : absente — le chiffre d'affaires" & _
                  " restera vide" & vbCrLf
        End If
    End If

    msg = msg & LigneTableI(NOM_TABLE_CLIENTS, "saisie assistée entreprise et nom")
    msg = msg & LigneTableI(NOM_TABLE_TEXTES, "textes de facture")

    ' --- cellules nommées -----------------------------------------------------
    msg = msg & vbCrLf & LigneNom(CEL_TITRE) & LigneNom(CEL_ANNEE)
    tuiles = TuilesStatistiques()
    For i = LBound(tuiles) To UBound(tuiles)
        msg = msg & LigneNom(CStr(tuiles(i)(1)))
    Next i

    ' --- graphique et image ---------------------------------------------------
    ' Le graphique est dessiné par le formulaire : ce ne sont plus les formes de
    ' la feuille qui comptent, mais la table qui porte ses valeurs.
    msg = msg & LigneTableI(NOM_TABLE_GRAPH, "valeurs du graphique du chiffre d'affaires")

    If Len(Interv_CheminImageTuile()) > 0 Then
        msg = msg & "[OK] Image " & IMAGE_TUILE & " : trouvée" & vbCrLf
    Else
        msg = msg & "[ ] Image " & IMAGE_TUILE & " : absente du sous-dossier " & DOSSIER_IMAGES & _
              " — les tuiles s'afficheront en blanc uni" & vbCrLf
    End If

    msg = msg & vbCrLf & "Formulaires : " & EtatFormI(NOM_FORM_INTERV) & _
          ", " & EtatFormI(NOM_FORM_CALENDRIER)

    MsgBox msg, vbInformation, "Interventions"
End Sub

'------------------------------------------------------------------------------
' Une ligne du diagnostic pour un tableau annexe.
'   usage : ce que le formulaire perd si ce tableau manque
'------------------------------------------------------------------------------
Private Function LigneTableI(ByVal nomTable As String, ByVal usage As String) As String
    Dim lo As ListObject
    Set lo = ObtenirTable(nomTable)
    If lo Is Nothing Then
        LigneTableI = "[ ] Tableau " & nomTable & " : absent - " & usage & " indisponible" & vbCrLf
    Else
        LigneTableI = "[OK] Tableau " & nomTable & " : " & lo.ListRows.Count & " lignes - " & _
                      usage & vbCrLf
    End If
End Function

'------------------------------------------------------------------------------
' Une ligne du diagnostic pour une cellule nommée, avec sa valeur si elle
' existe.
'------------------------------------------------------------------------------
Private Function LigneNom(ByVal nom As String) As String
    Dim v As Variant
    v = Interv_CelluleNommee(nom)
    If IsEmpty(v) Then
        LigneNom = "[X] Cellule nommée " & nom & " : introuvable" & vbCrLf
    Else
        LigneNom = "[OK] Cellule nommée " & nom & " = " & EnTexte(v) & vbCrLf
    End If
End Function

'------------------------------------------------------------------------------
' État d'un formulaire dans le projet : présent, à générer, ou indéterminable
' quand l'accès au projet VBA n'est pas autorisé.
'------------------------------------------------------------------------------
Private Function EtatFormI(ByVal nom As String) As String
    Dim vbProj As Object, vbComp As Object

    On Error Resume Next
    Set vbProj = ThisWorkbook.VBProject
    On Error GoTo 0
    If vbProj Is Nothing Then
        EtatFormI = nom & " : état inconnu"
        Exit Function
    End If

    On Error Resume Next
    Set vbComp = vbProj.VBComponents(nom)
    On Error GoTo 0
    EtatFormI = nom & IIf(vbComp Is Nothing, " : à générer", " : présent")
End Function
