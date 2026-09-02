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
' Diagnostic : ce que les contrôles du formulaire portent VRAIMENT comme police.
'
' À lancer quand l'aspect d'un texte ne correspond pas à ce que le thème
' demande. La lecture se fait dans le CONCEPTEUR, donc sur les mêmes valeurs
' que la fenêtre des propriétés du VBE — pas sur une copie chargée en mémoire.
'
' Le détail complet part dans la fenêtre d'exécution (Ctrl+G dans le VBE), d'où
' il se copie ; la boîte de dialogue n'affiche que les contrôles de saisie,
' ceux qui posent problème le plus souvent.
'
' Rappel utile pour lire le résultat : à 96 ppp un point vaut 0,75 pixel. Une
' taille qui n'est pas un multiple de 0,75 — 9,5 par exemple, soit 12,67 px —
' fait arrondir Windows, et le texte se dessine alors plus épais que demandé.
' Un texte qui PARAÎT gras avec Gras=Faux vient presque toujours de là.
'------------------------------------------------------------------------------
Public Sub DiagnostiquerPolicesInterv()
    Dim vbProj As Object, vbComp As Object, dsg As Object, ctl As Object
    Dim ligne As String, saisie As String, n As Long, nBancal As Long

    On Error Resume Next
    Set vbProj = ThisWorkbook.VBProject
    On Error GoTo 0
    If vbProj Is Nothing Then
        MsgBox "Excel refuse l'accès au projet VBA.", vbExclamation, "Polices"
        Exit Sub
    End If

    On Error Resume Next
    Set vbComp = vbProj.VBComponents(NOM_FORM_INTERV)
    On Error GoTo 0
    If vbComp Is Nothing Then
        MsgBox "Le formulaire " & NOM_FORM_INTERV & " n'existe pas encore.", _
               vbExclamation, "Polices"
        Exit Sub
    End If

    Set dsg = vbComp.Designer
    Debug.Print String$(70, "-")
    Debug.Print "Polices de " & NOM_FORM_INTERV
    Debug.Print "formulaire lui-même : " & DecrirePolice(dsg)
    Debug.Print String$(70, "-")

    For Each ctl In dsg.Controls
        ligne = ctl.Name & String$(IIf(Len(ctl.Name) < 24, 24 - Len(ctl.Name), 1), " ") & _
                TypeName(ctl) & vbTab & DecrirePolice(ctl)
        Debug.Print ligne
        n = n + 1
        If EstBancale(ctl) Then nBancal = nBancal + 1
        If TypeName(ctl) = "TextBox" Or TypeName(ctl) = "ComboBox" Then
            If Len(saisie) < 700 Then saisie = saisie & ligne & vbCrLf
        End If
    Next ctl

    MsgBox "Formulaire : " & DecrirePolice(dsg) & vbCrLf & vbCrLf & _
           "Contrôles de saisie :" & vbCrLf & saisie & vbCrLf & _
           n & " contrôles lus, dont " & nBancal & " à la taille bancale " & _
           "(pas un multiple de 0,75 pt : Windows arrondit et épaissit)." & vbCrLf & vbCrLf & _
           "Le détail complet est dans la fenêtre d'exécution (Ctrl+G).", _
           vbInformation, "Polices — " & NOM_FORM_INTERV
End Sub

'------------------------------------------------------------------------------
' « Segoe UI 9,5 gras » — ou la raison pour laquelle on n'a pas pu lire.
'------------------------------------------------------------------------------
Private Function DecrirePolice(ctl As Object) As String
    Dim s As String
    On Error GoTo SansPolice
    s = ctl.Font.Name & " " & CStr(ctl.Font.Size) & _
        IIf(ctl.Font.Bold, " GRAS", " maigre")
    If EstBancale(ctl) Then s = s & "   [taille bancale : " & _
        Format$(ctl.Font.Size / 0.75, "0.00") & " px]"
    DecrirePolice = s
    Exit Function
SansPolice:
    DecrirePolice = "(pas de police)"
End Function

'------------------------------------------------------------------------------
' True si le corps ne tombe pas sur un nombre entier de pixels.
'------------------------------------------------------------------------------
Private Function EstBancale(ctl As Object) As Boolean
    Dim px As Double
    On Error GoTo Fin
    px = ctl.Font.Size / 0.75
    EstBancale = (Abs(px - Int(px + 0.5)) > 0.000001)
Fin:
End Function

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
