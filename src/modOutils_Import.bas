Attribute VB_Name = "modOutils_Import"
Option Explicit
'==============================================================================
' modOutils_Import
'------------------------------------------------------------------------------
' Importe d'un coup tous les modules .bas d'un dossier dans ce classeur.
'
'   >>> Procédure à lancer : ImporterModulesFlecheNet
'
' POURQUOI CE MODULE
' L'éditeur VBA ne propose que « Fichier > Importer un fichier… » (Ctrl+M), un
' fichier à la fois : les quatorze modules de src/ demandent quatorze
' allers-retours. Une commande « Importer plusieurs fichiers » n'existe pas en
' standard dans ce menu ; celles qu'on y voit parfois sont posées par un
' complément, et repartent avec lui — un plantage d'Excel suffit à le faire
' désactiver, et l'entrée de menu disparaît.
'
' Cette macro rend le classeur autonome : elle vit dans le projet, se sauvegarde
' avec lui, et ne demande qu'un seul Ctrl+M pour être réinstallée.
'
' PREREQUIS : Fichier > Options > Centre de gestion de la confidentialité >
'             Paramètres du Centre de gestion de la confidentialité >
'             Paramètres des macros > cocher
'             « Accès approuvé au modèle d'objet du projet VBA ».
'             (le même que pour la génération des formulaires)
'
' DEUX POINTS DE VIGILANCE, hérités du fonctionnement de VBA :
'
'   * un module déjà présent est rechargé SUR PLACE — son code est vidé, puis
'     réécrit depuis le fichier. VBA ne libère le nom d'un composant supprimé
'     qu'au retour à Excel : le supprimer puis le réimporter dans la même
'     exécution donnerait un modClients_Theme1 à côté de l'ancien ;
'   * ce module ne se recharge jamais lui-même : réécrire le code en cours
'     d'exécution réinitialise le projet et interromprait l'import.
'==============================================================================

Private Const TITRE As String = "Import des modules"
Private Const MOI_MEME As String = "modOutils_Import"
Private Const FD_DOSSIER As Long = 4        ' msoFileDialogFolderPicker
Private Const CT_MODULE As Long = 1         ' vbext_ct_StdModule

' Sort de ChargerModule
Private Const AJOUTE As Long = 0
Private Const RECHARGE As Long = 1
Private Const ECHEC As Long = 2

'==============================================================================
' POINTS D'ENTREE
'==============================================================================
'------------------------------------------------------------------------------
' Demande le dossier des modules, puis y importe tous les fichiers .bas.
' À lancer par Alt+F8, ou depuis la fenêtre Exécution de l'éditeur VBA.
'------------------------------------------------------------------------------
Public Sub ImporterModulesFlecheNet()
    Dim dossier As String

    dossier = ChoisirDossier()
    If Len(dossier) = 0 Then Exit Sub       ' Annuler

    ImporterModulesDepuis dossier
End Sub

'------------------------------------------------------------------------------
' Importe tous les .bas d'un dossier connu, sans boîte de dialogue.
'   dossier : chemin du dossier src, avec ou sans barre oblique inverse finale
'
' Utilisable telle quelle depuis la fenêtre Exécution (Ctrl+G) :
'   ImporterModulesDepuis "D:\FlecheNet\src"
'
' Les fichiers sont énumérés en entier AVANT le premier import : Dir garde un
' état interne, et le relancer entre deux écritures dans le projet VBA ferait
' perdre le fil de l'énumération.
'------------------------------------------------------------------------------
Public Sub ImporterModulesDepuis(ByVal dossier As String)
    Dim vbProj As Object
    Dim fichiers() As String, n As Long, i As Long
    Dim texte As String, nomModule As String, erreur As String
    Dim ajoutes As String, recharges As String, ignores As String, echecs As String
    Dim nbAjoutes As Long, nbRecharges As Long, nbEchecs As Long

    ' --- accès au projet VBA --------------------------------------------------
    ' Sans l'option de confiance, cette seule ligne déclenche l'erreur 1004.
    ' On la neutralise pour afficher un message compréhensible.
    On Error Resume Next
    Set vbProj = ThisWorkbook.VBProject
    On Error GoTo 0
    If vbProj Is Nothing Then
        MsgBox "Excel refuse l'accès au projet VBA." & vbCrLf & vbCrLf & _
               "Activez d'abord l'option :" & vbCrLf & _
               "Fichier > Options > Centre de gestion de la confidentialité >" & vbCrLf & _
               "Paramètres du Centre de gestion de la confidentialité >" & vbCrLf & _
               "Paramètres des macros > " & Chr$(34) & "Accès approuvé au modèle d'objet du " & _
               "projet VBA" & Chr$(34) & vbCrLf & vbCrLf & _
               "Fermez puis rouvrez le classeur, et relancez cette procédure.", _
               vbCritical, TITRE
        Exit Sub
    End If

    ' --- le dossier ----------------------------------------------------------
    dossier = NormaliserDossier(dossier)
    If Len(dossier) = 0 Then
        MsgBox "Ce dossier est introuvable.", vbExclamation, TITRE
        Exit Sub
    End If

    n = ListerBas(dossier, fichiers)
    If n = 0 Then
        MsgBox "Aucun fichier .bas dans :" & vbCrLf & vbCrLf & dossier, vbExclamation, TITRE
        Exit Sub
    End If

    ' --- import --------------------------------------------------------------
    For i = 1 To n
        texte = LireFichier(dossier & fichiers(i))
        nomModule = NomDuModule(texte, fichiers(i))

        If StrComp(nomModule, MOI_MEME, vbTextCompare) = 0 Then
            ignores = ignores & vbCrLf & "   " & fichiers(i) & " (module en cours d'exécution)"
        Else
            Select Case ChargerModule(vbProj, dossier & fichiers(i), nomModule, texte, erreur)
            Case AJOUTE
                nbAjoutes = nbAjoutes + 1
                ajoutes = ajoutes & vbCrLf & "   " & nomModule
            Case RECHARGE
                nbRecharges = nbRecharges + 1
                recharges = recharges & vbCrLf & "   " & nomModule
            Case Else
                nbEchecs = nbEchecs + 1
                echecs = echecs & vbCrLf & "   " & nomModule & " : " & erreur
            End Select
        End If
    Next i

    MsgBox Rapport(dossier, nbAjoutes, ajoutes, nbRecharges, recharges, _
                   ignores, nbEchecs, echecs), _
           IIf(nbEchecs > 0, vbExclamation, vbInformation), TITRE
End Sub

'==============================================================================
' IMPORT D'UN MODULE
'==============================================================================
'------------------------------------------------------------------------------
' Met un fichier .bas dans le projet.
'   chemin    : le fichier
'   nomModule : le nom que portera le composant
'   texte     : le contenu du fichier, déjà lu
'   motif     : renseigné en sortie si l'opération échoue
'   renvoie   : AJOUTE, RECHARGE ou ECHEC
'
' Absent du projet     -> Import, qui crée le composant et le nomme d'après son
'                         attribut VB_Name ;
' déjà présent         -> le code est remplacé sur place (voir l'en-tête du
'                         module : la suppression ne libère pas le nom tout de
'                         suite).
'------------------------------------------------------------------------------
Private Function ChargerModule(vbProj As Object, ByVal chemin As String, _
                               ByVal nomModule As String, ByVal texte As String, _
                               ByRef motif As String) As Long
    Dim vbComp As Object

    ' « motif » plutôt que « erreur » : dans une même procédure, une variable
    ' ne peut pas porter le nom de l'étiquette de branchement Erreur.
    motif = ""

    On Error Resume Next
    Set vbComp = vbProj.VBComponents(nomModule)
    On Error GoTo 0

    On Error GoTo Erreur

    If vbComp Is Nothing Then
        vbProj.VBComponents.Import chemin
        ChargerModule = AJOUTE
        Exit Function
    End If

    ' Un composant de ce nom existe, mais ce n'est pas un module standard :
    ' feuille, classeur, UserForm… On n'y touche pas.
    If vbComp.Type <> CT_MODULE Then
        motif = "un composant de ce nom existe déjà et n'est pas un module standard"
        ChargerModule = ECHEC
        Exit Function
    End If

    With vbComp.CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        .AddFromString CodeSeul(texte)
    End With
    ChargerModule = RECHARGE
    Exit Function

Erreur:
    motif = Err.Number & " - " & Err.Description
    ChargerModule = ECHEC
End Function

'------------------------------------------------------------------------------
' Le contenu d'un .bas privé de son en-tête.
'   renvoie : le code seul, prêt pour AddFromString
'
' Les lignes « Attribute VB_Name = … » de tête décrivent le composant et non son
' code : laissées en place, elles apparaîtraient comme du texte dans le module.
' Import les consomme tout seul ; AddFromString, non.
'------------------------------------------------------------------------------
Private Function CodeSeul(ByVal texte As String) As String
    Dim p As Long

    p = 1
    Do While StrComp(Mid$(texte, p, 10), "Attribute ", vbTextCompare) = 0
        p = InStr(p, texte, vbCrLf)
        If p = 0 Then Exit Function          ' que des attributs : rien à charger
        p = p + 2
    Loop

    CodeSeul = Mid$(texte, p)
End Function

'------------------------------------------------------------------------------
' Le nom que le module portera dans le projet.
'   texte      : contenu du fichier
'   nomFichier : nom du fichier, avec son extension
'   renvoie    : l'attribut VB_Name s'il y en a un, sinon le nom du fichier
'                sans extension — ce que fait l'éditeur VBA lui-même.
'------------------------------------------------------------------------------
Private Function NomDuModule(ByVal texte As String, ByVal nomFichier As String) As String
    Dim p As Long, q As Long, r As Long

    p = InStr(1, texte, "Attribute VB_Name", vbTextCompare)
    If p > 0 Then
        q = InStr(p, texte, Chr$(34))
        If q > 0 Then
            r = InStr(q + 1, texte, Chr$(34))
            If r > q Then NomDuModule = Mid$(texte, q + 1, r - q - 1)
        End If
    End If

    If Len(NomDuModule) = 0 Then
        p = InStrRev(nomFichier, ".")
        If p > 1 Then NomDuModule = Left$(nomFichier, p - 1) Else NomDuModule = nomFichier
    End If
End Function

'==============================================================================
' FICHIERS ET DOSSIER
'==============================================================================
'------------------------------------------------------------------------------
' Les fichiers .bas d'un dossier.
'   fichiers : renseigné en sortie, indices 1 à n, noms sans le chemin
'   renvoie  : le nombre de fichiers trouvés
'------------------------------------------------------------------------------
Private Function ListerBas(ByVal dossier As String, ByRef fichiers() As String) As Long
    Dim nom As String, n As Long

    ReDim fichiers(1 To 64)

    nom = Dir$(dossier & "*.bas")
    Do While Len(nom) > 0
        n = n + 1
        If n > UBound(fichiers) Then ReDim Preserve fichiers(1 To n + 64)
        fichiers(n) = nom
        nom = Dir$
    Loop

    ListerBas = n
End Function

'------------------------------------------------------------------------------
' Contenu d'un fichier texte.
'
' Ouverture en mode texte : les octets Windows-1252 des .bas sont convertis avec
' la page de codes du poste, les accents des libellés restent lisibles. C'est le
' pendant de l'avertissement du dépôt : ne pas convertir les .bas en UTF-8.
'------------------------------------------------------------------------------
Private Function LireFichier(ByVal chemin As String) As String
    Dim h As Integer

    h = FreeFile
    Open chemin For Input As #h
    If LOF(h) > 0 Then LireFichier = Input$(LOF(h), #h)
    Close #h
End Function

'------------------------------------------------------------------------------
' Chemin de dossier terminé par une seule barre oblique inverse.
'   renvoie : le chemin normalisé, ou "" si le dossier n'existe pas
'------------------------------------------------------------------------------
Private Function NormaliserDossier(ByVal dossier As String) As String
    Dim att As Long

    dossier = Trim$(dossier)
    Do While Len(dossier) > 0 And Right$(dossier, 1) = "\"
        dossier = Left$(dossier, Len(dossier) - 1)
    Loop
    If Len(dossier) = 0 Then Exit Function

    ' GetAttr lève l'erreur 53 ou 76 sur un chemin absent : plus sûr que Dir,
    ' qui échoue aussi quand le lecteur n'existe pas.
    On Error Resume Next
    att = GetAttr(dossier)
    If Err.Number <> 0 Then att = 0
    Err.Clear
    On Error GoTo 0

    If (att And vbDirectory) = vbDirectory Then NormaliserDossier = dossier & "\"
End Function

'------------------------------------------------------------------------------
' Boîte de dialogue de choix du dossier.
'   renvoie : le dossier choisi, ou "" si l'utilisateur annule
'------------------------------------------------------------------------------
Private Function ChoisirDossier() As String
    Dim dlg As Object

    Set dlg = Application.FileDialog(FD_DOSSIER)
    dlg.Title = "Dossier des modules à importer (src)"
    If Len(ThisWorkbook.Path) > 0 Then dlg.InitialFileName = ThisWorkbook.Path & "\"

    If dlg.Show = -1 Then ChoisirDossier = dlg.SelectedItems(1)
End Function

'==============================================================================
' COMPTE RENDU
'==============================================================================
'------------------------------------------------------------------------------
' Le texte affiché à la fin de l'import.
'------------------------------------------------------------------------------
Private Function Rapport(ByVal dossier As String, _
                         ByVal nbAjoutes As Long, ByVal ajoutes As String, _
                         ByVal nbRecharges As Long, ByVal recharges As String, _
                         ByVal ignores As String, _
                         ByVal nbEchecs As Long, ByVal echecs As String) As String
    Dim msg As String

    msg = dossier & vbCrLf & String$(46, "-") & vbCrLf

    If nbAjoutes > 0 Then
        msg = msg & nbAjoutes & Pluriel(nbAjoutes, " module ajouté", " modules ajoutés") & _
              " :" & ajoutes & vbCrLf & vbCrLf
    End If

    If nbRecharges > 0 Then
        msg = msg & nbRecharges & Pluriel(nbRecharges, " module rechargé", " modules rechargés") & _
              " :" & recharges & vbCrLf & vbCrLf
    End If

    If Len(ignores) > 0 Then msg = msg & "Laissé de côté :" & ignores & vbCrLf & vbCrLf

    If nbEchecs > 0 Then
        msg = msg & nbEchecs & Pluriel(nbEchecs, " échec", " échecs") & " :" & echecs & vbCrLf & vbCrLf
    End If

    Rapport = msg & "Vérification : Débogage > Compiler VBAProject ne doit " & _
              "signaler aucune erreur."
End Function

'------------------------------------------------------------------------------
' Singulier ou pluriel selon le nombre.
'------------------------------------------------------------------------------
Private Function Pluriel(ByVal n As Long, ByVal un As String, ByVal plusieurs As String) As String
    If n > 1 Then Pluriel = plusieurs Else Pluriel = un
End Function
