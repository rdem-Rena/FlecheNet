Attribute VB_Name = "modDatas_Verrou"
Option Explicit
'==============================================================================
' modDatas_Verrou - UN SEUL POSTE À LA FOIS DANS LES DONNÉES
'------------------------------------------------------------------------------
' POURQUOI NOTRE PROPRE VERROU. Excel sait déjà refuser la seconde ouverture
' d'un classeur — mais pas sur OneDrive. Là, il ouvre à chacun sa copie, laisse
' les deux postes écrire, puis fabrique un « fichier en conflit » que personne
' ne relit jamais. Le verrou d'Excel ne nous protège donc de rien ici, et celui
' de OneDrive encore moins : il faut le nôtre.
'
' COMMENT. Un fichier posé à côté du classeur de données :
'
'   FlecheNettoyageSA-2026.xlsx   ->   FlecheNettoyageSA-2026.verrou
'
' Sa seule présence interdit l'écriture aux autres postes. Il tient sur une
' ligne — qui, sur quel poste, depuis quand :
'
'   rdem|PC-BUREAU|01.10.2026 14:32:05|46296.6056
'
' La date est écrite DEUX FOIS : en clair pour qui ouvre le fichier dans le
' Bloc-notes, et en nombre pour le calcul. Le nombre seul fait foi, car lui
' seul se relit à l'identique d'un poste à l'autre : Str$ et Val ignorent les
' réglages régionaux, là où CDbl et CDate lisent la virgule d'un poste français
' et le point d'un poste anglais.
'
' LA PÉREMPTION. Excel qui se ferme mal, un poste qu'on éteint, et le verrou
' resterait là pour toujours. Passé QUATRE HEURES il est tenu pour abandonné et
' se reprend sans rien demander. Quatre heures suffisent parce que LE VERROU SE
' RENOUVELLE À CHAQUE ÉCRITURE : un poste qui travaille vraiment le rajeunit
' sans cesse, et seul un poste parti le laisse vieillir.
'
' LA RELECTURE. OneDrive ne synchronise pas à l'instant : deux postes peuvent
' écrire le fichier presque en même temps et le croire chacun à eux. On relit
' donc ce qu'on vient d'écrire, après un instant. Si la signature n'est plus la
' nôtre, c'est que l'autre est passé après — et c'est lui qui a la main.
'
' CE QUE CE VERROU NE FAIT PAS. Il ne résiste pas à un poste qui écrit sans
' passer par l'application, ni à OneDrive hors ligne — un poste déconnecté ne
' voit pas le verrou des autres, et les autres ne voient pas le sien. C'était
' le besoin : empêcher deux personnes de se marcher dessus, pas se défendre.
'==============================================================================

'--- Le fichier ---------------------------------------------------------------
Public Const EXT_VERROU As String = ".verrou"

' Le séparateur des quatre champs. Aucun nom d'utilisateur ni de poste ne peut
' en contenir.
Private Const SEP As String = "|"

' Ce que rend Lire quand le fichier existe mais refuse de s'ouvrir. Un seul
' caractère : aucune marque véritable ne lui ressemble.
Private Const MARQUE_ILLISIBLE As String = "?"

'--- Les deux durées ----------------------------------------------------------
' Au-delà, le verrou est tenu pour abandonné.
Private Const PEREMPTION_HEURES As Double = 4#

' Le temps laissé à OneDrive avant de relire ce qu'on vient d'écrire. Une
' seconde est le pas d'Application.Wait, et c'est déjà beaucoup pour un disque
' local ; en dessous, on relirait son propre cache.
Private Const DELAI_RELECTURE As Long = 1

'--- Ce qu'on a écrit ---------------------------------------------------------
' LA MARQUE EXACTE, et non « est-ce mon nom ». Deux sessions d'Excel ouvertes
' par le même utilisateur sur le même poste porteraient le même nom et le même
' poste : seule la marque entière, horodatage compris, les distingue.
Private mMarque As String

'==============================================================================
' PRENDRE LE VERROU
'------------------------------------------------------------------------------
'   chemin  : le classeur de données, pas le fichier verrou
'   msg     : ce qu'il faut dire à l'utilisateur en cas de refus
'   renvoie : True si l'écriture est permise
'==============================================================================
Public Function Verrou_Prendre(ByVal chemin As String, ByRef msg As String) As Boolean
    Verrou_Prendre = Saisir(chemin, msg, True)
End Function

'==============================================================================
' LE CONFIRMER AVANT D'ÉCRIRE
'------------------------------------------------------------------------------
' Appelée par Datas_PeutEcrire, donc avant chaque ajout, chaque modification et
' chaque suppression. Elle rajeunit le verrou — c'est ce qui rend la péremption
' de quatre heures sans danger — et détecte le cas inverse : une application
' laissée ouverte toute la nuit, dont le verrou a été repris le matin par
' quelqu'un d'autre. Mieux vaut l'apprendre avant d'écrire qu'après.
'
' Sans question ni mot de passe : on ne va pas interroger l'utilisateur au
' milieu d'une saisie.
'==============================================================================
Public Function Verrou_Confirmer(ByVal chemin As String, ByRef msg As String) As Boolean
    Verrou_Confirmer = Saisir(chemin, msg, False)
End Function

'==============================================================================
' LE RENDRE
'------------------------------------------------------------------------------
' SEULEMENT S'IL EST ENCORE LE NÔTRE. Effacer le verrou d'un autre poste à la
' fermeture serait le pire des services : il se croirait seul et ne le serait
' plus.
'==============================================================================
Public Sub Verrou_Rendre(ByVal chemin As String)
    Dim v As String

    v = CheminVerrou(chemin)
    If Len(v) = 0 Then Exit Sub

    ' SEULEMENT S'IL EST ENCORE EXACTEMENT LE NÔTRE. Illisible, il ne s'efface
    ' pas non plus : on ne sait pas à qui on l'enlèverait.
    If StrComp(Lire(v), mMarque, vbBinaryCompare) = 0 Then Supprimer v

    mMarque = vbNullString
End Sub

'==============================================================================
' CE QU'IL EN EST, EN UNE LIGNE
'------------------------------------------------------------------------------
' Pour le diagnostic : dire qui tient les données répond à la seule question que
' l'utilisateur se pose quand on lui refuse l'écriture.
'==============================================================================
Public Function Verrou_Etat(ByVal chemin As String) As String
    Dim v As String, marque As String, age As Double

    v = CheminVerrou(chemin)
    If Len(v) = 0 Then
        Verrou_Etat = "(chemin inutilisable)"
        Exit Function
    End If

    marque = Lire(v)
    If Len(marque) = 0 Then
        Verrou_Etat = "libre"
        Exit Function
    End If

    If marque = MARQUE_ILLISIBLE Then
        Verrou_Etat = "présent, mais il refuse de s'ouvrir"
        Exit Function
    End If

    If StrComp(marque, mMarque, vbBinaryCompare) = 0 Then
        Verrou_Etat = "pris par ce poste depuis " & Morceau(marque, 2)
        Exit Function
    End If

    age = AgeHeures(marque)
    Verrou_Etat = "pris par " & Morceau(marque, 0) & " (poste " & _
                  Morceau(marque, 1) & ") depuis " & Morceau(marque, 2) & _
                  IIf(age > PEREMPTION_HEURES, "  -  périmé, reprenable", "")
End Function

'==============================================================================
' LE CHEMIN DU FICHIER VERROU
'------------------------------------------------------------------------------
' Le classeur sans son extension, plus la nôtre. Il est DANS LE DOSSIER PARTAGÉ,
' à côté des données : c'est le seul endroit que tous les postes voient.
'==============================================================================
Public Function CheminVerrou(ByVal chemin As String) As String
    Dim p As Long, c As String

    If Len(chemin) = 0 Then Exit Function
    c = chemin
    p = InStrRev(c, ".")
    If p > InStrRev(c, "\") Then c = Left$(c, p - 1)
    CheminVerrou = c & EXT_VERROU
End Function

'==============================================================================
' LE CORPS COMMUN
'------------------------------------------------------------------------------
'   interactif : True à l'ouverture — on peut poser une question et proposer de
'                prendre la main ; False avant une écriture, où l'on se contente
'                de constater
'==============================================================================
Private Function Saisir(ByVal chemin As String, ByRef msg As String, _
                        ByVal interactif As Boolean) As Boolean
    Dim v As String, marque As String, age As Double, apres As String
    Dim reprise As Boolean

    msg = vbNullString
    v = CheminVerrou(chemin)
    If Len(v) = 0 Then
        msg = "Chemin de données inutilisable : " & chemin
        Exit Function
    End If

    marque = Lire(v)

    ' REPRISE : la marque sur le disque n'est pas la nôtre. Le verrou est libre,
    ' périmé, ou tenu par quelqu'un d'autre — dans tous les cas il change de
    ' main, et c'est là, et seulement là, qu'il faut se méfier de OneDrive.
    reprise = (StrComp(marque, mMarque, vbBinaryCompare) <> 0) Or Len(mMarque) = 0

    If reprise And Len(marque) > 0 Then
        If marque = MARQUE_ILLISIBLE Then
            ' LE FICHIER EXISTE MAIS NE S'OUVRE PAS. C'est très exactement ce
            ' que donne un autre poste en train de l'écrire : on le respecte,
            ' au lieu de passer par-dessus. Si cela devait durer, la péremption
            ' de quatre heures y met fin toute seule.
            msg = "Le fichier verrou existe, mais refuse de s'ouvrir :" & vbCrLf & _
                  v & vbCrLf & vbCrLf & "Un autre poste est sans doute en train " & _
                  "de le prendre à l'instant même."
            If Not interactif Then Exit Function
            If Not PrendreLaMain(msg) Then Exit Function
        Else
            age = AgeHeures(marque)

            ' UN HORODATAGE ILLISIBLE VAUT PÉRIMÉ (AgeHeures rend -1). C'est ce
            ' que laisse une écriture interrompue, et le tenir pour valable
            ' bloquerait les données sans aucune issue.
            If age >= 0 And age <= PEREMPTION_HEURES Then
                msg = "Les données sont en cours d'utilisation par " & _
                      Morceau(marque, 0) & " (poste " & Morceau(marque, 1) & _
                      "), depuis " & Morceau(marque, 2) & "."
                If Not interactif Then Exit Function
                If Not PrendreLaMain(msg) Then Exit Function
            End If
        End If
    End If

    If Not Ecrire(v, NotreMarque()) Then
        msg = "Le fichier verrou n'a pas pu s'écrire :" & vbCrLf & v & vbCrLf & _
              vbCrLf & "Le dossier est-il accessible en écriture ?"
        mMarque = vbNullString
        Exit Function
    End If

    ' LE VERROU DÉJÀ NÔTRE NE SE RELIT PAS. Verrou_Confirmer passe ici à chaque
    ' enregistrement ; y attendre une seconde ferait traîner toute l'application
    ' pour une course qui n'a pas lieu — personne ne dispute un verrou qu'on
    ' tient déjà.
    If Not reprise Then
        Saisir = True
        Exit Function
    End If

    ' LA RELECTURE. Deux postes qui écrivent en même temps se croient chacun
    ' maîtres ; c'est OneDrive qui tranche, après coup et sans le dire. On lui
    ' laisse un instant, puis on regarde qui est écrit dans le fichier.
    Patienter DELAI_RELECTURE
    apres = Lire(v)
    If StrComp(apres, mMarque, vbBinaryCompare) <> 0 Then
        If apres = MARQUE_ILLISIBLE Then
            msg = "Le fichier verrou, qu'on vient d'écrire, ne se relit déjà " & _
                  "plus : un autre poste est en train de l'écrire à son tour."
        Else
            msg = "Les données viennent d'être prises par " & Morceau(apres, 0) & _
                  " (poste " & Morceau(apres, 1) & "), à l'instant même."
        End If
        mMarque = vbNullString
        Exit Function
    End If

    Saisir = True
End Function

'------------------------------------------------------------------------------
' La question du forçage, puis le mot de passe.
'
' DEUX BARRIÈRES PLUTÔT QU'UNE. La première explique ce qu'on risque — c'est
' elle qui compte, car l'utilisateur ne sait pas, lui, qu'un verrou peut être
' encore vivant. La seconde réserve le geste à qui connaît le mot de passe.
'------------------------------------------------------------------------------
Private Function PrendreLaMain(ByVal msg As String) As Boolean
    Dim mdp As String, saisi As String

    If MsgBox(msg & vbCrLf & vbCrLf & _
              "PRENDRE LA MAIN quand même ?" & vbCrLf & vbCrLf & _
              "À ne faire que si vous savez que ce poste a été éteint sans " & _
              "fermer l'application. S'il y travaille encore, l'un de vous " & _
              "deux perdra ce qu'il aura saisi — et ne l'apprendra pas.", _
              vbExclamation + vbYesNo + vbDefaultButton2, _
              "Données déjà prises") <> vbYes Then Exit Function

    mdp = MotDePasse()
    If Len(mdp) = 0 Then
        PrendreLaMain = True
        Exit Function
    End If

    saisi = InputBox("Mot de passe :", "Prendre la main sur les données")
    If Len(saisi) = 0 Then Exit Function                 ' annulé, ou vide
    If StrComp(saisi, mdp, vbBinaryCompare) <> 0 Then
        MsgBox "Mot de passe incorrect.", vbExclamation, _
               "Prendre la main sur les données"
        Exit Function
    End If

    PrendreLaMain = True
End Function

'==============================================================================
' LA MARQUE
'==============================================================================

'------------------------------------------------------------------------------
' Celle qu'on écrit, retenue au passage pour se reconnaître ensuite.
'------------------------------------------------------------------------------
Private Function NotreMarque() As String
    Dim maintenant As Double

    maintenant = CDbl(Now)
    mMarque = Environ$("USERNAME") & SEP & Environ$("COMPUTERNAME") & SEP & _
              Format$(Now, "dd.mm.yyyy hh:nn:ss") & SEP & Trim$(Str$(maintenant))
    NotreMarque = mMarque
End Function

'------------------------------------------------------------------------------
' Le champ de rang donné, 0 pour le premier. Une marque abîmée rend des champs
' vides plutôt qu'une erreur : on les affiche tels quels.
'------------------------------------------------------------------------------
Private Function Morceau(ByVal marque As String, ByVal rang As Long) As String
    Dim t As Variant

    If Len(marque) = 0 Then Exit Function
    t = Split(marque, SEP)
    If rang >= LBound(t) And rang <= UBound(t) Then Morceau = CStr(t(rang))
End Function

'------------------------------------------------------------------------------
' Depuis combien d'heures la marque a été posée, -1 si elle ne le dit pas.
'
' Val et Str$ et non CDbl et CStr : le nombre traverse ainsi des postes aux
' réglages régionaux différents. Un âge NÉGATIF — l'horloge de l'autre poste
' avance sur la nôtre — compte comme tout frais : on ne reprend pas un verrou
' sur un désaccord de pendules.
'------------------------------------------------------------------------------
Private Function AgeHeures(ByVal marque As String) As Double
    Dim quand As Double

    AgeHeures = -1
    quand = Val(Morceau(marque, 3))
    If quand <= 0 Then Exit Function
    AgeHeures = (CDbl(Now) - quand) * 24#
End Function

'==============================================================================
' LE FICHIER
'------------------------------------------------------------------------------
' Open et Print plutôt que FileSystemObject : aucune référence à ajouter, et le
' fichier reste lisible dans le Bloc-notes.
'==============================================================================
Private Function Ecrire(ByVal v As String, ByVal contenu As String) As Boolean
    Dim f As Integer

    f = 0
    On Error GoTo Erreur
    f = FreeFile
    Open v For Output As #f
    Print #f, contenu
    Close #f
    Ecrire = True
    Exit Function

Erreur:
    If f <> 0 Then
        On Error Resume Next
        Close #f
        On Error GoTo 0
    End If
End Function

Private Function Lire(ByVal v As String) As String
    Dim f As Integer, s As String

    If Not FichierExiste(v) Then Exit Function

    f = 0
    On Error GoTo Erreur
    f = FreeFile
    Open v For Input As #f
    If Not EOF(f) Then Line Input #f, s
    Close #f
    Lire = Trim$(s)
    Exit Function

Erreur:
    ' UN VERROU QU'ON NE PEUT PAS LIRE N'EST PAS UN VERROU LIBRE : c'est le plus
    ' souvent un autre poste qui l'écrit à cet instant même. On le dit, et
    ' Saisir le respecte — le tenir pour libre serait précisément écrire
    ' par-dessus l'autre.
    If f <> 0 Then
        On Error Resume Next
        Close #f
        On Error GoTo 0
    End If
    Lire = MARQUE_ILLISIBLE
End Function

Private Sub Supprimer(ByVal v As String)
    On Error Resume Next
    SetAttr v, vbNormal
    Kill v
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Attendre, sans API Windows. Application.Wait a la seconde pour pas ; c'est
' grossier, mais cela n'arrive qu'à l'ouverture.
'------------------------------------------------------------------------------
Private Sub Patienter(ByVal secondes As Long)
    On Error Resume Next
    Application.Wait Now + TimeSerial(0, 0, secondes)
    On Error GoTo 0
End Sub
