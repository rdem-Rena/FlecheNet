Attribute VB_Name = "modChemins"
Option Explicit
'==============================================================================
' modChemins - OÙ EST VRAIMENT LE CLASSEUR
'------------------------------------------------------------------------------
' LE PROBLÈME. Dans un dossier synchronisé par OneDrive ou SharePoint,
' ThisWorkbook.Path ne rend pas un chemin mais une ADRESSE WEB :
'
'   https://d.docs.live.net/a1b2c3d4/Documents/Flèche/FlecheNettoyageSA2026.xlsm
'   https://contoso.sharepoint.com/sites/Equipe/Documents partagés/Flèche/...
'
' Dir$, Open et LoadPicture ne savent pas lire cela. L'image de fond des tuiles
' était donc introuvable, en silence — le code se contentait de ne rien
' afficher — et le fichier de diagnostic des polices ne s'écrivait plus.
'
' LA SOLUTION. Le dossier existe pourtant bien sur le disque : OneDrive l'y
' synchronise. On retrouve donc le chemin local en confrontant la fin de
' l'adresse aux racines de synchronisation connues :
'
'   1. les racines viennent des variables d'environnement OneDrive,
'      OneDriveCommercial et OneDriveConsumer, que le client OneDrive pose ;
'   2. un site d'équipe SharePoint ne se synchronise pas à la racine mais sous
'      un dossier au nom du site — « Contoso\Equipe - Documents » — d'où
'      l'exploration de deux niveaux de sous-dossiers ;
'   3. pour chaque racine on essaie les fins de l'adresse, LA PLUS LONGUE
'      D'ABORD : c'est la plus spécifique, et la seule qui ne risque pas de
'      tomber sur un dossier homonyme.
'
' Le résultat est mis en cache : la recherche n'a lieu qu'une fois par session.
'
' CE MODULE NE DÉPEND DE RIEN. Il n'utilise que des fonctions VBA d'origine, et
' peut donc s'importer avant tous les autres.
'==============================================================================

' Le chemin trouvé, et le fait qu'on l'ait déjà cherché — une recherche
' infructueuse ne doit pas être refaite à chaque tuile.
Private mDossier As String
Private mResolu As Boolean

' Combien de sous-dossiers on accepte de lire dans une racine. Une racine
' OneDrive normale en compte quelques dizaines ; la borne protège d'un dossier
' pathologique, pas d'un usage courant.
Private Const MAX_ENFANTS As Long = 400

'==============================================================================
' LE DOSSIER DU CLASSEUR, EN CHEMIN LOCAL
'------------------------------------------------------------------------------
'   renvoie : un chemin utilisable par Dir$, Open et LoadPicture ; une chaîne
'             vide si le classeur n'a jamais été enregistré, ou s'il est ouvert
'             depuis le web sans copie synchronisée sur ce poste
'==============================================================================
Public Function DossierClasseur() As String
    Dim brut As String

    If mResolu Then
        DossierClasseur = mDossier
        Exit Function
    End If
    mResolu = True

    brut = ThisWorkbook.Path
    If Len(brut) = 0 Then Exit Function      ' classeur jamais enregistré

    If EstAdresseWeb(brut) Then
        mDossier = DossierLocalDepuisUrl(brut)
    Else
        mDossier = SansBarreFinale(brut)
    End If

    DossierClasseur = mDossier
End Function

'------------------------------------------------------------------------------
' Force une nouvelle recherche au prochain appel. Utile après avoir déplacé le
' classeur sans fermer Excel.
'------------------------------------------------------------------------------
Public Sub OublierDossierClasseur()
    mDossier = vbNullString
    mResolu = False
End Sub

'==============================================================================
' EXISTENCE
'------------------------------------------------------------------------------
' GetAttr plutôt que Dir$ : Dir$ possède un état interne partagé, et l'appeler
' au milieu d'un parcours de dossier interromprait ce parcours.
'==============================================================================
Public Function DossierExiste(ByVal chemin As String) As Boolean
    Dim a As Long

    a = -1
    On Error Resume Next
    a = GetAttr(chemin)
    On Error GoTo 0
    DossierExiste = (a <> -1) And ((a And vbDirectory) = vbDirectory)
End Function

Public Function FichierExiste(ByVal chemin As String) As Boolean
    Dim a As Long

    a = -1
    On Error Resume Next
    a = GetAttr(chemin)
    On Error GoTo 0
    FichierExiste = (a <> -1) And ((a And vbDirectory) = 0)
End Function

'==============================================================================
' DIAGNOSTIC
'------------------------------------------------------------------------------
' À lancer si une image ne s'affiche pas : le message dit ce que le classeur
' croit être son dossier, ce qu'on en a tiré, et si l'image y est.
'==============================================================================
Public Sub DiagnostiquerChemins()
    Dim brut As String, local As String, img As String, msg As String
    Dim rac As Variant, i As Long, n As Long

    brut = ThisWorkbook.Path
    OublierDossierClasseur
    local = DossierClasseur()

    msg = "Ce que le classeur annonce :" & vbCrLf & _
          IIf(Len(brut) > 0, brut, "(classeur jamais enregistré)") & vbCrLf & vbCrLf
    msg = msg & "Nature : " & IIf(EstAdresseWeb(brut), _
                                  "adresse web (OneDrive ou SharePoint)", _
                                  "chemin local") & vbCrLf & vbCrLf
    msg = msg & "Dossier local retenu :" & vbCrLf & _
          IIf(Len(local) > 0, local, "(introuvable)") & vbCrLf & vbCrLf

    If Len(local) > 0 Then
        img = local & Application.PathSeparator & "Images" & _
              Application.PathSeparator & "CartePremium.jpg"
        msg = msg & "Image des tuiles : " & _
              IIf(FichierExiste(img), "trouvée.", "ABSENTE de " & local & "\Images.")
    Else
        rac = Split(RacinesCandidates(), vbTab)
        For i = LBound(rac) To UBound(rac)
            If Len(rac(i)) > 0 Then n = n + 1
        Next i
        msg = msg & n & " racine(s) de synchronisation examinée(s)." & vbCrLf & vbCrLf
        If n = 0 Then
            msg = msg & "Aucune n'a été trouvée : le client OneDrive ne semble " & _
                  "pas installé sur ce poste, ou le classeur est ouvert depuis " & _
                  "le navigateur."
        Else
            msg = msg & "Aucune ne contient la fin de cette adresse. Ouvrez le " & _
                  "classeur depuis l'Explorateur, dans le dossier synchronisé, " & _
                  "plutôt que depuis le site web."
        End If
    End If

    MsgBox msg, vbInformation, "Emplacement du classeur"
End Sub

'==============================================================================
' RÉSOLUTION D'UNE ADRESSE WEB
'==============================================================================

Private Function EstAdresseWeb(ByVal chemin As String) As Boolean
    EstAdresseWeb = (LCase$(Left$(chemin, 7)) = "http://") Or _
                    (LCase$(Left$(chemin, 8)) = "https://")
End Function

'------------------------------------------------------------------------------
' Le protocole et le serveur n'ont aucun équivalent local : on ne garde que le
' chemin, décodé et retourné en séparateurs Windows, puis on cherche laquelle
' de ses fins existe sous une racine de synchronisation.
'------------------------------------------------------------------------------
Private Function DossierLocalDepuisUrl(ByVal url As String) As String
    Dim chemin As String, hote As String, p As Long
    Dim seg As Variant, racines As Variant
    Dim i As Long, j As Long, cand As String, fin As String
    Dim meilleur As String, meilleurScore As Long, sc As Long

    chemin = url
    p = InStr(chemin, "://")
    If p > 0 Then chemin = Mid$(chemin, p + 3)
    p = InStr(chemin, "/")
    If p = 0 Then Exit Function
    hote = Left$(chemin, p - 1)           ' le serveur sert à départager
    chemin = Mid$(chemin, p + 1)

    chemin = DecoderUrl(chemin)
    chemin = Replace$(chemin, "/", "\")
    chemin = SansBarreFinale(chemin)
    If Len(chemin) = 0 Then Exit Function

    seg = Split(chemin, "\")
    racines = Split(RacinesCandidates(), vbTab)

    ' LA FIN LA PLUS LONGUE D'ABORD, et à longueur égale la racine qui
    ' RESSEMBLE LE PLUS à l'adresse : deux dossiers synchronisés peuvent porter
    ' le même nom de projet, et prendre le premier venu donnerait, sans rien
    ' dire, les fichiers de l'autre.
    For i = LBound(seg) To UBound(seg)
        fin = DepuisSegment(seg, i)
        meilleur = vbNullString
        meilleurScore = -1
        For j = LBound(racines) To UBound(racines)
            If Len(racines(j)) > 0 Then
                cand = racines(j) & "\" & fin
                If DossierExiste(cand) Then
                    sc = Ressemblance(cand, hote, seg)
                    If sc > meilleurScore Then
                        meilleurScore = sc
                        meilleur = cand
                    End If
                End If
            End If
        Next j
        If Len(meilleur) > 0 Then
            DossierLocalDepuisUrl = meilleur
            Exit Function
        End If
    Next i
End Function

'------------------------------------------------------------------------------
' Combien de morceaux de l'adresse se retrouvent dans un chemin candidat.
'
' « Flèche » seul ne départage pas deux dossiers synchronisés qui portent tous
' deux ce nom. Le serveur et les dossiers intermédiaires, eux, le font :
' « Equipe » se retrouve dans « Equipe - Documents », et « contoso » dans
' « OneDrive - Contoso ».
'
' Les morceaux de moins de trois lettres sont ignorés : ils se retrouveraient
' partout et ne diraient rien.
'------------------------------------------------------------------------------
Private Function Ressemblance(ByVal chemin As String, ByVal hote As String, _
                              ByRef seg As Variant) As Long
    Dim k As Long, n As Long, mot As String

    mot = hote
    k = InStr(mot, ".")
    If k > 0 Then mot = Left$(mot, k - 1)
    If Len(mot) > 2 Then
        If InStr(1, chemin, mot, vbTextCompare) > 0 Then n = n + 1
    End If

    For k = LBound(seg) To UBound(seg)
        mot = CStr(seg(k))
        If Len(mot) > 2 Then
            If InStr(1, chemin, mot, vbTextCompare) > 0 Then n = n + 1
        End If
    Next k

    Ressemblance = n
End Function

'------------------------------------------------------------------------------
' Les segments d'indice debut et suivants, recollés en chemin.
'------------------------------------------------------------------------------
Private Function DepuisSegment(ByRef seg As Variant, ByVal debut As Long) As String
    Dim k As Long, res As String

    For k = debut To UBound(seg)
        res = res & IIf(Len(res) > 0, "\", "") & seg(k)
    Next k
    DepuisSegment = res
End Function

'==============================================================================
' LES RACINES DE SYNCHRONISATION
'------------------------------------------------------------------------------
' Les variables d'environnement d'abord, puis deux niveaux de sous-dossiers :
' un site d'équipe SharePoint se synchronise sous un dossier au nom du site, et
' la fin de l'adresse ne colle alors qu'à ce sous-dossier.
'
' Renvoyées séparées par des tabulations : un chemin ne peut pas en contenir.
'==============================================================================
Private Function RacinesCandidates() As String
    Dim base As Variant, i As Long, niveau As Long
    Dim tous As String, courant As String, suivant As String, liste As Variant

    base = Array(Environ$("OneDrive"), Environ$("OneDriveCommercial"), _
                 Environ$("OneDriveConsumer"), _
                 Environ$("USERPROFILE") & "\OneDrive")
    For i = LBound(base) To UBound(base)
        AjouterRacine courant, CStr(base(i))
    Next i
    tous = courant

    For niveau = 1 To 2
        suivant = vbNullString
        liste = Split(courant, vbTab)
        For i = LBound(liste) To UBound(liste)
            If Len(liste(i)) > 0 Then AjouterEnfants suivant, CStr(liste(i))
        Next i
        If Len(suivant) = 0 Then Exit For
        tous = tous & suivant
        courant = suivant
    Next niveau

    RacinesCandidates = tous
End Function

Private Sub AjouterRacine(ByRef liste As String, ByVal chemin As String)
    Dim c As String

    c = SansBarreFinale(chemin)
    If Len(c) = 0 Then Exit Sub
    If InStr(1, vbTab & liste, vbTab & c & vbTab, vbTextCompare) > 0 Then Exit Sub
    If Not DossierExiste(c) Then Exit Sub
    liste = liste & c & vbTab
End Sub

Private Sub AjouterEnfants(ByRef liste As String, ByVal racine As String)
    Dim noms As Variant, i As Long

    noms = Split(NomsDuDossier(racine), vbTab)
    For i = LBound(noms) To UBound(noms)
        If Len(noms(i)) > 0 Then AjouterRacine liste, racine & "\" & noms(i)
    Next i
End Sub

'------------------------------------------------------------------------------
' Les noms contenus dans un dossier, fichiers compris : le tri se fait ensuite,
' par DossierExiste. On ne l'appelle pas ici, car Dir$ ne supporte pas qu'on
' l'interrompe.
'------------------------------------------------------------------------------
Private Function NomsDuDossier(ByVal racine As String) As String
    Dim n As String, res As String, garde As Long

    On Error Resume Next
    n = Dir$(racine & "\*", vbDirectory)
    Do While Len(n) > 0 And garde < MAX_ENFANTS
        If n <> "." And n <> ".." Then res = res & n & vbTab
        n = Dir$()
        garde = garde + 1
    Loop
    On Error GoTo 0

    NomsDuDossier = res
End Function

'==============================================================================
' DÉCODAGE D'UNE ADRESSE
'------------------------------------------------------------------------------
' Une adresse OneDrive échappe l'espace en %20, et les accents en UTF-8 : le
' « è » de Flèche s'y écrit %C3%A8. Sans décodage, aucun dossier français ne se
' retrouverait sur le disque.
'
' Les séquences UTF-8 de un à trois octets suffisent : elles couvrent tout ce
' qu'un nom de dossier peut contenir.
'==============================================================================
Private Function DecoderUrl(ByVal s As String) As String
    Dim i As Long, n As Long, res As String
    Dim b1 As Long, b2 As Long, b3 As Long

    i = 1
    n = Len(s)
    Do While i <= n
        b1 = OctetUrl(s, i)
        If b1 < 0 Then
            res = res & Mid$(s, i, 1)
            i = i + 1
        ElseIf b1 < &H80 Then
            res = res & ChrW(b1)
            i = i + 3
        ElseIf (b1 And &HE0) = &HC0 Then
            b2 = OctetUrl(s, i + 3)
            If b2 < 0 Then
                res = res & Mid$(s, i, 1)
                i = i + 1
            Else
                res = res & ChrW((b1 And &H1F) * &H40 + (b2 And &H3F))
                i = i + 6
            End If
        ElseIf (b1 And &HF0) = &HE0 Then
            b2 = OctetUrl(s, i + 3)
            b3 = OctetUrl(s, i + 6)
            If b2 < 0 Or b3 < 0 Then
                res = res & Mid$(s, i, 1)
                i = i + 1
            Else
                res = res & ChrW((b1 And &HF) * &H1000 + _
                                 (b2 And &H3F) * &H40 + (b3 And &H3F))
                i = i + 9
            End If
        Else
            res = res & Mid$(s, i, 1)
            i = i + 1
        End If
    Loop

    DecoderUrl = res
End Function

'------------------------------------------------------------------------------
' L'octet codé par %XX à la position i, ou -1 si ce n'en est pas un.
'------------------------------------------------------------------------------
Private Function OctetUrl(ByVal s As String, ByVal i As Long) As Long
    OctetUrl = -1
    If i + 2 > Len(s) Then Exit Function
    If Mid$(s, i, 1) <> "%" Then Exit Function
    OctetUrl = DeuxHexa(Mid$(s, i + 1, 2))
End Function

'------------------------------------------------------------------------------
' Deux chiffres hexadécimaux en nombre, ou -1 si ce n'en sont pas.
'------------------------------------------------------------------------------
Private Function DeuxHexa(ByVal deux As String) As Long
    Dim k As Long, c As String, v As Long, h As Long

    DeuxHexa = -1
    For k = 1 To 2
        c = UCase$(Mid$(deux, k, 1))
        If c >= "0" And c <= "9" Then
            v = Asc(c) - 48
        ElseIf c >= "A" And c <= "F" Then
            v = Asc(c) - 55
        Else
            Exit Function
        End If
        h = h * 16 + v
    Next k
    DeuxHexa = h
End Function

'------------------------------------------------------------------------------
' Un chemin sans sa barre oblique finale : « C:\Dossier\ » devient
' « C:\Dossier ». Recoller un nom demande sinon de savoir s'il y en a une.
'------------------------------------------------------------------------------
Private Function SansBarreFinale(ByVal chemin As String) As String
    Dim c As String

    c = chemin
    Do While Len(c) > 0
        If Right$(c, 1) <> "\" And Right$(c, 1) <> "/" Then Exit Do
        c = Left$(c, Len(c) - 1)
    Loop
    SansBarreFinale = c
End Function
